import os, io, uuid, tempfile, shutil
from flask import Flask, render_template, request, jsonify, send_file, session
from contract_engine.generator import generate, get_docx

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "rm-psa-secret-2026")
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"]   = True
app.config["SESSION_COOKIE_HTTPONLY"] = True

MANAGER_PIN   = os.environ.get("MANAGER_PIN",   "RM2026")
MANAGER_EMAIL = os.environ.get("MANAGER_EMAIL", "wcanormg@gmail.com")

# Separate download token store: token -> job_id
# This lets downloads work without relying on sessions across workers
_dl_tokens: dict = {}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/ping")
def ping():
    return "ok"


@app.route("/verify-pin", methods=["POST"])
def verify_pin():
    data = request.get_json()
    if data.get("pin") == MANAGER_PIN:
        session["is_manager"] = True
        return jsonify({"ok": True})
    return jsonify({"ok": False}), 401


@app.route("/submit-rep", methods=["POST"])
def submit_rep():
    data = request.get_json()
    try:
        result = generate(
            data.get("fields", {}),
            data.get("services", []),
            send_to=MANAGER_EMAIL,
            rep_mode=True
        )
        return jsonify({"status": "ok", "email_sent": result["email_sent"],
                        "email_msg": result["email_msg"]})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/generate", methods=["POST"])
def generate_route():
    data = request.get_json()
    if not session.get("is_manager") and data.get("pin") != MANAGER_PIN:
        return jsonify({"status": "error", "message": "Unauthorized"}), 403

    fields   = data.get("fields", {})
    services = data.get("services", [])
    send_to  = fields.get("MANAGER_EMAIL", MANAGER_EMAIL)

    try:
        result    = generate(fields, services, send_to=send_to)
        job_id    = result["job_id"]
        candidate = fields.get("CANDIDATE_FULL_NAME", "Contract").replace(" ", "_")

        # Create a signed download token so download works across workers
        dl_token = uuid.uuid4().hex
        _dl_tokens[dl_token] = (job_id, candidate)
        # Keep last 50 tokens
        old = list(_dl_tokens.keys())[:-50]
        for k in old: del _dl_tokens[k]

        return jsonify({
            "status":     "ok",
            "docx_url":   f"/download/{dl_token}/{candidate}",
            "email_sent": result["email_sent"],
            "email_msg":  result["email_msg"],
            "candidate":  fields.get("CANDIDATE_FULL_NAME", "Contract"),
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/download/<token>/<candidate>")
def download(token, candidate):
    # Validate token
    entry = _dl_tokens.get(token)
    if not entry:
        return "Link expired — please generate the contract again", 404

    job_id, safe_name = entry
    docx_bytes, _ = get_docx(job_id)
    if not docx_bytes:
        return "File not found — please generate again", 404

    # Write to temp file and serve — more reliable than BytesIO on all platforms
    tmp = tempfile.NamedTemporaryFile(
        suffix=".docx", delete=False,
        dir=tempfile.gettempdir(), prefix="psa_dl_")
    try:
        tmp.write(docx_bytes)
        tmp.flush()
        tmp.close()
        return send_file(
            tmp.name,
            as_attachment=True,
            download_name=f"PSA_{safe_name}.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    finally:
        # Clean up after response
        try:
            os.unlink(tmp.name)
        except Exception:
            pass


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=False)
