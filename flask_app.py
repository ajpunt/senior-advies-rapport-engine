import base64, os, tempfile, traceback, secrets, time

from flask import Flask, request, jsonify, send_from_directory, abort

from rapport_generator import generate_rapport, TESTDATA

app = Flask(__name__)

GEHEIME_SLEUTEL = "senior-advies-2026"

# ---------------------------------------------------------------------------
# PDF-MAP VOOR PRINTENBIND
# ---------------------------------------------------------------------------
# PrintenBind haalt de PDF op via een publieke URL. Daarom bewaren we elke
# PDF kort op een vaste plek, onder een ONRAADBARE bestandsnaam (lange
# willekeurige reeks). De map ligt buiten /home/ajpunt zodat er nooit per
# ongeluk broncode of andere bestanden via deze weg bereikbaar zijn.

PDF_MAP = "/home/ajpunt/pdf_publiek"
os.makedirs(PDF_MAP, exist_ok=True)

# PDF's ouder dan dit aantal uur worden automatisch opgeruimd.
PDF_BEWAAR_UREN = 24

# Basisadres van de site (voor het opbouwen van de publieke pdf_url).
BASIS_URL = "https://ajpunt.pythonanywhere.com"

def ruimoude_pdfs_op():
    """Verwijder PDF's die ouder zijn dan PDF_BEWAAR_UREN. Faalt stil:
    opruimen mag het maken van een rapport nooit blokkeren."""
    try:
        grens = time.time() - (PDF_BEWAAR_UREN * 3600)
        for naam in os.listdir(PDF_MAP):
            pad = os.path.join(PDF_MAP, naam)
            if os.path.isfile(pad) and os.path.getmtime(pad) < grens:
                os.unlink(pad)
    except Exception:
        pass

@app.route("/", methods=["GET"])
def home():
    return "Senior Advies V2 doorgeefluik draait.", 200

@app.route("/rapport", methods=["POST"])
def rapport():
    try:
        data = request.get_json(force=True, silent=True) or {}

        if data.get("sleutel") != GEHEIME_SLEUTEL:
            return jsonify({"ok": False, "fout": "Ongeldige sleutel"}), 403

        data.pop("sleutel", None)

        for k, v in TESTDATA.items():
            if k not in data:
                data[k] = v

        # PDF aanmaken op een tijdelijke plek (zoals voorheen)
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            output_path = tmp.name

        if 'besparing_jaar1' not in data and 'voordeel_jaar1' in data:
            try:
                data['besparing_jaar1'] = '➬ ' + f"{int(float(data['voordeel_jaar1'])):,}".replace(',', '.')
            except:
                data['besparing_jaar1'] = str(data['voordeel_jaar1'])
        raw_voordeel = data.get("voordeel_jaar1")
        try:
            if isinstance(raw_voordeel, str):
                cleaned = raw_voordeel.strip()
                cleaned = cleaned.replace("€", "").replace(" ", "")
                if "," in cleaned and "." in cleaned:
                    cleaned = cleaned.replace(".", "").replace(",", ".")
                elif "," in cleaned:
                    cleaned = cleaned.replace(",", ".")
                voordeel = float(cleaned)
            else:
                voordeel = float(raw_voordeel or 0)
            data["rendabel"] = "JA" if voordeel >= 1000 else "NEE"
        except Exception:
            data["rendabel"] = "NEE"
        generate_rapport(data, output_path)

        # PDF inlezen voor de base64 (zoals voorheen, voor de e-mail)
        with open(output_path, "rb") as f:
            pdf_bytes = f.read()

        pdf_b64 = base64.b64encode(pdf_bytes).decode("ascii")

        # NIEUW: PDF ook bewaren onder een onraadbare naam voor PrintenBind
        pdf_url = ""
        try:
            ruimoude_pdfs_op()
            geheime_naam = secrets.token_urlsafe(24) + ".pdf"
            publiek_pad = os.path.join(PDF_MAP, geheime_naam)
            with open(publiek_pad, "wb") as f:
                f.write(pdf_bytes)
            pdf_url = BASIS_URL + "/pdf/" + geheime_naam
        except Exception:
            # Lukt het bewaren niet, dan blijft pdf_url leeg.
            # Code.gs slaat PrintenBind dan netjes over; de e-mail gaat door.
            pdf_url = ""

        # Tijdelijk bestand opruimen (zoals voorheen)
        os.unlink(output_path)

        return jsonify({"ok": True, "pdf_base64": pdf_b64, "pdf_url": pdf_url}), 200

    except Exception:
        return jsonify({"ok": False, "fout": traceback.format_exc()}), 500

@app.route("/pdf/<naam>", methods=["GET"])
def pdf_ophalen(naam):
    """Levert een opgeslagen PDF uit, maar alleen als de naam klopt met het
    veilige patroon (lange willekeurige reeks + .pdf). Zo kan niemand door de
    map bladeren of via slimme namen bij andere bestanden komen."""
    # alleen letters, cijfers, - en _ toegestaan, eindigend op .pdf
    if not naam.endswith(".pdf"):
        abort(404)
    kern = naam[:-4]
    if not kern or not all(c.isalnum() or c in "-_" for c in kern):
        abort(404)
    bestand = os.path.join(PDF_MAP, naam)
    if not os.path.isfile(bestand):
        abort(404)
    return send_from_directory(PDF_MAP, naam, mimetype="application/pdf")

@app.route("/pdf-opruimen", methods=["GET", "POST"])
def pdf_opruimen():
    """Handmatig of via een geplande taak alle oude PDF's opruimen.
    Beveiligd met dezelfde geheime sleutel."""
    sleutel = request.args.get("sleutel") or (request.get_json(force=True, silent=True) or {}).get("sleutel")
    if sleutel != GEHEIME_SLEUTEL:
        return jsonify({"ok": False, "fout": "Ongeldige sleutel"}), 403
    ruimoude_pdfs_op()
    return jsonify({"ok": True, "melding": "Oude PDF's opgeruimd"}), 200
