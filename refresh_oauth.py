"""Flow OAuth 2.0 manual con client credentials propias.
Paso 1: imprimir URL.
Paso 2: user pega código.
Paso 3: canjear por refresh_token y guardarlo como ADC."""
import json, os, sys, urllib.parse, urllib.request

CRED = json.load(open(r"C:\Users\tadeu\client_secret_42996714401-f48rbk1j429k93nmeahl3bvuvmjj3o4d.apps.googleusercontent.com.json"))["installed"]
CID = CRED["client_id"]
CSEC = CRED["client_secret"]
REDIR = "urn:ietf:wg:oauth:2.0:oob"  # out-of-band → muestra el código en la página
SCOPES = " ".join([
    "https://www.googleapis.com/auth/script.projects",
    "https://www.googleapis.com/auth/script.deployments",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "openid",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/cloud-platform",
])

if len(sys.argv) < 2:
    url = "https://accounts.google.com/o/oauth2/auth?" + urllib.parse.urlencode({
        "response_type": "code",
        "client_id": CID,
        "redirect_uri": REDIR,
        "scope": SCOPES,
        "access_type": "offline",
        "prompt": "consent",
    })
    print("ABRI ESTA URL, AUTORIZA, Y VOLVE CON EL CODIGO:\n")
    print(url)
    print("\nDespues corre: python refresh_oauth.py <CODIGO>")
    sys.exit(0)

code = sys.argv[1]
data = urllib.parse.urlencode({
    "code": code,
    "client_id": CID,
    "client_secret": CSEC,
    "redirect_uri": REDIR,
    "grant_type": "authorization_code",
}).encode()
resp = urllib.request.urlopen("https://oauth2.googleapis.com/token", data=data).read()
token = json.loads(resp)
print("Token recibido:", {k: (v if k != "refresh_token" else "***" + v[-10:]) for k, v in token.items()})

# Guardar como ADC credentials
adc_path = os.path.expanduser("~/AppData/Roaming/gcloud/application_default_credentials.json")
os.makedirs(os.path.dirname(adc_path), exist_ok=True)
adc = {
    "client_id": CID,
    "client_secret": CSEC,
    "refresh_token": token["refresh_token"],
    "type": "authorized_user",
}
json.dump(adc, open(adc_path, "w"))
print(f"\nADC guardado en {adc_path}")
print("Listo. Ahora podes deployar.")
