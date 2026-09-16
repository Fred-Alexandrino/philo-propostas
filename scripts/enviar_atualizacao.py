import json
import sys
import urllib.request

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
URL = "https://script.google.com/macros/s/AKfycbwaqTeyXPt9XHQB3ZFIvc8sLensJxxOt7ahLPFaOjNgFOzdJ_TT6-6KvW9s-xGst6fqzg/exec"
SENHA = "TROQUE_ESTA_SENHA"

with open("Code.gs", "r", encoding="utf-8") as f:
    codigo = f.read()

payload = {
    "action": "atualizarCodigo",
    "novoCodigoGs": codigo,
    "senha": SENHA,
}
body = json.dumps(payload).encode("utf-8")

print(f"Tamanho do payload: {len(body)} bytes", file=sys.stderr)
print(f"action = {payload['action']!r}", file=sys.stderr)

req = urllib.request.Request(
    URL,
    data=body,
    headers={
        "Content-Type": "text/plain;charset=utf-8",
        "User-Agent": UA,
    },
    method="POST",
)

class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None

opener = urllib.request.build_opener(NoRedirect)

try:
    resp = opener.open(req, timeout=60)
    print("Sem redirecionamento, status:", resp.status, file=sys.stderr)
    print(resp.read().decode("utf-8", "replace"))
except urllib.error.HTTPError as e:
    if e.code in (301, 302, 303):
        location = e.headers.get("Location")
        print(f"Redirecionado para: {location}", file=sys.stderr)
        req2 = urllib.request.Request(location, headers={"User-Agent": UA}, method="GET")
        resp2 = urllib.request.urlopen(req2, timeout=60)
        print("Status final:", resp2.status, file=sys.stderr)
        print(resp2.read().decode("utf-8", "replace"))
    else:
        print("Erro HTTP inesperado:", e.code, file=sys.stderr)
        print(e.read().decode("utf-8", "replace"))
