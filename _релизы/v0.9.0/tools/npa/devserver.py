# Локальный сервер для разработки: раздаёт папку проекта и принимает POST /save?name=<файл>
# (сохраняет тело запроса в tools/npa/out/<файл>). Только для извлечения данных из PDF в браузере.
import http.server, os, urllib.parse

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))  # корень проекта
OUT = os.path.join(HERE, "out")
os.makedirs(OUT, exist_ok=True)

class H(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=ROOT, **kw)

    def do_POST(self):
        u = urllib.parse.urlparse(self.path)
        if u.path != "/save":
            self.send_error(404); return
        name = os.path.basename(urllib.parse.parse_qs(u.query).get("name", ["out.json"])[0])
        body = self.rfile.read(int(self.headers.get("Content-Length", 0)))
        with open(os.path.join(OUT, name), "wb") as f:
            f.write(body)
        self.send_response(200); self.end_headers(); self.wfile.write(b"ok")

http.server.ThreadingHTTPServer(("127.0.0.1", 8766), H).serve_forever()
