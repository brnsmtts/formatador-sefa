"""Servidor local para testar o site no computador: python servir_local.py (ou SERVIR_LOCAL.bat)."""
import http.server
import socketserver
import threading
import webbrowser
from pathlib import Path

PORTA = 8000


class Manipulador(http.server.SimpleHTTPRequestHandler):
    extensions_map = {**http.server.SimpleHTTPRequestHandler.extensions_map,
                      ".wasm": "application/wasm", ".js": "text/javascript", ".mjs": "text/javascript",
                      ".whl": "application/zip", ".tsv": "text/plain; charset=utf-8"}

    def __init__(self, *a, **k):
        super().__init__(*a, directory=str(Path(__file__).resolve().parent), **k)

    def log_message(self, *a):
        pass


socketserver.TCPServer.allow_reuse_address = True
with socketserver.ThreadingTCPServer(("127.0.0.1", PORTA), Manipulador) as servidor:
    endereco = f"http://localhost:{PORTA}"
    print(f"Formatador SEFA web em {endereco}  (feche esta janela para parar)")
    threading.Timer(1.0, lambda: webbrowser.open(endereco)).start()
    servidor.serve_forever()
