from __future__ import annotations

import logging
import os
import ssl
import sys
import threading
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


def _bundle_root() -> Path:
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))


class _QuietStaticHandler(SimpleHTTPRequestHandler):
    def log_message(self, format: str, *args: object) -> None:
        logging.info("web " + format, *args)


def _serve_frontend(install_root: Path) -> None:
    handler = partial(
        _QuietStaticHandler,
        directory=str(install_root / "frontend"),
    )
    server = ThreadingHTTPServer(("127.0.0.1", 3000), handler)
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.load_cert_chain(
        install_root / "localhost.crt",
        install_root / "localhost.key",
    )
    server.socket = context.wrap_socket(server.socket, server_side=True)
    server.serve_forever()


def _data_root() -> Path:
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "Excel Bro"
    return Path(os.environ["LOCALAPPDATA"]) / "Excel Bro"


def main() -> None:
    install_root = Path(sys.executable).resolve().parent
    bundle_root = _bundle_root()
    data_root = _data_root()
    data_root.mkdir(parents=True, exist_ok=True)
    os.environ["EXCEL_BRO_CONFIG_DIR"] = str(data_root)
    os.environ["EXCEL_BRO_CAPABILITIES_PATH"] = str(
        bundle_root / "config" / "capabilities.json"
    )
    os.chdir(data_root)
    logging.basicConfig(
        filename=data_root / "excel-bro.log",
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )
    web_thread = threading.Thread(
        target=_serve_frontend,
        args=(install_root,),
        name="excel-bro-web",
        daemon=True,
    )
    web_thread.start()

    import uvicorn
    from server.app.main import app

    uvicorn.run(
        app,
        host="127.0.0.1",
        port=8765,
        log_config=None,
        access_log=False,
    )


if __name__ == "__main__":
    main()
