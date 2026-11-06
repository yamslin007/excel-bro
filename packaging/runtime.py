from __future__ import annotations

import logging
import json
import os
import ssl
import sys
import threading
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from server.app.app_data import application_data_root


def _bundle_root() -> Path:
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))


class _QuietStaticHandler(SimpleHTTPRequestHandler):
    def log_message(self, format: str, *args: object) -> None:
        logging.info("web " + format, *args)


def _frontend_server(install_root: Path, folder: str, port: int, *, tls: bool) -> ThreadingHTTPServer:
    handler = partial(
        _QuietStaticHandler,
        directory=str(install_root / folder),
    )
    server = ThreadingHTTPServer(("127.0.0.1", port), handler)
    if tls:
        context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        context.load_cert_chain(install_root / "localhost.crt", install_root / "localhost.key")
        server.socket = context.wrap_socket(server.socket, server_side=True)
    return server


def _wps_port(bundle_root: Path) -> int:
    config = json.loads((bundle_root / "config" / "capabilities.json").read_text(encoding="utf-8"))
    port = config["devServers"]["wpsPortCandidates"][0]
    if type(port) is not int or not 1 <= port <= 65535 or port == 3000:
        raise ValueError("WPS 安装服务端口配置无效。")
    return port


def _installer_command(install_root: Path, bundle_root: Path) -> bool:
    if len(sys.argv) == 3 and sys.argv[1] in ("--register-wps", "--unregister-wps"):
        from server.app.wps_installation import update_registration
        target = update_registration(Path(sys.argv[2]), _wps_port(bundle_root), remove=sys.argv[1] == "--unregister-wps")
        print(f"WPS registration: {target}")
        return True
    if sys.argv[1:] == ["--self-test"]:
        import subprocess
        import claude_agent_sdk
        from server.app.main import app
        cli = Path(claude_agent_sdk.__file__).parent / "_bundled" / ("claude.exe" if sys.platform == "win32" else "claude")
        if not cli.is_file():
            raise RuntimeError("SDK 原生 CLI 未打入安装包。")
        for resource in ("frontend/index.html", "wps-frontend/index.html", "wps-frontend/ribbon.xml", "manifest.xml", "localhost.crt", "localhost.key"):
            if not (install_root / resource).is_file():
                raise RuntimeError(f"安装资源缺失：{resource}")
        result = subprocess.run([str(cli), "--version"], check=True, capture_output=True, text=True, timeout=60)
        print(json.dumps({"backend": app.title, "wpsPort": _wps_port(bundle_root), "sdkCli": result.stdout.strip(), "frontends": ["excel", "wps"]}))
        return True
    if len(sys.argv) > 1:
        raise ValueError("不支持的安装器命令。")
    return False


def main() -> None:
    install_root = Path(sys.executable).resolve().parent
    bundle_root = _bundle_root()
    if _installer_command(install_root, bundle_root):
        return
    data_root = application_data_root()
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
    # Bind before starting the backend: a port conflict must fail startup, not
    # leave a healthy API pointing at an unrelated frontend process.
    servers = [_frontend_server(install_root, "frontend", 3000, tls=True)]
    if (install_root / "wps-frontend" / "index.html").is_file():
        servers.append(_frontend_server(install_root, "wps-frontend", _wps_port(bundle_root), tls=False))
    for server in servers:
        threading.Thread(target=server.serve_forever, name=f"excel-bro-web-{server.server_port}", daemon=True).start()

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
