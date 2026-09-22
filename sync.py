# -*- coding: utf-8 -*-
"""Синхронизация с сервером: push локальных изменений, pull чужих.

- Офлайн-первый: записи всегда сохраняются локально (equipment.json).
- Каждая запись имеет updated_at и deleted (надгробие).
- При появлении сети: push -> pull. Транспорт HTTPS (сертификат сервера),
  токен в заголовке Authorization.
"""

import json
import os
import ssl
import time
import uuid
from datetime import datetime, timezone

from kivy.utils import platform

try:
    from urllib.request import urlopen, Request, build_opener, ProxyHandler, HTTPSHandler
    from urllib.error import URLError, HTTPError
except ImportError:
    pass

STATE_FILE = "sync_state.json"

# На Android urlopen зависает при автоопределении системного прокси.
# Отключаем прокси явно через свой opener.
_opener = build_opener(ProxyHandler({}), HTTPSHandler(context=_ssl_ctx))

# Android не имеет системных CA-корней для Python. Используем certifi.
try:
    import certifi
    _ssl_ctx = ssl.create_default_context(cafile=certifi.where())
except Exception:
    _ssl_ctx = ssl.create_default_context()


# Адрес сервера и токен задаются при сборке (см. SERVER_URL / API_TOKEN)
SERVER_URL = "https://rezzonvoice.ru/api"
API_TOKEN = "hBqxlkwcoWrA65RuUaHstETn7ipOZ2801YIQD3GgSbeNmXJy"
APP_VERSION = "1.0"


def state_path():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), STATE_FILE)


def load_state():
    try:
        with open(state_path(), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_state(state):
    try:
        with open(state_path(), "w", encoding="utf-8") as f:
            json.dump(state, f)
    except Exception:
        pass


def device_id():
    st = load_state()
    if "device_id" not in st:
        st["device_id"] = uuid.uuid4().hex[:16]
        save_state(st)
    return st["device_id"]


def now_iso():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


class SyncClient:
    def __init__(self, app, server_url=None, token=None):
        self.app = app
        self.server_url = (server_url or SERVER_URL).rstrip("/")
        self.token = token or API_TOKEN
        st = load_state()
        self.cursor = st.get("cursor", "")
        self.last_sync = st.get("last_sync", "")
        self.last_error = st.get("last_error", "")

    def configured(self):
        return ("CHANGE-ME" not in self.server_url
                and "CHANGE_ME" not in self.token
                and bool(self.server_url) and bool(self.token))

    # ---------- низкий уровень ----------
    def _raw_connect_test(self):
        """Диагностика: DNS + TCP + SSL по шагам."""
        import socket
        host = "rezzonvoice.ru"
        try:
            print("NET-TEST: getaddrinfo...", flush=True)
            infos = socket.getaddrinfo(host, 443, socket.AF_INET, socket.SOCK_STREAM)
            print("NET-TEST: dns OK %s" % str(infos[0][4]), flush=True)
            s = socket.create_connection((host, 443), timeout=10)
            print("NET-TEST: tcp OK", flush=True)
            import ssl as _ssl
            ctx = _ssl_ctx
            ss = ctx.wrap_socket(s, server_hostname=host)
            print("NET-TEST: ssl OK %s" % ss.version(), flush=True)
            ss.sendall(b"GET /api/health HTTP/1.1\r\nHost: rezzonvoice.ru\r\nConnection: close\r\n\r\n")
            data = ss.recv(500)
            print("NET-TEST: response %s" % data[:80], flush=True)
            ss.close()
            return True
        except Exception as e:
            print("NET-TEST FAIL: %r" % e, flush=True)
            return False

    def _post(self, path, payload):
        print("SYNC: _post %s" % path, flush=True)
        if not self._raw_connect_test():
            return None
        req = Request(
            self.server_url + path,
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Content-Type": "application/json",
                "Authorization": "Bearer " + self.token,
            },
            method="POST")
        print("SYNC: _post request created", flush=True)
        ctx = _ssl_ctx
        with _opener.open(req, timeout=20) as resp:
            return json.loads(resp.read().decode("utf-8"))

    def _get(self, path):
        print("SYNC: _get %s" % path, flush=True)
        req = Request(
            self.server_url + path,
            headers={"Authorization": "Bearer " + self.token})
        ctx = _ssl_ctx
        with _opener.open(req, timeout=20) as resp:
            return json.loads(resp.read().decode("utf-8"))

    # ---------- синхронизация ----------
    def sync_now(self):
        """Push локальных -> pull серверных. Возвращает (ok, message)."""
        print("SYNC: starting, configured=%s" % self.configured(), flush=True)
        if not self.configured():
            return False, "Сервер не настроен"
        try:
            # PUSH: все записи, изменённые после последнего push
            st = load_state()
            since_push = st.get("last_push", "")
            dirty = [r for r in self.app.records
                     if (r.get("updated_at") or "") > since_push]
            payload = {"device_id": device_id(),
                       "records": dirty if dirty else self.app.records}
            print("SYNC: push sent, dirty=%d, total=%d" % (len(dirty), len(self.app.records)), flush=True)
            res = self._post("/push", payload)
            print("SYNC: push result:", res, flush=True)
            st["last_push"] = now_iso()
            save_state(st)

            # PULL: применяем чужие изменения
            pulled = self._pull_all()
            print("SYNC: pulled", pulled, flush=True)
            self.last_sync = now_iso()
            self.last_error = ""
            st = load_state()
            st["last_sync"] = self.last_sync
            st["last_error"] = ""
            st["cursor"] = pulled
            save_state(st)
            return True, "push %d, pull %d" % (res.get("pushed", 0), pulled)
        except HTTPError as e:
            msg = "HTTP %s: %s" % (e.code, e.read()[:200] if hasattr(e, 'read') else '')
            print("SYNC ERROR:", msg, flush=True)
            self.last_error = msg
            return False, msg
        except (URLError, ssl.SSLError, OSError) as e:
            msg = "Нет соединения: %s" % getattr(e, "reason", e)
            self.last_error = msg
            return False, msg
        except Exception as e:
            self.last_error = str(e)
            return False, str(e)

    def _pull_all(self):
        """Скачиваем изменения с курсора, применяем, повторяем до конца."""
        total = 0
        cursor = self.cursor
        for _ in range(10):  # максимум 10 страниц
            res = self._get("/pull?since=" + cursor)
            recs = res.get("records", [])
            if not recs:
                break
            self._apply_remote(recs)
            total += len(recs)
            new_cursor = res.get("cursor", "")
            if not new_cursor or new_cursor == cursor:
                break
            cursor = new_cursor
        self.cursor = cursor
        return total

    def _apply_remote(self, recs):
        """Слияние: серверная версия побеждает, если она новее локальной."""
        local = {r["id"]: r for r in self.app.records}
        changed = False
        for remote in recs:
            rid = remote["id"]
            loc = local.get(rid)
            if loc is None:
                if remote.get("deleted"):
                    continue
                local[rid] = remote
                changed = True
            else:
                r_ts = remote.get("updated_at", "")
                l_ts = loc.get("updated_at", "")
                if r_ts > l_ts:
                    local[rid] = remote
                    changed = True
        if changed:
            self.app.records = list(local.values())
            self.app.save_and_refresh()


# ------------------------- обновления приложения -------------------------

def check_update():
    """Проверить версию на сервере. Возвращает (new_version, apk_url) или
    (None, None), если обновление не требуется / сервер недоступен."""
    if "CHANGE-ME" in SERVER_URL:
        return None, None
    try:
        req = Request(SERVER_URL + "/version",
                      headers={"Authorization": "Bearer " + API_TOKEN})
        ctx = _ssl_ctx
        with _opener.open(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        ver = str(data.get("version", "")).strip()
        apk = str(data.get("apk_url", "")).strip()
        if ver and ver != APP_VERSION and apk:
            return ver, apk
    except Exception:
        pass
    return None, None


def download_update(apk_url, dest_path, progress_cb=None):
    """Скачать APK с сервера. progress_cb(loaded_bytes, total_bytes)."""
    req = Request(apk_url, headers={"Authorization": "Bearer " + API_TOKEN})
    with _opener.open(req, timeout=60) as resp:
        total = int(resp.headers.get("Content-Length") or 0)
        done = 0
        with open(dest_path, "wb") as f:
            while True:
                chunk = resp.read(65536)
                if not chunk:
                    break
                f.write(chunk)
                done += len(chunk)
                if progress_cb:
                    progress_cb(done, total)
    return dest_path


def install_update(apk_path):
    """Запустить установку APK (только Android)."""
    if platform != "android":
        return False
    try:
        from android.permissions import request_permissions, Permission
        request_permissions([Permission.WRITE_EXTERNAL_STORAGE])
    except Exception:
        pass
    try:
        from jnius import autoclass
        File = autoclass("java.io.File")
        Uri = autoclass("android.net.Uri")
        Intent = autoclass("android.content.Intent")
        Settings = autoclass("android.provider.Settings")
        pymt = autoclass("android.provider.Settings$ACTION_MANAGE_UNKNOWN_APP_SOURCES")
        ctx = autoclass("org.kivy.android.PythonActivity").mActivity
        # Разрешить установку из источника
        try:
            if not ctx.getPackageManager().canRequestPackageInstalls():
                i = Intent(pymt)
                i.setData(Uri.parse("package:" + ctx.getPackageName()))
                ctx.startActivity(i)
                return True
        except Exception:
            pass
        apk_file = File(apk_path)
        if int(android.os.Build.VERSION.SDK_INT) >= 24:
            uri = Uri.fromFile(apk_file)  # для file:// нужен FileProvider,
            # но на практиче пути приложения доступны и так через хранилище
        else:
            uri = Uri.fromFile(apk_file)
        intent = Intent(Intent.ACTION_VIEW)
        intent.setDataAndType(uri, "application/vnd.android.package-archive")
        intent.setFlags(Intent.FLAG_ACTIVITY_NEW_TASK)
        ctx.startActivity(intent)
        return True
    except Exception as e:
        print("install error:", e)
        return False
