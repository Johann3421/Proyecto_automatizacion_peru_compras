"""
Módulo de notificaciones via WhatsApp usando Evolution API (self-hosted).
Soporta envío de texto y archivos (Excel) sin costos de terceros.

Arquitectura de dos niveles:
  - wsp_server.json  → credenciales del servidor (admin, bundled en el .exe)
                       { "base_url": "...", "apikey": "...", "instance": "..." }
  - wsp_config.json  → configuración del usuario (solo teléfono + preferencias)
                       { "telefono": "51987654321", "notif_fin": true, "notif_error": true }
"""

from __future__ import annotations

import base64
import json
import logging
import threading
from pathlib import Path
from typing import Optional

log = logging.getLogger("peru_compras_bot")

# ---------------------------------------------------------------------------
# Rutas (se asignan desde gui.py antes de cualquier uso)
# ---------------------------------------------------------------------------
_SERVER_FILE: Optional[Path] = None   # wsp_server.json — bundled por el admin
_CONFIG_FILE: Optional[Path] = None   # wsp_config.json — escrito por el usuario


def set_server_path(path: Path) -> None:
    global _SERVER_FILE
    _SERVER_FILE = path


def set_config_path(path: Path) -> None:
    global _CONFIG_FILE
    _CONFIG_FILE = path


# ---------------------------------------------------------------------------
# Servidor (credenciales del admin — solo lectura)
# ---------------------------------------------------------------------------
def cargar_servidor() -> dict | None:
    """Lee wsp_server.json. Devuelve el dict o None si no existe / incompleto."""
    if _SERVER_FILE is None or not _SERVER_FILE.exists():
        return None
    try:
        with open(_SERVER_FILE, encoding="utf-8") as f:
            data = json.load(f)
        if all(k in data for k in ("base_url", "apikey", "instance")):
            return data
    except Exception as e:
        log.warning(f"[WSP] No se pudo leer wsp_server.json: {e}")
    return None


def servidor_configurado() -> bool:
    """True si wsp_server.json existe y tiene las claves necesarias."""
    return cargar_servidor() is not None


# ---------------------------------------------------------------------------
# Config de usuario (teléfono + preferencias)
# ---------------------------------------------------------------------------
def cargar_config() -> dict | None:
    """Lee wsp_config.json. Devuelve el dict o None si no existe / inválido."""
    if _CONFIG_FILE is None or not _CONFIG_FILE.exists():
        return None
    try:
        with open(_CONFIG_FILE, encoding="utf-8") as f:
            data = json.load(f)
        if "telefono" in data and data["telefono"]:
            return data
    except Exception as e:
        log.warning(f"[WSP] No se pudo leer wsp_config.json: {e}")
    return None


def guardar_config(
    telefono: str,
    notif_fin: bool = True,
    notif_error: bool = True,
) -> None:
    """Escribe la configuración del usuario en wsp_config.json."""
    if _CONFIG_FILE is None:
        return
    data = {
        "telefono": telefono,
        "notif_fin": notif_fin,
        "notif_error": notif_error,
    }
    try:
        with open(_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception as e:
        log.error(f"[WSP] No se pudo guardar wsp_config.json: {e}")


# ---------------------------------------------------------------------------
# Config combinada (servidor + usuario)
# ---------------------------------------------------------------------------
def _cfg_completa() -> dict | None:
    """Devuelve un dict con todos los campos necesarios para enviar, o None."""
    srv = cargar_servidor()
    usr = cargar_config()
    if srv is None or usr is None:
        return None
    return {**srv, **usr}


# ---------------------------------------------------------------------------
# Llamadas a Evolution API
# ---------------------------------------------------------------------------
def _headers(apikey: str) -> dict:
    return {"apikey": apikey, "Content-Type": "application/json"}


def _numero_limpio(telefono: str) -> str:
    return "".join(c for c in telefono if c.isdigit())


def _enviar_texto_raw(cfg: dict, mensaje: str, timeout: int = 15) -> bool:
    try:
        import requests
        payload = {
            "number": _numero_limpio(cfg["telefono"]),
            "text": mensaje,
        }
        resp = requests.post(
            f"{cfg['base_url']}/message/sendText/{cfg['instance']}",
            headers=_headers(cfg["apikey"]),
            json=payload,
            timeout=timeout,
        )
        if resp.status_code in (200, 201):
            log.info(f"[WSP] Mensaje enviado ({resp.status_code})")
            return True
        log.warning(f"[WSP] Respuesta inesperada texto: {resp.status_code} — {resp.text[:200]}")
        return False
    except Exception as e:
        log.error(f"[WSP] Error enviando texto: {e}")
        return False


def _enviar_archivo_raw(cfg: dict, ruta: Path, caption: str = "", timeout: int = 30) -> bool:
    try:
        import requests
        with open(ruta, "rb") as f:
            contenido_b64 = base64.b64encode(f.read()).decode()
        payload = {
            "number": _numero_limpio(cfg["telefono"]),
            "mediatype": "document",
            "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "caption": caption,
            "media": contenido_b64,
            "fileName": ruta.name,
        }
        resp = requests.post(
            f"{cfg['base_url']}/message/sendMedia/{cfg['instance']}",
            headers=_headers(cfg["apikey"]),
            json=payload,
            timeout=timeout,
        )
        if resp.status_code in (200, 201):
            log.info(f"[WSP] Archivo '{ruta.name}' enviado ({resp.status_code})")
            return True
        log.warning(f"[WSP] Respuesta inesperada archivo: {resp.status_code} — {resp.text[:200]}")
        return False
    except Exception as e:
        log.error(f"[WSP] Error enviando archivo: {e}")
        return False


# ---------------------------------------------------------------------------
# API pública
# ---------------------------------------------------------------------------
def testear_conexion_servidor(telefono: str) -> tuple[bool, str]:
    """
    Envía un mensaje de prueba usando las credenciales del servidor.
    Devuelve (True, "") si funcionó o (False, mensaje_error).
    """
    srv = cargar_servidor()
    if srv is None:
        return False, "wsp_server.json no encontrado. Contacta al administrador del sistema."
    try:
        import requests  # noqa: F401
    except ImportError:
        return False, "La librería 'requests' no está instalada."

    cfg = {**srv, "telefono": telefono}
    ok = _enviar_texto_raw(cfg, "✅ Peru Compras Bot — Conexión verificada correctamente.", timeout=10)
    if ok:
        return True, ""
    return False, "No se pudo enviar el mensaje. Verifica que tu número sea correcto y que el servidor esté activo."


def enviar_notificacion(
    evento: str,
    total: int,
    exitos: int,
    fallos: int,
    modo: str,
    ruta_excel: Path | None = None,
) -> None:
    """
    Envía la notificación en un hilo separado (no bloquea la GUI).
    evento: "FIN_OK" | "FIN_FALLOS" | "ERROR_CRITICO" | "INTERRUMPIDO"
    """
    cfg = _cfg_completa()
    if cfg is None:
        return  # no configurado

    if evento in ("FIN_OK", "FIN_FALLOS") and not cfg.get("notif_fin", True):
        return
    if evento in ("ERROR_CRITICO", "INTERRUMPIDO") and not cfg.get("notif_error", True):
        return

    def _enviar():
        from datetime import datetime
        ahora = datetime.now().strftime("%d/%m/%Y %H:%M")
        etiqueta = {
            "FIN_OK":        "✅ Proceso finalizado",
            "FIN_FALLOS":    "⚠️ Proceso finalizado con errores",
            "ERROR_CRITICO": "❌ Error crítico — proceso interrumpido",
            "INTERRUMPIDO":  "⏹️ Proceso detenido manualmente",
        }.get(evento, f"ℹ️ {evento}")

        lineas = [
            "*Peru Compras Bot*",
            etiqueta,
            "",
            f"📅 {ahora}",
            f"📋 Modo: {modo}",
            f"📦 Total: {total}",
            f"✔️ Exitosos: {exitos}",
            f"✖️ Fallidos: {fallos}",
        ]
        if ruta_excel and ruta_excel.exists():
            lineas.append(f"📊 Reporte: {ruta_excel.name}")

        ok_texto = _enviar_texto_raw(cfg, "\n".join(lineas))
        if ok_texto and ruta_excel and ruta_excel.exists():
            caption = f"Reporte {ahora} — {exitos}/{total} exitosos"
            _enviar_archivo_raw(cfg, ruta_excel, caption=caption)

    threading.Thread(target=_enviar, daemon=True, name="wsp-notif").start()

