#!/usr/bin/env python3
"""
Vigila los cupos disponibles del reporte Power BI "Disponibilidad Vértice"
y avisa por Discord, Telegram (o por consola) cuando cambian.

Uso:
    python3 vigilar_cupos.py          # ejecución normal (para cron)
    python3 vigilar_cupos.py --ver    # solo muestra la tabla actual, sin avisar ni guardar
    python3 vigilar_cupos.py --test   # envía un mensaje de prueba

Solo usa la librería estándar de Python 3.8+ (no hay que instalar nada).
"""
import json
import os
import sys
import time
import uuid
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path

# ============================ CONFIGURACIÓN ============================
FECHA_DESDE = "2026-11-24"            # primer día a vigilar
FECHA_HASTA = "2026-11-26"            # último día a vigilar (incluido)
ALOJAMIENTOS = ["Paine Grande", "Grey"]  # puedes poner varios
EXCLUIR_DETALLE = ["Sitios Grupales"] # mismo filtro que usa el reporte
SOLO_AVISAR_SI_AUMENTA = False        # True = avisar solo cuando se liberan cupos

# Alarma fuerte cuando se liberan cupos: se repite N veces por cada DÍA que libera cupos.
# {n} = cupos liberados ese día, {sitio} = sitio en mayúsculas, {dia} = "24-11".
MENSAJE_LIBERACION = "CSM ! SE LIBERARON {n} EN {sitio} EL {dia} CSM"
REPETIR_LIBERACION = 10

# True = en CADA ejecución llega un mensaje por sitio con sus días y cupos,
#        aunque no haya cambios (con la tarea cada 10 min = 1 mensaje por sitio cada 10 min).
# False = el mensaje por sitio llega solo cuando ese sitio cambia.
RESUMEN_EN_CADA_EJECUCION = True

# False = los resúmenes "sin cambios" llegan sin mencionarte (sin notificación al
# celular). Las alertas de cupos liberados, los cambios y los errores siempre te mencionan.
MENCIONAR_EN_RESUMEN = False

# Discord (opcional): URL del webhook del canal donde quieres las alertas.
# Se toma del secret DISCORD_WEBHOOK_URL de GitHub (no lo escribas aquí).
DISCORD_WEBHOOK_URL = os.environ.get("DISCORD_WEBHOOK_URL", "")
DISCORD_MENCION = os.environ.get("DISCORD_MENCION") or "<@350749732122918912>"  # @rulz1

# Opcional: un canal de Discord distinto por sitio. Los sitios que no estén aquí
# (o que tengan "") usan DISCORD_WEBHOOK_URL.
DISCORD_WEBHOOK_POR_SITIO = {
    "Grey": "",
    "Paine Grande": "",
}

# Telegram (opcional). Si no hay ningún canal definido, los avisos se imprimen en consola.
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

ARCHIVO_ESTADO = Path(__file__).resolve().with_name("estado_cupos.json")
# =======================================================================

URL = ("https://wabi-south-central-us-api.analysis.windows.net"
       "/public/reports/querydata?synchronous=true")
RESOURCE_KEY = "d052e2cd-6777-420b-937a-1c6335461a01"
DATASET_ID = "0ee7a608-6552-4b6d-bda2-17ee6f495a3c"
REPORT_ID = "ca488c99-3cb8-4604-ba59-c792bbb90da9"
VISUAL_ID = "10d537cc789a78b9036c"
MODEL_ID = 13163927

MESES = {m: i for i, m in enumerate(
    ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
     "agosto", "septiembre", "octubre", "noviembre", "diciembre"], start=1)}


# --------------------------- Construir consulta ---------------------------
def _col(src, prop):
    return {"Column": {"Expression": {"SourceRef": {"Source": src}}, "Property": prop}}


def _lit(valor):
    return {"Literal": {"Value": valor}}


def _texto(s):
    return "'" + s.replace("'", "''") + "'"


def _nivel_fecha(nivel):
    return {
        "HierarchyLevel": {
            "Expression": {"Hierarchy": {
                "Expression": {"PropertyVariationSource": {
                    "Expression": {"SourceRef": {"Source": "c"}},
                    "Name": "Variación", "Property": "fecha"}},
                "Hierarchy": "Jerarquía de fechas"}},
            "Level": nivel},
        "Name": f"Consulta1.fecha.Variación.Jerarquía de fechas.{nivel}",
        "NativeReferenceName": f"fecha {nivel}",
    }


def construir_body():
    desde = date.fromisoformat(FECHA_DESDE)
    hasta_excl = date.fromisoformat(FECHA_HASTA) + timedelta(days=1)
    where = [
        {"Condition": {"And": {
            "Left": {"Comparison": {"ComparisonKind": 2, "Left": _col("c", "fecha"),
                                    "Right": _lit(f"datetime'{desde}T00:00:00'")}},
            "Right": {"Comparison": {"ComparisonKind": 3, "Left": _col("c", "fecha"),
                                     "Right": _lit(f"datetime'{hasta_excl}T00:00:00'")}},
        }}},
        {"Condition": {"In": {
            "Expressions": [_col("c", "alojamiento")],
            "Values": [[_lit(_texto(a))] for a in ALOJAMIENTOS]}}},
    ]
    if EXCLUIR_DETALLE:
        where.append({"Condition": {"Not": {"Expression": {"In": {
            "Expressions": [_col("m", "Detalle")],
            "Values": [[_lit(_texto(d))] for d in EXCLUIR_DETALLE]}}}}})

    query = {
        "Version": 2,
        "From": [{"Name": "c", "Entity": "Consulta1", "Type": 0},
                 {"Name": "m", "Entity": "Maestro", "Type": 0}],
        "Select": [
            _nivel_fecha("Año"), _nivel_fecha("Mes"), _nivel_fecha("Día"),
            {**_col("c", "alojamiento"), "Name": "Consulta1.alojamiento",
             "NativeReferenceName": "Ubicación"},
            {**_col("m", "Tipo"), "Name": "Maestro.Tipo", "NativeReferenceName": "Tipo"},
            {"Aggregation": {"Expression": _col("c", "cupos_disponibles"), "Function": 0},
             "Name": "Sum(Consulta1.cupos_disponibles)",
             "NativeReferenceName": "Suma de cupos_disponibles"},
        ],
        "Where": where,
    }
    binding = {
        "Primary": {"Groupings": [{"Projections": [3]}, {"Projections": [4]}]},
        "Secondary": {"Groupings": [{"Projections": [0]}, {"Projections": [1]},
                                    {"Projections": [2, 5]}]},
        "DataReduction": {"DataVolume": 3, "Primary": {"Window": {"Count": 100}},
                          "Secondary": {"Top": {"Count": 100}}},
        "Version": 1,
    }
    return {
        "version": "1.0.0",
        "queries": [{
            "Query": {"Commands": [{"SemanticQueryDataShapeCommand": {
                "Query": query, "Binding": binding, "ExecutionMetricsKind": 1}}]},
            "QueryId": "",
            "ApplicationContext": {"DatasetId": DATASET_ID,
                                   "Sources": [{"ReportId": REPORT_ID, "VisualId": VISUAL_ID}]},
        }],
        "cancelQueries": [],
        "modelId": MODEL_ID,
    }


def consultar():
    datos = json.dumps(construir_body()).encode("utf-8")
    req = urllib.request.Request(URL, data=datos, method="POST", headers={
        "Content-Type": "application/json;charset=UTF-8",
        "Accept": "application/json, text/plain, */*",
        "Origin": "https://app.powerbi.com",
        "Referer": "https://app.powerbi.com/",
        "X-PowerBI-ResourceKey": RESOURCE_KEY,
        "ActivityId": str(uuid.uuid4()),
        "RequestId": str(uuid.uuid4()),
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36",
    })
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.loads(r.read().decode("utf-8"))


# ----------------------------- Parsear DSR -----------------------------
def _items(lista, clave, esquemas, dicts):
    """Recorre una lista DSR resolviendo esquema (S), repetidos (R), nulos (Ø) y diccionarios."""
    previo = {}
    for it in lista:
        if "S" in it:
            esquemas[clave] = it["S"]
        esquema = esquemas.get(clave, [])
        rep, nul = it.get("R", 0), it.get("Ø", 0)
        vals = {}
        for i, s in enumerate(esquema):
            n = s["N"]
            if rep >> i & 1:
                v = previo.get(n)
            elif nul >> i & 1:
                v = None
            else:
                v = it.get(n)
                if "DN" in s and isinstance(v, int):
                    v = dicts[s["DN"]][v]
            vals[n] = v
        previo = vals
        yield it, vals


def parsear(respuesta):
    """Devuelve {"Alojamiento|Tipo|AAAA-MM-DD": cupos}."""
    res = respuesta["results"][0]["result"]
    if "error" in res or "data" not in res:
        raise RuntimeError(f"Power BI devolvió un error: {json.dumps(res)[:500]}")
    ds = res["data"]["dsr"]["DS"][0]
    if "odata.error" in ds:
        raise RuntimeError(f"Error en la consulta: {ds['odata.error']}")
    dicts = ds.get("ValueDicts", {})
    esquemas = {}

    # Columnas: año → mes → día
    fechas = []
    for y_it, y in _items(ds.get("SH", [{}])[0].get("DM2", []), "DM2", esquemas, dicts):
        for m_grp in y_it.get("M", []):
            for m_it, m in _items(m_grp.get("DM3", []), "DM3", esquemas, dicts):
                mes = m["G3"]
                mes_num = MESES[mes.lower()] if isinstance(mes, str) else int(mes)
                for d_grp in m_it.get("M", []):
                    for _, d in _items(d_grp.get("DM4", []), "DM4", esquemas, dicts):
                        fechas.append(date(int(y["G2"]), mes_num, int(d["G4"])).isoformat())

    # Filas: alojamiento → tipo, con celdas en X
    cupos = {}
    for a_it, a in _items(ds.get("PH", [{}])[0].get("DM0", []), "DM0", esquemas, dicts):
        for t_grp in a_it.get("M", []):
            for t_it, t in _items(t_grp.get("DM1", []), "DM1", esquemas, dicts):
                col = 0
                for celda_it, celda in _items(t_it.get("X", []), "X", esquemas, dicts):
                    if "I" in celda_it:
                        col = celda_it["I"]
                    if col < len(fechas):
                        cupos[f"{a['G0']}|{t['G1']}|{fechas[col]}"] = celda.get("M0")
                    col += 1
                # Celdas no informadas = sin datos
                for f in fechas:
                    cupos.setdefault(f"{a['G0']}|{t['G1']}|{f}", None)
    return cupos


# ------------------------------ Avisos ------------------------------
def _webhook(sitio=None):
    return (DISCORD_WEBHOOK_POR_SITIO.get(sitio) or DISCORD_WEBHOOK_URL) if sitio else DISCORD_WEBHOOK_URL


def _enviar_discord(url, texto, mencionar=True):
    contenido = texto
    if mencionar and DISCORD_MENCION:
        contenido += f"\n{DISCORD_MENCION}"
    contenido = contenido[:2000]  # límite de Discord
    cuerpo = json.dumps({"content": contenido,
                         "allowed_mentions": {"parse": ["everyone", "users", "roles"]}}).encode()
    for intento in range(5):
        req = urllib.request.Request(
            url, data=cuerpo, method="POST",
            headers={"Content-Type": "application/json",
                     # Discord rechaza el User-Agent por defecto de Python
                     "User-Agent": "vigilar-cupos/1.0"})
        try:
            urllib.request.urlopen(req, timeout=10).read()
            return
        except urllib.error.HTTPError as e:
            if e.code == 429:  # demasiados mensajes seguidos: esperar lo que pide Discord
                try:
                    espera = float(json.loads(e.read().decode()).get("retry_after", 2))
                except Exception:
                    espera = 2
                time.sleep(espera + 0.2)
                continue
            print(f"No se pudo enviar a Discord: {e}", file=sys.stderr)
            return
        except Exception as e:
            print(f"No se pudo enviar a Discord: {e}", file=sys.stderr)
            return


def avisar(texto, sitio=None, mencionar=True):
    """Envía un aviso. Con sitio, usa el canal de ese sitio si está configurado."""
    print(texto)
    url = _webhook(sitio)
    if url:
        _enviar_discord(url, texto, mencionar)
    elif sitio is None:
        # Aviso general sin canal general: mandarlo a los canales de cada sitio
        for u in dict.fromkeys(u for u in DISCORD_WEBHOOK_POR_SITIO.values() if u):
            _enviar_discord(u, texto, mencionar)
    if not (TELEGRAM_TOKEN and TELEGRAM_CHAT_ID):
        return
    datos = urllib.parse.urlencode({"chat_id": TELEGRAM_CHAT_ID, "text": texto}).encode()
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    try:
        urllib.request.urlopen(url, data=datos, timeout=10).read()
    except Exception as e:
        print(f"No se pudo enviar a Telegram: {e}", file=sys.stderr)


def _sitio(clave):
    return clave.split("|")[0]


DIAS_SEMANA = ["lun", "mar", "mié", "jue", "vie", "sáb", "dom"]


def _dia(fecha_iso):
    """'2026-11-25' -> 'mié 25/11'."""
    d = date.fromisoformat(fecha_iso)
    return f"{DIAS_SEMANA[d.weekday()]} {d.day:02d}/{d.month:02d}"


def _fmt_clave(clave):
    """'Camping · mié 25/11' (el sitio va en el título del mensaje)."""
    _, tipo, f = clave.split("|")
    return f"{tipo} · {_dia(f)}"


def _por_sitio(cupos):
    grupos = {}
    for k, v in sorted(cupos.items()):
        grupos.setdefault(_sitio(k), {})[k] = v
    return grupos


def tabla(grupo, marcas=None):
    """Tabla de un sitio: una fila por día, una columna por tipo.
    marcas = {clave: "↑" o "↓"} para señalar las celdas que cambiaron."""
    if not grupo:
        return "_(sin datos)_"
    marcas = marcas or {}
    sitio = _sitio(next(iter(grupo)))
    tipos = sorted({k.split("|")[1] for k in grupo})
    fechas = sorted({k.split("|")[2] for k in grupo})
    anchos = [max(len(t), 3) + 2 for t in tipos]
    lineas = [" " * 9 + "".join((t + " ").rjust(w) for t, w in zip(tipos, anchos))]
    for f in fechas:
        fila = _dia(f).ljust(9)
        for t, w in zip(tipos, anchos):
            k = f"{sitio}|{t}|{f}"
            v = grupo.get(k)
            fila += (("—" if v is None else str(v)) + marcas.get(k, " ")).rjust(w)
        lineas.append(fila.rstrip())
    return "```\n" + "\n".join(lineas) + "\n```"


# ------------------------------ Principal ------------------------------
def cargar_estado():
    try:
        return json.loads(ARCHIVO_ESTADO.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def guardar_estado(estado):
    ARCHIVO_ESTADO.write_text(json.dumps(estado, ensure_ascii=False, indent=2), encoding="utf-8")


def main():
    if "--test" in sys.argv:
        for sitio in ALOJAMIENTOS:
            avisar(f"### ✅ Prueba de alerta · {sitio}\nSi ves esto, las notificaciones funcionan.", sitio)
        return

    estado = cargar_estado()
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M")

    try:
        cupos = parsear(consultar())
        if not cupos:
            raise RuntimeError("La consulta no devolvió datos (¿fechas sin disponibilidad cargada?)")
    except Exception as e:
        print(f"[{ahora}] ERROR: {e}", file=sys.stderr)
        # Avisar el error solo una vez para no llenarte de mensajes
        if "--ver" not in sys.argv and not estado.get("error"):
            avisar(f"## ⚠️ El monitor de cupos falló\n```\n{e}\n```\n"
                   "Puede que el reporte haya cambiado. Te aviso cuando vuelva a funcionar.")
            estado["error"] = True
            guardar_estado(estado)
        sys.exit(1)

    if "--ver" in sys.argv:
        for sitio, grupo in _por_sitio(cupos).items():
            print(f"📍 {sitio}\n{tabla(grupo)}\n")
        return

    anteriores = estado.get("cupos")
    if estado.get("error"):
        avisar("### ✅ El monitor de cupos volvió a funcionar")

    anteriores = anteriores or {}
    primera_vez = not anteriores
    cambios_por_sitio, liberados_por_sitio, marcas = {}, {}, {}
    if not primera_vez:
        for k in sorted(set(cupos) | set(anteriores)):
            antes, ahora_v = anteriores.get(k), cupos.get(k)
            if antes == ahora_v:
                continue
            subio = (ahora_v or 0) > (antes or 0)
            if subio:
                fecha = k.split("|")[2]
                por_dia = liberados_por_sitio.setdefault(_sitio(k), {})
                por_dia[fecha] = por_dia.get(fecha, 0) + (ahora_v or 0) - (antes or 0)
            if SOLO_AVISAR_SI_AUMENTA and not subio:
                continue
            flecha = "⬆️" if subio else "⬇️"
            marcas[k] = "↑" if subio else "↓"
            cambios_por_sitio.setdefault(_sitio(k), []).append(
                f"{flecha} **{_fmt_clave(k)}:** {antes if antes is not None else '—'}"
                f" → **{ahora_v if ahora_v is not None else '—'}**")

    grupos = _por_sitio(cupos)
    hora = datetime.now().strftime("%H:%M")
    for sitio in ALOJAMIENTOS:
        grupo = grupos.get(sitio)
        if not grupo:
            avisar(f"### ⚠️ {sitio}\nEl reporte no devolvió datos para este sitio. "
                   "Revisa que el nombre esté escrito igual que en el Power BI.", sitio)
            continue

        # 1) Alertas extra: 10 mensajes por cada día que liberó cupos
        for fecha, n in sorted(liberados_por_sitio.get(sitio, {}).items()):
            d = date.fromisoformat(fecha)
            grito = "## 🚨 " + MENSAJE_LIBERACION.format(n=n, sitio=sitio.upper(),
                                                        dia=f"{d.day:02d}-{d.month:02d}")
            for _ in range(REPETIR_LIBERACION):
                avisar(grito, sitio)
                time.sleep(1)  # para no chocar con el límite de Discord

        # 2) Notificación del sitio con sus días y números
        cambios = cambios_por_sitio.get(sitio)
        sitio_marcas = {k: m for k, m in marcas.items() if _sitio(k) == sitio}
        if primera_vez:
            texto = f"## 🏕️ {sitio}\n-# Monitor iniciado · {hora}\n{tabla(grupo)}"
            mencionar = MENCIONAR_EN_RESUMEN
        elif cambios:
            texto = (f"## 🔔 {sitio} · hay cambios\n" + "\n".join(cambios)
                     + f"\n{tabla(grupo, sitio_marcas)}\n-# {hora}")
            mencionar = True
        elif RESUMEN_EN_CADA_EJECUCION:
            texto = f"### 📍 {sitio}\n-# Sin cambios · {hora}\n{tabla(grupo)}"
            mencionar = MENCIONAR_EN_RESUMEN
        else:
            continue
        avisar(texto, sitio, mencionar)

    if not primera_vez and not cambios_por_sitio:
        print(f"[{ahora}] Sin cambios.")

    guardar_estado({"cupos": cupos, "actualizado": ahora, "error": False})


if __name__ == "__main__":
    main()
