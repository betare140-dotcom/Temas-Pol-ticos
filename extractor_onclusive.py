import io
import re
import html
import unicodedata
from collections import Counter
from urllib.parse import urlparse, urlunparse

import pandas as pd
import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

from sentimiento_local import (
    SKLEARN_DISPONIBLE,
    POSITIVA_INFORMATIVA,
    NEGATIVA_CRITICA,
    REVISAR,
    entrenar_desde_archivos,
    clasificar_registros,
    resumen_entrenamiento,
    csv_revision,
    guardar_aprendizaje_interno,
    sincronizar_base_con_github,
    estado_base_interna,
)

try:
    import requests
except Exception:  # requests es opcional para resolución remota de usuarios
    requests = None


REDES_PERMITIDAS = {"X", "FACEBOOK", "INSTAGRAM", "TIKTOK"}


def quitar_acentos(texto):
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return ""
    return "".join(
        c
        for c in unicodedata.normalize("NFD", str(texto))
        if unicodedata.category(c) != "Mn"
    ).lower()


def obtener_campo(row, candidatos):
    columnas = {quitar_acentos(str(c).strip()): c for c in row.index}
    for candidato in candidatos:
        original = columnas.get(quitar_acentos(candidato))
        if original is None:
            continue
        val = row[original]
        if val is None or pd.isna(val):
            continue
        texto = str(val).strip()
        if texto and texto.lower() not in {"nan", "none", "null"}:
            return texto
    return ""


def cargar_onclusive(file):
    raw = file.read()
    file.seek(0)
    try:
        hojas = pd.read_excel(io.BytesIO(raw), sheet_name=None)
        if hojas:
            return pd.concat(hojas.values(), ignore_index=True)
    except Exception:
        pass

    for encoding in ("utf-8-sig", "utf-8", "latin1"):
        try:
            return pd.read_csv(io.BytesIO(raw), encoding=encoding, on_bad_lines="skip")
        except Exception:
            continue
    raise ValueError("No fue posible leer el archivo como Excel o CSV.")


def normalizar_red(valor, link=""):
    t = quitar_acentos(valor)
    u = (link or "").lower()

    if "twitter" in t or re.search(r"(^|\W)x($|\W)", t) or "x.com/" in u or "twitter.com/" in u:
        return "X"
    if "facebook" in t or "fb" == t.strip() or "facebook.com/" in u or "fb.watch/" in u:
        return "FACEBOOK"
    if "instagram" in t or "instagram.com/" in u:
        return "INSTAGRAM"
    if "tiktok" in t or "tiktok.com/" in u or "vm.tiktok.com/" in u:
        return "TIKTOK"
    return ""


def detectar_red(row):
    link = obtener_link_principal(row)
    for campo in [
        "Media type", "Media Type", "Tipo de Medio", "Tipo de medio",
        "Social Network", "Social network", "Red Social", "Fuente", "Canal",
        "Source type", "Platform", "Plataforma"
    ]:
        val = obtener_campo(row, [campo])
        red = normalizar_red(val, link)
        if red:
            return red
    return normalizar_red("", link)


def obtener_link_principal(row):
    candidatos = [
        "Link URL Medio", "Link de Nota", "URL", "Url", "Enlace", "Link",
        "Permalink", "Post URL", "Post Url", "Original URL"
    ]
    urls = []
    for c in candidatos:
        val = obtener_campo(row, [c])
        if val.startswith("http") and val not in urls:
            urls.append(val)
    if not urls:
        for val in row.values:
            if val is None or pd.isna(val):
                continue
            s = str(val).strip()
            if s.startswith("http") and s not in urls:
                urls.append(s)
    if not urls:
        return ""

    externos = [u for u in urls if "hanakua.mx" not in u.lower()]
    return (externos or urls)[0]


def canonicalizar_link(url):
    if not url:
        return ""
    url = html.unescape(str(url).strip())
    try:
        p = urlparse(url)
        host = p.netloc.lower().replace("www.", "")
        path = p.path
        # En X/Facebook /photo/1 o /video/1 puede ser solo una variante del mismo post.
        # En TikTok /video/<id> es parte esencial del permalink y debe conservarse.
        if host in {"x.com", "twitter.com", "mobile.twitter.com", "facebook.com", "m.facebook.com"}:
            path = re.sub(r"/(photo|video)/\d+/?$", "", path, flags=re.I)
        path = re.sub(r"/+", "/", path).rstrip("/")
        if host == "twitter.com":
            host = "x.com"
        return urlunparse((p.scheme or "https", host, path, "", "", ""))
    except Exception:
        return re.sub(r"[?#].*$", "", url).rstrip("/")


def limpiar_texto_publicacion(texto):
    if not texto:
        return ""
    t = str(texto)
    # URLs dentro del contenido: se conserva únicamente el enlace principal por separado.
    t = re.sub(r"https?://\S+", "", t)
    # Sufijos de variante que a veces quedan sueltos al exportar contenido.
    t = re.sub(r"(?<!\w)/(?:photo|video)/\d+\b", "", t, flags=re.I)
    # Emojis escritos como códigos :purple_heart:, :warning:, etc.
    t = re.sub(r":[A-Za-z0-9_+\-|]+:", "", t)
    # IDs numéricos de Facebook como @840385099165296
    t = re.sub(r"(?<!\w)@\d{7,}\b", "", t)
    t = re.sub(r"[ \t]+", " ", t)
    t = re.sub(r"\s*\n\s*", " ", t)
    return t.strip(" -\t\n")


def texto_de_fila(row):
    titulo = obtener_campo(row, ["Titulo", "Título", "Title", "Encabezado", "Tema"])
    contenido = obtener_campo(
        row,
        ["Contenido", "Detail", "Summary", "Síntesis", "Sintesis", "Nota", "Text", "Texto"]
    )

    # Onclusive suele exportar un Title truncado con "..." y después el Detail completo.
    # Antes se unían ambos y el Word repetía el inicio de la publicación. Ahora se detecta
    # cuando el título es solo un prefijo/resumen del contenido y se conserva el Detail.
    if contenido and titulo:
        titulo_core = re.split(r"(?:\.{3,}|…)", html.unescape(titulo), maxsplit=1)[0].strip()
        t_core = quitar_acentos(titulo_core)
        c_norm = quitar_acentos(html.unescape(contenido))
        t_norm = quitar_acentos(html.unescape(titulo))

        titulo_ya_esta = False
        if len(t_core) >= 24 and t_core in c_norm:
            titulo_ya_esta = True
        elif len(t_norm) >= 24 and t_norm in c_norm:
            titulo_ya_esta = True

        bruto = contenido if titulo_ya_esta else f"{titulo}. {contenido}"
    else:
        bruto = contenido or titulo

    return limpiar_texto_publicacion(bruto)


def normalizar_texto_dup(texto):
    t = quitar_acentos(texto)
    t = re.sub(r"\s+", " ", t)
    t = re.sub(r"[^a-z0-9#@ ]", "", t)
    return t.strip()


# Palabras funcionales que no ayudan a definir un tema.
STOPWORDS_TEMA = {
    "a", "al", "ante", "bajo", "como", "con", "contra", "de", "del", "desde",
    "durante", "e", "el", "ella", "ellas", "ellos", "en", "entre", "es", "esta",
    "este", "esto", "hacia", "hasta", "la", "las", "lo", "los", "o", "para", "pero",
    "por", "que", "se", "sin", "sobre", "su", "sus", "un", "una", "y"
}


def normalizar_busqueda(texto):
    """Normalización tolerante a acentos, hashtags, HTML y signos."""
    t = html.unescape(str(texto or ""))
    t = quitar_acentos(t)
    t = re.sub(r"https?://\\S+", " ", t)
    t = re.sub(r"[^a-z0-9@#_ ]+", " ", t)
    t = re.sub(r"\\s+", " ", t).strip()
    return t


def compactar_busqueda(texto):
    """Sirve para detectar #LaBonitaSánchez, @bonita_sanchez_oficial, etc."""
    return re.sub(r"[^a-z0-9]+", "", normalizar_busqueda(texto))


def parsear_lista_reglas(texto):
    """Acepta reglas separadas por coma, punto y coma o salto de línea."""
    if not texto:
        return []
    partes = re.split(r"[\\n,;]+", str(texto))
    salida = []
    vistos = set()
    for p in partes:
        p = p.strip()
        if not p:
            continue
        clave = compactar_busqueda(p)
        if clave and clave not in vistos:
            vistos.add(clave)
            salida.append(p)
    return salida


def terminos_desde_tema(tema):
    """Extrae automáticamente las palabras útiles del tema escrito por el usuario."""
    tokens = re.findall(r"[a-z0-9áéíóúñü]+", str(tema or "").lower())
    salida = []
    vistos = set()
    for token in tokens:
        n = quitar_acentos(token)
        if len(n) < 3 or n in STOPWORDS_TEMA or n in vistos:
            continue
        vistos.add(n)
        salida.append(n)
    return salida


def coincidencias_reglas(texto, reglas):
    """Devuelve las reglas encontradas, tolerando hashtags y guiones bajos."""
    normal = normalizar_busqueda(texto)
    compacto = compactar_busqueda(texto)
    encontrados = []
    for regla in reglas:
        rn = normalizar_busqueda(regla).strip()
        rc = compactar_busqueda(regla)
        if not rn:
            continue
        # Frases normales o versiones pegadas: #GabyBonitaSanchez / bonita_sanchez.
        if rn in normal or (len(rc) >= 4 and rc in compacto):
            encontrados.append(regla)
    return encontrados


def evaluar_relevancia_tema(
    texto,
    tema,
    aliases=None,
    terminos_extra=None,
    exclusiones=None,
    min_coincidencias=2,
):
    """
    Decide si una publicación pertenece al tema.

    Reglas:
    1) Si se proporcionan aliases/actor, debe aparecer al menos uno.
    2) El tema se convierte automáticamente en palabras útiles y se cuentan coincidencias.
    3) Los términos extra funcionan como sinónimos o conceptos adicionales.
    4) Las exclusiones son absolutas y se aplican al final.

    Retorna: (es_relevante, motivo, coincidencias)
    """
    aliases = aliases or []
    terminos_extra = terminos_extra or []
    exclusiones = exclusiones or []

    # Exclusiones explícitas: si el usuario las escribe, prevalecen.
    hits_exclusion = coincidencias_reglas(texto, exclusiones)
    if hits_exclusion:
        return False, f"Exclusión: {', '.join(hits_exclusion[:3])}", hits_exclusion

    hits_alias = coincidencias_reglas(texto, aliases)
    if aliases and not hits_alias:
        return False, "No aparece el actor/alias requerido", []

    terminos_base = terminos_desde_tema(tema)
    reglas_tema = list(terminos_base) + list(terminos_extra)
    # Evita que alias muy obvios cuenten dos veces como concepto del tema cuando el usuario
    # ya decidió exigirlos por separado.
    hits_tema = coincidencias_reglas(texto, reglas_tema)

    # En temas de una sola palabra basta una coincidencia; en los demás, por defecto dos.
    minimo = max(1, int(min_coincidencias or 1))
    if len(reglas_tema) == 1:
        minimo = 1

    if len(hits_tema) < minimo:
        detalle = ", ".join(hits_tema[:4]) if hits_tema else "ninguna"
        return False, f"Fuera de tema: {len(hits_tema)}/{minimo} coincidencias ({detalle})", hits_tema

    motivo = f"Tema: {len(hits_tema)} coincidencias"
    if hits_alias:
        motivo += f" · alias: {hits_alias[0]}"
    return True, motivo, hits_tema


def es_rt(row, texto, red):
    if red != "X":
        return False
    tipo = " ".join(
        obtener_campo(row, [c])
        for c in ["Tipo de Nota", "Post type", "Type", "Publication type", "Interaction type"]
    )
    combinado = quitar_acentos(f"{tipo} {texto}").strip()
    return bool(
        re.match(r"^rt\s+@", combinado)
        or " retweet" in f" {combinado}"
        or " retuit" in f" {combinado}"
        or combinado.startswith("retweet")
        or combinado.startswith("retuit")
    )


def parsear_fecha(valor):
    if valor is None or pd.isna(valor):
        return pd.NaT
    return pd.to_datetime(valor, dayfirst=True, errors="coerce")


def obtener_autor(row):
    return obtener_campo(
        row,
        ["Author name", "Autor", "Author", "Nombre del Medio", "Media name", "Fuente", "Page name", "Account name", "Canal"]
    )


def obtener_handle_exportado(row):
    return obtener_campo(
        row,
        ["Author handle (@username)", "Handle", "Username", "Screen Name", "Account", "Perfil", "User name"]
    ).lstrip("@")


def usuario_desde_url(red, url):
    if not url:
        return ""
    try:
        p = urlparse(url)
        host = p.netloc.lower().replace("www.", "")
        partes = [x for x in p.path.split("/") if x]
    except Exception:
        return ""

    if red == "X" and host in {"x.com", "twitter.com", "mobile.twitter.com"} and partes:
        candidato = partes[0]
        if candidato.lower() not in {"i", "intent", "search", "home", "share"}:
            return candidato.lstrip("@")

    if red == "TIKTOK":
        for parte in partes:
            if parte.startswith("@") and len(parte) > 1:
                return parte[1:]

    if red == "INSTAGRAM" and partes:
        reservado = {"p", "reel", "reels", "tv", "stories", "explore", "accounts", "share"}
        if partes[0].lower() not in reservado:
            return partes[0].lstrip("@")

    if red == "FACEBOOK" and partes:
        reservado = {"share", "reel", "watch", "photo", "groups", "events", "story.php", "permalink.php", "profile.php"}
        candidato = partes[0]
        if candidato.lower() not in reservado and not candidato.isdigit():
            return candidato.lstrip("@")
    return ""


@st.cache_data(ttl=86400, show_spinner=False)
def resolver_usuario_publico(url, red):
    """Mejor esfuerzo. No usa login ni intenta evadir bloqueos."""
    if not url or requests is None or red not in {"INSTAGRAM", "TIKTOK"}:
        return ""
    try:
        resp = requests.get(
            url,
            timeout=5,
            allow_redirects=True,
            headers={
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 Chrome/124 Safari/537.36"
                )
            },
        )
        if not resp.ok:
            return ""
        final_url = resp.url
        local = usuario_desde_url(red, final_url)
        if local:
            return local

        cuerpo = resp.text[:500000]
        metas = " ".join(re.findall(r'<meta[^>]+(?:content|name|property)=["\'][^"\']+["\'][^>]*>', cuerpo, flags=re.I))
        texto = html.unescape(re.sub(r"<[^>]+>", " ", metas))

        if red == "INSTAGRAM":
            patrones = [
                r"@([A-Za-z0-9._]{2,30})",
                r"([A-Za-z0-9._]{2,30})\s+on\s+Instagram",
            ]
        else:
            patrones = [r"@([A-Za-z0-9._-]{2,40})"]

        for patron in patrones:
            m = re.search(patron, texto, flags=re.I)
            if m:
                return m.group(1)
    except Exception:
        return ""
    return ""


def recortar_por_tema(texto, tema, max_chars=520):
    texto = (texto or "").strip()
    if len(texto) <= max_chars:
        return texto

    tokens = [
        x for x in re.findall(r"[a-z0-9áéíóúñü]+", tema.lower())
        if len(x) >= 4
    ]
    texto_low = texto.lower()
    posiciones = [texto_low.find(t) for t in tokens if texto_low.find(t) >= 0]
    centro = min(posiciones) if posiciones else 0

    inicio = max(0, centro - 120)
    fin = min(len(texto), inicio + max_chars)
    frag = texto[inicio:fin].strip()
    if inicio > 0:
        frag = "…" + frag
    if fin < len(texto):
        frag += "…"
    return frag


def etiqueta_lista(red, autor, handle):
    if red in {"X", "INSTAGRAM", "TIKTOK"}:
        return f"@{handle}" if handle else (autor or red.title())
    return autor or (f"@{handle}" if handle else "Facebook")


def fuente_desglose(red, autor, handle):
    if red == "X":
        if autor and handle:
            return f"{autor} @{handle}"
        return f"@{handle}" if handle else autor or "X"
    if red == "FACEBOOK":
        return autor or (f"@{handle}" if handle else "Facebook")
    if red in {"INSTAGRAM", "TIKTOK"}:
        if autor and handle:
            return f"{autor} @{handle}"
        return f"@{handle}" if handle else autor or red.title()
    return autor or handle or red


def procesar_onclusive(
    df,
    tema,
    resolver_remoto=False,
    aliases=None,
    terminos_extra=None,
    exclusiones=None,
    cuentas_excluir=None,
    min_coincidencias=2,
    deduplicar_texto=False,
    recortar_texto=False,
    redes_permitidas=None,
):
    """
    Procesa Onclusive con filtro temático real.

    Diferencia clave respecto a la versión anterior:
    - `tema` ya no se usa solo para recortar el texto: ahora decide qué entra y qué se excluye.
    - Por defecto NO elimina publicaciones distintas que tengan el mismo texto. Solo elimina
      el mismo permalink repetido. Esto conserva republicaciones válidas en X/Facebook/Instagram.
    """
    registros = []
    excluidos = []
    stats = Counter()
    vistos_links = set()
    vistos_texto_fuente = set()

    aliases = aliases or []
    terminos_extra = terminos_extra or []
    exclusiones = exclusiones or []
    cuentas_excluir = {compactar_busqueda(x) for x in (cuentas_excluir or []) if x}
    redes_ok = set(redes_permitidas or REDES_PERMITIDAS)

    for idx, row in df.iterrows():
        link_raw = obtener_link_principal(row)
        red = detectar_red(row)
        if red not in redes_ok:
            stats["otras_redes"] += 1
            continue

        texto = texto_de_fila(row)
        if not texto:
            stats["sin_texto"] += 1
            continue
        if es_rt(row, texto, red):
            stats["rt"] += 1
            continue

        link = canonicalizar_link(link_raw)
        autor = obtener_autor(row)
        handle = obtener_handle_exportado(row) or usuario_desde_url(red, link_raw)

        if not handle and resolver_remoto and red in {"INSTAGRAM", "TIKTOK"}:
            handle = resolver_usuario_publico(link_raw, red)

        if red == "FACEBOOK" and re.fullmatch(r"@?\d{7,}", autor or ""):
            autor = "Facebook"

        cuenta_key = compactar_busqueda(handle or autor or "")
        if cuenta_key and cuenta_key in cuentas_excluir:
            stats["cuenta_excluida"] += 1
            excluidos.append({
                "orden": idx, "red": red, "autor": autor, "handle": handle,
                "texto": texto, "link": link or link_raw,
                "motivo": "Cuenta excluida"
            })
            continue

        # Para decidir relevancia usamos texto + autor + handle. Esto permite reconocer
        # menciones como @bonita_sanchez_oficial aunque el cuerpo no repita el nombre.
        texto_clasificar = " ".join(x for x in [texto, autor, handle] if x)
        relevante, motivo_tema, hits_tema = evaluar_relevancia_tema(
            texto_clasificar,
            tema,
            aliases=aliases,
            terminos_extra=terminos_extra,
            exclusiones=exclusiones,
            min_coincidencias=min_coincidencias,
        )
        if not relevante:
            stats["fuera_tema"] += 1
            excluidos.append({
                "orden": idx, "red": red, "autor": autor, "handle": handle,
                "texto": texto, "link": link or link_raw,
                "motivo": motivo_tema,
            })
            continue

        # Mismo permalink = mismo post: sí se depura.
        if link and link in vistos_links:
            stats["dup_link"] += 1
            continue
        if link:
            vistos_links.add(link)

        # Mismo texto en otro post o plataforma NO se elimina por defecto.
        # Se deja como opción porque en monitoreo suele representar otro impacto real.
        if deduplicar_texto:
            fuente_key = quitar_acentos(handle or autor or red)
            texto_key = normalizar_texto_dup(texto)
            if texto_key and (fuente_key, texto_key) in vistos_texto_fuente:
                stats["dup_texto"] += 1
                continue
            if texto_key:
                vistos_texto_fuente.add((fuente_key, texto_key))

        fecha_raw = obtener_campo(
            row,
            ["Publish date", "Fecha", "Date", "Fecha de publicación", "Fecha de publicacion", "Published", "Publication date"]
        )
        fecha = parsear_fecha(fecha_raw)
        texto_salida = recortar_por_tema(texto, tema) if recortar_texto else texto

        registros.append({
            "orden": idx,
            "red": red,
            "autor": autor,
            "handle": handle,
            "texto": texto_salida,
            "link": link or link_raw,
            "fecha": fecha,
            "motivo_tema": motivo_tema,
            "coincidencias_tema": hits_tema,
        })

    registros.sort(
        key=lambda r: (
            pd.Timestamp.max if pd.isna(r["fecha"]) else r["fecha"].normalize(),
            r["orden"],
        )
    )
    return registros, stats, excluidos


def agregar_hipervinculo(parrafo, texto, url):
    part = parrafo.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(color)
    rPr.append(u)
    new_run.append(rPr)

    text = OxmlElement("w:t")
    text.text = texto
    new_run.append(text)
    hyperlink.append(new_run)
    parrafo._p.append(hyperlink)



def _titulo_sentimiento(etiqueta):
    mapa = {
        POSITIVA_INFORMATIVA: "POSITIVAS / INFORMATIVAS",
        NEGATIVA_CRITICA: "NEGATIVAS / CRÍTICAS",
        REVISAR: "REVISAR MANUALMENTE",
        "POSITIVA": "POSITIVAS",
        "INFORMATIVA": "INFORMATIVAS",
        "NEGATIVA": "NEGATIVAS / CRÍTICAS",
    }
    return mapa.get(
        str(etiqueta).upper(),
        str(etiqueta).replace("_", " ")
    )


def _orden_sentimientos(registros):
    valores = [
        str(r.get("sentimiento", "")).upper()
        for r in registros
        if r.get("sentimiento")
    ]

    preferido = [
        POSITIVA_INFORMATIVA,
        "POSITIVA",
        "INFORMATIVA",
        NEGATIVA_CRITICA,
        "NEGATIVA",
        REVISAR,
    ]

    salida = []
    for etiqueta in preferido + valores:
        if etiqueta in valores and etiqueta not in salida:
            salida.append(etiqueta)
    return salida


def crear_word_clasificado(registros, tema):
    doc = Document()
    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(10)

    p = doc.add_paragraph()
    r = p.add_run(f"REDES SOCIALES ({len(registros)})")
    r.bold = True

    # RESUMEN
    for etiqueta in _orden_sentimientos(registros):
        grupo = [
            x for x in registros
            if str(x.get("sentimiento", "")).upper() == etiqueta
        ]

        p = doc.add_paragraph()
        rr = p.add_run(
            f"{_titulo_sentimiento(etiqueta)} ({len(grupo)})"
        )
        rr.bold = True
        if etiqueta in {NEGATIVA_CRITICA, "NEGATIVA"}:
            rr.font.color.rgb = RGBColor(192, 0, 0)

        fuentes = [
            etiqueta_lista(
                x["red"], x["autor"], x["handle"]
            )
            for x in grupo
        ]
        conteos = Counter(fuentes)
        vistos = set()
        for fuente in fuentes:
            if fuente in vistos:
                continue
            vistos.add(fuente)
            pp = doc.add_paragraph()
            pp.paragraph_format.space_after = Pt(0)
            pp.add_run(
                f"{fuente} ({conteos[fuente]})"
                if conteos[fuente] > 1
                else fuente
            )

    p = doc.add_paragraph()
    r = p.add_run("DESGLOSE")
    r.bold = True

    # DESGLOSE
    for etiqueta in _orden_sentimientos(registros):
        grupo = [
            x for x in registros
            if str(x.get("sentimiento", "")).upper() == etiqueta
        ]

        p = doc.add_paragraph()
        rr = p.add_run(
            f"{_titulo_sentimiento(etiqueta)} ({len(grupo)})"
        )
        rr.bold = True
        if etiqueta in {NEGATIVA_CRITICA, "NEGATIVA"}:
            rr.font.color.rgb = RGBColor(192, 0, 0)

        fecha_actual = object()
        for item in grupo:
            fecha_key = (
                None
                if pd.isna(item["fecha"])
                else item["fecha"].strftime("%d.%m.%y")
            )

            if fecha_key != fecha_actual:
                fecha_actual = fecha_key
                if fecha_key:
                    pp = doc.add_paragraph()
                    run = pp.add_run(fecha_key)
                    run.bold = True

            pp = doc.add_paragraph()
            fuente = fuente_desglose(
                item["red"],
                item["autor"],
                item["handle"],
            )
            run = pp.add_run(fuente)
            run.bold = True
            run.add_break()

            pp.add_run(item["texto"])
            pp.add_run().add_break()

            if item["link"]:
                agregar_hipervinculo(
                    pp, item["link"], item["link"]
                )

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out


def crear_word(registros, tema):
    if registros and any(r.get("sentimiento") for r in registros):
        return crear_word_clasificado(registros, tema)

    doc = Document()
    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(10)

    p = doc.add_paragraph()
    r = p.add_run(f"REDES SOCIALES ({len(registros)})")
    r.bold = True

    etiquetas = [etiqueta_lista(r["red"], r["autor"], r["handle"]) for r in registros]
    conteos = Counter(etiquetas)
    vistos = set()
    for et in etiquetas:
        if et in vistos:
            continue
        vistos.add(et)
        p = doc.add_paragraph()
        p.add_run(f"{et} ({conteos[et]})" if conteos[et] > 1 else et)

    p = doc.add_paragraph()
    r = p.add_run("DESGLOSE")
    r.bold = True

    fecha_actual = object()
    for item in registros:
        fecha_key = None if pd.isna(item["fecha"]) else item["fecha"].strftime("%d.%m.%y")
        if fecha_key != fecha_actual:
            fecha_actual = fecha_key
            if fecha_key:
                p = doc.add_paragraph()
                rr = p.add_run(fecha_key)
                rr.bold = True

        # Un solo párrafo por mención. add_break() = salto manual tipo Shift+Enter.
        p = doc.add_paragraph()
        fuente = fuente_desglose(item["red"], item["autor"], item["handle"])
        rr = p.add_run(fuente)
        rr.bold = True
        rr.add_break()
        p.add_run(item["texto"])
        p.add_run().add_break()
        if item["link"]:
            agregar_hipervinculo(p, item["link"], item["link"])

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out


def crear_txt(registros):
    lineas = [f"REDES SOCIALES ({len(registros)})"]
    etiquetas = [etiqueta_lista(r["red"], r["autor"], r["handle"]) for r in registros]
    conteos = Counter(etiquetas)
    vistos = set()
    for et in etiquetas:
        if et not in vistos:
            vistos.add(et)
            lineas.append(f"{et} ({conteos[et]})" if conteos[et] > 1 else et)

    lineas += ["", "DESGLOSE"]
    fecha_actual = object()
    for item in registros:
        fecha_key = None if pd.isna(item["fecha"]) else item["fecha"].strftime("%d.%m.%y")
        if fecha_key != fecha_actual:
            fecha_actual = fecha_key
            if fecha_key:
                lineas += ["", fecha_key]
        lineas += [
            fuente_desglose(item["red"], item["autor"], item["handle"]),
            item["texto"],
            item["link"],
            "",
        ]
    return "\n".join(lineas).encode("utf-8")


def crear_html(registros):
    etiquetas = [etiqueta_lista(r["red"], r["autor"], r["handle"]) for r in registros]
    conteos = Counter(etiquetas)
    vistos = set()
    partes = [
        "<!doctype html><html><head><meta charset='utf-8'>",
        "<style>body{font-family:Arial,sans-serif;font-size:14px;line-height:1.35} .m{margin-bottom:14px}</style>",
        "</head><body>",
        f"<p><strong>REDES SOCIALES ({len(registros)})</strong></p>",
    ]
    for et in etiquetas:
        if et in vistos:
            continue
        vistos.add(et)
        txt = f"{et} ({conteos[et]})" if conteos[et] > 1 else et
        partes.append(f"<div>{html.escape(txt)}</div>")

    partes.append("<p><strong>DESGLOSE</strong></p>")
    fecha_actual = object()
    for item in registros:
        fecha_key = None if pd.isna(item["fecha"]) else item["fecha"].strftime("%d.%m.%y")
        if fecha_key != fecha_actual:
            fecha_actual = fecha_key
            if fecha_key:
                partes.append(f"<p><strong>{fecha_key}</strong></p>")
        fuente = html.escape(fuente_desglose(item["red"], item["autor"], item["handle"]))
        texto = html.escape(item["texto"])
        link = html.escape(item["link"] or "")
        link_html = f'<a href="{link}">{link}</a>' if link else ""
        partes.append(f"<div class='m'><strong>{fuente}</strong><br>{texto}<br>{link_html}</div>")
    partes.append("</body></html>")
    return "".join(partes).encode("utf-8")



def crear_excel_entrenamiento(registros, tema=""):
    """
    Genera un Excel listo para volver a cargarse como entrenamiento adicional.

    - La primera hoja se llama 'Entrenamiento' para que pd.read_excel()
      la lea directamente en sentimiento_local.py.
    - Las filas que siguen en REVISAR quedan con etiqueta vacía y no
      entrenarán al modelo hasta que el usuario las resuelva.
    """
    filas = []
    for r in registros:
        sentimiento_actual = str(r.get("sentimiento", "") or "").upper()
        etiqueta = "" if sentimiento_actual == REVISAR else sentimiento_actual

        usuario = ""
        if r.get("handle"):
            usuario = f"@{r['handle']}"
        elif r.get("autor"):
            usuario = r.get("autor", "")

        filas.append({
            "actor": "",
            "tema": tema,
            "red": r.get("red", ""),
            "fuente": r.get("autor", "") or usuario,
            "usuario": usuario,
            "texto": r.get("texto", ""),
            "enlace": r.get("link", ""),
            "etiqueta": etiqueta,
            "clasificacion_inicial": r.get(
                "sentimiento_inicial",
                r.get("sentimiento_predicho", "")
            ),
            "prediccion_modelo": r.get("sentimiento_predicho", ""),
            "confianza": round(
                float(r.get("confianza_sentimiento", 0) or 0), 4
            ),
            "origen": "revision_manual_streamlit",
        })

    df = pd.DataFrame(filas)

    resumen = pd.DataFrame([
        ["Tema", tema],
        ["Total de publicaciones", len(df)],
        [
            "POSITIVA_INFORMATIVA",
            int((df["etiqueta"] == POSITIVA_INFORMATIVA).sum())
        ],
        [
            "NEGATIVA_CRITICA",
            int((df["etiqueta"] == NEGATIVA_CRITICA).sum())
        ],
        [
            "Pendientes / REVISAR",
            int((df["etiqueta"] == "").sum())
        ],
    ], columns=["Concepto", "Valor"])

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        # Debe ir primero: sentimiento_local.py usa la primera hoja.
        df.to_excel(writer, sheet_name="Entrenamiento", index=False)
        resumen.to_excel(writer, sheet_name="Resumen", index=False)

        ws = writer.book["Entrenamiento"]
        ws.freeze_panes = "A2"
        anchos = {
            "A": 24, "B": 36, "C": 14, "D": 28, "E": 24,
            "F": 80, "G": 55, "H": 24, "I": 24, "J": 24,
            "K": 14, "L": 28,
        }
        for col, width in anchos.items():
            ws.column_dimensions[col].width = width

        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)

        ws2 = writer.book["Resumen"]
        ws2.freeze_panes = "A2"
        ws2.column_dimensions["A"].width = 30
        ws2.column_dimensions["B"].width = 45
        for cell in ws2[1]:
            cell.font = cell.font.copy(bold=True)

    out.seek(0)
    return out.getvalue()


def _usuario_visible(registro):
    if registro.get("handle"):
        return f"@{registro['handle']}"
    return str(registro.get("autor", "") or "").strip()


def _aplicar_usuario_editado(registro, valor):
    """
    Permite completar/corregir el usuario desde la tabla.

    Casos admitidos:
    - @usuario
    - usuario
    - Nombre visible @usuario
    """
    valor = str(valor or "").strip()

    if not valor:
        registro["handle"] = ""
        registro["autor"] = ""
        return

    coincidencia = re.search(r"@([A-Za-z0-9_.-]+)", valor)
    if coincidencia:
        registro["handle"] = coincidencia.group(1).strip()
        nombre = valor[:coincidencia.start()].strip(" -|")
        if nombre:
            registro["autor"] = nombre
        return

    # Si no incluye @, se conserva como nombre visible.
    registro["autor"] = valor
    if not registro.get("handle"):
        registro["handle"] = ""


def _sentimiento_para_descarga(registro):
    """
    El estado REVISAR es operativo, no bloquea la clasificación.

    Si el usuario todavía no revisó una fila de baja confianza,
    la descarga conserva la predicción original del modelo.
    """
    actual = str(registro.get("sentimiento", "") or "").upper().strip()

    if actual == REVISAR:
        pred = str(
            registro.get("sentimiento_predicho", "") or ""
        ).upper().strip()
        if pred in {POSITIVA_INFORMATIVA, NEGATIVA_CRITICA}:
            return pred

    return actual


def _registros_para_descarga(registros):
    salida = []
    for r in registros:
        nuevo = dict(r)
        final = _sentimiento_para_descarga(nuevo)
        if final:
            nuevo["sentimiento"] = final
        salida.append(nuevo)
    return salida


def _vista_registros(registros, con_sentimiento=False):
    filas = []
    for i, r in enumerate(registros):
        fila = {
            "N.º": i + 1,
            "Red": r.get("red", ""),
            "Usuario": _usuario_visible(r),
            "Texto": r.get("texto", ""),
        }

        if con_sentimiento:
            requiere_revision = (
                str(r.get("sentimiento", "") or "").upper() == REVISAR
            )
            fila.update({
                "Sentimiento corregido": r.get("sentimiento", ""),
                "Predicción del modelo": r.get(
                    "sentimiento_predicho", ""
                ),
                "Revisión": (
                    "REVISAR" if requiere_revision else "OK"
                ),
                "Confianza": float(
                    r.get("confianza_sentimiento", 0) or 0
                ),
            })

        fila.update({
            "Motivo": r.get("motivo_tema", ""),
            "Link": r.get("link", ""),
        })
        filas.append(fila)

    return pd.DataFrame(filas)


def _contenido_dialogo_nota(registro, con_sentimiento=False):
    usuario = _usuario_visible(registro) or "Usuario no identificado"
    st.caption(
        f"{registro.get('red', '')} · {usuario}"
    )

    if con_sentimiento:
        pred = str(
            registro.get("sentimiento_predicho", "") or ""
        ).replace("_", " ")
        conf = float(
            registro.get("confianza_sentimiento", 0) or 0
        )
        final = _sentimiento_para_descarga(registro).replace("_", " ")

        c1, c2 = st.columns(2)
        c1.metric("Clasificación para descarga", final)
        c2.metric("Confianza del modelo", f"{conf:.1%}")

        if str(registro.get("sentimiento", "")).upper() == REVISAR:
            st.info(
                "Esta publicación está marcada para revisión, pero si no "
                "la modificas se descargará con la predicción original "
                "del modelo."
            )

    st.markdown("##### Nota completa")
    st.write(str(registro.get("texto", "") or ""))

    motivo = str(registro.get("motivo_tema", "") or "").strip()
    if motivo:
        st.caption(f"Coincidencia temática: {motivo}")

    link = str(registro.get("link", "") or "").strip()
    if link:
        st.link_button("Abrir publicación original", link)


def _abrir_dialogo_nota(registro, con_sentimiento=False):
    _contenido_dialogo_nota(registro, con_sentimiento)


# st.dialog está disponible en las versiones recientes de Streamlit.
# Si no existe, la aplicación usa un expander como alternativa.
if hasattr(st, "dialog"):
    _abrir_dialogo_nota = st.dialog(
        "Vista completa de la publicación",
        width="large",
    )(_abrir_dialogo_nota)


def render_extractor_onclusive():
    st.subheader("Extracción de menciones Onclusive por tema")
    st.caption(
        "Versión 2.3 · filtro temático + sentimiento local + "
        "revisión manual editable"
    )
    st.caption(
        "Filtra por tema real, conserva impactos distintos aunque repitan texto, "
        "elimina RT/permalinks duplicados y permite corregir el sentimiento "
        "antes de generar entregables."
    )

    # -----------------------------------------------------------------
    # Entradas de extracción
    # -----------------------------------------------------------------
    tema = st.text_input(
        "Tema principal",
        placeholder=(
            "Ej. Gaby La Bonita Sánchez como aspirante "
            "por la alcaldía de Puebla"
        ),
        key="onclusive_tema",
        help=(
            "El programa extrae automáticamente las palabras útiles "
            "del tema y las usa para decidir qué publicaciones pertenecen a él."
        ),
    ).strip()

    archivo = st.file_uploader(
        "Sube el Excel o CSV exportado desde Onclusive",
        type=["xlsx", "xls", "csv"],
        key="onclusive_archivo",
    )

    if tema:
        detectados = terminos_desde_tema(tema)
        st.caption(
            "Palabras detectadas del tema: "
            + (
                ", ".join(detectados)
                if detectados
                else "sin términos suficientes"
            )
        )
    else:
        detectados = []

    with st.expander("Afinar tema y exclusiones", expanded=True):
        aliases_txt = st.text_area(
            "Actor, nombre y alias (opcional, recomendado cuando el tema es una persona)",
            placeholder=(
                "Gaby Sánchez, Gabriela Sánchez, La Bonita, "
                "Bonita Sánchez, Bonita_sanchez\n"
                "Si escribes alias aquí, la publicación deberá "
                "mencionar al menos uno."
            ),
            key="onclusive_aliases",
            height=95,
        )
        extras_txt = st.text_area(
            "Sinónimos o conceptos adicionales del tema (opcional)",
            placeholder=(
                "presidencia municipal, candidatura, proceso interno, "
                "levanta la mano, se destapa, contender, 2027"
            ),
            key="onclusive_terminos_extra",
            height=85,
        )
        exclusiones_txt = st.text_area(
            "Excluir siempre si aparece alguna de estas frases (opcional)",
            placeholder=(
                "Escribe solo exclusiones absolutas, "
                "una por línea o separadas por coma"
            ),
            key="onclusive_exclusiones",
            height=75,
            help=(
                "Estas exclusiones tienen prioridad incluso si la publicación "
                "también coincide con el tema."
            ),
        )
        cuentas_txt = st.text_area(
            "Cuentas/usuarios a excluir (opcional)",
            placeholder="@DeporteGobPue, @cuenta_propia",
            key="onclusive_cuentas_excluir",
            height=70,
        )

        max_slider = max(
            1,
            min(
                5,
                len(detectados)
                + len(parsear_lista_reglas(extras_txt))
            )
        )

        if max_slider <= 1:
            min_coincidencias = 1
            st.caption(
                "Coincidencias mínimas con el tema: 1. "
                "Al agregar más términos podrás ajustar este valor."
            )
        else:
            min_coincidencias = st.slider(
                "Coincidencias mínimas con el tema",
                min_value=1,
                max_value=max_slider,
                value=2,
                help=(
                    "Con 2, una publicación que solo diga 'Puebla' no entra. "
                    "Sube el valor para temas muy amplios."
                ),
                key="onclusive_min_coincidencias",
            )

    redes_seleccionadas = st.multiselect(
        "Redes a incluir",
        options=["X", "FACEBOOK", "INSTAGRAM", "TIKTOK"],
        default=["X", "FACEBOOK", "INSTAGRAM", "TIKTOK"],
        key="onclusive_redes",
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        dedup_texto = st.checkbox(
            "Eliminar mismo texto del mismo usuario",
            value=False,
            key="onclusive_dedup_texto",
        )
    with c2:
        recortar = st.checkbox(
            "Recortar textos largos",
            value=False,
            key="onclusive_recortar",
        )
    with c3:
        resolver = st.checkbox(
            "Resolver usuarios IG/TikTok",
            value=False,
            key="onclusive_resolver_usuarios",
        )

    # -----------------------------------------------------------------
    # Sentimiento opcional
    # -----------------------------------------------------------------
    st.markdown("### Análisis de sentimiento (opcional)")
    analizar_sentimiento = st.checkbox(
        "Realizar análisis de sentimiento local",
        value=False,
        key="onclusive_analizar_sentimiento",
        help=(
            "Solo se ejecuta si lo activas. Usa la base interna acumulada "
            "y no depende de una API de inteligencia artificial."
        ),
    )

    umbral_sentimiento = 0.70

    if analizar_sentimiento:
        if not SKLEARN_DISPONIBLE:
            st.error(
                "Falta scikit-learn. Agrega "
                "'scikit-learn>=1.4,<2' a requirements.txt."
            )

        estado_modelo = estado_base_interna()
        st.info(
            "Aprendizaje interno disponible: "
            f"{estado_modelo['total']} ejemplos · "
            f"positivas/informativas {estado_modelo['positivas']} · "
            f"negativas/críticas {estado_modelo['negativas']}."
        )

        umbral_sentimiento = st.slider(
            "Confianza mínima para aceptar la clasificación",
            min_value=0.50,
            max_value=0.95,
            value=0.70,
            step=0.05,
            key="onclusive_umbral_sentimiento",
            help=(
                "Debajo del umbral, la publicación queda como REVISAR."
            ),
        )

    # -----------------------------------------------------------------
    # Ejecutar extracción y GUARDAR resultado en session_state.
    # Esto permite que st.data_editor sobreviva a cada edición/rerun.
    # -----------------------------------------------------------------
    ejecutar = st.button(
        "Extraer menciones",
        type="primary",
        key="onclusive_extraer",
        disabled=not (archivo and tema),
    )

    if ejecutar:
        with st.spinner(
            "Identificando el tema, aplicando exclusiones y "
            "preparando resultados..."
        ):
            try:
                df = cargar_onclusive(archivo)
                aliases = parsear_lista_reglas(aliases_txt)
                extras = parsear_lista_reglas(extras_txt)
                exclusiones = parsear_lista_reglas(exclusiones_txt)
                cuentas_excluir = parsear_lista_reglas(cuentas_txt)

                registros, stats, excluidos = procesar_onclusive(
                    df,
                    tema,
                    resolver_remoto=resolver,
                    aliases=aliases,
                    terminos_extra=extras,
                    exclusiones=exclusiones,
                    cuentas_excluir=cuentas_excluir,
                    min_coincidencias=min_coincidencias,
                    deduplicar_texto=dedup_texto,
                    recortar_texto=recortar,
                    redes_permitidas=set(redes_seleccionadas),
                )

                if not registros:
                    st.warning(
                        "No quedaron menciones después del filtro temático."
                    )
                    st.session_state["onclusive_resultado"] = {
                        "registros": [],
                        "excluidos": excluidos,
                        "stats": dict(stats),
                        "tema": tema,
                        "redes": list(redes_seleccionadas),
                        "analizar_sentimiento": False,
                        "resumen_entrenamiento": {},
                    }
                else:
                    resumen_modelo = {}

                    if analizar_sentimiento:
                        if not SKLEARN_DISPONIBLE:
                            st.error(
                                "No se puede ejecutar sentimiento hasta "
                                "instalar scikit-learn."
                            )
                            return

                        with st.spinner(
                            "Entrenando con el aprendizaje interno "
                            "y clasificando..."
                        ):
                            modelo_sentimiento, df_entrenamiento = (
                                entrenar_desde_archivos(None)
                            )
                            registros = clasificar_registros(
                                registros,
                                modelo_sentimiento,
                                umbral=umbral_sentimiento,
                            )

                        resumen_modelo = resumen_entrenamiento(
                            df_entrenamiento
                        )

                        # Guardar estado inicial antes de cualquier corrección manual.
                        for r in registros:
                            r["sentimiento_inicial"] = r.get(
                                "sentimiento", ""
                            )

                    st.session_state["onclusive_resultado"] = {
                        "registros": registros,
                        "excluidos": excluidos,
                        "stats": dict(stats),
                        "tema": tema,
                        "redes": list(redes_seleccionadas),
                        "analizar_sentimiento": bool(
                            analizar_sentimiento
                        ),
                        "resumen_entrenamiento": resumen_modelo,
                    }

                    # Fuerza una nueva identidad del editor para la nueva extracción.
                    st.session_state.pop(
                        "onclusive_editor_sentimiento", None
                    )
                    st.session_state.pop(
                        "onclusive_editor_principal", None
                    )
                    st.session_state.pop(
                        "onclusive_nota_seleccionada", None
                    )

            except Exception as exc:
                st.error(f"Error al procesar el archivo: {exc}")
                return

    # -----------------------------------------------------------------
    # Resultado persistente: se mantiene aunque el editor provoque reruns.
    # -----------------------------------------------------------------
    resultado = st.session_state.get("onclusive_resultado")

    if not resultado:
        return

    registros = resultado.get("registros", [])
    excluidos = resultado.get("excluidos", [])
    stats = Counter(resultado.get("stats", {}))
    tema_resultado = resultado.get("tema", tema)
    redes_resultado = resultado.get("redes", [])
    con_sentimiento = bool(
        resultado.get("analizar_sentimiento", False)
    )

    if not registros:
        if excluidos:
            st.dataframe(
                pd.DataFrame(excluidos),
                use_container_width=True,
                hide_index=True,
            )
        return

    total_por_red = Counter(r["red"] for r in registros)
    st.success(
        f"Se extrajeron {len(registros)} menciones relacionadas con el tema."
    )
    st.write(
        " · ".join(
            f"{red.title() if red != 'X' else 'X'}: "
            f"{total_por_red.get(red, 0)}"
            for red in ["X", "FACEBOOK", "INSTAGRAM", "TIKTOK"]
            if red in redes_resultado
        )
    )

    st.caption(
        "Omitidas/depuradas: "
        f"fuera de tema {stats['fuera_tema']}, "
        f"cuentas excluidas {stats['cuenta_excluida']}, "
        f"otras redes {stats['otras_redes']}, RT {stats['rt']}, "
        f"duplicados por link {stats['dup_link']}, "
        f"duplicados por texto {stats['dup_texto']}, "
        f"sin texto {stats['sin_texto']}."
    )

    if con_sentimiento:
        resumen_modelo = resultado.get(
            "resumen_entrenamiento", {}
        )
        if resumen_modelo:
            st.caption(
                "Entrenamiento usado: "
                + " · ".join(
                    f"{k}: {v}"
                    for k, v in resumen_modelo.items()
                )
            )

    # -----------------------------------------------------------------
    # VISTA DE CONTROL / EDICIÓN
    # -----------------------------------------------------------------
    tab_in, tab_out = st.tabs([
        f"🧾 Vista previa ({len(registros)})",
        f"🚫 Excluidas ({len(excluidos)})",
    ])

    with tab_in:
        st.markdown("#### Mesa de revisión")
        st.caption(
            "Puedes completar o corregir el usuario directamente en la tabla. "
            "El texto se conserva completo y puede abrirse en una ventana de "
            "lectura. Si activaste sentimiento, también puedes corregirlo."
        )

        faltan_usuario = sum(
            1 for r in registros
            if not str(_usuario_visible(r)).strip()
        )

        pendientes = (
            sum(
                1 for r in registros
                if str(r.get("sentimiento", "")).upper() == REVISAR
            )
            if con_sentimiento
            else 0
        )

        # Las métricas de sentimiento representan lo que saldrá en la descarga.
        registros_previa_descarga = _registros_para_descarga(registros)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Publicaciones", len(registros))
        m2.metric("Usuarios por completar", faltan_usuario)

        if con_sentimiento:
            m3.metric(
                "Positivas / informativas",
                sum(
                    1 for r in registros_previa_descarga
                    if str(r.get("sentimiento", "")).upper()
                    == POSITIVA_INFORMATIVA
                ),
            )
            m4.metric(
                "Negativas / críticas",
                sum(
                    1 for r in registros_previa_descarga
                    if str(r.get("sentimiento", "")).upper()
                    == NEGATIVA_CRITICA
                ),
            )
            if pendientes:
                st.info(
                    f"{pendientes} publicación(es) están marcadas para revisar "
                    "por baja confianza. La revisión es opcional para descargar: "
                    "si no las modificas, se conservará la predicción del modelo."
                )
        else:
            m3.metric("Análisis de sentimiento", "No aplicado")
            m4.metric("Vista completa", "Disponible")

        vista_in = _vista_registros(
            registros,
            con_sentimiento=con_sentimiento,
        )

        columnas_bloqueadas = [
            "N.º",
            "Red",
            "Texto",
            "Motivo",
            "Link",
        ]

        config_columnas = {
            "N.º": st.column_config.NumberColumn(
                "N.º",
                width="small",
            ),
            "Red": st.column_config.TextColumn(
                "Red",
                width="small",
            ),
            "Usuario": st.column_config.TextColumn(
                "Usuario",
                width="medium",
                help=(
                    "Editable. Puedes escribir @usuario o "
                    "Nombre visible @usuario."
                ),
            ),
            "Texto": st.column_config.TextColumn(
                "Vista previa de la nota",
                width="large",
                help=(
                    "El texto completo se puede abrir debajo de la tabla "
                    "con el botón 'Ver nota completa'."
                ),
            ),
            "Motivo": st.column_config.TextColumn(
                "Coincidencia temática",
                width="medium",
            ),
            "Link": st.column_config.LinkColumn(
                "Publicación",
                display_text="Abrir",
                width="small",
            ),
        }

        if con_sentimiento:
            columnas_bloqueadas += [
                "Predicción del modelo",
                "Revisión",
                "Confianza",
            ]
            config_columnas.update({
                "Sentimiento corregido":
                    st.column_config.SelectboxColumn(
                        "Sentimiento",
                        options=[
                            POSITIVA_INFORMATIVA,
                            NEGATIVA_CRITICA,
                            REVISAR,
                        ],
                        required=True,
                        width="medium",
                        help=(
                            "REVISAR indica baja confianza. Si no cambias "
                            "esa fila, la descarga utilizará la predicción "
                            "original del modelo."
                        ),
                    ),
                "Predicción del modelo":
                    st.column_config.TextColumn(
                        "Predicción del modelo",
                        width="medium",
                    ),
                "Revisión":
                    st.column_config.TextColumn(
                        "Estado",
                        width="small",
                    ),
                "Confianza":
                    st.column_config.NumberColumn(
                        "Confianza",
                        format="%.1f%%",
                        width="small",
                    ),
            })

        editada = st.data_editor(
            vista_in,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            row_height=68,
            height=min(760, 115 + max(1, len(vista_in)) * 68),
            key="onclusive_editor_principal",
            disabled=columnas_bloqueadas,
            column_config=config_columnas,
        )

        # Aplicar inmediatamente los cambios del usuario y del sentimiento.
        if len(editada) == len(registros):
            for i, r in enumerate(registros):
                usuario_editado = editada.iloc[i]["Usuario"]
                _aplicar_usuario_editado(
                    r,
                    usuario_editado,
                )

                if con_sentimiento:
                    valor = str(
                        editada.iloc[i]["Sentimiento corregido"]
                        or REVISAR
                    ).upper()
                    if valor not in {
                        POSITIVA_INFORMATIVA,
                        NEGATIVA_CRITICA,
                        REVISAR,
                    }:
                        valor = REVISAR
                    r["sentimiento"] = valor

            resultado["registros"] = registros
            st.session_state["onclusive_resultado"] = resultado

        # -------------------------------------------------------------
        # Lector de nota completa: funciona con o sin sentimiento.
        # -------------------------------------------------------------
        st.markdown("##### Lectura completa")
        opciones = list(range(len(registros)))

        seleccion = st.selectbox(
            "Selecciona una publicación",
            options=opciones,
            key="onclusive_nota_seleccionada",
            format_func=lambda i: (
                f"{i + 1}. "
                f"{_usuario_visible(registros[i]) or 'Usuario no identificado'}"
                f" — "
                f"{str(registros[i].get('texto', '') or '')[:105]}"
                f"{'…' if len(str(registros[i].get('texto', '') or '')) > 105 else ''}"
            ),
        )

        if st.button(
            "🔎 Ver nota completa",
            key="onclusive_ver_nota_completa",
        ):
            if hasattr(st, "dialog"):
                _abrir_dialogo_nota(
                    registros[seleccion],
                    con_sentimiento,
                )
            else:
                st.session_state[
                    "onclusive_mostrar_nota_fallback"
                ] = seleccion

        if (
            not hasattr(st, "dialog")
            and st.session_state.get(
                "onclusive_mostrar_nota_fallback"
            ) is not None
        ):
            idx_fallback = st.session_state[
                "onclusive_mostrar_nota_fallback"
            ]
            with st.expander(
                "Vista completa de la publicación",
                expanded=True,
            ):
                _contenido_dialogo_nota(
                    registros[idx_fallback],
                    con_sentimiento,
                )

        # Resumen de cambios operativos.
        usuarios_completados = sum(
            1 for r in registros
            if str(_usuario_visible(r)).strip()
        )
        st.caption(
            f"Usuarios identificados/completados: "
            f"{usuarios_completados}/{len(registros)}. "
            "Los cambios de usuario y sentimiento se aplican "
            "automáticamente a los archivos de descarga."
        )

    with tab_out:
        if excluidos:
            vista_out = pd.DataFrame([
                {
                    "Red": r["red"],
                    "Usuario": (
                        f"@{r['handle']}"
                        if r.get("handle")
                        else r.get("autor")
                    ),
                    "Motivo de exclusión": r.get("motivo", ""),
                    "Texto": r.get("texto", ""),
                    "Link": r.get("link", ""),
                }
                for r in excluidos
            ])
            st.dataframe(
                vista_out,
                use_container_width=True,
                hide_index=True,
                row_height=58,
                column_config={
                    "Texto": st.column_config.TextColumn(
                        "Texto",
                        width="large",
                    ),
                    "Link": st.column_config.LinkColumn(
                        "Publicación",
                        display_text="Abrir",
                    ),
                },
            )
        else:
            st.info(
                "No hubo publicaciones excluidas por el filtro temático."
            )

    # -----------------------------------------------------------------
    # APRENDIZAJE INTERNO
    # -----------------------------------------------------------------
    if con_sentimiento:
        st.markdown("### Aprendizaje interno")

        resueltos = sum(
            1
            for r in registros
            if str(r.get("sentimiento", "")).upper()
            in {POSITIVA_INFORMATIVA, NEGATIVA_CRITICA}
        )
        pendientes_guardado = sum(
            1
            for r in registros
            if str(r.get("sentimiento", "")).upper()
            == REVISAR
        )

        st.caption(
            "Cuando termines la revisión, guarda las etiquetas corregidas. "
            "La app las agregará a su base interna y la próxima extracción "
            "entrenará con ese aprendizaje acumulado."
        )

        if st.button(
            "💾 Guardar correcciones en el aprendizaje interno",
            type="primary",
            key="onclusive_guardar_aprendizaje_interno",
            disabled=(resueltos == 0),
        ):
            try:
                resultado_guardado = guardar_aprendizaje_interno(
                    registros,
                    tema=tema_resultado,
                )

                # Reentrenar inmediatamente para dejar el modelo actualizado
                # dentro de la sesión actual.
                modelo_nuevo, df_nuevo = entrenar_desde_archivos(None)
                st.session_state[
                    "modelo_sentimiento_actualizado"
                ] = modelo_nuevo

                st.success(
                    "Aprendizaje actualizado: "
                    f"{resultado_guardado['guardados']} registros procesados. "
                    f"Base interna actual: {resultado_guardado['total']} ejemplos "
                    f"({resultado_guardado['positivas']} positivas/informativas · "
                    f"{resultado_guardado['negativas']} negativas/críticas)."
                )

                if pendientes_guardado:
                    st.info(
                        f"{pendientes_guardado} publicación(es) siguen en REVISAR "
                        "y no se guardaron como entrenamiento."
                    )

                # Persistencia automática opcional para Streamlit Cloud.
                # No usa una API de IA; solo guarda el CSV de entrenamiento
                # en el repositorio GitHub para sobrevivir reinicios/deploys.
                try:
                    gh_token = str(
                        st.secrets.get("GITHUB_TOKEN", "")
                    ).strip()
                    gh_repo = str(
                        st.secrets.get("GITHUB_REPO", "")
                    ).strip()
                    gh_branch = str(
                        st.secrets.get("GITHUB_BRANCH", "main")
                    ).strip() or "main"
                    gh_path = str(
                        st.secrets.get(
                            "GITHUB_TRAINING_PATH",
                            "entrenamiento_sentimiento.csv",
                        )
                    ).strip() or "entrenamiento_sentimiento.csv"
                except Exception:
                    gh_token = ""
                    gh_repo = ""
                    gh_branch = "main"
                    gh_path = "entrenamiento_sentimiento.csv"

                if gh_token and gh_repo:
                    with st.spinner(
                        "Guardando aprendizaje de forma persistente..."
                    ):
                        sincronizar_base_con_github(
                            token=gh_token,
                            repo=gh_repo,
                            branch=gh_branch,
                            ruta_repo=gh_path,
                            mensaje=(
                                "Actualizar aprendizaje de sentimiento "
                                f"desde Streamlit: {tema_resultado[:80]}"
                            ),
                        )
                    st.success(
                        "Base de aprendizaje guardada también en GitHub. "
                        "Se conservará después de reinicios y nuevos deploys."
                    )
                else:
                    st.warning(
                        "La base quedó actualizada dentro de la instancia actual. "
                        "En Streamlit Cloud ese archivo puede perderse al reiniciar "
                        "la app. Para conservarlo automáticamente entre reinicios, "
                        "configura GITHUB_TOKEN y GITHUB_REPO en Secrets."
                    )

            except Exception as exc:
                st.error(
                    f"No fue posible guardar el aprendizaje: {exc}"
                )

    # -----------------------------------------------------------------
    # Descargas
    # -----------------------------------------------------------------
    st.markdown("### Descargas")

    if con_sentimiento:
        excel_entrenamiento = crear_excel_entrenamiento(
            registros,
            tema_resultado,
        )
        st.download_button(
            "📥 Descargar Excel corregido para entrenamiento",
            data=excel_entrenamiento,
            file_name="entrenamiento_sentimiento_corregido.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="onclusive_excel_entrenamiento",
            help=(
                "Archivo de respaldo con las etiquetas revisadas. "
                "Las filas todavía marcadas REVISAR quedan sin etiqueta "
                "de entrenamiento."
            ),
        )

        st.download_button(
            "📥 Descargar CSV corregido para entrenamiento",
            data=csv_revision(registros, tema_resultado),
            file_name="entrenamiento_sentimiento_corregido.csv",
            mime="text/csv",
            key="onclusive_csv_entrenamiento",
        )

    faltan_handle = sum(
        1
        for r in registros
        if r["red"] in {"X", "INSTAGRAM", "TIKTOK"}
        and not r.get("handle")
    )
    if faltan_handle:
        st.info(
            f"{faltan_handle} publicación(es) no permitieron identificar "
            "un @usuario. Se conservó el nombre disponible."
        )

    base = re.sub(
        r"[^A-Za-z0-9ÁÉÍÓÚáéíóúÑñ_-]+",
        "_",
        tema_resultado,
    ).strip("_") or "tema"

    # IMPORTANTE:
    # - Se respetan los usuarios/sentimientos corregidos en la tabla.
    # - Si una fila sigue en REVISAR, la descarga usa la predicción
    #   original del modelo para que la revisión manual no sea obligatoria.
    registros_descarga = _registros_para_descarga(registros)
    word = crear_word(registros_descarga, tema_resultado)
    html_bytes = crear_html(registros_descarga)
    txt_bytes = crear_txt(registros_descarga)

    d1, d2, d3 = st.columns(3)
    with d1:
        st.download_button(
            "📥 Descargar Word",
            data=word,
            file_name=f"Extraccion_{base}.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
        )
    with d2:
        st.download_button(
            "📥 Descargar HTML",
            data=html_bytes,
            file_name=f"Extraccion_{base}.html",
            mime="text/html",
        )
    with d3:
        st.download_button(
            "📥 Descargar TXT",
            data=txt_bytes,
            file_name=f"Extraccion_{base}.txt",
            mime="text/plain",
        )

    if st.button(
        "🗑️ Limpiar resultado actual",
        key="onclusive_limpiar_resultado",
    ):
        st.session_state.pop("onclusive_resultado", None)
        st.session_state.pop(
            "onclusive_editor_sentimiento", None
        )
        st.rerun()
