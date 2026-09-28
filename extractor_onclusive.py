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


def render_extractor_onclusive():
    st.subheader("Extracción de menciones Onclusive por tema")
    st.caption("Versión 2.2 · filtro temático + análisis de sentimiento local opcional")
    st.caption(
        "Filtra por tema real, conserva impactos distintos aunque repitan texto, elimina RT/permalinks duplicados "
        "y genera Word, HTML y TXT."
    )

    tema = st.text_input(
        "Tema principal",
        placeholder="Ej. Gaby La Bonita Sánchez como aspirante por la alcaldía de Puebla",
        key="onclusive_tema",
        help=(
            "El programa extrae automáticamente las palabras útiles del tema y las usa para decidir qué publicaciones pertenecen a él."
        ),
    ).strip()

    archivo = st.file_uploader(
        "Sube el Excel o CSV exportado desde Onclusive",
        type=["xlsx", "xls", "csv"],
        key="onclusive_archivo",
    )

    if tema:
        detectados = terminos_desde_tema(tema)
        st.caption("Palabras detectadas del tema: " + (", ".join(detectados) if detectados else "sin términos suficientes"))
    else:
        detectados = []

    with st.expander("Afinar tema y exclusiones", expanded=True):
        aliases_txt = st.text_area(
            "Actor, nombre y alias (opcional, recomendado cuando el tema es una persona)",
            placeholder=(
                "Gaby Sánchez, Gabriela Sánchez, La Bonita, Bonita Sánchez, Bonita_sanchez\n"
                "Si escribes alias aquí, la publicación deberá mencionar al menos uno."
            ),
            key="onclusive_aliases",
            height=95,
        )
        extras_txt = st.text_area(
            "Sinónimos o conceptos adicionales del tema (opcional)",
            placeholder="presidencia municipal, candidatura, proceso interno, levanta la mano, se destapa, contender, 2027",
            key="onclusive_terminos_extra",
            height=85,
        )
        exclusiones_txt = st.text_area(
            "Excluir siempre si aparece alguna de estas frases (opcional)",
            placeholder="Escribe solo exclusiones absolutas, una por línea o separadas por coma",
            key="onclusive_exclusiones",
            height=75,
            help=(
                "Estas exclusiones tienen prioridad incluso si la publicación también coincide con el tema. "
                "Evita poner términos muy generales como 'deporte' si una nota política también puede mencionarlos."
            ),
        )
        cuentas_txt = st.text_area(
            "Cuentas/usuarios a excluir (opcional)",
            placeholder="@DeporteGobPue, @cuenta_propia",
            key="onclusive_cuentas_excluir",
            height=70,
        )

        max_slider = max(1, min(5, len(detectados) + len(parsear_lista_reglas(extras_txt))))

        # Streamlit no permite un slider cuando min_value == max_value.
        # Esto ocurre, por ejemplo, cuando todavía no se ha escrito el tema
        # o cuando solo se detecta una palabra/concepto útil.
        if max_slider <= 1:
            min_coincidencias = 1
            st.caption(
                "Coincidencias mínimas con el tema: 1. "
                "Al agregar más términos al tema podrás ajustar este valor."
            )
        else:
            min_coincidencias = st.slider(
                "Coincidencias mínimas con el tema",
                min_value=1,
                max_value=max_slider,
                value=2,
                help=(
                    "Con 2, una publicación que solo diga 'Puebla' no entra. "
                    "Sube el valor para temas muy amplios; bájalo si el tema "
                    "tiene pocas palabras específicas."
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
            help=(
                "Déjalo desactivado para trabajar como la extracción revisada: dos posts distintos cuentan como dos impactos aunque tengan el mismo texto."
            ),
            key="onclusive_dedup_texto",
        )
    with c2:
        recortar = st.checkbox(
            "Recortar textos largos",
            value=False,
            help="Desactivado conserva el texto completo limpio, como en la extracción revisada.",
            key="onclusive_recortar",
        )
    with c3:
        resolver = st.checkbox(
            "Resolver usuarios IG/TikTok",
            value=False,
            help=(
                "Intento opcional para publicaciones públicas. No usa login ni intenta evadir restricciones."
            ),
            key="onclusive_resolver_usuarios",
        )


    st.markdown("### Análisis de sentimiento (opcional)")
    analizar_sentimiento = st.checkbox(
        "Realizar análisis de sentimiento local",
        value=False,
        key="onclusive_analizar_sentimiento",
        help=(
            "Solo se ejecuta si lo activas. "
            "No utiliza API: entrena un modelo local con "
            "entrenamiento_sentimiento.csv."
        ),
    )

    umbral_sentimiento = 0.70
    dataset_sentimiento_extra = None

    if analizar_sentimiento:
        if not SKLEARN_DISPONIBLE:
            st.error(
                "Falta scikit-learn. Agrega "
                "'scikit-learn>=1.4,<2' a requirements.txt."
            )

        umbral_sentimiento = st.slider(
            "Confianza mínima para aceptar la clasificación",
            min_value=0.50,
            max_value=0.95,
            value=0.70,
            step=0.05,
            key="onclusive_umbral_sentimiento",
            help=(
                "Si la confianza queda por debajo del umbral, "
                "la publicación se marca como REVISAR."
            ),
        )

        dataset_sentimiento_extra = st.file_uploader(
            "Dataset de sentimiento corregido/adicional (opcional)",
            type=["csv", "xlsx", "xls"],
            key="onclusive_dataset_sentimiento_extra",
            help=(
                "Debe tener como mínimo las columnas texto y etiqueta. "
                "Se mezcla con la base inicial de La Bonita."
            ),
        )

    if archivo and tema and st.button("Extraer menciones", type="primary", key="onclusive_extraer"):
        with st.spinner("Identificando el tema, aplicando exclusiones y preparando entregables..."):
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
                        "No quedaron menciones después del filtro temático. Revisa aliases, exclusiones o baja las coincidencias mínimas."
                    )
                    if excluidos:
                        st.dataframe(pd.DataFrame(excluidos), use_container_width=True, hide_index=True)
                    return

                if analizar_sentimiento:
                    if not SKLEARN_DISPONIBLE:
                        st.error(
                            "No se puede ejecutar sentimiento hasta "
                            "instalar scikit-learn."
                        )
                        return

                    with st.spinner(
                        "Entrenando el modelo local y clasificando..."
                    ):
                        modelo_sentimiento, df_entrenamiento = (
                            entrenar_desde_archivos(
                                dataset_sentimiento_extra
                            )
                        )
                        registros = clasificar_registros(
                            registros,
                            modelo_sentimiento,
                            umbral=umbral_sentimiento,
                        )

                    resumen_modelo = resumen_entrenamiento(
                        df_entrenamiento
                    )
                    st.caption(
                        "Entrenamiento usado: "
                        + " · ".join(
                            f"{k}: {v}"
                            for k, v in resumen_modelo.items()
                        )
                    )

                    conteo_sent = Counter(
                        r.get("sentimiento", "")
                        for r in registros
                    )
                    st.write(
                        "Clasificación: "
                        + " · ".join(
                            f"{_titulo_sentimiento(k)}: {v}"
                            for k, v in conteo_sent.items()
                        )
                    )

                    st.info(
                        "La base inicial contiene 99 publicaciones "
                        "del caso de Gaby 'La Bonita' Sánchez. "
                        "Es una semilla: conviene revisar los casos "
                        "marcados como REVISAR y agregar ejemplos "
                        "de otros actores y temas."
                    )

                total_por_red = Counter(r["red"] for r in registros)
                st.success(f"Se extrajeron {len(registros)} menciones relacionadas con el tema.")
                st.write(
                    " · ".join(
                        f"{red.title() if red != 'X' else 'X'}: {total_por_red.get(red, 0)}"
                        for red in ["X", "FACEBOOK", "INSTAGRAM", "TIKTOK"]
                        if red in redes_seleccionadas
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

                # Vista de control: permite comprobar qué entró y qué quedó fuera antes de descargar.
                tab_in, tab_out = st.tabs([
                    f"Incluidas ({len(registros)})",
                    f"Excluidas por tema/reglas ({len(excluidos)})",
                ])
                with tab_in:
                    vista_in = pd.DataFrame([
                        {
                            "Red": r["red"],
                            "Usuario": (f"@{r['handle']}" if r["handle"] else r["autor"]),
                            "Texto": r["texto"],
                            "Sentimiento": r.get("sentimiento", ""),
                            "Confianza": (
                                f"{r.get('confianza_sentimiento', 0):.1%}"
                                if analizar_sentimiento
                                else ""
                            ),
                            "Motivo": r.get("motivo_tema", ""),
                            "Link": r["link"],
                        }
                        for r in registros
                    ])
                    st.dataframe(vista_in, use_container_width=True, hide_index=True)
                with tab_out:
                    if excluidos:
                        vista_out = pd.DataFrame([
                            {
                                "Red": r["red"],
                                "Usuario": (f"@{r['handle']}" if r.get("handle") else r.get("autor")),
                                "Motivo de exclusión": r.get("motivo", ""),
                                "Texto": r.get("texto", ""),
                                "Link": r.get("link", ""),
                            }
                            for r in excluidos
                        ])
                        st.dataframe(vista_out, use_container_width=True, hide_index=True)
                    else:
                        st.info("No hubo publicaciones excluidas por el filtro temático.")

                faltan_handle = sum(
                    1 for r in registros if r["red"] in {"X", "INSTAGRAM", "TIKTOK"} and not r["handle"]
                )
                if faltan_handle:
                    st.info(
                        f"{faltan_handle} publicación(es) no permitieron identificar un @usuario con los datos disponibles. "
                        "Se conservó el nombre del autor/página cuando estaba disponible."
                    )

                if analizar_sentimiento:
                    st.download_button(
                        "📥 Descargar CSV para corregir y seguir entrenando",
                        data=csv_revision(registros, tema),
                        file_name="revision_sentimiento.csv",
                        mime="text/csv",
                        key="onclusive_revision_sentimiento",
                        help=(
                            "Corrige la columna etiqueta y vuelve a cargar "
                            "ese archivo como dataset adicional en una "
                            "ejecución futura."
                        ),
                    )

                base = re.sub(r"[^A-Za-z0-9ÁÉÍÓÚáéíóúÑñ_-]+", "_", tema).strip("_") or "tema"
                word = crear_word(registros, tema)
                html_bytes = crear_html(registros)
                txt_bytes = crear_txt(registros)

                st.download_button(
                    "📥 Descargar Word",
                    data=word,
                    file_name=f"Extraccion_{base}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
                st.download_button(
                    "📥 Descargar HTML",
                    data=html_bytes,
                    file_name=f"Extraccion_{base}.html",
                    mime="text/html",
                )
                st.download_button(
                    "📥 Descargar TXT",
                    data=txt_bytes,
                    file_name=f"Extraccion_{base}.txt",
                    mime="text/plain",
                )
            except Exception as exc:
                st.error(f"Error al procesar el archivo: {exc}")


