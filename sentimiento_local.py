
import io
import re
import unicodedata
from pathlib import Path
import base64
import json


import pandas as pd

try:
    import requests
except Exception:
    requests = None

try:
    from sklearn.pipeline import Pipeline, FeatureUnion
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.linear_model import LogisticRegression
    SKLEARN_DISPONIBLE = True
except Exception:
    Pipeline = FeatureUnion = TfidfVectorizer = LogisticRegression = None
    SKLEARN_DISPONIBLE = False


POSITIVA_INFORMATIVA = "POSITIVA_INFORMATIVA"
NEGATIVA_CRITICA = "NEGATIVA_CRITICA"
REVISAR = "REVISAR"


def _sin_acentos(texto):
    texto = str(texto or "")
    return "".join(
        c
        for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    ).lower()


def _normalizar_texto(texto):
    t = _sin_acentos(texto)
    t = re.sub(r"https?://\S+", " ", t)
    t = re.sub(r"\s+", " ", t)
    return t.strip()


def ruta_base_entrenamiento():
    return Path(__file__).resolve().parent / "entrenamiento_sentimiento.csv"


def cargar_entrenamiento(archivo_extra=None):
    """
    Carga la base incluida y, si el usuario proporciona otra base,
    la mezcla con la inicial.

    Columnas obligatorias:
    - texto
    - etiqueta
    """
    frames = []

    ruta = ruta_base_entrenamiento()
    if ruta.exists():
        try:
            frames.append(pd.read_csv(ruta, encoding="utf-8-sig"))
        except Exception:
            frames.append(pd.read_csv(ruta))

    if archivo_extra is not None:
        raw = archivo_extra.read()
        archivo_extra.seek(0)
        nombre = (getattr(archivo_extra, "name", "") or "").lower()

        if nombre.endswith((".xlsx", ".xls")):
            extra = pd.read_excel(io.BytesIO(raw))
        else:
            try:
                extra = pd.read_csv(io.BytesIO(raw), encoding="utf-8-sig")
            except Exception:
                extra = pd.read_csv(io.BytesIO(raw), encoding="latin1")
        frames.append(extra)

    if not frames:
        raise FileNotFoundError(
            "No se encontró entrenamiento_sentimiento.csv junto al programa."
        )

    df = pd.concat(frames, ignore_index=True)

    # Tolera mayúsculas/minúsculas en encabezados.
    columnas = {_sin_acentos(c).strip(): c for c in df.columns}
    if "texto" not in columnas or "etiqueta" not in columnas:
        raise ValueError(
            "El dataset de sentimiento debe tener columnas 'texto' y 'etiqueta'."
        )

    df = df.rename(
        columns={
            columnas["texto"]: "texto",
            columnas["etiqueta"]: "etiqueta",
        }
    )

    df["texto"] = df["texto"].fillna("").astype(str).str.strip()
    df["etiqueta"] = (
        df["etiqueta"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # REVISAR no es una etiqueta de entrenamiento.
    df = df[
        (df["texto"] != "")
        & (df["etiqueta"] != "")
        & (df["etiqueta"] != REVISAR)
    ].copy()

    # Si un texto fue corregido después, conserva la última etiqueta.
    df["_key"] = df["texto"].map(_normalizar_texto)
    df = (
        df.drop_duplicates(subset=["_key"], keep="last")
        .drop(columns=["_key"])
        .reset_index(drop=True)
    )

    if df["etiqueta"].nunique() < 2:
        raise ValueError(
            "El dataset necesita al menos dos etiquetas distintas."
        )

    return df



def guardar_aprendizaje_interno(registros, tema="", ruta=None):
    """
    Guarda las correcciones resueltas directamente dentro de
    entrenamiento_sentimiento.csv.

    - Ignora REVISAR.
    - Si el mismo texto ya existe, conserva la etiqueta más reciente.
    - Devuelve estadísticas del guardado.
    """
    ruta = Path(ruta or ruta_base_entrenamiento())

    filas = []
    for r in registros:
        etiqueta = str(r.get("sentimiento", "") or "").strip().upper()
        texto = str(r.get("texto", "") or "").strip()

        if not texto or etiqueta in {"", REVISAR}:
            continue

        fuente = r.get("autor", "") or (
            f"@{r.get('handle')}" if r.get("handle") else ""
        )

        filas.append({
            "actor": r.get("actor", ""),
            "tema": tema,
            "red": r.get("red", ""),
            "fuente": fuente,
            "usuario": (
                f"@{r.get('handle')}" if r.get("handle") else ""
            ),
            "texto": texto,
            "enlace": r.get("link", ""),
            "etiqueta": etiqueta,
            "prediccion_original": r.get(
                "sentimiento_predicho", ""
            ),
            "confianza": round(
                float(r.get("confianza_sentimiento", 0) or 0),
                4,
            ),
            "origen": "revision_manual_guardada_en_app",
        })

    nuevos = pd.DataFrame(filas)
    if nuevos.empty:
        return {
            "guardados": 0,
            "total": 0,
            "positivas": 0,
            "negativas": 0,
            "ruta": str(ruta),
        }

    if ruta.exists():
        try:
            base = pd.read_csv(ruta, encoding="utf-8-sig")
        except Exception:
            base = pd.read_csv(ruta)
    else:
        base = pd.DataFrame()

    combinada = pd.concat([base, nuevos], ignore_index=True, sort=False)

    if "texto" not in combinada.columns:
        raise ValueError(
            "La base interna no contiene la columna 'texto'."
        )
    if "etiqueta" not in combinada.columns:
        raise ValueError(
            "La base interna no contiene la columna 'etiqueta'."
        )

    combinada["texto"] = combinada["texto"].fillna("").astype(str)
    combinada["etiqueta"] = (
        combinada["etiqueta"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    combinada["_key"] = combinada["texto"].map(_normalizar_texto)
    combinada = (
        combinada[
            (combinada["_key"] != "")
            & (combinada["etiqueta"] != "")
            & (combinada["etiqueta"] != REVISAR)
        ]
        .drop_duplicates(subset=["_key"], keep="last")
        .drop(columns=["_key"])
        .reset_index(drop=True)
    )

    ruta.parent.mkdir(parents=True, exist_ok=True)
    combinada.to_csv(
        ruta,
        index=False,
        encoding="utf-8-sig",
    )

    conteos = combinada["etiqueta"].value_counts().to_dict()
    return {
        "guardados": len(nuevos),
        "total": len(combinada),
        "positivas": int(
            conteos.get(POSITIVA_INFORMATIVA, 0)
        ),
        "negativas": int(
            conteos.get(NEGATIVA_CRITICA, 0)
        ),
        "ruta": str(ruta),
    }


def bytes_base_entrenamiento(ruta=None):
    ruta = Path(ruta or ruta_base_entrenamiento())
    if not ruta.exists():
        return b""
    return ruta.read_bytes()


def sincronizar_base_con_github(
    token,
    repo,
    branch="main",
    ruta_repo="entrenamiento_sentimiento.csv",
    mensaje="Actualizar aprendizaje de sentimiento",
    contenido=None,
):
    """
    Persistencia opcional para Streamlit Cloud.

    El modelo sigue siendo local. GitHub se usa únicamente para
    conservar entrenamiento_sentimiento.csv entre reinicios/deploys.
    """
    if not token or not repo:
        raise ValueError("Faltan token o repositorio de GitHub.")
    if requests is None:
        raise RuntimeError("requests no está disponible.")

    contenido = contenido if contenido is not None else bytes_base_entrenamiento()
    if not contenido:
        raise ValueError("La base interna está vacía.")

    url = f"https://api.github.com/repos/{repo}/contents/{ruta_repo}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }

    sha = None
    get_resp = requests.get(
        url,
        headers=headers,
        params={"ref": branch},
        timeout=20,
    )
    if get_resp.status_code == 200:
        sha = get_resp.json().get("sha")
    elif get_resp.status_code != 404:
        raise RuntimeError(
            f"GitHub GET {get_resp.status_code}: "
            f"{get_resp.text[:300]}"
        )

    payload = {
        "message": mensaje,
        "content": base64.b64encode(contenido).decode("ascii"),
        "branch": branch,
    }
    if sha:
        payload["sha"] = sha

    put_resp = requests.put(
        url,
        headers=headers,
        json=payload,
        timeout=30,
    )
    if put_resp.status_code not in {200, 201}:
        raise RuntimeError(
            f"GitHub PUT {put_resp.status_code}: "
            f"{put_resp.text[:500]}"
        )

    data = put_resp.json()
    return {
        "ok": True,
        "commit": data.get("commit", {}).get("sha", ""),
        "ruta_repo": ruta_repo,
        "repo": repo,
        "branch": branch,
    }


def estado_base_interna(ruta=None):
    ruta = Path(ruta or ruta_base_entrenamiento())
    if not ruta.exists():
        return {
            "total": 0,
            "positivas": 0,
            "negativas": 0,
        }

    try:
        df = pd.read_csv(ruta, encoding="utf-8-sig")
    except Exception:
        df = pd.read_csv(ruta)

    if "etiqueta" not in df.columns:
        return {
            "total": len(df),
            "positivas": 0,
            "negativas": 0,
        }

    etiquetas = (
        df["etiqueta"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )
    return {
        "total": len(df),
        "positivas": int(
            (etiquetas == POSITIVA_INFORMATIVA).sum()
        ),
        "negativas": int(
            (etiquetas == NEGATIVA_CRITICA).sum()
        ),
    }


def entrenar_modelo(df):
    """
    Modelo local:
    TF-IDF de palabras + TF-IDF de caracteres + Regresión Logística.

    No utiliza ninguna API ni conexión externa.
    """
    if not SKLEARN_DISPONIBLE:
        raise RuntimeError(
            "Falta scikit-learn. Agrega scikit-learn>=1.4,<2 "
            "a requirements.txt."
        )

    features = FeatureUnion([
        (
            "palabras",
            TfidfVectorizer(
                lowercase=True,
                strip_accents="unicode",
                ngram_range=(1, 2),
                min_df=1,
                max_features=30000,
                sublinear_tf=True,
            ),
        ),
        (
            "caracteres",
            TfidfVectorizer(
                lowercase=True,
                strip_accents="unicode",
                analyzer="char_wb",
                ngram_range=(3, 5),
                min_df=1,
                max_features=30000,
                sublinear_tf=True,
            ),
        ),
    ])

    modelo = Pipeline([
        ("features", features),
        (
            "clasificador",
            LogisticRegression(
                max_iter=2500,
                class_weight="balanced",
                solver="liblinear",
                random_state=42,
            ),
        ),
    ])

    modelo.fit(
        df["texto"].astype(str),
        df["etiqueta"].astype(str),
    )
    return modelo


def entrenar_desde_archivos(archivo_extra=None):
    df = cargar_entrenamiento(archivo_extra)
    modelo = entrenar_modelo(df)
    return modelo, df


def clasificar_registros(registros, modelo, umbral=0.70):
    """
    Añade:
    - sentimiento_predicho
    - sentimiento
    - confianza_sentimiento

    Si la confianza es menor al umbral, usa REVISAR.
    """
    if not registros:
        return registros

    textos = [r.get("texto", "") for r in registros]
    proba = modelo.predict_proba(textos)
    clases = list(modelo.classes_)

    salida = []
    for item, probs in zip(registros, proba):
        mejor = max(range(len(probs)), key=lambda i: probs[i])
        pred = str(clases[mejor])
        confianza = float(probs[mejor])

        nuevo = dict(item)
        nuevo["sentimiento_predicho"] = pred
        nuevo["confianza_sentimiento"] = confianza
        nuevo["sentimiento"] = pred if confianza >= umbral else REVISAR
        salida.append(nuevo)

    return salida


def resumen_entrenamiento(df):
    return df["etiqueta"].value_counts().to_dict()


def csv_revision(registros, tema=""):
    """
    Crea un CSV que el usuario puede corregir y volver a cargar después
    como entrenamiento adicional.
    """
    filas = []
    for r in registros:
        etiqueta = r.get("sentimiento", "")
        filas.append({
            "actor": "",
            "tema": tema,
            "fuente": r.get("autor", "") or r.get("handle", ""),
            "texto": r.get("texto", ""),
            "enlace": r.get("link", ""),
            "etiqueta": "" if etiqueta == REVISAR else etiqueta,
            "prediccion_original": r.get("sentimiento_predicho", ""),
            "confianza": round(
                float(r.get("confianza_sentimiento", 0)), 4
            ),
            "origen": "revision_usuario",
        })

    return (
        pd.DataFrame(filas)
        .to_csv(index=False)
        .encode("utf-8-sig")
    )
