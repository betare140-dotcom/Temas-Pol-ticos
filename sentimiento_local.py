
import io
import re
import unicodedata
from pathlib import Path

import pandas as pd

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
