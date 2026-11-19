"""Consulta local del directorio institucional para verificar responsables."""
from collections import defaultdict
from email.utils import parseaddr
import re
import unicodedata


def clave(valor):
    texto = unicodedata.normalize("NFKD", str(valor or "")).casefold()
    return " ".join(re.findall(r"\w+", "".join(c for c in texto if not unicodedata.combining(c))))


def cargar_directorio(ruta):
    from openpyxl import load_workbook

    libro = load_workbook(ruta, read_only=True, data_only=True)
    personas = defaultdict(set)
    unidades = defaultdict(set)
    try:
        for nombre, correo, unidad, *_ in list(libro["Funcionarios"].values)[1:]:
            if nombre:
                ficha = (str(nombre), parseaddr(str(correo or ""))[1].lower(), str(unidad or ""))
                for indice in (clave(nombre), ficha[1]):
                    if indice:
                        personas[indice].add(ficha)
            if unidad:
                unidades[clave(unidad)].add((str(unidad), ""))
        for fila in list(libro["Directorio de dependencias"].values)[1:]:
            departamento, programa = fila[:2]
            nombre = programa or departamento
            if nombre:
                unidades[clave(nombre)].add((str(nombre), str(fila[9] or "").strip()))
        for departamento, _, correo, *_ in list(libro["Correos oficiales"].values)[1:]:
            if departamento:
                unidades[clave(departamento)].add((str(departamento), str(correo or "").strip()))
    finally:
        libro.close()
    return personas, unidades


def resolver(actividad, directorio):
    personas, unidades = directorio
    responsable = actividad["responsable"].strip()
    correo = actividad["correo_responsable"].strip().lower()
    por_nombre = personas.get(clave(responsable), set())
    por_correo = personas.get(correo, set())
    candidatos = (por_nombre & por_correo if responsable and correo else por_nombre or por_correo)
    dependencia = actividad["dependencia"]
    if dependencia:
        candidatos = {f for f in candidatos if not f[2] or clave(f[2]) == clave(dependencia)}
    actividad["verificacion_responsable"] = "sin_evidencia" if not (responsable or correo) else "requiere_revision"
    if len(candidatos) == 1:
        nombre, email, unidad = next(iter(candidatos))
        actividad.update(responsable=nombre, correo_responsable=email or correo,
                         verificacion_responsable="coincidencia_directorio")
        actividad["dependencia_directorio"] = unidad
    fichas = unidades.get(clave(dependencia), set())
    correos = {f[1] for f in fichas if f[1]}
    actividad["correo_dependencia"] = next(iter(correos)) if len(correos) == 1 else ""
    actividad["verificacion_dependencia"] = "coincidencia_directorio" if fichas and len(correos) <= 1 else "requiere_revision"
