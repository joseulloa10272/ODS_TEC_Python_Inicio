#!/usr/bin/env python3
"""Inventario de la extracción ODS TEC. Solo biblioteca estándar, Python 3.10+."""
import argparse
from collections import Counter, defaultdict
from datetime import date, datetime, timezone
import csv
from email import policy
from email.parser import BytesParser
import hashlib
import json
from pathlib import Path
import re
import sys
from urllib.parse import unquote, urlsplit

VERSION = '1.0.0'

def huellas_eml(ruta):
    with ruta.open('rb') as archivo:
        mensaje = BytesParser(policy=policy.default).parse(archivo)
    huellas = set()
    for parte in mensaje.walk():
        if parte.is_multipart():
            continue
        if parte.get_filename() or parte.get('Content-ID') or parte.get_content_disposition() in {'inline', 'attachment'}:
            datos = parte.get_payload(decode=True)
            if datos and not parte.defects:
                huellas.add(hashlib.sha256(datos).hexdigest())
    return huellas


def normalizar(nombre):
    return re.sub(r'[^a-z0-9]', '', nombre.lower())

def carpetas(raiz):
    dirs = {normalizar(p.name): p for p in raiz.iterdir() if p.is_dir() and not p.is_symlink()}
    if '01correos' not in dirs or '02adjuntos' not in dirs:
        raise ValueError('Selecciona la carpeta que contiene 01_Correos y 02_Adjuntos (también se admiten nombres sin guion).')
    return dirs['01correos'], dirs['02adjuntos'], dirs.get('03control')

def archivos(carpeta):
    if carpeta is None:
        return []
    # No seguir enlaces de directorios ni archivos, ni incluir metadatos del sistema.
    result = []
    for p in sorted(carpeta.rglob('*')):
        parts = p.relative_to(carpeta).parts
        if any(x.startswith('.') or x == '__MACOSX' for x in parts):
            continue
        if p.is_file() and not p.is_symlink():
            result.append(p)
    return result

def nombre_ruta(valor):
    if not isinstance(valor, str) or not valor:
        return ''
    valor = valor.replace('\\', '/')
    if valor.startswith(('https://', 'http://')):
        valor = urlsplit(valor).path
    return Path(unquote(valor)).name

def sha_archivo(path):
    sha = hashlib.sha256()
    with path.open('rb') as f:
        for block in iter(lambda: f.read(1024 * 1024), b''):
            sha.update(block)
    return sha.hexdigest()

def valor_csv(value):
    # Evitar que Excel interprete el contenido de un correo como una fórmula.
    if isinstance(value, str) and value.startswith(('=', '+', '-', '@', '\t', '\r')):
        return "'" + value
    return value

def escribir_csv(path, filas, campos):
    with path.open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=campos, delimiter=';', extrasaction='ignore')
        writer.writeheader()
        for row in filas:
            writer.writerow({k: valor_csv(row.get(k, '')) for k in campos})

def inventariar(raiz, salida, inicio='2026-06-10', fin='2026-11-16', huellas=True, exclusiones=None, progreso=True, verificar_eml=False):
    raiz = Path(raiz).expanduser().resolve()
    salida = Path(salida).expanduser().resolve()
    inicio_d, fin_d = date.fromisoformat(inicio), date.fromisoformat(fin)
    if inicio_d > fin_d:
        raise ValueError('La fecha inicial supera la final.')
    correos_dir, adjuntos_dir, control_dir = carpetas(raiz)
    for entrada in (correos_dir, adjuntos_dir, control_dir):
        if entrada and (salida == entrada or entrada in salida.parents):
            raise ValueError('La salida debe quedar fuera de las carpetas originales 01, 02 y 03.')
    if salida.exists() and any(salida.iterdir()):
        raise ValueError('La carpeta de salida ya contiene archivos; utiliza otra para conservar el informe anterior.')
    salida.mkdir(parents=True, exist_ok=True)
    excluidos = set(exclusiones or [])
    incidentes = []
    def incidencia(tipo, archivo, detalle, id_correo=''):
        incidentes.append({'tipo': tipo, 'archivo': str(archivo), 'id_correo': id_correo, 'detalle': detalle})

    controles = {}
    controles_conflictivos = set()
    for p in archivos(control_dir):
        if not p.name.lower().startswith('control_') or p.suffix.lower() != '.json':
            continue
        try:
            d = json.loads(p.read_text(encoding='utf-8-sig'))
            ids = d.get('idsCompletados')
            if d.get('schemaVersion') != '1.0' or not isinstance(ids, list) or any(not isinstance(x, str) for x in ids):
                raise ValueError('Control sin esquema 1.0 o lista válida de IDs.')
            date.fromisoformat(d['fecha'])
            key = (d['fecha'], d['buzon'], d['carpetaId'])
            if key in controles_conflictivos:
                raise ValueError('El control de esta fecha y origen tiene copias contradictorias.')
            if key in controles and controles[key] != set(ids):
                controles_conflictivos.add(key)
                controles.pop(key, None)
                raise ValueError('Dos controles de la misma fecha y origen contienen IDs distintos.')
            controles[key] = set(ids)
        except (ValueError, KeyError, TypeError, OSError) as e:
            incidencia('control_no_utilizable', p, str(e))

    fs_correos, fs_adjuntos = archivos(correos_dir), archivos(adjuntos_dir)
    emls = {p.resolve() for p in fs_correos if p.suffix.lower() == '.eml'}
    emls_por_nombre = defaultdict(list)
    for p in emls:
        emls_por_nombre[p.name].append(p)
    adjuntos_por_nombre = defaultdict(list)
    inventario_adjuntos = []
    sha_grupos = defaultdict(list)
    for i, p in enumerate(fs_adjuntos, 1):
        adjuntos_por_nombre[p.name].append(p)
        try:
            size = p.stat().st_size
            sha = sha_archivo(p) if huellas else ''
            if sha:
                sha_grupos[sha].append(str(p))
            inventario_adjuntos.append({'archivo': str(p), 'nombre': p.name, 'extension': p.suffix.lower(), 'bytes': size, 'sha256': sha})
        except OSError as e:
            incidencia('adjunto_no_legible', p, str(e))
            inventario_adjuntos.append({'archivo': str(p), 'nombre': p.name, 'extension': p.suffix.lower(), 'bytes': '', 'sha256': ''})
        if progreso and i % 250 == 0:
            print(f'Adjuntos revisados: {i}/{len(fs_adjuntos)}', flush=True)

    filas, grupos, todos_ids = [], defaultdict(list), set()
    adjuntos_con_huella = {r['archivo'] for r in inventario_adjuntos if r['sha256']}
    huellas_locales = {r['archivo']: r['sha256'] for r in inventario_adjuntos}
    emls_referidos, adjuntos_referidos = set(), set()
    diferencias_confirmadas = 0
    for p in (f for f in fs_correos if f.suffix.lower() == '.json'):
        row = {'archivo_json': str(p), 'archivo_eml': '', 'id_correo': '', 'internet_message_id': '',
               'clave_archivo': '', 'asunto': '', 'fecha_consulta': '', 'estado_json': '',
               'control_confirmado': False, 'adjuntos_manifestados': 0, 'adjuntos_localizados': 0,
               'estado_inventario': 'NoUtilizable', 'seleccionado': False, 'motivos': ''}
        motivos = []
        huellas_del_mensaje = None
        try:
            d = json.loads(p.read_text(encoding='utf-8-sig'))
            if not isinstance(d, dict) or not isinstance(d.get('correo'), dict):
                raise ValueError('No es un registro de correo con objeto correo.')
            correo = d['correo']
            ident = correo.get('id')
            if not isinstance(ident, str) or not ident:
                raise ValueError('Falta correo.id, no se sustituye por el GUID del archivo.')
            todos_ids.add(ident)
            clave = d.get('claveArchivo')
            row.update({'id_correo': ident, 'internet_message_id': correo.get('internetMessageId') or '',
                        'clave_archivo': clave or '', 'asunto': correo.get('subject') or '',
                        'fecha_consulta': d.get('fechaConsulta') or '', 'estado_json': d.get('estado') or ''})
            if d.get('estado') != 'Completo':
                motivos.append('estado_distinto_de_Completo')
            if ident in excluidos:
                motivos.append('exclusion_explicita')
            fecha = row['fecha_consulta']
            try:
                if not (inicio_d <= date.fromisoformat(fecha) <= fin_d):
                    motivos.append('fuera_del_periodo')
            except (ValueError, TypeError):
                motivos.append('fecha_consulta_invalida')
            control_key = (fecha, d.get('buzon'), d.get('carpetaId'))
            row['control_confirmado'] = ident in controles.get(control_key, set())
            if not row['control_confirmado']:
                motivos.append('id_no_confirmado_en_control_diario')
            # Primero la pareja con el mismo nombre local; después la ruta guardada.
            candidatos_eml = []
            pareja = p.with_suffix('.eml').resolve()
            if pareja in emls:
                candidatos_eml = [pareja]
            else:
                nombre = nombre_ruta(d.get('rutaEML')) or ((str(clave) + '.eml') if clave else '')
                candidatos_eml = emls_por_nombre.get(nombre, [])
            if len(candidatos_eml) != 1:
                motivos.append('eml_ausente_o_ambiguo')
            else:
                row['archivo_eml'] = str(candidatos_eml[0])
                emls_referidos.add(candidatos_eml[0])
                if candidatos_eml[0].stat().st_size == 0:
                    motivos.append('eml_vacio')
            manifiesto = d.get('adjuntos')
            if not isinstance(manifiesto, list):
                manifiesto = []
                motivos.append('manifiesto_adjuntos_invalido')
            row['adjuntos_manifestados'] = len(manifiesto)
            if correo.get('hasAttachments') is True and not manifiesto:
                motivos.append('correo_indica_adjuntos_sin_manifiesto')
            for adj in manifiesto:
                if not isinstance(adj, dict):
                    motivos.append('entrada_adjunto_invalida')
                    continue
                nombre = nombre_ruta(adj.get('rutaGuardada'))
                candidatos = adjuntos_por_nombre.get(nombre, []) if nombre else []
                if len(candidatos) != 1:
                    motivos.append('adjunto_ausente_o_ambiguo')
                    continue
                archivo = candidatos[0]
                adjuntos_referidos.add(archivo)
                row['adjuntos_localizados'] += 1
                if adj.get('estado') != 'Guardado':
                    motivos.append('adjunto_no_marcado_Guardado')
                expected = adj.get('size')
                real = archivo.stat().st_size
                if isinstance(expected, int) and expected >= 0 and real != expected:
                    if verificar_eml and row['archivo_eml']:
                        if huellas_del_mensaje is None:
                            huellas_del_mensaje = huellas_eml(Path(row['archivo_eml']))
                        huella = huellas_locales.get(str(archivo)) or sha_archivo(archivo)
                        if huella in huellas_del_mensaje:
                            diferencias_confirmadas += 1
                        else:
                            motivos.append('tamano_adjunto_no_coincide')
                    else:
                        motivos.append('tamano_adjunto_no_coincide')
                if real == 0:
                    motivos.append('adjunto_vacio')
                if huellas and str(archivo) not in adjuntos_con_huella:
                    motivos.append('adjunto_no_legible')
            row['_huella_contenido'] = hashlib.sha256(json.dumps(correo, ensure_ascii=False, sort_keys=True, separators=(',', ':')).encode()).hexdigest()
            row['_extraido'] = d.get('extraidoUTC') or ''
            row['estado_inventario'] = 'Candidato' if not motivos else 'PendienteOExcluido'
            grupos[ident].append(row)
        except (ValueError, KeyError, TypeError, OSError) as e:
            motivos.append('json_no_utilizable')
            incidencia('json_no_utilizable', p, str(e), row['id_correo'])
        row['motivos'] = ' | '.join(sorted(set(motivos)))
        filas.append(row)

    seleccionados, duplicados = [], []
    for ident, versiones in grupos.items():
        aptas = [r for r in versiones if r['estado_inventario'] == 'Candidato']
        if len({r['_huella_contenido'] for r in aptas}) > 1:
            for r in aptas:
                r['estado_inventario'] = 'RevisarVersiones'
                r['motivos'] = 'contenido_diferente_en_copias_completas_del_mismo_id'
        elif aptas:
            elegida = max(aptas, key=lambda r: (r['_extraido'], r['archivo_json']))
            elegida['seleccionado'] = True
            elegida['estado_inventario'] = 'CompletoVerificado'
            seleccionados.append(elegida)
            for r in aptas:
                if r is not elegida:
                    r['estado_inventario'] = 'CopiaDuplicada'
        if len(versiones) > 1:
            for r in versiones:
                duplicados.append({'id_correo': ident, 'archivo_json': r['archivo_json'], 'estado': r['estado_inventario'], 'seleccionado': r['seleccionado']})

    por_internet = defaultdict(set)
    for r in seleccionados:
        if r['internet_message_id']:
            por_internet[r['internet_message_id']].add(r['id_correo'])
    for mid, ids in por_internet.items():
        if len(ids) > 1:
            incidencia('internetMessageId_compartido', '', 'IDs de Outlook diferentes: ' + ' | '.join(sorted(ids)), mid)
    for p in sorted(emls - emls_referidos):
        incidencia('eml_sin_json_relacionado', p, 'No eliminar; puede corresponder a una extracción parcial.')
    for p in fs_adjuntos:
        if p not in adjuntos_referidos:
            incidencia('adjunto_sin_referencia_localizada', p, 'No eliminar; puede corresponder a una extracción parcial o una copia de descarga.')
    conteo_sha = Counter(r['sha256'] for r in inventario_adjuntos if r['sha256'])
    for r in inventario_adjuntos:
        r['referenciado'] = Path(r['archivo']) in adjuntos_referidos
        r['copias_mismo_contenido'] = conteo_sha.get(r['sha256'], '') if r['sha256'] else ''
    campos = ['archivo_json', 'archivo_eml', 'id_correo', 'internet_message_id', 'clave_archivo', 'asunto', 'fecha_consulta', 'estado_json', 'control_confirmado', 'adjuntos_manifestados', 'adjuntos_localizados', 'estado_inventario', 'seleccionado', 'motivos']
    escribir_csv(salida/'correos.csv', filas, campos)
    escribir_csv(salida/'correos_seleccionados.csv', seleccionados, campos)
    escribir_csv(salida/'adjuntos.csv', inventario_adjuntos, ['archivo', 'nombre', 'extension', 'bytes', 'sha256', 'referenciado', 'copias_mismo_contenido'])
    escribir_csv(salida/'duplicados_correos.csv', duplicados, ['id_correo', 'archivo_json', 'estado', 'seleccionado'])
    escribir_csv(salida/'incidencias.csv', incidentes, ['tipo', 'archivo', 'id_correo', 'detalle'])
    manifest = {'schemaVersion': VERSION, 'tipo': 'SeleccionParaAnalisis', 'correos': [
        {'idCorreo': r['id_correo'], 'json': r['archivo_json'], 'eml': r['archivo_eml'], 'claveArchivo': r['clave_archivo']} for r in seleccionados]}
    (salida/'seleccion_analisis.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    (salida/'duplicados_adjuntos.json').write_text(json.dumps({k: v for k, v in sha_grupos.items() if len(v) > 1}, ensure_ascii=False, indent=2), encoding='utf-8')
    resumen = {'schemaVersion': VERSION, 'generadoUTC': datetime.now(timezone.utc).isoformat(), 'raiz': str(raiz),
        'fechaInicio': inicio, 'fechaFin': fin, 'archivos01': len(fs_correos), 'json01': len(filas), 'eml01': len(emls),
        'archivos02': len(fs_adjuntos), 'bytes01': sum(p.stat().st_size for p in fs_correos),
        'bytes02': sum(r['bytes'] for r in inventario_adjuntos if isinstance(r['bytes'], int)),
        'correosUnicosPorId': len(todos_ids), 'correosSeleccionados': len(seleccionados),
        'idsSinVersionCompletaVerificada': len(todos_ids - {r['id_correo'] for r in seleccionados}),
        'copiasJsonAdicionalesPorId': sum(max(0, len(rs)-1) for rs in grupos.values()),
        'estadosInventario': dict(Counter(r['estado_inventario'] for r in filas)),
        'archivosConIncidencia': len(incidentes), 'tiposIncidencia': dict(Counter(r['tipo'] for r in incidentes)),
        'controlesDiariosUtilizables': len(controles), 'huellasAdjuntosCalculadas': huellas,
        'contenidosUnicosAdjuntos': len(sha_grupos) if huellas else None,
        'copiasBinariasAdicionales': sum(len(v)-1 for v in sha_grupos.values()) if huellas else None,
        'extensionesAdjuntos': dict(Counter(r['extension'] or '(sin extension)' for r in inventario_adjuntos)),
        'exclusionesExplicitasSolicitadas': len(excluidos),
        'diferenciasTamanoConfirmadasPorEML': diferencias_confirmadas,
        'alcance': 'Comprueba registros JSON, controles, parejas EML y archivos referenciados; con verificación EML confirma por SHA-256 los adjuntos con tamaño discrepante. El OCR y la clasificación son pasos posteriores.'}
    (salida/'resumen_inventario.json').write_text(json.dumps(resumen, ensure_ascii=False, indent=2), encoding='utf-8')
    return resumen

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--raiz', required=True, type=Path, help='Carpeta que contiene 01_Correos, 02_Adjuntos y 03_Control.')
    parser.add_argument('--salida', type=Path, help='Carpeta vacía para los informes; predeterminado: resultados_inventario/fechaUTC.')
    parser.add_argument('--inicio', default='2026-06-10')
    parser.add_argument('--fin', default='2026-11-16')
    parser.add_argument('--verificar-eml', action='store_true', help='Confirmar por SHA-256 los adjuntos cuyo tamaño difiere del JSON.')
    parser.add_argument('--sin-huellas', action='store_true', help='Omitir lectura binaria SHA-256 de adjuntos; no detecta duplicados de contenido.')
    parser.add_argument('--exclusiones', type=Path, help='JSON opcional con ids_correo: lista de IDs originales.')
    args = parser.parse_args()
    if sys.version_info < (3, 10):
        parser.error('Se requiere Python 3.10 o posterior.')
    salida = args.salida or Path.cwd()/'resultados_inventario'/datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ')
    try:
        ex = []
        if args.exclusiones:
            obj = json.loads(args.exclusiones.read_text(encoding='utf-8-sig'))
            ex = obj['ids_correo']
            if not isinstance(ex, list) or any(not isinstance(x, str) or not x for x in ex):
                raise ValueError('ids_correo debe ser una lista de identificadores de Outlook no vacíos.')
        r = inventariar(args.raiz, salida, args.inicio, args.fin, not args.sin_huellas, ex, verificar_eml=args.verificar_eml)
        print(json.dumps(r, ensure_ascii=False, indent=2))
        print('\nInformes guardados en: ' + str(salida.resolve()))
        return 0
    except (OSError, ValueError, KeyError, TypeError) as e:
        print('No se completó el inventario: ' + str(e), file=sys.stderr)
        return 2

if __name__ == '__main__':
    raise SystemExit(main())
