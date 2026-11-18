import copy
import json
from pathlib import Path
import tempfile
import unittest

from inventario_ods import inventariar

class InventarioTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.base = Path(self.tmp.name)
        self.root = self.base/'entrada'
        for x in ['01_Correos', '02_Adjuntos', '03_Control']:
            (self.root/x).mkdir(parents=True)
        self.ids = []

    def tearDown(self):
        self.tmp.cleanup()

    def guardar(self, key='a', ident='id-a', estado='Completo', adjuntos=None, cuerpo='Texto', fecha='2026-06-10'):
        d = {'schemaVersion':'1.0','claveArchivo':key,'extraidoUTC':'2026-11-16T18:00:00Z',
             'buzon':'ods-tec@itcr.ac.cr','carpetaId':'carpeta-test','fechaConsulta':fecha,
             'estado':estado,'correo':{'id':ident,'internetMessageId':'<'+ident+'>','subject':'Taller ñ',
                                      'body':cuerpo,'hasAttachments':bool(adjuntos)},'adjuntos':adjuntos or [],
             'rutaEML':'/ODS/01_Correos/'+key+'.eml'}
        p=self.root/'01_Correos'/f'{key}.json'
        p.write_text(json.dumps(d),encoding='utf-8')
        p.with_suffix('.eml').write_bytes(b'Subject: prueba\r\n\r\nMensaje')
        self.ids.append(ident)
        self.control(fecha=fecha)
        return p,d

    def control(self,ids=None,fecha='2026-06-10',nombre=None):
        d={'schemaVersion':'1.0','fecha':fecha,'buzon':'ods-tec@itcr.ac.cr','carpetaId':'carpeta-test','idsCompletados':ids if ids is not None else self.ids}
        name=nombre or 'control_'+fecha.replace('-','')+'.json'
        (self.root/'03_Control'/name).write_text(json.dumps(d))

    def ejecutar(self,**kw):
        return inventariar(self.root,self.base/'salida',progreso=False,**kw)

    def test_pareja_completa_control_y_sin_cambios(self):
        self.guardar()
        antes={str(p):p.read_bytes() for p in self.root.rglob('*') if p.is_file()}
        r=self.ejecutar()
        despues={str(p):p.read_bytes() for p in self.root.rglob('*') if p.is_file()}
        self.assertEqual(r['correosSeleccionados'],1)
        self.assertEqual(antes,despues)

    def test_reintentos_y_seleccion_una_version(self):
        self.guardar('a','id-a',estado='Pendiente')
        self.guardar('b','id-a')
        self.guardar('c','id-a')
        r=self.ejecutar()
        self.assertEqual(r['correosUnicosPorId'],1)
        self.assertEqual(r['correosSeleccionados'],1)
        self.assertEqual(r['copiasJsonAdicionalesPorId'],2)
        self.assertEqual(r['estadosInventario']['CopiaDuplicada'],1)

    def test_pendiente_no_se_analiza(self):
        self.guardar(estado='Pendiente')
        self.assertEqual(self.ejecutar()['correosSeleccionados'],0)

    def test_no_confunde_guid_con_id(self):
        p,d=self.guardar();d['correo'].pop('id');p.write_text(json.dumps(d))
        r=self.ejecutar()
        self.assertEqual(r['correosUnicosPorId'],0)
        self.assertEqual(r['correosSeleccionados'],0)

    def test_adjunto_presente_con_bytes_verificados(self):
        (self.root/'02_Adjuntos'/'a_0001.pdf').write_bytes(b'123')
        self.guardar(adjuntos=[{'estado':'Guardado','rutaGuardada':'/ODS/02_Adjuntos/a_0001.pdf','size':3}])
        self.assertEqual(self.ejecutar()['correosSeleccionados'],1)

    def test_adjunto_faltante_o_tamano_incorrecto(self):
        (self.root/'02_Adjuntos'/'a_0001.pdf').write_bytes(b'123')
        self.guardar(adjuntos=[{'estado':'Guardado','rutaGuardada':'/ODS/02_Adjuntos/a_0001.pdf','size':8},
                               {'estado':'Guardado','rutaGuardada':'/ODS/02_Adjuntos/inexistente.pdf','size':1}])
        self.assertEqual(self.ejecutar()['correosSeleccionados'],0)

    def test_copias_con_contenido_distinto_requieren_revision(self):
        self.guardar('a','mismo',cuerpo='Primera versión')
        self.guardar('b','mismo',cuerpo='Segunda versión')
        r=self.ejecutar()
        self.assertEqual(r['correosSeleccionados'],0)
        self.assertEqual(r['estadosInventario']['RevisarVersiones'],2)

    def test_control_ausente_no_se_inventa(self):
        self.guardar()
        for p in (self.root/'03_Control').glob('*'):p.unlink()
        self.assertEqual(self.ejecutar()['correosSeleccionados'],0)

    def test_controles_contradictorios_no_confirman(self):
        self.guardar()
        self.control(ids=[],nombre='control_copia.json')
        self.control(nombre='control_otro.json')
        r=self.ejecutar()
        self.assertEqual(r['correosSeleccionados'],0)
        self.assertEqual(r['controlesDiariosUtilizables'],0)

    def test_exclusion_explicita(self):
        self.guardar()
        self.assertEqual(self.ejecutar(exclusiones=['id-a'])['correosSeleccionados'],0)

    def test_eml_huerfano_y_json_corrupto(self):
        (self.root/'01_Correos'/'solo.eml').write_bytes(b'EML')
        (self.root/'01_Correos'/'malo.json').write_text('no json')
        r=self.ejecutar()
        self.assertEqual(r['tiposIncidencia']['eml_sin_json_relacionado'],1)
        self.assertEqual(r['tiposIncidencia']['json_no_utilizable'],1)

    def test_hashes_duplicados_y_archivos_sin_referencia(self):
        for name in ['a.pdf','b.pdf']:(self.root/'02_Adjuntos'/name).write_bytes(b'iguales')
        r=self.ejecutar()
        self.assertEqual(r['contenidosUnicosAdjuntos'],1)
        self.assertEqual(r['copiasBinariasAdicionales'],1)
        self.assertEqual(r['tiposIncidencia']['adjunto_sin_referencia_localizada'],2)

    def test_fuera_del_periodo(self):
        self.guardar(fecha='2026-11-17')
        self.assertEqual(self.ejecutar()['correosSeleccionados'],0)

    def test_nombres_sin_guiones(self):
        self.guardar()
        for a,b in [('01_Correos','01correos'),('02_Adjuntos','02adjuntos'),('03_Control','03control')]:
            (self.root/a).rename(self.root/b)
        self.assertEqual(self.ejecutar()['correosSeleccionados'],1)

    def test_salida_no_sobrescribe_informes(self):
        self.guardar();self.ejecutar()
        with self.assertRaises(ValueError):self.ejecutar()

    def test_rutas_con_espacios_y_url_codificada(self):
        (self.root/'02_Adjuntos'/'a uno.pdf').write_bytes(b'123')
        self.guardar(adjuntos=[{'estado':'Guardado','rutaGuardada':'https://example.test/ODS/a%20uno.pdf?download=1','size':3}])
        self.assertEqual(self.ejecutar()['correosSeleccionados'],1)

    def test_adjunto_con_nombre_ambiguo(self):
        (self.root/'02_Adjuntos'/'sub').mkdir()
        for path in ['a.pdf','sub/a.pdf']:(self.root/'02_Adjuntos'/path).write_bytes(b'123')
        self.guardar(adjuntos=[{'estado':'Guardado','rutaGuardada':'/ODS/a.pdf','size':3}])
        self.assertEqual(self.ejecutar()['correosSeleccionados'],0)

    def test_correo_indica_adjuntos_pero_sin_manifestar(self):
        p,d=self.guardar();d['correo']['hasAttachments']=True;p.write_text(json.dumps(d))
        self.assertEqual(self.ejecutar()['correosSeleccionados'],0)

if __name__=='__main__':
    unittest.main()
