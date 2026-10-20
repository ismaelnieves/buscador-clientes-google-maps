# Buscador de clientes en Google Maps

Herramienta en Python para freelancers y agencias web: busca negocios de un rubro en una ciudad (por ejemplo, *pizzerías en Córdoba, Argentina*), revisa la página web de cada uno y detecta cuáles **no tienen web, dependen solo de redes sociales o tienen una web vieja**. Con eso arma un Excel ordenado por prioridad, con plantillas de correo y WhatsApp listas para personalizar.

**No envía nada**: vos revisás la lista y contactás a mano.

| Nombre | Estado web | Problema principal | Puntaje | Prioridad | WhatsApp | Plantilla WhatsApp |
|---|---|---|---|---|---|---|
| Maderas Viejas | Vieja | Sin celular | 7 | Alta | 098 765 432 | Hola, ¿qué tal? Soy … |
| Carpintería Sin Web | No tiene | Sin web | 7 | Alta | 099 123 456 | Hola, ¿qué tal? Soy … |
| Muebles Insta | Solo redes o app de terceros | Solo redes | 6 | Alta | No figura | Hola, ¿qué tal? Soy … |

*Ejemplo con datos inventados. El Excel completo está en [ejemplo-salida-simulada.xlsx](ejemplo-salida-simulada.xlsx).*

## Usar el ejecutable de Windows (sin instalar nada)

1. Descargá `buscar_clientes.exe` desde [Releases](https://github.com/ismaelnieves/buscador-clientes-google-maps/releases).
2. Al lado del `.exe`, creá un `config.json` copiando [config.example.json](config.example.json) y poné tu clave de Google Places (ver abajo).
3. Hacé doble clic en `buscar_clientes.exe`. Te pregunta el rubro, la ciudad y el país, y deja el Excel en esa misma carpeta.

Windows puede mostrar "Windows protegió su PC" porque el programa no está firmado: tocá *Más información* → *Ejecutar de todas formas*. El `.exe` se compiló desde `buscar_clientes.py` con PyInstaller y no trae ninguna clave adentro.

## Usar con Python

1. Instalá Python 3.9 o más nuevo desde [python.org](https://www.python.org/downloads/) (en Windows marcá *Add Python to PATH*).
2. Instalá las dependencias:
   ```
   pip install -r requirements.txt
   ```
3. Copiá `config.example.json` con el nombre `config.json` y completá:
   - `google_api_key`: tu clave de Google Places (Google Cloud Console → activar **Places API (New)** → Credenciales → Crear clave).
   - `remitente_nombre`, `remitente_descripcion` y `remitente_contacto`: se usan en la firma de las plantillas.

   `config.json` queda solo en tu computadora (está en `.gitignore`). No lo compartas.
4. Buscá:
   ```
   python buscar_clientes.py "carpinterías" "Córdoba" "Argentina"
   ```
   Al terminar deja un archivo como `clientes-carpinterías-córdoba-argentina-2026-10-19.xlsx`. Con `--salida OTRA_CARPETA` lo guarda en otro lado.

En `config.json` también podés cambiar `paginas_max` (cada página trae hasta 20 negocios, máximo 3) y `puntaje_minimo` (por defecto 4; bajalo para ver más negocios).

### Probar sin gastar consultas

```
python prueba_simulada.py
```

Usa negocios inventados, comprueba que todo dé lo esperado y genera `ejemplo-salida-simulada.xlsx`.

## Países

Funciona con cualquier país: el país se agrega a la búsqueda de Google. Para estos, además, se ajusta la región de búsqueda y se detecta si el teléfono es celular (y por lo tanto probable WhatsApp):

Uruguay, Argentina, Chile, Paraguay, Colombia, Perú y España.

México y Estados Unidos tienen región, pero sin detección de celular, porque allí no se distingue por el número. Para agregar un país, sumalo a `PAISES` en `buscar_clientes.py`.

## Qué hace, paso a paso

1. Busca `"<rubro> en <ciudad>, <país>"` en Google Places (hasta 3 páginas de 20).
2. Descarta los negocios cerrados y los repetidos.
3. Si el único enlace es Facebook, Instagram, Linktree, AgendaPro, etc., lo marca como *Solo redes*. Si no hay enlace, *No tiene*.
4. Si tiene web propia, la descarga y anota problemas: no carga, sin HTTPS, no adaptada a celulares, año viejo en el pie, tecnología antigua, sin WhatsApp, sin formulario, sin descripción para Google. Si encuentra un correo en la página, lo guarda.
5. Calcula el puntaje, descarta los de puntaje bajo y ordena de mayor a menor.
6. Arma el Excel con los datos, los problemas detectados y las plantillas de *asunto*, *correo* y *WhatsApp*.

### Puntaje

| Estado de la web | Puntos |
|---|---|
| Solo redes o app de terceros | 5 |
| No tiene | 4 |
| Caída o no responde | 4 |
| Vieja | 3 |
| Mejorable | 1 |
| Correcta | 0 |

A eso se suma: **+1** si tiene teléfono (**−2** si no tiene), **+1** si el teléfono es celular, **+1** si tiene 20 reseñas o más (negocio activo) y **+1** si se encontró un correo.

Prioridad: **Alta** desde 6 puntos, **Media** desde 4, **Baja** desde 2.

Para cambiar los textos de las plantillas, editá `problema_principal` y `armar_fila` en `buscar_clientes.py`.

## Costos y cuidado

- Cada búsqueda completa son hasta 3 consultas a Google Places (Text Search). Hay una cuota gratis por mes; poné una alerta de presupuesto en Google Cloud.
- Google Places no da correos: solo aparecen si están escritos en la web del negocio.
- Revisá cada mensaje antes de mandarlo, mandá pocos y personalizados, y si alguien pide que no lo contactes más, no lo contactes. Respetá la ley de protección de datos de tu país (en Uruguay, la Ley 18.331).

## Licencia

[MIT](LICENSE)
