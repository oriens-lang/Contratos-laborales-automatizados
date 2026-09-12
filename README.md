# ORIENS | Generador Inteligente de Contratos

Aplicación interna de Oriens Abogados para cargar un Excel de trabajadores, validar sus datos y generar contratos individuales de trabajo por tiempo indeterminado en Word, con el perfil de puesto de cada trabajador como ANEXO UNO.

## Cómo abrirla

Doble clic en `iniciar.command`, o bien desde la terminal:

```bash
.venv/bin/python app.py
```

Luego abre <http://127.0.0.1:5050>. Solo funciona en este equipo (no queda expuesta a la red).

## Plantilla del contrato (no incluida en el repositorio)

El formato de contrato del despacho no se publica. Para usar la aplicación, copia la plantilla oficial como `plantillas/FORMATO CONTRATO TIEMPO INDETERMINADO (original).docx` y crea la copia marcada:

```bash
.venv/bin/python plantillas/marcar_plantilla.py
```

## Estructura

```
generador-contratos/
├── app.py                  Servidor Flask: rutas de la página y de la API
├── config.py               Puerto, límite del archivo, plantilla, perfiles y salidas
├── servicios/              Lógica de negocio (sin nada de interfaz)
│   ├── archivos.py         Revisión del formato del archivo recibido
│   ├── lector_excel.py     Lectura de las hojas Trabajadores y Patron
│   ├── perfiles.py         Relación PUESTO → perfil y extracción de las cinco actividades
│   ├── validador.py        Datos faltantes, inconsistencias, puestos y perfiles
│   ├── mapeo_contrato.py   Mapeo Excel → marcadores de la plantilla oficial
│   └── generador.py        Contrato por trabajador + perfil como ANEXO UNO
├── templates/index.html    Estructura de la pantalla
├── static/                 Diseño (css) y comportamiento (js) de la pantalla
├── plantillas/             Plantilla oficial (original y copia marcada) y script de marcado
├── perfiles de puesto/     Un «Perfil de Puesto <puesto>.docx» por puesto
├── salidas/                Contratos generados: una carpeta por lote (no se versiona)
├── pruebas/                Archivos de prueba con datos simulados o ficticios
├── requirements.txt        Dependencias de Python
└── iniciar.command         Lanzador con doble clic (macOS)
```

## API

| Ruta | Qué hace |
|---|---|
| `GET /` | Pantalla principal |
| `POST /api/validar` | Lee el `.xlsx` (en memoria, sin guardarlo) y devuelve resumen, puestos y perfiles, patrón, trabajadores y validación |
| `POST /api/perfiles` | Recibe el perfil (.docx) de un puesto, revisa sus cinco actividades y lo guarda en `perfiles de puesto/` |
| `POST /api/generar` | Genera un `.docx` por trabajador con perfil en `salidas/Contratos AAAA-MM-DD HH.MM.SS/` |
| `POST /api/salidas/<lote>/abrir` | Abre en Finder la carpeta de un lote (solo dentro de `salidas/`) |

## Cómo lee el Excel

- Busca las hojas **Trabajadores** y **Patron** por su nombre (sin importar mayúsculas ni acentos); las demás se reportan como omitidas.
- Detecta la fila de encabezados y, si existe, la fila de grupos superior con celdas combinadas (DOMICILIO, SALARIO DIARIO IMSS, ACTA CONSTITUTIVA). Muestra los encabezados tal como están en el Excel.
- Ignora, y lo informa, las filas vacías, las que solo repiten texto del formato (p. ej. "FECHA DE INGRESO"), los encabezados repetidos, las notas o instrucciones y los datos sueltos sin nombre ni identificación.
- Una celda vacía nunca provoca error: se muestra como `____________` y genera una advertencia.
- Si una celda de un campo que no es fecha (EDAD, SALARIO, NÚMERO…) tiene formato de fecha, recupera el número capturado en lugar de mostrar una fecha falsa (p. ej. `29/01/1900` en vez de `29`).
- Revisa, solo como advertencia, el formato de CURP, RFC, NSS y correo, y que CURP, RFC y NSS no se repitan entre trabajadores.

## Perfiles de puesto

- **Desde la pantalla:** en «Puestos y perfiles», botón **Subir perfil** (o **Reemplazar**) del puesto. El Word se revisa (debe tener las cinco actividades) y se guarda con el nombre correcto; la validación se actualiza sola.
- También se pueden copiar a mano a `perfiles de puesto/` con el nombre **`Perfil de Puesto <puesto>.docx`** (p. ej. `Perfil de Puesto Ejecutivo de ventas.docx`). El puesto se compara con la columna PUESTO sin distinguir mayúsculas, acentos ni espacios repetidos.
- Solo el perfil de demostración (Diseñadora) se publica en el repositorio; los demás perfiles quedan fuera de git.
- De cada perfil se toman, sin modificarlas, las cinco actividades numeradas de la sección **«Cinco actividades principales»**; si no hay exactamente cinco, el perfil se marca como no utilizable.
- Un mismo perfil sirve a todos los trabajadores con ese puesto. Si falta el perfil, la validación muestra `Perfil de puesto pendiente: <puesto>` y **ese contrato no se genera**.

## Cómo genera los contratos

- Plantilla: `plantillas/Contrato tiempo indeterminado (marcado).docx`, copia de la plantilla oficial con marcadores `{{ campo }}` solo en los espacios variables (ver `plantillas/LEEME.md`).
- Por trabajador: datos del trabajador y del patrón, cinco actividades de SU perfil en la cláusula PRIMERA y SU perfil completo como ANEXO UNO en una sección nueva que conserva la página, los márgenes y el pie del perfil.
- Nombres de archivo: `Contrato Daniela Sofía Martínez López.docx` (espacios normales, sin guiones bajos; si se repite, `Contrato X (2).docx`).
- Cada lote incluye `Resumen de generación.txt`: perfil usado, datos en blanco y trabajadores no generados.
- Los datos pendientes no bloquean: se escriben como `____________` y la pantalla pide confirmación antes de generar.

## Pruebas

```bash
.venv/bin/python pruebas/crear_excel_ficticio.py   # pruebas/Base de prueba FICTICIA.xlsx (datos inventados)
```

`pruebas/Base ORIENS Datos Simulados Ejercicio.xlsx` es la base simulada del ejercicio (5 trabajadores, 4 puestos).

## Hoja de ruta

1. ✔ Interfaz inicial.
2. ✔ Lectura del Excel y validación de datos faltantes.
3. ✔ Generación en Word con la plantilla oficial y perfiles de puesto como ANEXO UNO.
4. Pendiente: definir los espacios de la plantilla sin dato en el Excel y agregar los perfiles faltantes.
