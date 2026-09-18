# ORIENS | Generador Inteligente de Contratos

Aplicación interna de Oriens Abogados para cargar un Excel de trabajadores, validar sus datos y generar contratos individuales de trabajo por tiempo indeterminado en Word, con el perfil de puesto de cada trabajador como ANEXO UNO.

## Uso diario

- **En la Mac del titular:** doble clic en **ORIENS Contratos** (Escritorio). Enciende el servidor si hace falta y abre la aplicación.
- **Abogada asistente, desde su computadora:** abrir en el navegador la dirección que aparece al pie de la pantalla («Acceso en la red del despacho», por ejemplo `http://192.168.1.20:5050/`). Ambas computadoras deben estar en la red de la oficina y la Mac del titular, encendida.
- **Contraseña:** la primera vez se crea en la Mac del titular; después la usan ambos. Para restablecerla, borrar `acceso.json` y volver a entrar desde esa Mac.

Flujo:
1. **Plantilla del contrato:** elegir una de las guardadas o subir una nueva (no hay plantilla predeterminada; el Excel se habilita al elegirla).
2. **Datos del Excel:** cargar el Excel del cliente (el **formato de captura** se descarga desde ese mismo paso).
3. **Validar información** y subir los perfiles de puesto que falten.
4. **Fecha de firma:** una misma fecha para todos (predeterminada) o la FECHA ALTA IMSS de cada trabajador (solo para contratos de nuevo ingreso; nunca en los de regularización) → **Generar contratos** → **Descargar contratos (.zip)**.

### Varias plantillas

- En el paso 1, **Subir plantilla nueva** acepta cualquier Word que tenga marcadores `{{ campo }}`. La guía «Cómo preparar una plantilla nueva» (en ese mismo paso) lista todos los marcadores con un botón **Copiar**.
- Si la plantilla usa las actividades del perfil (`{{ actividad_1 }}` … `{{ actividad_5 }}`), cada contrato lleva el perfil de puesto como ANEXO UNO; si no las usa (p. ej., un convenio), no se agrega anexo.
- El resumen de cada lote solo reporta en blanco los datos que usa la plantilla elegida.

## Plantilla del contrato (no incluida en el repositorio)

El formato de contrato del despacho no se publica. Para instalar la aplicación en otra Mac, copia la plantilla oficial como `plantillas/FORMATO CONTRATO TIEMPO INDETERMINADO (original).docx` y crea la copia marcada:

```bash
.venv/bin/python plantillas/marcar_plantilla.py
```

Variante de regularización (relaciones de trabajo ya vigentes): copia su formato como `plantillas/FORMATO CONTRATO TIEMPO INDETERMINADO REGULARIZACION (original).docx` y ejecuta `marcar_plantilla.py regularizacion`.

Para cambiar el texto de la plantilla, ver `plantillas/LEEME.md`.

## Estructura

```
generador-contratos/
├── app.py                  Servidor Flask: acceso con contraseña, pantalla y API
├── config.py               Puerto, red, carpetas y archivo de acceso
├── servicios/              Lógica de negocio (sin nada de interfaz)
│   ├── acceso.py           Contraseña compartida (solo se guarda su hash)
│   ├── archivos.py         Revisión del formato de los archivos recibidos
│   ├── lector_excel.py     Lectura de las hojas Trabajadores y Patron
│   ├── perfiles.py         Relación PUESTO → perfil y extracción de las cinco actividades
│   ├── validador.py        Datos faltantes, inconsistencias, puestos y perfiles
│   ├── mapeo_contrato.py   Mapeo Excel → marcadores de la plantilla
│   ├── plantillas.py       Plantillas disponibles, subida y revisión
│   └── generador.py        Contrato por trabajador + perfil como ANEXO UNO
├── templates/              Pantallas (principal y acceso)
├── static/                 Diseño (css) y comportamiento (js)
├── plantillas/             Plantilla oficial, copia marcada, formato de captura y scripts
├── perfiles de puesto/     Un «Perfil de Puesto <puesto>.docx» por puesto
├── contratos generados/    Una carpeta por cliente y fecha (no se versiona)
├── pruebas/                Archivos de prueba con datos simulados o ficticios
└── requirements.txt        Dependencias de Python
```

## API (requiere sesión iniciada)

| Ruta | Qué hace |
|---|---|
| `GET /entrar`, `GET /salir` | Acceso con la contraseña del despacho |
| `GET /api/formato` | Descarga el formato de captura para el cliente |
| `POST /api/validar` | Lee el `.xlsx` (en memoria) y devuelve resumen, puestos y perfiles, patrón, trabajadores y validación |
| `POST /api/generar` | Genera un `.docx` por trabajador (campos: `archivo`, `plantilla`, `fecha_firma`) |
| `GET /api/plantillas` · `POST /api/plantillas` | Lista las plantillas · sube una plantilla marcada (se revisan sus marcadores) |
| `GET /api/plantillas/<nombre>/descargar` | Descarga una plantilla para editarla |
| `POST /api/perfiles` | Recibe el perfil (.docx) de un puesto y revisa sus cinco actividades |
| `GET /api/salidas/<lote>/zip` | Descarga los contratos de un lote |
| `POST /api/salidas/<lote>/abrir` | Abre la carpeta en Finder (solo en la Mac donde corre) |

## Cómo lee el Excel

- Busca las hojas **Trabajadores** y **Patron** por su nombre; las demás se reportan como omitidas.
- Detecta los encabezados en uno o dos niveles (grupos combinados como DOMICILIO o HORARIO SÁBADO) y los muestra tal como están.
- Ignora, y lo informa, las filas vacías, las que solo repiten texto del formato, los encabezados repetidos, las notas y los datos sueltos sin nombre ni identificación.
- Una celda vacía nunca provoca error: se muestra como `____________` y genera una advertencia.
- Si una celda que no es fecha tiene formato de fecha, recupera el número capturado (p. ej. `29` y no `29/01/1900`).
- Revisa, solo como advertencia, CURP, RFC, NSS y correo, y que no se repitan entre trabajadores.

## Cómo llena el contrato

- Datos del trabajador y del patrón; lugar y fecha de nacimiento y sexo (columnas del formato nuevo).
- **Declaración II b):** la experiencia se llena con el PUESTO.
- **Colonia:** se completa con el municipio, porque la plantilla no tiene espacio propio para él.
- **Cláusula CUARTA:** entrada, inicio y fin de comida y salida de lunes a viernes, y horario del sábado. Si faltan las columnas de entrada y salida, se toman las dos horas de JORNADA DE TRABAJO.
- **Cláusula PRIMERA:** las cinco actividades del perfil del puesto; el perfil completo va como ANEXO UNO.
- **Fecha y lugar de firma:** la fecha elegida en la pantalla (o la FECHA ALTA IMSS de cada trabajador); el lugar, cuando la plantilla lo usa, es el domicilio del patrón.
- **Fecha de ingreso:** la columna FECHA DE INGRESO; si viene vacía, la FECHA ALTA IMSS (se anota en el resumen del lote).
- **Horas a la semana:** `{{ horas_semana }}` = (salida − entrada − comida) × 5; p. ej., 8:30 a 18:00 con 30 minutos → 45.
- **Edad:** `{{ edad }}` escribe «52 años de edad».
- **Comida:** `{{ comida_duracion }}` escribe la duración que indica el Excel («30 min» → «30 minutos»).
- **Domicilio en una sola columna** (DOMICILIO): se escribe tal como se capturó con `{{ domicilio }}`.
- **Firma de la empresa:** nombre del representante legal y razón social (`{{ patron_representante }}`).
- Lo que falte queda como `____________` y se lista en `Resumen de generación.txt`.

## Perfiles de puesto

- Desde la pantalla, en «Puestos y perfiles», botón **Subir perfil** (o **Reemplazar**). Debe tener la sección «Cinco actividades principales» con cinco actividades numeradas.
- Se guardan como `Perfil de Puesto <puesto>.docx`. Solo el de demostración (Diseñadora) se publica en el repositorio.
