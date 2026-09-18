/* ORIENS | Generador Inteligente de Contratos — lógica de la interfaz */
(() => {
  "use strict";

  const LIMITE_MB = Number(document.body.dataset.limiteMb) || 10;
  const ES_LOCAL = document.body.dataset.esLocal === "1";
  const NOTA_INICIAL = "Elige la plantilla y carga un archivo .xlsx para iniciar la validación.";
  const FALTANTE = "____________";
  const TIPO_WORD = ".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  const $ = (id) => document.getElementById(id);
  const ui = {
    plantilla: $("plantilla"),
    notaPlantilla: $("nota-plantilla"),
    descargarPlantilla: $("enlace-descargar-plantilla"),
    subirPlantilla: $("boton-subir-plantilla"),
    tarjetaCarga: $("tarjeta-carga"),
    avisoBloqueo: $("aviso-bloqueo"),
    zona: $("zona-carga"),
    entrada: $("entrada-archivo"),
    seleccionado: $("archivo-seleccionado"),
    nombre: $("archivo-nombre"),
    detalle: $("archivo-detalle"),
    quitar: $("boton-quitar"),
    error: $("error-carga"),
    resumen: $("resumen"),
    validacion: $("validacion"),
    seccionPatron: $("seccion-patron"),
    patron: $("patron"),
    notaPatron: $("nota-patron"),
    seccionTrabajadores: $("seccion-trabajadores"),
    trabajadores: $("trabajadores"),
    notaTrabajadores: $("nota-trabajadores"),
    seccionPerfiles: $("seccion-perfiles"),
    perfiles: $("perfiles"),
    notaPerfiles: $("nota-perfiles"),
    validar: $("boton-validar"),
    generar: $("boton-generar"),
    nota: $("nota-acciones"),
    fechaFirma: $("fecha-firma"),
    modoFirma: $("modo-firma"),
    campoFechaFirma: $("campo-fecha-firma"),
    seccionGenerados: $("seccion-generados"),
    notaGenerados: $("nota-generados"),
    listaGenerados: $("lista-generados"),
    descargarZip: $("boton-zip"),
    abrir: $("boton-abrir"),
    pasos: document.querySelectorAll(".paso"),
    banda: {
      trabajadores: $("banda-trabajadores"),
      trabajadoresDetalle: $("banda-trabajadores-detalle"),
      puestos: $("banda-puestos"),
      perfiles: $("banda-perfiles"),
      perfilesDetalle: $("banda-perfiles-detalle"),
      contratos: $("banda-contratos"),
      contratosDetalle: $("banda-contratos-detalle"),
    },
  };

  // plantillas: descripción de /api/plantillas; resultado: última respuesta de /api/validar;
  // lote: carpeta del último lote generado.
  const estado = {
    plantillas: [], archivo: null, resultado: null, lote: null, puestoPerfil: null,
    validando: false, generando: false,
  };

  // Contenido inicial de los paneles, para restaurarlo al quitar el archivo.
  const VACIO_RESUMEN = ui.resumen.innerHTML;
  const VACIO_VALIDACION = ui.validacion.innerHTML;

  /* ---------- Utilidades ---------- */

  function crear(etiqueta, clase, texto) {
    const nodo = document.createElement(etiqueta);
    if (clase) nodo.className = clase;
    if (texto !== undefined) nodo.textContent = texto;
    return nodo;
  }

  function formatearTamano(bytes) {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  function mostrarError(mensaje) {
    ui.error.textContent = mensaje;
    ui.error.hidden = !mensaje;
  }

  // Mensaje de la barra de acciones.
  function avisar(texto, esError = false) {
    ui.nota.textContent = texto;
    ui.nota.classList.toggle("barra-acciones__nota--error", esError);
  }

  function errorLegible(error) {
    return error instanceof TypeError
      ? "No fue posible conectar con el servidor. Verifica que la aplicación siga en ejecución."
      : error.message;
  }

  // Petición a la API: devuelve el JSON o lanza un Error con el mensaje del servidor.
  async function solicitar(url, opciones) {
    const respuesta = await fetch(url, opciones);
    if (respuesta.status === 401) {
      window.location.href = "/entrar";
      throw new Error("La sesión expiró; vuelve a entrar.");
    }
    const json = await respuesta.json().catch(() => ({ error: "Respuesta inesperada del servidor." }));
    if (!respuesta.ok || !json.ok) throw new Error(json.error || `Error ${respuesta.status}.`);
    return json;
  }

  // Selector de archivos oculto (para subir perfiles y plantillas).
  function selectorArchivo(tipos, alElegir) {
    const entrada = crear("input");
    entrada.type = "file";
    entrada.accept = tipos;
    entrada.hidden = true;
    document.body.append(entrada);
    entrada.addEventListener("change", () => {
      const archivo = entrada.files[0];
      entrada.value = "";
      if (archivo) alElegir(archivo);
    });
    return entrada;
  }

  function marcarPaso(actual) {
    ui.pasos.forEach((paso) => {
      const numero = Number(paso.dataset.paso);
      paso.classList.toggle("paso--activo", numero === actual);
      paso.classList.toggle("paso--hecho", numero < actual);
    });
  }

  // Paso en el que está el usuario según lo que ya eligió o cargó.
  function pasoActual() {
    if (!ui.plantilla.value) return 1;
    if (!estado.archivo) return 2;
    return estado.resultado?.puede_generar ? 4 : 3;
  }

  // Banda de estado superior: valores ausentes se muestran como «—».
  function actualizarBanda(valores) {
    for (const [clave, nodo] of Object.entries(ui.banda)) {
      const valor = valores[clave];
      nodo.textContent = valor === undefined || valor === null ? (clave.endsWith("Detalle") ? "" : "—") : String(valor);
    }
  }

  function reiniciarPaneles() {
    estado.resultado = null;
    ui.resumen.innerHTML = VACIO_RESUMEN;
    ui.validacion.innerHTML = VACIO_VALIDACION;
    ui.seccionPatron.hidden = true;
    ui.seccionTrabajadores.hidden = true;
    ui.seccionPerfiles.hidden = true;
    ui.seccionGenerados.hidden = true;
    ui.generar.disabled = true;
    actualizarBanda({});
  }

  function ponerCargando(boton, activo, textoCargando) {
    if (activo) {
      boton.dataset.texto = boton.textContent;
      boton.replaceChildren(crear("span", "girador"), textoCargando);
    } else {
      boton.textContent = boton.dataset.texto;
    }
    boton.disabled = activo;
    boton.classList.toggle("cargando", activo);
  }

  // Tarjeta de dato: etiqueta + valor; si el valor falta, muestra la línea en blanco.
  function crearDato(etiqueta, valor, clases = "") {
    const vacio = valor === undefined || valor === null || valor === "";
    const dato = crear("div", `dato ${clases}${vacio ? " dato--faltante" : ""}`);
    dato.append(crear("dt", "", etiqueta), crear("dd", "", vacio ? FALTANTE : String(valor)));
    return dato;
  }

  function tituloPlantilla() {
    return ui.plantilla.selectedOptions[0]?.textContent || ui.plantilla.value;
  }

  /* ---------- Paso 1 · Plantilla ---------- */

  async function cargarPlantillas(seleccion) {
    try {
      const json = await solicitar("/api/plantillas");
      estado.plantillas = json.plantillas;
      const elegida = seleccion ?? ui.plantilla.value;
      const opciones = [crear("option", "", "— Elige una plantilla —")];
      opciones[0].value = "";
      for (const p of json.plantillas) {
        const opcion = crear("option", "", p.error ? `${p.titulo} (no utilizable)` : p.titulo);
        opcion.value = p.nombre;
        opcion.disabled = Boolean(p.error);
        opcion.selected = p.nombre === elegida && !p.error;
        opciones.push(opcion);
      }
      ui.plantilla.replaceChildren(...opciones);
      alCambiarPlantilla();
    } catch (error) {
      ui.notaPlantilla.textContent = errorLegible(error);
    }
  }

  function alCambiarPlantilla() {
    const nombre = ui.plantilla.value;
    const descripcion = estado.plantillas.find((p) => p.nombre === nombre);
    ui.descargarPlantilla.hidden = !nombre;
    ui.descargarPlantilla.href = nombre ? `/api/plantillas/${encodeURIComponent(nombre)}/descargar` : "#";
    ui.notaPlantilla.textContent = descripcion
      ? `Usa ${descripcion.campos} datos del Excel · ` +
        (descripcion.anexo ? "incluye las cinco actividades y el perfil de puesto como ANEXO UNO."
                           : "no usa perfil de puesto (sin ANEXO UNO).")
      : `Todavía no has elegido plantilla. Hay ${estado.plantillas.filter((p) => !p.error).length} disponibles.`;

    // El Excel se carga solo después de elegir la plantilla.
    const bloqueado = !nombre;
    ui.tarjetaCarga.classList.toggle("tarjeta--bloqueada", bloqueado);
    ui.avisoBloqueo.hidden = !bloqueado;
    ui.entrada.disabled = bloqueado;
    if (!estado.archivo) avisar(bloqueado ? NOTA_INICIAL : "Plantilla elegida. Carga el Excel de trabajadores (paso 2).");
    marcarPaso(pasoActual());
  }

  const entradaPlantilla = selectorArchivo(TIPO_WORD, subirPlantilla);

  async function subirPlantilla(archivo) {
    if (!archivo.name.toLowerCase().endsWith(".docx")) {
      ui.notaPlantilla.textContent = "La plantilla debe ser un documento Word (.docx).";
      return;
    }
    if (estado.plantillas.some((p) => p.nombre === archivo.name) &&
        !window.confirm(`Ya existe la plantilla «${archivo.name}». ¿Deseas reemplazarla por la versión que elegiste?`)) {
      return;
    }
    const datos = new FormData();
    datos.append("archivo", archivo);
    ui.notaPlantilla.textContent = `Revisando la plantilla «${archivo.name}»…`;
    try {
      const json = await solicitar("/api/plantillas", { method: "POST", body: datos });
      await cargarPlantillas(json.nombre);
      ui.notaPlantilla.textContent = `Plantilla «${json.nombre.replace(/\.docx$/i, "")}» revisada y elegida. ` +
        ui.notaPlantilla.textContent;
    } catch (error) {
      ui.notaPlantilla.textContent = errorLegible(error);
    }
  }

  async function copiarMarcador(boton) {
    const texto = boton.dataset.marcador;
    try {
      await navigator.clipboard.writeText(texto);
    } catch {
      // Sin acceso al portapapeles (p. ej., desde otra computadora de la red): copia a la antigua.
      const area = crear("textarea");
      area.value = texto;
      document.body.append(area);
      area.select();
      document.execCommand("copy");
      area.remove();
    }
    boton.textContent = "Copiado";
    setTimeout(() => { boton.textContent = "Copiar"; }, 1200);
  }

  /* ---------- Paso 2 · Excel ---------- */

  function seleccionarArchivo(archivo) {
    mostrarError("");
    if (!archivo) return;
    if (!ui.plantilla.value) {
      ui.entrada.value = "";
      mostrarError("Primero elige la plantilla del contrato (paso 1).");
      return;
    }
    if (!archivo.name.toLowerCase().endsWith(".xlsx")) {
      ui.entrada.value = "";
      mostrarError(`“${archivo.name}” no es un archivo .xlsx. Carga un libro de Excel en ese formato.`);
      return;
    }
    if (archivo.size > LIMITE_MB * 1024 * 1024) {
      ui.entrada.value = "";
      mostrarError(`El archivo pesa ${formatearTamano(archivo.size)}; el límite es de ${LIMITE_MB} MB.`);
      return;
    }

    estado.archivo = archivo;
    ui.nombre.textContent = archivo.name;
    ui.detalle.textContent = `${formatearTamano(archivo.size)} · listo para validar`;
    ui.zona.hidden = true;
    ui.seleccionado.hidden = false;
    reiniciarPaneles();
    avisar("Archivo cargado. Pulsa “Validar información” para revisarlo.");
    marcarPaso(pasoActual());
  }

  function quitarArchivo() {
    estado.archivo = null;
    ui.entrada.value = "";
    ui.seleccionado.hidden = true;
    ui.zona.hidden = false;
    mostrarError("");
    reiniciarPaneles();
    alCambiarPlantilla();
  }

  /* ---------- Paso 3 · Validación ---------- */

  async function validar() {
    if (estado.validando || estado.generando) return;
    if (!ui.plantilla.value) {
      avisar("Primero elige la plantilla del contrato (paso 1).", true);
      return;
    }
    if (!estado.archivo) {
      mostrarError("Primero carga el archivo Excel (.xlsx) de trabajadores.");
      return;
    }

    estado.validando = true;
    mostrarError("");
    reiniciarPaneles();
    ponerCargando(ui.validar, true, "Validando…");

    const datos = new FormData();
    datos.append("archivo", estado.archivo);

    try {
      const json = await solicitar("/api/validar", { method: "POST", body: datos });
      estado.resultado = json;
      pintarResumen(json.resumen, json.hojas);
      pintarValidacion(json.validacion);
      pintarPerfiles(json.perfiles);
      pintarPatron(json.patron);
      pintarTrabajadores(json.trabajadores);
      ui.generar.disabled = !json.puede_generar;
      marcarPaso(pasoActual());
      const r = json.resumen;
      actualizarBanda({
        trabajadores: r.trabajadores,
        trabajadoresDetalle: `${r.completos} completos · ${r.con_faltantes} con pendientes`,
        puestos: r.puestos,
        perfiles: r.perfiles_disponibles,
        perfilesDetalle: `${r.perfiles_pendientes} pendientes`,
      });
      if (!r.trabajadores) {
        avisar("No hay trabajadores capturados; no hay contratos que generar.");
      } else {
        avisar(`Se generarán ${r.trabajadores} contratos con la plantilla «${tituloPlantilla()}». ` +
               `Los datos pendientes quedarán como ${FALTANTE}.`);
      }
    } catch (error) {
      pintarValidacion({ estado: "error", observaciones: [errorLegible(error)] });
    } finally {
      estado.validando = false;
      ponerCargando(ui.validar, false);
    }
  }

  /* ---------- Paso 4 · Generación ---------- */

  async function generar() {
    const resultado = estado.resultado;
    if (estado.generando || estado.validando || !estado.archivo || !resultado?.puede_generar) return;
    if (!ui.plantilla.value) {
      avisar("Elige la plantilla del contrato (paso 1).", true);
      return;
    }

    const r = resultado.resumen;
    const conAnexo = estado.plantillas.find((p) => p.nombre === ui.plantilla.value)?.anexo;
    const partes = [];
    if (conAnexo && r.generables < r.trabajadores) {
      partes.push(`${r.trabajadores - r.generables} trabajador(es) sin perfil de puesto: su contrato saldrá ` +
                  `con las actividades como ${FALTANTE} y sin ANEXO UNO.`);
    }
    if (r.con_faltantes) {
      partes.push(`${r.con_faltantes} trabajador(es) con datos faltantes: se escribirán como ${FALTANTE} ` +
                  "para completarlos a mano antes de firmar.");
    }
    const conFechaUnica = ui.modoFirma.value === "fecha";
    if (conFechaUnica && !ui.fechaFirma.value) {
      partes.push(`No elegiste fecha de firma: quedará como ${FALTANTE}.`);
    }
    if (partes.length && !window.confirm(`${partes.join("\n\n")}\n\n¿Deseas generar los contratos?`)) return;

    estado.generando = true;
    ponerCargando(ui.generar, true, "Generando…");
    const datos = new FormData();
    datos.append("archivo", estado.archivo);
    datos.append("plantilla", ui.plantilla.value);
    datos.append("fecha_firma", conFechaUnica ? ui.fechaFirma.value : "imss");

    try {
      const json = await solicitar("/api/generar", { method: "POST", body: datos });
      pintarGenerados(json);
      avisar(`Se ${json.total === 1 ? "generó 1 contrato" : `generaron ${json.total} contratos`} ` +
             `con la plantilla «${tituloPlantilla()}».`);
      marcarPaso(5);
    } catch (error) {
      avisar(errorLegible(error), true);
    } finally {
      estado.generando = false;
      ponerCargando(ui.generar, false);
    }
  }

  function pintarGenerados(lote) {
    estado.lote = lote.lote;
    ui.banda.contratos.textContent = String(lote.total);
    ui.banda.contratosDetalle.textContent = `en «${lote.lote}»`;
    ui.notaGenerados.textContent = `Carpeta: ${lote.carpeta}`;
    const items = [...lote.archivos, "Resumen de generación.txt"].map((nombre) => {
      const item = crear("li", "", nombre);
      item.dataset.tipo = nombre.split(".").pop().toUpperCase();  // etiqueta DOCX / TXT
      return item;
    });
    for (const pendiente of lote.sin_perfil || []) {
      const item = crear("li", "omitido",
        `${pendiente.nombre}: perfil de puesto pendiente («${pendiente.puesto || "sin puesto"}»); ` +
        "contrato sin actividades ni ANEXO UNO");
      item.dataset.tipo = "AVISO";
      items.push(item);
    }
    ui.listaGenerados.replaceChildren(...items);
    ui.abrir.hidden = !ES_LOCAL;
    ui.seccionGenerados.hidden = false;
    ui.seccionGenerados.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }

  function descargarLote() {
    if (estado.lote) window.location.href = `/api/salidas/${encodeURIComponent(estado.lote)}/zip`;
  }

  async function abrirCarpeta() {
    if (!estado.lote) return;
    try {
      await solicitar(`/api/salidas/${encodeURIComponent(estado.lote)}/abrir`, { method: "POST" });
    } catch (error) {
      avisar(errorLegible(error), true);
    }
  }

  /* ---------- Perfiles de puesto ---------- */

  const entradaPerfil = selectorArchivo(TIPO_WORD, subirPerfil);

  function elegirPerfil(puesto) {
    estado.puestoPerfil = puesto;
    entradaPerfil.click();
  }

  async function subirPerfil(archivo) {
    const puesto = estado.puestoPerfil;
    if (!puesto) return;
    if (!archivo.name.toLowerCase().endsWith(".docx")) {
      avisar("El perfil debe ser un documento Word (.docx).", true);
      return;
    }
    const datos = new FormData();
    datos.append("puesto", puesto);
    datos.append("archivo", archivo);
    avisar(`Cargando el perfil de «${puesto}»…`);
    try {
      const json = await solicitar("/api/perfiles", { method: "POST", body: datos });
      await validar();  // vuelve a relacionar puestos y perfiles con el perfil nuevo
      avisar(`Perfil de «${json.puesto}» cargado (${json.actividades.length} actividades). ${ui.nota.textContent}`);
    } catch (error) {
      avisar(errorLegible(error), true);
    }
  }

  /* ---------- Paneles ---------- */

  function pintarResumen(resumen, hojas) {
    const lista = crear("dl", "datos datos--tres");
    const patron = crearDato("Patrón detectado", resumen.patron || "No detectado", "dato--ancho");
    if (!resumen.patron) patron.classList.add("dato--pendiente");
    lista.append(
      patron,
      crearDato("Trabajadores", resumen.trabajadores),
      crearDato("Completos", resumen.completos),
      crearDato("Con pendientes", resumen.con_faltantes),
      crearDato("Puestos distintos", resumen.puestos),
      crearDato("Perfiles disponibles", resumen.perfiles_disponibles),
      crearDato("Perfiles pendientes", resumen.perfiles_pendientes),
    );

    let texto = `Hojas leídas: ${hojas.leidas.join(", ")}`;
    if (hojas.omitidas.length) texto += ` · Omitidas: ${hojas.omitidas.join(", ")}`;
    ui.resumen.replaceChildren(lista, crear("p", "panel-nota", texto));
  }

  function pintarValidacion(validacion) {
    const ETIQUETAS = { correcto: "Sin pendientes", pendiente: "Con pendientes", error: "Con errores" };
    const tipo = ETIQUETAS[validacion.estado] ? validacion.estado : "pendiente";
    const nodos = [crear("span", `estado estado--${tipo}`, ETIQUETAS[tipo])];

    const generales = crear("ul", "observaciones");
    for (const texto of validacion.observaciones || []) generales.append(crear("li", "", texto));
    nodos.push(generales);

    const incidencias = validacion.por_trabajador || [];
    if (incidencias.length) {
      nodos.push(crear("p", "subtitulo-panel", `Pendientes por trabajador (${incidencias.length})`));
      const lista = crear("ul", "incidencias");
      for (const t of incidencias) {
        const item = crear("li", "incidencia");
        item.append(crear("p", "incidencia__titulo", `Fila ${t.fila} · ${t.nombre}`));
        if (t.faltantes.length) {
          item.append(crear("p", "incidencia__detalle", `Faltan (${t.faltantes.length}): ${t.faltantes.join(", ")}.`));
        }
        for (const aviso of t.advertencias) {
          item.append(crear("p", "incidencia__detalle incidencia__detalle--aviso", aviso));
        }
        lista.append(item);
      }
      nodos.push(lista);
    }
    ui.validacion.replaceChildren(...nodos);
  }

  const ESTADOS_PERFIL = {
    disponible: ["correcto", "Disponible"],
    pendiente: ["pendiente", "Pendiente"],
    invalido: ["error", "No utilizable"],
    sin_puesto: ["pendiente", "Sin puesto"],
  };

  function pastillaPerfil(estadoPerfil) {
    const [clase, texto] = ESTADOS_PERFIL[estadoPerfil] || ESTADOS_PERFIL.pendiente;
    return crear("span", `pastilla pastilla--${clase}`, texto);
  }

  function pintarPerfiles(perfiles) {
    const tabla = crear("table", "tabla");
    const encabezado = crear("tr");
    for (const texto of ["Puesto (Excel)", "Trabajadores", "Perfil de puesto", "Estado", "Acción"]) {
      const th = crear("th", "", texto);
      th.scope = "col";
      encabezado.append(th);
    }
    const thead = crear("thead");
    thead.append(encabezado);

    const tbody = crear("tbody");
    for (const p of perfiles.puestos) {
      const fila = crear("tr");
      const estadoCelda = crear("td");
      estadoCelda.append(pastillaPerfil(p.estado));
      const accionCelda = crear("td");
      const boton = crear("button", "boton-accion", p.estado === "disponible" ? "Reemplazar" : "Subir perfil");
      boton.type = "button";
      boton.addEventListener("click", () => elegirPerfil(p.puesto));
      accionCelda.append(boton);
      fila.append(
        crear("td", "", p.puesto),
        crear("td", "tabla__lista", p.trabajadores.map((t) => `${t.nombre} (fila ${t.fila})`).join("\n")),
        crear("td", p.perfil ? "" : "tabla__faltante", p.perfil || `Falta «Perfil de Puesto ${p.puesto}.docx»`),
        estadoCelda,
        accionCelda,
      );
      tbody.append(fila);
    }
    tabla.append(thead, tbody);
    const contenedor = crear("div", "tabla-contenedor");
    contenedor.append(tabla);
    const nodos = [contenedor, crear("p", "panel-nota",
      "Sube el perfil en Word (.docx) de cada puesto. Debe incluir la sección «Cinco actividades principales» " +
      "con cinco actividades numeradas; se guardará como «Perfil de Puesto <puesto>.docx».")];

    // Las cinco actividades que irán a la cláusula PRIMERA, por perfil disponible.
    for (const p of perfiles.puestos.filter((x) => x.estado === "disponible")) {
      const detalles = crear("details", "actividades");
      detalles.append(crear("summary", "", `Cinco actividades del perfil «${p.puesto}» (cláusula PRIMERA)`));
      const lista = crear("ol");
      for (const actividad of p.actividades) lista.append(crear("li", "", actividad));
      detalles.append(lista);
      nodos.push(detalles);
    }

    const pendientes = perfiles.puestos.filter((x) => x.estado !== "disponible").length;
    ui.notaPerfiles.textContent = `${perfiles.puestos.length} puestos · ` +
      `${perfiles.puestos.length - pendientes} con perfil · ${pendientes} pendientes`;
    ui.perfiles.replaceChildren(...nodos);
    ui.seccionPerfiles.hidden = false;
  }

  function pintarPatron(patron) {
    if (!patron) {
      ui.notaPatron.textContent = "";
      ui.patron.replaceChildren(crear("p", "vacio", "El libro no tiene una hoja con datos del patrón."));
      ui.seccionPatron.hidden = false;
      return;
    }

    // Agrupa los campos como en el Excel (DOMICILIO, ACTA CONSTITUTIVA…).
    const grupos = new Map();
    for (const campo of patron.campos) {
      const nombre = campo.grupo || "Datos generales";
      if (!grupos.has(nombre)) grupos.set(nombre, []);
      grupos.get(nombre).push(campo);
    }

    const bloques = [];
    for (const [nombre, campos] of grupos) {
      const bloque = crear("div", "grupo-datos");
      const lista = crear("dl", "datos datos--auto");
      for (const campo of campos) lista.append(crearDato(campo.encabezado, campo.valor));
      bloque.append(crear("p", "subtitulo-panel", nombre), lista);
      bloques.push(bloque);
    }

    ui.notaPatron.textContent = patron.encontrado
      ? `Hoja «${patron.hoja}» · fila ${patron.fila}`
      : `Hoja «${patron.hoja}» · sin datos capturados`;
    ui.patron.replaceChildren(...bloques);
    ui.seccionPatron.hidden = false;
  }

  function pintarTrabajadores(datos) {
    const { columnas, registros } = datos;
    const hayGrupos = columnas.some((c) => c.grupo);

    const th = (texto, clase = "", filas = 1, columnasOcupa = 1) => {
      const celda = crear("th", clase, texto);
      celda.scope = columnasOcupa > 1 ? "colgroup" : "col";
      if (filas > 1) celda.rowSpan = filas;
      if (columnasOcupa > 1) celda.colSpan = columnasOcupa;
      return celda;
    };

    // Encabezado en dos niveles, igual que en el Excel.
    const filaGrupos = crear("tr");
    const filaEncabezados = crear("tr");
    const niveles = hayGrupos ? 2 : 1;
    filaGrupos.append(th("Fila", "", niveles), th("Estado", "", niveles), th("Perfil", "", niveles));
    for (let i = 0; i < columnas.length;) {
      const col = columnas[i];
      if (!col.grupo) {
        filaGrupos.append(th(col.encabezado, "", niveles));
        i += 1;
        continue;
      }
      let n = 1;
      while (i + n < columnas.length && columnas[i + n].grupo === col.grupo) n += 1;
      filaGrupos.append(th(col.grupo, "tabla__grupo", 1, n));
      for (let k = 0; k < n; k += 1) filaEncabezados.append(th(columnas[i + k].encabezado));
      i += n;
    }
    const thead = crear("thead");
    thead.append(filaGrupos);
    if (hayGrupos) thead.append(filaEncabezados);

    const tbody = crear("tbody");
    if (!registros.length) {
      const fila = crear("tr");
      const celda = crear("td", "tabla__vacia");
      celda.append(crear("span", "", "No hay trabajadores capturados en la hoja."));
      celda.colSpan = columnas.length + 3;
      fila.append(celda);
      tbody.append(fila);
    }
    for (const reg of registros) {
      const fila = crear("tr");
      const pendientes = reg.faltantes.length;
      const pastilla = crear(
        "span",
        `pastilla pastilla--${pendientes ? "pendiente" : "correcto"}`,
        pendientes ? `${pendientes} pendiente${pendientes === 1 ? "" : "s"}` : "Completo",
      );
      const celdaEstado = crear("td");
      celdaEstado.append(pastilla);
      const celdaPerfil = crear("td");
      celdaPerfil.append(pastillaPerfil(reg.perfil));
      fila.append(crear("td", "tabla__fila", String(reg.fila)), celdaEstado, celdaPerfil);

      for (const col of columnas) {
        const valor = reg.valores[col.clave];
        const celda = crear("td", valor ? "" : "tabla__faltante", valor || FALTANTE);
        if (valor) celda.title = valor;
        fila.append(celda);
      }
      tbody.append(fila);
    }

    const tabla = crear("table", "tabla");
    tabla.append(thead, tbody);
    ui.trabajadores.replaceChildren(tabla);
    ui.notaTrabajadores.textContent =
      `Hoja «${datos.hoja}» · ${columnas.length} columnas · ${registros.length} trabajador${registros.length === 1 ? "" : "es"}`;
    ui.seccionTrabajadores.hidden = false;
  }

  /* ---------- Eventos ---------- */

  ui.plantilla.addEventListener("change", alCambiarPlantilla);
  ui.modoFirma.addEventListener("change", () => {
    ui.campoFechaFirma.hidden = ui.modoFirma.value !== "fecha";
  });
  ui.subirPlantilla.addEventListener("click", () => entradaPlantilla.click());
  document.addEventListener("click", (e) => {
    const boton = e.target.closest(".boton-copiar");
    if (boton) copiarMarcador(boton);
  });
  ui.entrada.addEventListener("change", () => seleccionarArchivo(ui.entrada.files[0]));
  ui.quitar.addEventListener("click", quitarArchivo);
  ui.validar.addEventListener("click", validar);
  ui.generar.addEventListener("click", generar);
  ui.descargarZip.addEventListener("click", descargarLote);
  ui.abrir.addEventListener("click", abrirCarpeta);

  ["dragenter", "dragover"].forEach((evento) =>
    ui.zona.addEventListener(evento, (e) => {
      e.preventDefault();
      ui.zona.classList.add("arrastrando");
    })
  );
  ["dragleave", "drop"].forEach((evento) =>
    ui.zona.addEventListener(evento, (e) => {
      e.preventDefault();
      ui.zona.classList.remove("arrastrando");
    })
  );
  ui.zona.addEventListener("drop", (e) => seleccionarArchivo(e.dataTransfer.files[0]));

  // Evita que el navegador abra el archivo si se suelta fuera de la zona de carga.
  ["dragover", "drop"].forEach((evento) => window.addEventListener(evento, (e) => e.preventDefault()));

  alCambiarPlantilla();
  cargarPlantillas();
})();
