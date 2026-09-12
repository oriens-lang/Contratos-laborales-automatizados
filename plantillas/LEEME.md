- `FORMATO CONTRATO TIEMPO INDETERMINADO (original).docx`: plantilla oficial tal como la entregó el despacho. **No se modifica.**
- `Contrato tiempo indeterminado (marcado).docx`: copia con marcadores `{{ campo }}` en los espacios variables. Es la que usa la aplicación.
- `Formato de captura de trabajadores.xlsx`: Excel que se envía al cliente (se descarga desde la pantalla).
- `marcar_plantilla.py`: crea la copia marcada a partir del original y verifica que el texto jurídico quedó idéntico. No sobrescribe una copia existente salvo con `--forzar`.
- `crear_formato_captura.py`: vuelve a crear el formato de captura.

## Cambiar el texto de una plantilla

1. En la pantalla, **Descargar para editar** (o abre el archivo marcado de esta carpeta en Word).
2. Cambia lo que necesites **sin tocar los marcadores** `{{ … }}`.
3. En la pantalla, **Subir plantilla**. La aplicación revisa los marcadores antes de aceptarla.

Una plantilla nueva (otro tipo de contrato) necesita marcadores; mientras no los tenga, la aplicación la rechaza.
