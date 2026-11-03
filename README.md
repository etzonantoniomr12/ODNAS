# ODNA's

PWA móvil para convertir Excel de Drive, revisar ítems ODNA, aplicar una actividad ofertada o medida cautelar y generar PDF.

## Archivos

- `index.html`, `styles.css`, `app.js`: aplicación web estática.
- `manifest.json`, `sw.js`: instalación PWA y caché.
- `apps-script/Code.gs`: backend de Google Apps Script.
- `apps-script/appsscript.json`: manifiesto y permisos del backend.

## Instalación del backend

1. Crear un proyecto Apps Script independiente en la cuenta propietaria.
2. Copiar `apps-script/Code.gs` y reemplazar el manifiesto por `apps-script/appsscript.json`.
3. Implementar como **Aplicación web**:
   - Ejecutar como: el propietario.
   - Acceso: usuarios autorizados según la política de la cuenta.
4. Copiar la URL terminada en `/exec` dentro de `config.js`.
5. Publicar los archivos web en GitHub Pages u otro alojamiento HTTPS.

## Flujo

1. El usuario selecciona la carpeta origen y la carpeta destino exacta dentro de la raíz autorizada.
2. Registra todos los `.xlsx` de la carpeta origen y sus subcarpetas.
3. Convierte un archivo por petición usando la copia/conversión interna de Drive v2, evitando `getBlob()`.
4. Procesa únicamente la hoja principal identificada por `ITEM` y `DEFECTO` en la fila 7.
5. Lee las actividades desde la hoja `ACT`, rango `C2:C52`, del catálogo central configurado.
6. Aplica el texto rojo dentro de `B` o pinta `A:F` en rojo para medida cautelar.
7. Genera un PDF con todas las hojas visibles.

## Fotografías incrustadas

El backend lee parcialmente el contenedor ZIP del `.xlsx`, sin cargar el archivo completo en memoria. Extrae las fotografías asociadas a las filas `5, 9, 13...`, primero izquierda y luego derecha, y las guarda en una subcarpeta privada del proyecto. Este flujo fue validado con un Excel de 80.4 MB.

## Publicación web

La interfaz es estática y puede publicarse con GitHub Pages. Debe mantenerse la URL `/exec` vigente de Apps Script dentro de `config.js`. La instalación PWA usa el nombre **ODNA's** y los iconos ubicados en `assets/`.
