# Excel → PowerPoint | Generador de “Ficha N°2” (VBA)

Macro en VBA que genera automáticamente presentaciones PowerPoint (“Ficha N°2”) a partir de una tabla en Excel, reemplazando placeholders en una plantilla PPTX y guardando un archivo por cada fila.

# ¿Qué hace?
- Lee datos desde la hoja **"Ficha 2"**, tabla **"TFicha"**.
- Recorre filas desde `START_ROW` (configurable).
- Si la columna **Estado** es `"OK"`, **salta** esa fila.
- Si **N°** está vacío, **detiene** el proceso.
- Abre una **plantilla PPTX**, reemplaza placeholders (ej: `{N°}`, `{Problemática}`) y guarda el archivo como:
  - `Ficha N°2 - {N°} - {Nombre Iniciativa}.pptx`
- Marca **Estado = OK** al finalizar cada fila procesada.

## Requisitos
- Windows + Microsoft Excel de escritorio con VBA habilitado.
- Microsoft PowerPoint instalado.
- Un archivo Excel con:
  - Hoja: **Ficha 2**
  - Tabla: **TFicha** (encabezados en A1:O1)
  - Columna: **Estado** (para controlar “OK”)
- Una plantilla PowerPoint `.pptx` con placeholders.

## Placeholders soportados (según plantilla)
Ejemplos típicos:
- `{N°}`
- `{Nombre Iniciativa}`
- `{Líder de Idea}`
- `{Cargo líder}` 
- `{Problemática}`
- `{Solución y Beneficios}`
- `{¿Cómo se aborda el problema hoy en día?}` 
…y otros definidos en el diccionario de alias del código.

## Configuración rápida
En el módulo VBA, edita el bloque **CONFIG**:

- `START_ROW`: fila desde donde comenzar a procesar.
- `TEMPLATE_PPT_PATH`: ruta completa de la plantilla PPTX.
- `OUTPUT_FOLDER`: carpeta de salida para las fichas generadas.
- `SOURCE_WB_PATH`: (opcional) ruta del Excel fuente si no es el mismo archivo.

## Uso
1. Abre el archivo `.xlsm`.
2. Habilita macros (y agrega la carpeta como “Ubicación de confianza” si aplica).
3. Ejecuta `Generar_Fichas_PPT_Ficha2`.
4. Revisa la carpeta de salida y confirma que **Estado** quedó en `"OK"`.

## Reprocesar una fila
Si deseas regenerar una ficha:
1. Borra el `"OK"` en la columna **Estado** para esa fila.
2. Ejecuta nuevamente la macro.
3. Se volverá a generar y sobrescribirá el archivo.

## Notas
- Los nombres de archivo se limpian con `SafeFileName` para evitar caracteres inválidos en Windows.
- El reemplazo usa `TextRange.Replace` (incluye texto normal, tablas y shapes agrupados).

## Roadmap (opcional)
- Soporte para imagen/ilustración (insertar desde ruta).
- Log de ejecución (errores por fila, tiempo, etc.).
- Modo “solo pendientes” desde el primer NO OK (automático).

## Ejemplo de plantilla PPT