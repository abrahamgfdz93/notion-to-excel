---
name: notion-to-excel
description: Compila una parrilla de contenido exportada de Notion a un Excel (.xlsx) con una hoja TODO y una pestaña por tenant/marca, filtrando por mes. Lee el export .zip más reciente de la carpeta notion/. Usar cuando se pida convertir/compilar/exportar contenido de Notion a Excel, armar la parrilla de un mes por tenant, o pasar un calendario de contenido a Sheets/Drive. Trigger con /notion-to-excel.
allowed-tools: Read, Write, Edit, Bash, AskUserQuestion
---

# Notion → Excel — Parrilla de contenido por tenant

Compila el export de una base de datos de Notion (una parrilla / calendario de contenido) a un
`.xlsx` listo para Drive/Sheets, dividido en una hoja **TODO** + una **pestaña por tenant** (marca),
filtrado por el mes que pida el usuario.

## Cómo funciona
- Toma **automáticamente el `.zip` más reciente** de la carpeta `notion/` (junto al script).
  Al subir un export nuevo, lo usa solo — sin tocar código.
- Maneja el zip-dentro-de-zip típico de Notion y detecta el CSV principal por sus columnas.
- Filtra por mes con `--mes` (acepta "Julio" o "Jul"; sin `--mes` incluye todos los meses).
- Genera: hoja **TODO** + una hoja por cada tenant presente + hoja **_SIN-TENANT** si hay filas
  sin tenant asignado (aviso de calidad de datos).
- Guarda con consecutivo (`_01`, `_02`...) en `output/` para no pisar archivos previos.

## Requisitos del export de Notion
La base de datos de Notion debe tener (como mínimo) una columna **`Tenant`** (la marca) y una
columna **`Mes`** (Notion la guarda como "Jul 26", "Jun 26", etc.). El script detecta el CSV
principal por la columna `Dirección de Arte`; si tu base usa otros nombres, ajusta `COLS_BASE`
y la detección en `notion_to_excel.py`.

## Flujo a seguir
1. **Confirmar el mes.** Si el usuario no lo dijo, pregúntale (AskUserQuestion) qué mes compilar.
2. **Verificar la fuente.** Lista los `.zip` en `notion/` y confirma cuál es el más reciente
   (`ls -t notion/*.zip | head -1`). Si el usuario menciona una versión que NO es la más reciente,
   adviérteselo con datos antes de continuar.
3. **Generar:**
   ```bash
   cd ~/.claude/skills/notion-to-excel
   python3 notion_to_excel.py --mes "<Mes>"
   ```
4. **Reportar:** ruta del archivo, nº de publicaciones, desglose por tenant y cualquier aviso
   (filas sin tenant → hoja _SIN-TENANT, sugerir corregir en Notion).

## Notas
- Cualquier tenant nuevo aparece automáticamente con su clave tal cual.
- Si el export no tiene publicaciones del mes pedido, el script lo dice y no genera archivo.
- Requiere `openpyxl` (`pip install openpyxl`). El `install.sh` lo instala en un venv.
- **Privacidad:** la carpeta `notion/` y la salida `output/` están en `.gitignore`. Nunca subas
  exports reales ni Excel generados a un repo público — contienen tu contenido.
