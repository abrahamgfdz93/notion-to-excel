# 📊 notion-to-excel

Compila una parrilla de contenido exportada de Notion a un Excel (`.xlsx`) con una hoja **TODO** y una **pestaña por tenant/marca**, filtrando por mes. Skill para [Claude Code](https://docs.claude.com/claude-code).

## ✨ Qué hace

- Toma **automáticamente el `.zip` más reciente** que dejes en la carpeta `notion/` (al subir un export nuevo, lo usa solo).
- Maneja el **zip-dentro-de-zip** típico de Notion y detecta el CSV principal por sus columnas (no por nombre, que Notion genera con hashes).
- **Filtra por mes** (`--mes Julio`, acepta también `Jul`; sin argumento incluye todos los meses).
- Genera un Excel con hoja **TODO** + una **pestaña por cada tenant** presente, con estilos (encabezado, filtros, zebra, ajuste de texto).
- Marca las filas **sin tenant** en una hoja `_SIN-TENANT` como aviso de calidad de datos.
- Guarda con consecutivo (`_01`, `_02`...) para no pisar archivos previos.

## 🎯 Casos de uso

- Pasar el calendario de contenido mensual de Notion a un Excel compartible en Drive/Sheets.
- Repartir la parrilla por marca/cuenta (una pestaña por tenant) sin trabajo manual.
- Generar un snapshot fechado del plan de contenido de un mes.

## 📋 Requisitos

- [Claude Code](https://docs.claude.com/claude-code)
- Python 3.8+
- `openpyxl` (lo instala `install.sh` en un venv)
- Una base de datos de Notion con columnas `Tenant` y `Mes` (Notion guarda el mes como "Jul 26").

## 🚀 Instalación

```bash
git clone https://github.com/abrahamgfdz93/notion-to-excel.git
cd notion-to-excel
./install.sh
```

El instalador te **preguntará cuál es tu carpeta principal de Claude** (la que contiene tu carpeta `projects`, ya sea en Desktop, Documents o donde la tengas) e instalará todo ahí:

```
<tu-carpeta>/projects/notion-to-excel/
   ├── notion_to_excel.py      ← el script
   ├── notion/                 ← deja aquí tu export .zip de Notion
   └── output/                 ← aquí aparecen los Excel
```

La skill se registra dentro de **tu propia carpeta** (`<tu-carpeta>/.claude/skills/`), no en `~/.claude`. Por eso `/notion-to-excel` funciona cuando abres Claude Code en tu carpeta principal.

## 💻 Uso

1. Exporta tu base de datos de Notion (**Export → Markdown & CSV**) y deja el `.zip` en la carpeta `notion/` de la instalación.
2. Abre Claude Code en tu carpeta principal y escribe:

```
/notion-to-excel
```

O directo por terminal:

```bash
python3 notion_to_excel.py --mes Julio
```

El Excel queda en `output/parrilla_<Mes>_<NN>.xlsx`.

## ❓ FAQ

**¿Mi contenido de Notion se sube al repo?**
No. Las carpetas `notion/` (exports) y `output/` (Excel generados) están en `.gitignore`. Solo se versiona la herramienta.

**¿Y si mis columnas de Notion se llaman distinto?**
El script detecta el CSV por la columna `Dirección de Arte` y ordena por `COLS_BASE`. Ajusta esos valores en `notion_to_excel.py` para tu esquema.

## 📄 Licencia

[MIT](LICENSE) — úsalo libremente, modifícalo, distribúyelo.

---

Hecho con ☕ por [@abrahamgfdz93](https://github.com/abrahamgfdz93)
