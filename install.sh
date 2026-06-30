#!/bin/bash
# Instalador de la skill notion-to-excel.
# Pregunta tu carpeta principal de Claude (la que contiene 'projects') e instala
# TODO ahí: <base>/projects/notion-to-excel/ (script + notion/ + output/).
# La skill se registra dentro de TU carpeta principal (<base>/.claude/skills/),
# NO en ~/.claude — así nada queda enterrado en carpetas ocultas del sistema.

set -e

GREEN='\033[0;32m'; BLUE='\033[0;34m'; YELLOW='\033[1;33m'; RED='\033[0;31m'; NC='\033[0m'
SKILL_NAME="notion-to-excel"
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
echo -e "${BLUE}  Instalador: notion-to-excel${NC}"
echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
echo ""

# --- 1. Preguntar la carpeta principal de Claude ---
echo -e "${YELLOW}¿Cuál es tu carpeta principal de Claude (donde tienes/quieres tu carpeta 'projects')?${NC}"
echo "   Ejemplos:  ~/Desktop/CLAUDE   |   ~/Documents/CLAUDE   |   \"~/Desktop/MI CLAUDE\""
read -r -p "Ruta [Enter para usar ~/Desktop/CLAUDE]: " BASE_INPUT
BASE_INPUT="${BASE_INPUT:-$HOME/Desktop/CLAUDE}"
# Expandir ~ al inicio
BASE="${BASE_INPUT/#\~/$HOME}"

# Si la ruta termina en /projects, subir un nivel (el usuario apuntó a projects directo)
case "$(basename "$BASE")" in projects) BASE="$(dirname "$BASE")";; esac

mkdir -p "$BASE/projects"
DEST="$BASE/projects/$SKILL_NAME"
echo ""
echo -e "Se instalará en: ${GREEN}$DEST${NC}"
echo ""

# --- 2. Copiar la herramienta a projects/ (sin install.sh, .git, ni datos) ---
mkdir -p "$DEST/notion" "$DEST/output"
for f in notion_to_excel.py SKILL.md requirements.txt README.md LICENSE .gitignore; do
    [ -f "$SCRIPT_DIR/$f" ] && cp "$SCRIPT_DIR/$f" "$DEST/$f"
done
[ -f "$SCRIPT_DIR/notion/README.md" ] && cp "$SCRIPT_DIR/notion/README.md" "$DEST/notion/README.md"
echo -e "${GREEN}✓${NC} Herramienta copiada a projects/"

# --- 3. Dependencias de Python en un venv local ---
if [ -f "$DEST/requirements.txt" ]; then
    echo -e "${YELLOW}Instalando dependencias (openpyxl)...${NC}"
    [ -d "$DEST/venv" ] || python3 -m venv "$DEST/venv"
    # shellcheck disable=SC1091
    source "$DEST/venv/bin/activate"
    pip install --upgrade pip --quiet
    pip install -r "$DEST/requirements.txt" --quiet
    deactivate
    echo -e "${GREEN}✓${NC} Dependencias instaladas en $DEST/venv"
fi

# --- 4. Registrar la skill DENTRO de la carpeta principal (no en ~/.claude) ---
mkdir -p "$BASE/.claude/skills"
ln -sfn "$DEST" "$BASE/.claude/skills/$SKILL_NAME"
echo -e "${GREEN}✓${NC} Skill registrada en $BASE/.claude/skills/$SKILL_NAME"

echo ""
echo -e "${GREEN}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
echo -e "${GREEN}  ✅ Instalación completa${NC}"
echo -e "${GREEN}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
echo ""
echo "Cómo usarla:"
echo "  1. Deja tu export .zip de Notion en:  $DEST/notion/"
echo "  2. Abre Claude Code en tu carpeta principal:  $BASE"
echo "  3. Escribe:  /$SKILL_NAME"
echo ""
echo "Los Excel se guardan en:  $DEST/output/"
