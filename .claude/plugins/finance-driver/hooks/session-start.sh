#!/bin/bash
# DRIVER Session Start Hook
# Injects the using-driver skill at the start of every session

# Derive plugin root from this script's location if not set
if [ -z "$CLAUDE_PLUGIN_ROOT" ]; then
    SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
    CLAUDE_PLUGIN_ROOT="$(dirname "$SCRIPT_DIR")"
fi

SKILL_PATH="${CLAUDE_PLUGIN_ROOT}/skills/using-driver/SKILL.md"

if [ -f "$SKILL_PATH" ]; then
    echo "<EXTREMELY-IMPORTANT>"
    cat "$SKILL_PATH"
    echo "</EXTREMELY-IMPORTANT>"
fi
