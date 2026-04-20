#!/usr/bin/env bash
# =============================================================================
# Excel Pipeline launcher for macOS / Linux
# Checks for uv, pulls latest logic from GitHub, syncs deps, starts Streamlit.
# =============================================================================
set -e

cd "$(dirname "$0")"

# --- Check for uv -----------------------------------------------------------
if ! command -v uv &> /dev/null; then
    echo "Installing uv (one-time)..."
    curl -LsSf https://astral.sh/uv/install.sh | sh
    # Source uv's env if available, else add to PATH manually
    if [ -f "$HOME/.local/bin/env" ]; then
        # shellcheck source=/dev/null
        source "$HOME/.local/bin/env"
    else
        export PATH="$HOME/.local/bin:$PATH"
    fi
fi

# --- Pull latest logic from GitHub -----------------------------------------
echo "Checking for updates..."
if ! git pull --ff-only; then
    echo "WARNING: Could not pull updates. Continuing with local version."
fi

# --- Sync dependencies -----------------------------------------------------
echo "Syncing dependencies..."
uv sync

# --- Check for .env --------------------------------------------------------
if [ ! -f .env ]; then
    cat <<EOF

==========================================================
  .env file not found.
  Copy .env.example to .env and add your ANTHROPIC_API_KEY.
  Get a key at https://console.anthropic.com
==========================================================

EOF
    exit 1
fi

# --- Launch ----------------------------------------------------------------
echo "Starting Excel Pipeline..."
echo "(Press Ctrl+C to stop the app.)"
uv run streamlit run app.py
