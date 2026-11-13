#!/bin/bash

# Al primo avvio, rilancia sé stesso sotto systemd-inhibit (blocca idle e sospensione)
if [ -z "$INHIBITED" ]; then
    export INHIBITED=1
    if [ -f /.flatpak-info ] && command -v flatpak-spawn >/dev/null 2>&1; then
        # VS Code Flatpak: il blocco va chiesto al sistema host, non alla sandbox
        flatpak-spawn --host --watch-bus systemd-inhibit --what=idle:sleep \
            --who="run_orari" --why="Esecuzione script Python in corso" sleep infinity &
        INH_PID=$!
        trap 'kill $INH_PID 2>/dev/null' EXIT
        sleep 1
        if ! kill -0 "$INH_PID" 2>/dev/null; then
            echo "ATTENZIONE: blocco sospensione non attivo (permessi Flatpak?)." >&2
        fi
    else
        INHIBIT_BIN="$(command -v systemd-inhibit || true)"
        [ -z "$INHIBIT_BIN" ] && [ -x /usr/bin/systemd-inhibit ] && INHIBIT_BIN=/usr/bin/systemd-inhibit
        if [ -n "$INHIBIT_BIN" ]; then
            exec "$INHIBIT_BIN" --what=idle:sleep --who="run_orari" \
                --why="Esecuzione script Python in corso" "$0" "$@"
        else
            echo "ATTENZIONE: systemd-inhibit non trovato, il PC potrebbe sospendersi." >&2
        fi
    fi
fi

cd "$(dirname "$0")" || exit 1

for x in 2 3 4; do
    echo "[$(date '+%F %T')] Avvio main_orari_$x.py"
    out="$(mktemp)"
    python -u "main_orari_$x.py" 2>&1 | tee "$out"
    rc=${PIPESTATUS[0]}
    obj="$(grep -oP '>>> Miglior risultato: objective = \K[^ ]+' "$out" | tail -1)"
    echo "main_orari_$x.py: ${obj:-N/D}" >> results.txt
    rm -f "$out"
    echo "[$(date '+%F %T')] Terminato main_orari_$x.py (exit code: $rc)"
done

echo "[$(date '+%F %T')] Tutti gli script completati."