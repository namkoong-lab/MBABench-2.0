#!/bin/bash
# Container entrypoint: install the egress firewall as root, then drop to the
# unprivileged agent user and exec the agent command (passed as docker CMD).
# Exit 97 = sandbox setup failure (the runner classifies it as infra_failure).
set -uo pipefail

if ! /usr/local/bin/init-firewall.sh; then
    echo "entrypoint: firewall setup FAILED — aborting attempt (fail-safe)" >&2
    exit 97
fi

if [ "$#" -eq 0 ]; then
    echo "entrypoint: no agent command given" >&2
    exit 97
fi

chown agent:agent /workspace 2>/dev/null || true

if [ -n "${TRAJ_UPSTREAM:-}" ]; then
    mkdir -p /trajectory && chown agent:agent /trajectory
    runuser -u agent -- python3 /usr/local/bin/traj_relay.py >>/trajectory/relay.log 2>&1 &
    relay_up=0
    for i in $(seq 1 30); do
        if python3 -c "import socket; socket.create_connection(('127.0.0.1', ${TRAJ_PORT:-9877}), 1)" 2>/dev/null; then
            relay_up=1; break
        fi
        sleep 0.3
    done
    # The agent is pointed at the relay; without it every API call fails and
    # the run would burn a trial as agent_failure. Abort as infra instead.
    if [ "$relay_up" -ne 1 ]; then
        echo "entrypoint: trajectory relay did not come up — aborting attempt (fail-safe)" >&2
        cat /trajectory/relay.log >&2 2>/dev/null || true
        exit 97
    fi
fi

exec runuser -u agent -- "$@"
