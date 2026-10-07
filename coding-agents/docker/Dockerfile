# The per-attempt sandbox image. Build it once, tagged as the run configs and
# coding_agent/config.py (SandboxConfig.image) expect — the tag is the CLI
# version pin every attempt row records (extra_configs.sandbox_image):
#
#   docker build -t mbabench-coding-agent:v2 coding-agents/docker/
#
# Pin the CLI versions for a reproducible image, e.g.
#   docker build --build-arg CLAUDE_CODE_VERSION=2.1.251 --build-arg CODEX_VERSION=0.155.1 \
#       -t mbabench-coding-agent:v2 coding-agents/docker/
# Retag any rebuild that changes contents; a run config's sandbox.image names
# the tag it runs under.
FROM node:22-bookworm-slim

ARG CLAUDE_CODE_VERSION=latest
ARG CODEX_VERSION=latest
ARG INCLUDE_LIBREOFFICE=true

RUN apt-get update && apt-get install -y --no-install-recommends \
        python3 python3-pip python3-venv \
        iptables ipset dnsutils ca-certificates procps \
    && if [ "$INCLUDE_LIBREOFFICE" = "true" ]; then \
        apt-get install -y --no-install-recommends libreoffice-calc; fi \
    && rm -rf /var/lib/apt/lists/*

RUN pip3 install --break-system-packages --no-cache-dir openpyxl pandas

RUN npm install -g "@anthropic-ai/claude-code@${CLAUDE_CODE_VERSION}" \
                   "@openai/codex@${CODEX_VERSION}" \
    && npm cache clean --force

RUN useradd -m -u 1001 -s /bin/bash agent

# --chmod: COPY otherwise keeps the host's file mode, and a 640 checkout left
# traj_relay.py unreadable by the unprivileged agent user that runs it.
COPY --chmod=755 init-firewall.sh /usr/local/bin/init-firewall.sh
COPY --chmod=755 entrypoint.sh /usr/local/bin/entrypoint.sh
COPY --chmod=755 traj_relay.py /usr/local/bin/traj_relay.py

WORKDIR /workspace
ENTRYPOINT ["/usr/local/bin/entrypoint.sh"]
