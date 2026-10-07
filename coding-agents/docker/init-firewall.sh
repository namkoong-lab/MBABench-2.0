#!/bin/bash
# Default-deny egress firewall. Allows DNS + established traffic + HTTPS to the
# domains in $ALLOWED_DOMAINS (comma-separated). Anything else is dropped —
# including all plain HTTP. Must run as root with CAP_NET_ADMIN.
#
# FAIL-SAFE: any error exits non-zero; the entrypoint then aborts the attempt
# rather than running the agent with an open network.
set -euo pipefail

if [ -z "${ALLOWED_DOMAINS:-}" ]; then
    echo "init-firewall: ALLOWED_DOMAINS is empty — refusing to run open" >&2
    exit 1
fi

ipset create allowed-hosts hash:ip -exist
ipset flush allowed-hosts

IFS=',' read -ra DOMAINS <<< "$ALLOWED_DOMAINS"
for domain in "${DOMAINS[@]}"; do
    domain="$(echo "$domain" | xargs)"
    [ -z "$domain" ] && continue
    ips="$(dig +short A "$domain" | grep -E '^[0-9.]+$' || true)"
    if [ -z "$ips" ]; then
        echo "init-firewall: could not resolve $domain" >&2
        exit 1
    fi
    for ip in $ips; do
        ipset add allowed-hosts "$ip" -exist
    done
    echo "init-firewall: allowed $domain -> $(echo $ips | tr '\n' ' ')"
done

iptables -F OUTPUT
iptables -P OUTPUT DROP
iptables -A OUTPUT -o lo -j ACCEPT
iptables -A OUTPUT -m state --state ESTABLISHED,RELATED -j ACCEPT
iptables -A OUTPUT -p udp --dport 53 -j ACCEPT
iptables -A OUTPUT -p tcp --dport 53 -j ACCEPT
iptables -A OUTPUT -p tcp --dport 443 -m set --match-set allowed-hosts dst -j ACCEPT

echo "init-firewall: default-deny egress active ($(ipset list allowed-hosts | grep -c '^[0-9]') hosts allowed)"
