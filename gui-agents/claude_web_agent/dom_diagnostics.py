"""Self-documenting DOM diagnostics for the artifact-download step.

Whenever a provider's UI changes how it presents the finished .xlsx — file
card, icon-only download button, sandbox ``/mnt`` link, or a plain inline
``<a>`` link — the download selectors can silently miss it and the run saves
nothing. This dumps, at the download decision point:

  (a) every download-ish control on the page, found *selector-independently*
      (any <a>/<button>/[download] whose href/aria/title/text hits a file
      keyword) — the part that survives UI drift; and
  (b) the last assistant message's outerHTML (capped) for full context.

So the next format drift shows up in the engine log as a clear ``[DOM-DIAG]``
block instead of a silent zero-file result. This is diagnostic only and must
never break a run: every path is wrapped and swallows its own errors.
"""

from __future__ import annotations

# Substrings that mark an element as a plausible download control / file
# reference. Matched case-insensitively against href/aria-label/title/
# download/text. Intentionally broad — this is a net, not a precise selector.
_FILE_KEYWORDS = [
    ".xlsx", ".xls", "sandbox:", "blob:", "/mnt/", "download-file",
    "wiggle/download", "spreadsheet", "excel", "download",
]

_ARTIFACT_DIAG_JS = r"""
(args) => {
  const KW = args.keywords;
  const cap = args.cap;
  const messageSelectors = args.messageSelectors;
  const low = (s) => (s || '').toLowerCase();
  const hit = (s) => { const t = low(s); return KW.some(k => t.includes(k)); };

  function chain(el) {
    const parts = [];
    let cur = el;
    while (cur && cur !== document.body && parts.length < 6) {
      let lab = cur.tagName.toLowerCase();
      const role = cur.getAttribute && cur.getAttribute('data-message-author-role');
      const tid = cur.getAttribute && cur.getAttribute('data-testid');
      if (role) lab += '[role=' + role + ']';
      if (tid) lab += '[testid=' + tid + ']';
      parts.push(lab);
      cur = cur.parentElement;
    }
    return parts.join('>');
  }

  // (a) selector-independent scan of candidate download controls.
  const controls = [];
  const nodes = Array.from(
    document.querySelectorAll('a, button, [download], [role=button]')
  );
  for (const el of nodes) {
    const href  = (el.getAttribute && (el.getAttribute('href') || '')) || '';
    const aria  = (el.getAttribute && (el.getAttribute('aria-label') || '')) || '';
    const title = (el.getAttribute && (el.getAttribute('title') || '')) || '';
    const dl    = (el.getAttribute && (el.getAttribute('download') || '')) || '';
    const text  = (el.textContent || '').trim().substring(0, 80);
    if (hit(href) || hit(aria) || hit(title) || hit(dl) || hit(text)) {
      controls.push({
        tag: el.tagName.toLowerCase(),
        href: href.substring(0, 300),
        download: dl.substring(0, 120),
        ariaLabel: aria.substring(0, 120),
        title: title.substring(0, 120),
        text: text,
        where: chain(el),
      });
      if (controls.length >= 40) break;
    }
  }

  // (b) last assistant message outerHTML (first selector that matches wins).
  let lastMsg = null, usedSel = null;
  for (const sel of messageSelectors) {
    let els;
    try { els = document.querySelectorAll(sel); } catch (e) { continue; }
    if (els && els.length) { lastMsg = els[els.length - 1]; usedSel = sel; break; }
  }
  const html = lastMsg ? lastMsg.outerHTML : '';
  return {
    url: location.href,
    controlCount: controls.length,
    controls: controls,
    messageSelector: usedSel,
    messageHtmlLen: html.length,
    messageHtml: html.substring(0, cap),
  };
}
"""


async def dump_final_message_dom(
    page, logger, provider, message_selectors, cap: int = 25000
) -> None:
    """Log a ``[DOM-DIAG]`` block describing how the finished file is presented.

    `page` is a Playwright page; `message_selectors` is an ordered list of CSS
    selectors for the assistant message container (most specific first). Any
    failure is logged and swallowed — this never propagates.
    """
    try:
        report = await page.evaluate(
            _ARTIFACT_DIAG_JS,
            {
                "keywords": _FILE_KEYWORDS,
                "cap": cap,
                "messageSelectors": message_selectors,
            },
        )
    except Exception as e:  # noqa: BLE001 — diagnostic must never raise
        logger.info(f"[DOM-DIAG] {provider}: capture failed ({type(e).__name__}: {e})")
        return

    try:
        controls = report.get("controls", []) or []
        logger.info(
            f"[DOM-DIAG] {provider}: {report.get('controlCount', 0)} candidate "
            f"download control(s) at {report.get('url', '')}; last message via "
            f"{report.get('messageSelector')!r} "
            f"(outerHTML {report.get('messageHtmlLen', 0)} chars)"
        )
        for i, c in enumerate(controls):
            logger.info(
                f"[DOM-DIAG] {provider} control[{i}] <{c.get('tag')}> "
                f"aria={c.get('ariaLabel')!r} download={c.get('download')!r} "
                f"href={c.get('href')!r} text={c.get('text')!r} @ {c.get('where')}"
            )
        html = report.get("messageHtml") or ""
        if html:
            logger.info(
                f"[DOM-DIAG] {provider} final-message outerHTML "
                f"(capped {len(html)} chars) >>>\n{html}\n"
                f"<<< [DOM-DIAG] end {provider}"
            )
    except Exception as e:  # noqa: BLE001
        logger.info(f"[DOM-DIAG] {provider}: log-format error ({type(e).__name__}: {e})")
