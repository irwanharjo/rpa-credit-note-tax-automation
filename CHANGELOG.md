# Changelog

The public `otomasi_cn_tax.py` is the July 2026 snapshot. The production version roughly
doubled in size; it stays private because it is bound to a real client's mailbox, sheets and
portal, but the lessons below are generic and reusable.

## 2026-08 → 2026-11 — Production hardening (summary, code not published)

Some of these live in the main robot, others in its companion email-triage robot that reads
client replies and decides what the pipeline should do.

- **Zero-VAT gate** — documents whose recalculated VAT is zero are routed to a separate path
  instead of producing an empty tax request.
- **Dedup within the same batch** — the request-sheet writer already skipped documents present
  in the sheet, but two different email threads about the same document in ONE run could both
  pass. The writer now also dedups against the rows it is about to write.
- **"Status update" emails are not requests** — replies that only say "being verified / will
  re-check" (no attachment, no request keywords) are now classified as *no action*, instead of
  re-sending a stale earlier result found in the same thread.
- **Already-answered guard** — if the newest attachment in a thread is older than our own last
  reply, the thread is held instead of answered twice.
- **Internal-only emails ignored** — a sent message counts as "already handled manually" only
  if it went to the client; alerts we sent to ourselves no longer count.
  Gotcha found on the way: Gmail's `to:` search operator also matches **Cc**, so the check reads
  the `To` header in code instead of relying on search operators.
- **Mass-mailing threads skipped** — Gmail search also matches email bodies, so a mass
  announcement mentioning many document numbers could be mistaken for a document's thread.
  Such candidates are skipped; replies target exactly one `threadId` with correct
  `In-Reply-To` / `References` headers.
- **API retry** for transient Google API errors, and a single, throttled worker for
  mailbox sweeps after parallel sweeps exhausted the per-user-per-minute quota used by the live
  robot.
- **Dispute ledger** — open disputes are tracked in one place so the robot doesn't act on a
  document that a human is still arguing about.
- **Buyer-specific VAT collection rule** — above a VAT threshold some buyers collect the VAT
  themselves, so no tax request is raised; one buyer entity is always treated as
  non-collecting, per the client's written instruction.

## 2026-07 — Initial public version

- Credit note pipeline: Excel input → portal → recalculation → Gmail thread reply → sheet log.
