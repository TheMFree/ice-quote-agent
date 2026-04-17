# ICE Quote Agent

Autonomous email-to-proposal pipeline for ICE Contractors, Inc.

A team member emails raw quote info to a dedicated mailbox (text in the
body, PDFs, Word docs, or photos). The agent extracts the information
with Claude, fills the ICE proposal template, and emails the draft
`.docx` to the reviewer (Michael) for approval.

## Deploy

On the target VPS (as root), run:

```bash
curl -fsSL https://raw.githubusercontent.com/TheMFree/ice-quote-agent/main/deploy/bootstrap.sh | bash
```

Then edit `/opt/ice-quote-agent/.env` with real secrets and:

```bash
systemctl enable --now ice-quote-agent.service
```
