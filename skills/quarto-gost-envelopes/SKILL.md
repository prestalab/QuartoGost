---
name: quarto-gost-envelopes
description: Use when generating mailing envelopes in this repository from a TSV address list for dissertation abstract or document distribution scenarios.
---

# QuartoGost Envelopes

Use this skill when the user needs a printable DOCX with mailing pages built
from a TSV recipient list.

## Build command

`powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType envelopes -AddressList <file.tsv> -OutputDir <dir> -Name <name> -SenderName <sender> -SenderAddress <address>`

## Rules

1. The TSV file must have five tab-separated columns.
2. No header row is required.
3. Empty recipient cells are allowed.
4. Keep sender data outside the TSV and pass it through command parameters.

Read [tsv-format.md](references/tsv-format.md) before editing the mailing list.

