# Client asset handover template (Excel + CSV)

A free spreadsheet for recording who owns each client's domain, hosting, email and
accounts, and how they get back into them without you. No email address, no signup,
no thanks-for-subscribing page.

**[Download the Excel template (.xlsx)](client-asset-handover-template.xlsx)** —
the readiness column is a working formula, with a percentage per client beside the table.

**[Download the CSV](client-asset-handover-template.csv)** — same eleven columns as
plain text, for Google Sheets, Numbers, or an import into Notion.

Both ship with six example rows so the shape is obvious. Delete them and start with the
client you would least like to lose.

## One rule

**Never put a password in this file.** Not in a hidden column, not in a note, not
temporarily.

A file with live credentials gets forwarded, and it sits in inboxes that outlive the
relationship. It also goes stale the moment anyone rotates anything, so it ends up both
dangerous and wrong. Record *where the credential lives* — "client 1Password vault",
"reset via admin@theirbusiness.com" — and let the password manager do its job.

## What's in the repository

| File | |
|---|---|
| `client-asset-handover-template.xlsx` | Two sheets: **Handover** (the table) and **How to use** (a column-by-column reference). Frozen header, filters, and conditional formatting on the readiness column. |
| `client-asset-handover-template.csv` | The same eleven columns and six rows in plain text. |
| `build.py` | The script that generates the workbook. |

## The columns

One row per asset, not per client. A client with five accounts has five rows.

| Column | Example |
|---|---|
| Client | `Northline Cafe` |
| Asset | `Domain`, `Hosting`, `DNS`, `Email`, `Analytics`, `Repository` |
| Provider | `Namecheap`, `Vercel`, `Cloudflare`, `Google Workspace` |
| Owner email | `admin@northlinecafe.com` — the address the account belongs to |
| Ownership | `client owned` · `needs transfer` · `freelancer managed` |
| Login method | `email + password`, `Google sign-in`, `SSO` |
| Renews | `2027-03-12`. Blank means you do not know, which is worth knowing |
| Cost | `18`. Annual, in whatever currency you bill |
| Where the password lives | `Client 1Password vault` — a pointer, never the password |
| Ready? | Calculated. Do not type in this column |
| Notes | `Registered on my card in 2019. They have no registrar account yet.` |

Owner email is the column that decides everything, because every provider resolves a
lost account by sending a reset link to the address on the account. If you fill in one
column this month, fill in that one.

The reasoning behind each of the eleven, and the three ways a spreadsheet eventually
stops working, is in the guide:
**[A free client asset handover template](https://geteverkey.com/blog/client-asset-handover-template)**.

## The readiness formula

The `Ready?` column is already wired up in the Excel file:

```excel
=IF(AND(E2="client owned", D2<>"", F2<>""), "yes", "no")
```

An asset is ready when the client owns it, the owning address is written down, and the
way in is written down. Miss any one of the three and the answer is `no`, because any
one of them missing is enough to strand somebody.

For a percentage per client:

```excel
=IFERROR(COUNTIFS(A:A,"Northline Cafe",J:J,"yes") / COUNTIF(A:A,"Northline Cafe"), "")
```

A client at 100% can be handed over this afternoon. A client at 40% is a project.

With the six example rows in place, the readiness column reads
`yes / yes / no / no / no / no`, and the summary reads **Northline Cafe 50%**,
**Acme Dental 0%**. If you open the file and see something else, something has shifted.

> **The `Ready?` cells look empty until your spreadsheet calculates them.** The file
> stores formulas without cached results. Excel, Google Sheets, Numbers and LibreOffice
> all calculate on open, so the values appear immediately. Plain-text previews that do
> not evaluate formulas will show the column blank.

## Rebuilding it

The workbook is generated, not hand-edited, so you can audit every cell and change the
example rows, the columns or the styling without unpacking a binary:

```bash
pip install -r requirements.txt
python3 build.py
```

Writing to a different path:

```bash
python3 build.py --out my-version.xlsx
```

`build.py` rebuilds `client-asset-handover-template.xlsx` cell for cell — the shipped
file and a fresh build are identical across all 78 populated cells and both sheets.

If you change the grid, keep the header on row 1 and the eleven columns in order. The
published guide quotes the cell references `E2` / `D2` / `F2` by name, so moving them
makes that guide wrong.

The CSV is maintained alongside the workbook rather than generated from it; the two
differ deliberately in one place, where the CSV writes `18 USD` and the workbook writes
`18` with a currency-less number format.

## Moving an asset to the client

Filling in `needs transfer` is the easy half. The provider-specific part — who has to
own the account, what the transfer locks are, and what breaks when the account is still
yours — is written up per provider:

- [GoDaddy domains](https://geteverkey.com/transfer/godaddy/domain)
- [Namecheap domains](https://geteverkey.com/transfer/namecheap/domain)
- [Cloudflare DNS](https://geteverkey.com/transfer/cloudflare/dns)
- [Google Workspace email](https://geteverkey.com/transfer/google-workspace/email)
- [All providers](https://geteverkey.com/transfer)

## Where this came from

Built for [Everkey](https://geteverkey.com), which is the version of this that reminds
you before a renewal and produces a one-page recovery document for each client. It
stores no passwords either, for the same reason this file does not.

You do not need it to use the template. The spreadsheet is genuinely better than what
most freelancers have, which is nothing written down at all.

## Licence

[CC0 1.0 Universal](LICENSE) — public domain. Use it, sell work built on it, rebrand it,
put it in your own onboarding pack. No attribution required.
