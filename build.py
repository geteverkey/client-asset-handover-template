"""Generate the Excel handover template.

The CSV beside it is the canonical plain-text version; this produces the .xlsx,
which is the format people actually search for.

Both files must keep the same eleven columns, the same six example rows, and the
header on row 1 — the published guide at
geteverkey.com/blog/client-asset-handover-template quotes the cell references
E2/D2/F2 by name, so shifting the grid would make that guide wrong.

openpyxl writes formulas without cached results, so the Ready? column reads blank
until a spreadsheet application opens the file and calculates. Excel, Google
Sheets, Numbers and LibreOffice all do that on open.

Requires openpyxl:  pip install -r requirements.txt
Run:                python3 build.py [--out FILE]
"""
import argparse
from datetime import date

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = "client-asset-handover-template.xlsx"

HEADERS = [
    "Client", "Asset", "Provider", "Owner email", "Ownership", "Login method",
    "Renews", "Cost", "Where the password lives", "Ready?", "Notes",
]

# Six example rows. Each one demonstrates a different way an asset fails the
# readiness test, which is the whole point of shipping examples at all.
ROWS = [
    ["Northline Cafe", "Domain", "Namecheap", "admin@northlinecafe.com", "client owned",
     "email + password", date(2027, 3, 12), 18, "Client 1Password vault", None,
     "Transferred 2026-02"],
    ["Northline Cafe", "Hosting", "Vercel", "admin@northlinecafe.com", "client owned",
     "Google sign-in", date(2027, 4, 18), 240, "Client 1Password vault", None, ""],
    ["Northline Cafe", "DNS", "Cloudflare", "me@myagency.com", "needs transfer",
     "email + password", None, 0, "My 1Password vault", None,
     "Still on my account. Move at next renewal."],
    ["Northline Cafe", "Email", "Google Workspace", "admin@northlinecafe.com", "client owned",
     "", date(2027, 1, 9), 432, "Client 1Password vault", None,
     "Owner is right but I never wrote down how they sign in."],
    ["Acme Dental", "Domain", "GoDaddy", "me@myagency.com", "needs transfer",
     "email + password", date(2026, 11, 30), 22, "My 1Password vault", None,
     "Registered on my card in 2019. They have no registrar account yet."],
    ["Acme Dental", "Analytics", "Google Analytics", "", "freelancer managed",
     "Google sign-in", None, 0, "My Google account", None,
     "Attached to my personal Google login. Leaves with me."],
]

WIDTHS = [16, 12, 18, 26, 18, 18, 12, 9, 26, 9, 46]

ARIAL = "Arial"
INK = "1F2421"
RULE = "D8DCD6"
HEAD_FILL = "EEF1EC"
YES_FILL = "E4F0E4"
NO_FILL = "F7E7E4"

thin = Side(style="thin", color=RULE)
box = Border(bottom=thin)


def build(out: str = OUT) -> None:
    wb = Workbook()

    ws = wb.active
    ws.title = "Handover"

    for i, head in enumerate(HEADERS, start=1):
        c = ws.cell(row=1, column=i, value=head)
        c.font = Font(name=ARIAL, size=10, bold=True, color=INK)
        c.fill = PatternFill("solid", fgColor=HEAD_FILL)
        c.border = box
        c.alignment = Alignment(vertical="center")

    for r, row in enumerate(ROWS, start=2):
        for i, value in enumerate(row, start=1):
            c = ws.cell(row=r, column=i)
            if i == 10:
                # The readiness test, exactly as published in the article.
                c.value = f'=IF(AND(E{r}="client owned",D{r}<>"",F{r}<>""),"yes","no")'
                c.alignment = Alignment(horizontal="center")
            else:
                c.value = value
            c.font = Font(name=ARIAL, size=10, color=INK)
            if i == 7:
                c.number_format = "yyyy-mm-dd"
            if i == 8:
                c.number_format = "#,##0;(#,##0);-"
            if i == 11:
                c.alignment = Alignment(vertical="top", wrap_text=True)

    for i, w in enumerate(WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:K{1 + len(ROWS)}"

    ws.conditional_formatting.add(
        "J2:J500",
        CellIsRule(operator="equal", formula=['"yes"'],
                   fill=PatternFill("solid", fgColor=YES_FILL)),
    )
    ws.conditional_formatting.add(
        "J2:J500",
        CellIsRule(operator="equal", formula=['"no"'],
                   fill=PatternFill("solid", fgColor=NO_FILL)),
    )

    # Readiness summary, kept off the A:K grid so the article's cell references hold.
    ws["M1"] = "Readiness by client"
    ws["M1"].font = Font(name=ARIAL, size=10, bold=True, color=INK)
    for offset, client in enumerate(["Northline Cafe", "Acme Dental"]):
        r = 2 + offset
        ws[f"M{r}"] = client
        ws[f"N{r}"] = f'=IFERROR(COUNTIFS(A:A,M{r},J:J,"yes")/COUNTIF(A:A,M{r}),"")'
        ws[f"N{r}"].number_format = "0%"
        for col in ("M", "N"):
            ws[f"{col}{r}"].font = Font(name=ARIAL, size=10, color=INK)
    ws["M5"] = "Type any client name in column M to get their percentage."
    ws["M5"].font = Font(name=ARIAL, size=9, italic=True, color="6B7280")
    ws.column_dimensions["M"].width = 22
    ws.column_dimensions["N"].width = 10

    legend = wb.create_sheet("How to use")
    legend.column_dimensions["A"].width = 30
    legend.column_dimensions["B"].width = 92

    lines: list[tuple[str, str]] = [
        ("Client asset handover template", ""),
        ("", "Delete the six example rows and start with the client you would least like to lose."),
        ("", ""),
        ("Rule", "Never put a password in this file. Record where the credential lives, not the credential."),
        ("", ""),
        ("Client", "One row per asset, not per client. A client with five accounts has five rows."),
        ("Asset", "Domain, hosting, DNS, email, SaaS, analytics, repository."),
        ("Provider", "The company you would ring. Namecheap, Vercel, Cloudflare, Google Workspace."),
        ("Owner email", "The address the account belongs to. This is the column that decides everything."),
        ("Ownership", "client owned / needs transfer / freelancer managed. Intent is not ownership."),
        ("Login method", "Password, Google sign-in, SSO. Determines whether a reset is even possible."),
        ("Renews", "The date. Blank means you do not know, which is worth knowing."),
        ("Cost", "Annual cost, in whatever currency you bill. Unit is yours to keep consistent."),
        ("Where the password lives", "A pointer, never the password itself. \"Client 1Password vault\" is the right shape."),
        ("Ready?", "Calculated. Do not type in this column."),
        ("Notes", "The context you will have forgotten in a year."),
        ("", ""),
        ("Readiness test", '=IF(AND(E2="client owned",D2<>"",F2<>""),"yes","no")'),
        ("", "An asset is ready when the client owns it, the owning address is written down, and the way in is written down."),
        ("", ""),
        ("Readiness per client", '=IFERROR(COUNTIFS(A:A,"Northline Cafe",J:J,"yes")/COUNTIF(A:A,"Northline Cafe"),"")'),
        ("", "100% can be handed over this afternoon. 40% is a project."),
        ("", ""),
        ("Source", "geteverkey.com/blog/client-asset-handover-template"),
    ]
    for r, (label, body) in enumerate(lines, start=1):
        a = legend.cell(row=r, column=1, value=label)
        b = legend.cell(row=r, column=2, value=body)
        a.font = Font(name=ARIAL, size=10, bold=(r == 1 or bool(label)), color=INK)
        b.font = Font(name=ARIAL, size=10, color=INK)
        b.alignment = Alignment(vertical="top", wrap_text=True)
    legend["A1"].font = Font(name=ARIAL, size=13, bold=True, color=INK)

    wb.save(out)
    print("wrote " + out)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=OUT, help=f"output path (default: {OUT})")
    build(parser.parse_args().out)
