---
name: "conference-expense-report"
description: "Turn a folder of travel/conference receipts (photos, PDFs, screenshots — any country, any language, any currency) into a simple reimbursement Excel workbook: a chronological Summary sheet where every item links to its own receipt sheet with the embedded image, an editable exchange-rate table, category subtotals, and comments. Use when the user asks to organize receipts for reimbursement, make an expense report/spreadsheet for a trip, or says \"put my receipts into Excel\"."
---

# Conference / Travel Expense Report (receipts → linked Excel)

Builds a reimbursement-ready `.xlsx` from a folder of receipts from **any trip, country, language or currency**.

**Keep it simple.** The user prefers this exact layout. Do not add extra blocks, columns or sections such as per-row rate columns, separate vendor, city or payment columns, or a totals block above the table, unless the user asks for them.

## Output format (user-approved)

- **Summary** sheet:
  - Title (conference, city, country, dates) and a traveler line.
  - A small **exchange-rate table**: `Currency | <Home> per 1 unit | Source`. Foreign rates are editable yellow cells, and the home currency is 1.
  - The main table, sorted as a **timeline**: `# | Date | Category | Item | Currency | Amount (original) | <Home> | Receipt | Comments`
    - Home = `=ROUND(Amount*INDEX(rates,MATCH(Currency,codes,0)),2)`.
    - **There is no Rate column**, because the rate lives only in the rate table.
  - Total row, subtotal by category (SUMIF) with a check sum, a spent-by-currency block if more than one currency was used, and a notes line.
- Each **Receipt** cell is a hyperlink to that receipt's own sheet (`R01_Registration`, `R05_Dinner_0312`, …).
- Each receipt sheet has:
  - A `← Back to Summary` link in A1.
  - A one-line header: #, date, category, item, currency and amount.
  - The comment.
  - The receipt image(s), embedded and stacked. A multi-page PDF becomes one image per page.
- Arial font throughout. Values copied from receipts are blue, and formulas are black. Landscape printing, fit to page width.

Home currency defaults to **USD**. If the institution reimburses in another currency, set `HOME`/`HOME_SYM`.

## Workflow

1. **Inventory the folder**: run `ls -la` and include subfolders.
   - If the user gives an example workbook, match its conventions. They take priority.
   - Ask (AskUserQuestion) only for things the receipts can't tell you: the conference name and dates, the reimbursement currency, and any required rate policy.
2. **Read every receipt. Never guess amounts.**
   - PDFs: run `pdftotext -layout file.pdf -` and `pdfinfo`. To embed pages, render them with `pdftoppm -r 110 -png file.pdf /tmp/img/<name>`.
   - Photos and screenshots: make downscaled copies (PIL `ImageOps.exif_transpose`, thumbnail ~1100px, JPEG q80) in the outputs folder and view each one with the Read tool.
   - Record for each receipt: date, time, vendor, city, total, currency, payment method, card last 4 digits, and anything unusual.
3. **Interpret receipts in any language.**

   | Meaning | Examples |
   |---|---|
   | Total | Total, Gesamt/Summe, Totale, Totaal, Итого, 合計/合计, 총액, Toplam |
   | Cash | Contado/Efectivo (ES), Espèces (FR), Bar (DE), Contanti (IT), Dinheiro (PT), Contant (NL), 現金/现金, 현금, Наличные |
   | Card | Tarjeta, Carte/CB, Karte/EC, Carta, Cartão, Pin/Kaart, カード, 刷卡/银行卡, 카드, Visa/MC/Amex slip |
   | Tip / service | Propina, Pourboire/Service, Trinkgeld, Mancia/Coperto, 小费/服务费, チップ |
   | Tax | VAT, IVA, TVA, MwSt/USt, BTW, GST/HST, 消費税, 增值税, 부가세 |
   | Change given | Cambio, Monnaie rendue, Rückgeld, Resto, 找零, お釣り |

   - **Date formats:**
     - US receipts use MM/DD/YYYY.
     - Most of Europe, Latin America, India and Australia use DD/MM/YYYY.
     - China, Japan, Korea and ISO style use YYYY/MM/DD.
     - Check ambiguous dates against the trip dates. Write all dates as MM/DD/YYYY.
   - **Currency:** read the symbol or code on the receipt, not the country. Online bookings and DCC slips may be in USD.
   - **Shared or split bills:**
     - Fractional quantities (0.5, 0.25) mean the traveler's share.
     - If the card slip is less than the bill, claim the card amount and note "Total bill X for group; my card payment Y".
     - If the receipt shows items for more than one person, ask the user and note it (e.g. "Covers two people").
   - **Proof of payment:** a proforma with no card slip gets a note. For cash, say how you know (e.g. "'Contado' printed on receipt = cash").
   - Itemize fees when it helps (e.g. "Ticket 20.00 + booking fee 3.00 + seat 2.00").
   - Price-comparison screenshots (e.g. a corporate-portal quote) go on the same sheet as the item they support, with a comment.
   - Keep comments factual and neutral, since an accountant will read them. If a draft contains personal remarks, point them out before submission.
4. **Order and categorize.**
   - Pre-trip purchases go first, sorted by payment date: registration, visa, insurance. Everything else is sorted by date of service.
   - Multi-day items use a range, "MM/DD/YYYY – MM/DD/YYYY".
   - Categories: Registration, Flight, Hotel, Ground transport, Meal, Insurance, Visa, Communication, Other.
5. **Exchange rates.**
   - Use the institution's method if it names one.
   - Otherwise, use WebSearch to find an official mid-trip rate: the Fed H.10, the ECB, or the trip country's central bank.
   - Cite the sources in your reply.
6. **Build** with the script below and save to the user's folder. Don't overwrite a file the user may have edited.
7. **Recalculate and verify.**
   - Run the xlsx skill's `scripts/recalc.py <file> 60`. It must report `status: success` with 0 errors.
   - Read back the total with `data_only=True` and spot-check two or three rows.
   - If recalc reports "Permission denied", the file is open in Excel. Ask the user to close it, then re-run.
   - Delete any leftover `.recalc-*.xlsx` temp files.
8. **Report briefly:** the total, the rate(s) used, and a short list of things to double-check, such as missing receipts, shared bills, estimated amounts, missing proof of payment, and side-trip items.

## Build script (fill the marked section)

```python
import openpyxl, os
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from PIL import Image, ImageOps

# ---------- fill these in ----------
OUT      = "/path/to/user/folder/Expense_<Conference>_<City>.xlsx"
TITLE    = "<Conference> (<City>, <Country>, <Dates>) – Travel Expense Summary"
TRAVELER = "Traveler: <Name>, <Dept>, <Institution>"
HOME     = "USD"          # reimbursement currency
HOME_SYM = "$"
# one row per currency used on the trip (home currency rate = 1)
RATES = [  # (code, units of HOME per 1 unit, source note)
    ("USD", 1.0,    "Home currency"),
    ("EUR", 1.0000, "<source, e.g. ECB reference rate, MM/DD/YYYY>")  # replace 1.0000 with the real rate,
]
rows = [
 # (date, category, item, currency, amount, sheet_name(<=31), [image paths], comment)
 ("MM/DD/YYYY","Registration","Conference registration fee","EUR",100.00,"R01_Registration",["/tmp/img/registration-1.png"],""),
 ("MM/DD/YYYY – MM/DD/YYYY","Flight","Round-trip flight","USD",500.00,"R02_Flight",["/tmp/img/flight-1.png"],""),
]
# -----------------------------------

F="Arial"; thin=Side(style="thin",color="999999"); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
HDR=PatternFill("solid",fgColor="1F4E78"); YEL=PatternFill("solid",fgColor="FFFF00"); TOT=PatternFill("solid",fgColor="DDEBF7")
wb=openpyxl.Workbook(); ws=wb.active; ws.title="Summary"
ws["A1"]=TITLE; ws["A1"].font=Font(name=F,bold=True,size=13)
ws["A2"]=TRAVELER; ws["A2"].font=Font(name=F,size=10)

# exchange-rate table
ws["A4"]="Exchange rates"; ws["A4"].font=Font(name=F,bold=True)
for i,h in enumerate(["Currency",f"{HOME} per 1 unit","Source"],2):
    c=ws.cell(4,i,h); c.font=Font(name=F,bold=True,color="FFFFFF"); c.fill=HDR; c.border=bd
r0=5
for k,(code,rate,src) in enumerate(RATES):
    rr=r0+k
    ws.cell(rr,2,code).font=Font(name=F,size=10)
    c=ws.cell(rr,3,rate); c.number_format="0.0000"; c.font=Font(name=F,size=10,color="0000FF",bold=True)
    if code!=HOME: c.fill=YEL
    ws.cell(rr,4,src).font=Font(name=F,italic=True,size=9,color="666666")
    for col in (2,3,4): ws.cell(rr,col).border=bd
rlast=r0+len(RATES)-1
RNG_C=f"$B${r0}:$B${rlast}"; RNG_R=f"$C${r0}:$C${rlast}"

hr=rlast+2
H=["#","Date","Category","Item","Currency","Amount (original)",f"{HOME} ({HOME_SYM})","Receipt","Comments"]
for i,h in enumerate(H,1):
    c=ws.cell(hr,i,h); c.font=Font(name=F,bold=True,color="FFFFFF"); c.fill=HDR; c.border=bd
    c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
first=r=hr+1
for n,(d,cat,item,cur,amt,sh,imgs,com) in enumerate(rows,1):
    vals=[n,d,cat,item,cur,amt,f"=ROUND(F{r}*INDEX({RNG_R},MATCH(E{r},{RNG_C},0)),2)",sh,com]
    for i,v in enumerate(vals,1):
        c=ws.cell(r,i,v); c.font=Font(name=F,size=10); c.border=bd
        c.alignment=Alignment(vertical="center",wrap_text=(i in (4,9)))
    ws.cell(r,5).alignment=Alignment(horizontal="center",vertical="center")
    ws.cell(r,6).number_format="#,##0.00"; ws.cell(r,6).font=Font(name=F,size=10,color="0000FF")
    ws.cell(r,7).number_format=f'"{HOME_SYM}"#,##0.00'
    g=ws.cell(r,8); g.hyperlink=f"#'{sh}'!A1"; g.font=Font(name=F,size=10,color="0563C1",underline="single")
    r+=1
last=r-1
ws.cell(r,4,"Total")
c=ws.cell(r,7,f"=SUM(G{first}:G{last})"); c.number_format=f'"{HOME_SYM}"#,##0.00'
for i in range(1,10):
    c=ws.cell(r,i); c.border=bd; c.fill=TOT; c.font=Font(name=F,bold=True)
# subtotals by category and by currency
r+=2; ws.cell(r,3,f"Subtotal by category ({HOME})").font=Font(name=F,bold=True); s0=r+1
for cat in dict.fromkeys(x[1] for x in rows):
    r+=1; ws.cell(r,3,cat).font=Font(name=F,size=10)
    c=ws.cell(r,7,f"=SUMIF($C${first}:$C${last},C{r},$G${first}:$G${last})"); c.number_format=f'"{HOME_SYM}"#,##0.00'; c.font=Font(name=F,size=10)
r+=1; ws.cell(r,3,"Check (should equal total)").font=Font(name=F,italic=True,size=9)
c=ws.cell(r,7,f"=SUM(G{s0}:G{r-1})"); c.number_format=f'"{HOME_SYM}"#,##0.00'; c.font=Font(name=F,italic=True,size=9)
if len(RATES)>1:
    r+=2; ws.cell(r,3,"Spent by currency (original amount)").font=Font(name=F,bold=True)
    for code,_,_ in RATES:
        r+=1; ws.cell(r,3,code).font=Font(name=F,size=10)
        c=ws.cell(r,6,f"=SUMIF($E${first}:$E${last},C{r},$F${first}:$F${last})"); c.number_format="#,##0.00"; c.font=Font(name=F,size=10)
r+=2; ws.cell(r,1,f"Notes: Blue values are copied from receipts. {HOME} = original amount × rate from the exchange-rate table (edit the yellow cells to change). Click a Receipt link to open its sheet; each receipt sheet links back here.").font=Font(name=F,italic=True,size=9)
for col,w in zip("ABCDEFGHI",[4,22,16,50,9,12,12,18,46]): ws.column_dimensions[col].width=w
ws.freeze_panes=f"A{hr+1}"
ws.page_setup.orientation="landscape"; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0; ws.sheet_properties.pageSetUpPr.fitToPage=True

# receipt sheets
os.makedirs("/tmp/build/sm",exist_ok=True)
for n,(d,cat,item,cur,amt,sh,imgs,com) in enumerate(rows,1):
    s=wb.create_sheet(sh)
    s["A1"]="← Back to Summary"; s["A1"].hyperlink="#'Summary'!A1"; s["A1"].font=Font(name=F,color="0563C1",underline="single",bold=True)
    s["A2"]=f"#{n}  {d}  |  {cat}  |  {item}  |  {cur} {amt:,.2f}"; s["A2"].font=Font(name=F,bold=True,size=11)
    if com: s["A3"]=com; s["A3"].font=Font(name=F,italic=True,size=10)
    row=5
    for k,p in enumerate(imgs):
        im=ImageOps.exif_transpose(Image.open(p)).convert("RGB"); im.thumbnail((1200,1600))
        out=f"/tmp/build/sm/{sh}_{k}.jpg"; im.save(out,quality=78)
        xi=XLImage(out); W=620 if im.width<im.height else 760
        xi.height=int(im.height*W/im.width); xi.width=W
        s.add_image(xi,f"A{row}"); row+=int(xi.height/20)+3
wb.save(OUT)
print("saved",OUT)
```

## Notes

- Sheet names must be ≤31 characters with none of `[]:*?/\`. Use `R##_<Short>` so they sort in timeline order.
- Keep images compressed (≤1200px, q≈78) so the workbook stays a few MB.
- For a domestic, single-currency trip, `RATES` is just the home currency.
- Never invent receipts or amounts. List anything missing for the user instead.

