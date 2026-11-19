# The Master's General Account in One Workbook

**How I turned a pile of linked spreadsheets into one checked Excel file**

By **John Winsome Munar (Trozovka)**, Master Mariner
[trozovka.github.io](https://trozovka.github.io) · [GitHub](https://github.com/Trozovka)

---

Every month the Master sends the office a general account: crew wages, allotments
to families, cash advances, bonded store sales, allowances paid in cash, and the
cash left in the safe. Most companies give you an Excel form for it. Most Masters
then build their own helper files around that form, because the form alone does
not help you get the numbers right.

I did the same for years. My helper files grew until they caused more mistakes
than they prevented. This page explains what was wrong, what I changed, and the
rules I now follow, so another Master (or a developer helping one) can do the same
with their own company's form.

Nothing here is copied from any company form. All numbers are made up.

## Contents

1. [What the monthly account covers](#1-what-the-monthly-account-covers)
2. [Where my old setup went wrong](#2-where-my-old-setup-went-wrong)
3. [The rules I set myself](#3-the-rules-i-set-myself)
4. [The sheets](#4-the-sheets)
5. [Wages: pro-rata by 30 days](#5-wages-pro-rata-by-30-days)
6. [Let the allotment calculate itself](#6-let-the-allotment-calculate-itself)
7. [Bonded store and the handler's commission](#7-bonded-store-and-the-handlers-commission)
8. [Getting the numbers into the company form](#8-getting-the-numbers-into-the-company-form)
9. [Checks you can see](#9-checks-you-can-see)
10. [Testing it with a real past month](#10-testing-it-with-a-real-past-month)
11. [The Excel functions I used](#11-the-excel-functions-i-used)
12. [Lessons](#12-lessons)
13. [Disclaimer](#13-disclaimer)

## 1. What the monthly account covers

| Part | What it is |
|---|---|
| Wages | Each crew member's earnings for the month, by rank and days on board |
| Other incomes | Bonuses and allowances that go through the payroll |
| Deductions | Leave pay, union fee, bonded store, cash advance, phone, others |
| Allotments | What is left, sent home to one or more beneficiaries |
| Cash advances | Cash given on board, split by reason |
| Cash allowances | Hold cleaning and similar work paid in cash |
| Bonded store | Cigarettes, drinks and snacks sold to crew or used for entertainment |
| Cash account | Cash received, cash paid out, cash carried forward |
| Payslips | One receipt per crew member |

## 2. Where my old setup went wrong

- **Links to other files.** My helper files pulled numbers from the company file
  by external links. Every month the file got a new name or moved to a new folder,
  and the links either broke or, worse, kept reading last month's file without any
  warning. One of my files was still linked to a copy on a USB stick.
- **The same crew list typed six times.** Cash advances, cash received, bonuses,
  allowances, bonded store and the company form each had their own list of names.
  One spelling change and the totals stopped matching.
- **Rates typed inside formulas.** Each wage cell held the monthly rate inside the
  formula itself. When the company changed the scale, every row had to be edited
  by hand.
- **The allotment was a typed number.** I worked it out on the side so that
  nothing was carried forward, then typed it in. Any later change to a deduction
  meant doing it again.
- **Checks that depended on the links.** My cross-check file was only right when
  its links happened to point at the right file.
- **Rounding everywhere and nowhere.** Some sheets rounded to cents, others did
  not, so totals that should agree were a few cents apart and I never knew which
  one was right.

## 3. The rules I set myself

1. **The company form is never changed.** It is their document. My file only
   prepares numbers for it.
2. **One file per month.** No links to any other file.
3. **Type each fact once.** A crew member's name, rank and dates are typed in one
   place; every other sheet reads them from there.
4. **Plain formulas only.** No macros (company laptops often block them) and only
   functions that work in old Excel versions, so any Master can open it.
5. **Same cell addresses as the company sheet** wherever numbers are copied
   across, so a block can be pasted straight in.
6. **Checks on screen.** Green when right, red when wrong, amber when it might be
   intended.
7. **One rounding rule.** Money that leaves the ship (allotments) is rounded to
   cents; the tiny remainder is carried forward and stays visible.

## 4. The sheets

| Sheet | What you do there |
|---|---|
| START | Vessel, month, Master, and the steps for the month |
| RATES | The company wage scale by rank, typed once, changed only when the company changes it |
| CREW | One line per person: rank, name, nationality, sign on, sign off, personal allowance, balances |
| INCOME | One line per extra income item, with a type, so you can see later what each amount was |
| ADVANCES | Cash advance per person, split into parts (on board, port orders, home office, bonded, other) |
| WAGES | The full payroll, calculated |
| CASH ALLOW | Cash allowances per person |
| BOND | Bonded store quantities per person and per article, stock in and out |
| CASH | Cash received, cash paid, balance, and a box to count the safe |
| ALLOTMENTS | Beneficiary lines, with fixed amounts and "the rest" |
| PAYSLIP | Pick a name, print, sign |
| CHECKS | Every check in one place |
| TRANSFER | The finished numbers laid out exactly like the company sheet |

## 5. Wages: pro-rata by 30 days

Many companies pay a full month as 30 days, whatever the calendar says, and a part
month by the actual days on board. With a made-up basic wage of 900 a month:

- On board the whole month: 900 × 30 / 30 = **900**
- Joined on the 6th of a 31-day month (26 days): 900 × 26 / 30 = **780**

The days fill in by themselves from the sign-on and sign-off dates:

```
=IF(name="","",
  IF(AND(OR(signon="",signon<=month_start), OR(signoff="",signoff>=EOMONTH(month_start,0))),
     30,
     MIN(IF(signoff="",EOMONTH(month_start,0),signoff),EOMONTH(month_start,0))
     - MAX(IF(signon="",month_start,signon),month_start) + 1))
```

Each wage part is then a lookup into the RATES sheet by rank, times days / 30.
One part (a personal allowance) differs from person to person even within the
same rank, so it sits on the CREW sheet instead.

## 6. Let the allotment calculate itself

The regular allotment is simply what is left:

```
allotment = earned wage + other incomes + last month's balance
          - all deductions - anything the crew member wants kept on board
```

rounded to cents. If a crew member ends up owing money (for example a large cash
advance on a short month), the allotment goes negative. A negative allotment cannot
be sent, so the CHECKS sheet shows it in amber: it might be intended, but you
should look.

Many crew send money to more than one person: a fixed amount to one, the rest to
another. Each beneficiary is one line. Type the fixed amount on the first line,
leave the second blank, and the blank line takes whatever is left:

```
=IF(fixed<>"", fixed,
   IF(first line of this person without a fixed amount,
      ROUND(person's allotment - SUMIFS(fixed amounts of this person), 2), 0))
```

The lines of all crew must add up to the payroll total. That is one of the checks.

## 7. Bonded store and the handler's commission

Bonded items are bought from a different supplier in each port, so the price is
typed per article, once. Quantities are typed per person. Each person's total is
`SUMPRODUCT(quantities, prices)`, which goes straight into their deductions. No
long hand-made `qty×price + qty×price + ...` formulas that silently miss a column.

If the selling price is cost plus a markup that pays the store handler, the
commission can be worked out from total sales. With a 5% markup:

```
commission = total sales / 1.05 × 0.05
```

Stock carried forward is stock last month + purchased − sold − used for
entertainment. A negative number means a typing error, and it turns red.

## 8. Getting the numbers into the company form

The TRANSFER sheet uses the same cell addresses as the company sheet. You select
a white block, copy, click the same cell in the company file, and use
**Paste Special → Values**. Columns that are formulas in the company form are
shaded grey, and you never paste over them.

This is why "same cell addresses" is one of the rules: there is nothing to map by
hand, so nothing to get wrong.

## 9. Checks you can see

| Check | Why |
|---|---|
| Every rank exists in the rate table | A typo in a rank would quietly pay zero |
| Days between 1 and 30 | Catches wrong dates |
| Every income line belongs to a crew member | Catches name spelling differences |
| Every income line has a type | So the payslip can be explained later |
| Negative allotment (amber) | Crew owing money |
| Carried forward = what you wanted | The allotment really took everything else |
| Income total and cash advance total agree across sheets | Nothing lost between sheets |
| Allotment lines = payroll allotments | Every dollar has a beneficiary |
| Allotment > 0 but no beneficiary line | Someone would not get paid |
| Negative bonded stock | Typing error |
| Cash counted = cash balance | The safe matches the books |
| Company totals typed back in | After pasting, the company form agrees with this file |

The last check replaces the external links I used before. After pasting into the
company form, you type its handful of totals back into the CHECKS sheet. Six
numbers typed by hand are more reliable than six links that may point at the wrong
file.

## 10. Testing it with a real past month

The only test I trust: take a month that was already sent and accepted, put its
data into the new file, and require every line to match the old result **to the
cent**. Wages, incomes, allotments, bonded, cash advances, the commission and the
final cash balance.

It matched on every line. Along the way the checks also showed two things the old
files had hidden: a crew member who left owing money (I had planted that one on
purpose, to see if it would be caught), and two parts of the company form that
disagreed with each other about that same person. Both would have been caught
before sending.

## 11. The Excel functions I used

All available since Excel 2010, so they work on old office laptops:

| Function | Used for |
|---|---|
| `INDEX` + `MATCH` | Looking up a rank's rates, a person's allotment |
| `SUMIFS`, `COUNTIFS` | Totals per person, "first line of this person" |
| `SUMPRODUCT` | Quantity × price across a row, counting problems in a range |
| `EOMONTH` | Month end for the day count |
| `ROUND` | One rounding rule, in one place |
| Data validation lists | Rank and name drop-downs, so spelling always matches |
| Conditional formatting | Green / red / amber checks |

No `XLOOKUP`, no `FILTER`, no macros.

## 12. Lessons

- **Type each fact once.** Most of my old mistakes were the same fact typed twice
  and drifting apart.
- **Links between files are a trap** for anything renamed monthly. One file per
  month, no links.
- **A typed balancing number is a mistake waiting to happen.** Calculate it.
- **Test against a month you already know is right.** "It looks right" is not a test.
- **Make checks visible.** A check hidden in another file is a check nobody runs.

## 13. Disclaimer

This is how I organise my own monthly account. It is not any company's official
form or procedure. Your company's rules, forms and deadlines always come first;
if your company counts days, rounds or deducts differently, change the file to
match them.

---

**About the author.** John Winsome Munar (Trozovka) is a Master Mariner who builds
practical software for seafarers, including WX weather apps for slow ship internet,
PocketMOB Alert (man-overboard alarm), PocketVDR (voyage data recorder) and OpenCPN
GPS Server. More at [trozovka.github.io](https://trozovka.github.io).

Text © 2026 John Winsome Munar. Shared under
[CC BY 4.0](https://creativecommons.org/licenses/by/4.0/).
