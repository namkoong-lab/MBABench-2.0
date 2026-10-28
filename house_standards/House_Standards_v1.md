# Financial Modelling — House Standards

October 2026

These are the conventions particular to us. Ordinary professional practice is assumed and not restated here. Where the engagement brief or instructions conflict with anything below, they govern.
You are also allowed to edit input files to ensure that the house standards are also applied to those tabs.

---

# 1 · Structure

A single-entity model recalculates in under five seconds and stays under 10MB. Beyond the sheets any model needs, every model carries a linked contents page, a checks tab, and a change log recording what changed, when and by whom.

**The cover** states scope — including what the model deliberately leaves out — the legend of colours and abbreviations (glossary), the approach and the main design decisions behind it, and a short note on what may be changed, how a case is selected, and where the answer is read.

**Build for the second use.** Where continued use is foreseeable, build for it rather than for the first answer alone: unless the instructions rule it out, adding or removing a business unit is a matter of inputs and switches rather than a structural rebuild.

**Periods.** A given period sits in the same column on every tab. Headers carry the period and its nature — FY2024A, FY2025E, Q1-26B. Periods follow the entity's fiscal year; where none is given, calendar year, stated on the cover. One labelled anchor date drives all date logic. Discount stubs run actual/365; instruments accrue on their own stated basis, recorded beside the input.

---

# 2 · Building

**Units.** $mm unless the instructions say otherwise, stated at the head of every block and column with the currency alongside. Anything in another unit — per-unit prices, headcount, per-share figures, multiples — is labelled where it appears.

**Rounding.** State the convention beside the figures it governs, in words that match what the model does, e.g. "rounded to $0.01mm" where numbers are actually rounded.

**Signs.** Costs and outflows negative, revenues and inflows positive, everywhere — including movements in working capital.

**Assumptions.** An input block on a build sheet is acceptable only as the sole source for that driver, signposted from the assumptions tab. Each assumption carries a source precise enough to re-find — "10-K FY24 p.45", "brief §2", "instructions line 23" — or is marked a judgment with an owner and date. Named ranges are reserved for a handful of global inputs.

**Formulas.** Prefer XLOOKUP/XMATCH to VLOOKUP, IFS or SWITCH to nested IFs, LET for repeated intermediates; INDEX/XMATCH is fine; legacy constructs where compatibility requires, with the reason noted. Use the '3-clause rule': no more than three functions nested in one formula. Mark any intentional break in an otherwise consistent row or block.

**Cases** (if applicable) run off one labelled selector on the assumptions tab that pulls the live set; structural switches — capitalise/expense, include/exclude a unit — work the same way, and no sheet carries a private toggle. Results for every case appear side by side in one view, not one case at a time.

**Checks** are consolidated on the checks tab, with a single flag on the cover that turns red on any failure. The sense check sits there too: key outputs set against outside benchmarks or external comparables, in the workbook rather than in the covering note.

---

# 3 · Presentation

**Colour.** Blue font: hardcoded inputs · black: formulas · green: links to another sheet/tab (also when used inside formulas) · red: links to another workbook. Inputs also carry a pale blue fill; key inputs carry data validation to prevent future users from entering nonsensical or bad inputs. Beyond these, at most two or three additional colours across fonts, fills and tab colours, none of them resembling the four. Yellow means "unfinished" or "review me" and nothing else — nothing unfinished ships, and any yellow left in the delivered file is explained in the legend.

**Numbers.** Zeros as dashes, negatives in parentheses. $mm to two decimals, percentages to one, multiples to one (7.5x), counts to none — held constant across comparable values.

**Grid.** Arial 10 throughout. No merged cells — Center Across Selection instead. A buffer row inside every total range.

---

# 4 · Delivery

Group rather than hide rows and columns, and nothing is hidden anywhere else either. Workbooks are self-contained: no links to other workbooks unless the brief or instructions require them, and no macros without disclosure on the cover and a reason native Excel cannot do the job.

Tab order:  Cover → Contents → Assumptions → Calculations → Outputs/Summary → Checks, Instructions before Assumptions, Questions are Outputs, no executive summary upfront

Circularity only where the economics require it, each instance noted and fitted with a breaker switch, with iterative calculation enabled and stated.

Non-obvious formulas carry a note giving the reasoning — why this method — not a restatement of the formula.

One zoom level throughout, cursor at A1 on every sheet, print areas and page breaks set.

No unresolved cell warnings, generally adhere to a good file hygiene.