# psy342-data-generator
For Emotion Psych Data File Creation
# PSY 342 simulated dataset generator

R script that turns each research group's data-spec spreadsheet into a simulated dataset (CSV) for their project.

- **Strict:** the data follow the spec exactly: sample size, % female, variable ranges, condition labels, and whole numbers on Likert-type scales.
- **Not too strict:** the planned effects are moderate and realistic, so the results are usable without looking too clean.
- **The TA decides** which hypotheses come out supported. Every group gets at least one supported effect, but not necessarily all of them.

## Folder setup

```
Desktop/
├── generate_datasets.R
├── specs/          <- put the groups' .xlsx spec files here (do NOT commit)
└── output/         <- created by the script (do NOT commit)
```

## Running it

1. **First time only:** in the RStudio Console, run `install.packages(c("readxl", "haven"))`.
2. Put the groups' `.xlsx` spec files in `specs/`.
3. Edit the `EFFECTS` section of the script (section 2) for each group, based on their letter of intent.
4. In the Console, run `setwd("/Users/<your-name>/Desktop")`. Use the folder that **contains** `specs/`.
5. Click **Source**.

If you get `No .xlsx files found in 'specs'`, the working directory is wrong. Run `getwd()` and repeat step 4.

## Writing effects (section 2)

```r
EFFECTS <- list(
  "Group3" = list(                                   # matched against the file name
    diff_effect("PostPE", higher = "Positive", lower = "Neutral", d = 0.5),  # supported
    cor_effect("Extraversion", "PrePE", r = 0)                               # not supported
  )
)
```

| Helper | Meaning |
|---|---|
| `diff_effect(dv, higher, lower, d)` | Group difference on `dv`. If `dv` is a post measure with a matching pre measure (`_T1`/`_T2`, `Pre`/`Post`), this is the Condition × Time interaction. Also works with `"Female"`/`"Male"`. |
| `cor_effect(var1, var2, r)` | Correlation between two continuous variables. |
| `time_effect(pre, post, d)` | Everyone changes from pre to post. |
| `moderation_effect(dv, higher, lower, moderator, d_per_sd)` | The group difference grows with a continuous moderator. |

Default sizes are d = 0.5 and r = .30. Setting an effect to `0` makes the script guarantee that it is **not** significant.

## Output (in `output/`)

| File | What it is |
|---|---|
| `<group>.csv` | The dataset to send to the group. It opens in SPSS, Excel, R, or JASP. |
| `<group>_report.txt` | What was read and auto-fixed, planned vs. observed effects, and descriptives. |
| `_summary.csv` | One row per group: OK / WARN / ERROR. |
| `_contact_groups.txt` | Spec problems to ask each group about. Groups with errors get no dataset until they fix their spec. |

## Built-in checks

- **Spec errors** (no dataset is generated): missing sample size, a range with only one value, a grouping variable with only one level.
- **Auto-fixed and noted in the report:** a proportion written as 0.5 instead of 50%, a range entered upper value first, spaces or symbols in variable or condition names.
- **Effect check:** each planned effect is re-tested on the generated data. The script redraws until every effect is significant in the right direction and within ±0.15 (d) or ±0.10 (r) of its planned value. Effects set to 0 must come out non-significant.

The script also warns when a group's N is too small for the planned effect, when an effect is set unrealistically large, or when no effect is set as supported.

## Settings (section 1)

`PREPOST_R` (pre–post correlation, default .50), `SCALE_SPREAD`, `TOLERANCE_D`, `TOLERANCE_R`, `MAX_TRIES`, `DECIMALS`, and `ALSO_WRITE_SAV` (set to `TRUE` to also write SPSS `.sav` files).
