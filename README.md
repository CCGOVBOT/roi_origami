# roi_origami

Workers' compensation metrics sourced from Origami Risk. Origami is primarily utilized
by Risk Management.

## Dependencies

- `pandas`
- `roi_common`
- `re`
- `python-dotenv`

## Data source

Email attachments from `notifications@origamirisk.com` (Excel files).

- Sender filter: `sendername_contains="notifications@origamirisk.com"` (resolved at runtime).
- Attachment filename must contain a date in `YYYY-MM-DD` format; this date is used as
  `report_date`.
- All sheets in each attachment are parsed with `header=4` (row 5 as header).
- Intermediate combined files are written locally to `email_attachments/combined_{sheet_name}.xlsx`.

No date range filter on email retrieval — all matching emails are processed.

## Processing notes

### Sheets consumed (by cleaned sheet name)

| Sheet key | Content |
|---|---|
| `open_wc` | Open Workers' Comp claims |
| `closed_wc` | Closed Workers' Comp claims |
| `average_adjuster_caseload` | Per-adjuster claim count |

### Filtering

Rows where `Coverage` or `Adjuster User` contains `"Grand Totals"` are removed.

### Calculations

- `closing_ratio = closed_claims / open_claims` grouped by `report_date`.
- `average_caseload = total_claims / num_adjusters` grouped by `report_date`.
- `end_date = start_date + MonthEnd(0)`.

### Deduplication / Validation

Relies on `roi_merge_metric_data` for key-conflict detection.

## Metrics written to `roi.metric_data`

| DB name (roi.metric_defs) | metric_id | Period |
|---|---|---|
| Workers' compensation claim closing ratio | 12210 | Monthly |
| Worker's compensation claims opened | 12495 | NULL |
| Worker's compensation claims closed | 12496 | NULL |
| Open workers' compensation claims per adjuster | 12871 | Monthly |
