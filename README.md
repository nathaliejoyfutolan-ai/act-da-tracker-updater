# ACT DA tracker updater

Updates the private `ACT_DA_Opportunity_Tracker_Interactive.xlsx` workbook every ten minutes while the workbook is closed.

The workflow retrieves public ACT development-application data, maps it to the workbook's existing 28-column live-feed table, and writes it through Microsoft Graph. The workbook and Microsoft credentials are not stored in this repository.

Required encrypted repository secrets:

- `MS_TENANT_ID`
- `MS_CLIENT_ID`
- `MS_REFRESH_TOKEN`

The workflow also supports a manual run from the Actions tab.

The `Email ACT DA daily report` workflow refreshes the workbook and emails the Daily Report to
`nathalie@decarbysolar.com.au` at 5:30 AM Philippine time each day. Its Microsoft delegated
permissions must include `Mail.Send` in addition to the workbook permissions above.

## Decarby awards monitoring

The `Monitor and email Decarby awards` workflow checks 19 official award sources at 6:00 AM
Philippine time every Monday and Friday. Friday refreshes the saved scan state and report. Monday
also sends a formatted HTML briefing and attaches the filterable awards workbook to
`nathalie@decarbysolar.com.au` using the same authorised Microsoft Graph connection.

The workflow can also be run manually from GitHub Actions with email delivery enabled or disabled.
Its latest machine-readable and HTML reports are saved under `reports/`, while the last page hashes
are stored in `data/awards-scan-state.json` for change detection.
