# ACT DA tracker updater

Updates the private `ACT_DA_Opportunity_Tracker_Interactive.xlsx` workbook every ten minutes while the workbook is closed.

The workflow retrieves public ACT development-application data, maps it to the workbook's existing 28-column live-feed table, and writes it through Microsoft Graph. The workbook and Microsoft credentials are not stored in this repository.

Required encrypted repository secrets:

- `MS_TENANT_ID`
- `MS_CLIENT_ID`
- `MS_REFRESH_TOKEN`

The workflow also supports a manual run from the Actions tab.
