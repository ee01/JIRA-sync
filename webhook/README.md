# JIRA-sync

## JIRA changes trigger

Create a webhook for any issue changes in JIRA automation.

1. Field value changed
    * Summary
    * fixVersion/s
    * ... (list all fields to watch)
2. Send web request
    * Content-Type: application/json
    * HTTP method: POST
    * Webhook body: Custom data
    * Custom data: (Attach the JIRA webhook.json here)


## Get JIRA data webhook

Create a webhook to get issue data in JIRA automation.

This imcoming webhook will trigger the changes webhook above to sync data back to sheet.

1. Field value changed
    * Summary
    * fixVersion/s
    * ... (list all fields to watch)
2. Send web request
    * Content-Type: application/json
    * HTTP method: POST
    * Webhook body: Custom data
    * Custom data: (Attach the JIRA webhook.json here)


RC JIRA imcoming webhook mapping to JIRA project:
- MTR: https://jira.ringcentral.com/rest/cb-automation/latest/hooks/9edd6e55ec9b7da28206ab927562da913f5532bf
- FIJI: https://jira.ringcentral.com/rest/cb-automation/latest/hooks/12044d178a8091e40b447d27a20ec08efe3c7ef0
- EOINT: https://jira.ringcentral.com/rest/cb-automation/latest/hooks/13da6faff84caa9a2815e6c17daefd68d063a801
- EW: https://jira.ringcentral.com/rest/cb-automation/latest/hooks/5721109a14fa4d21b3d4ca4354cbb229d1b965b5