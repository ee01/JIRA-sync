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
(moved into changelog/Code.js)