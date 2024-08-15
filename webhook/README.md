# JIRA-sync

Create a webhook in JIRA automation.

1. Field value changed
    * Summary
    * fixVersion/s
    * ... (list all fields to watch)
2. Send web request
    * Content-Type: application/json
    * HTTP method: POST
    * Webhook body: Custom data
    * Custom data: (Attach the JIRA webhook.json here)


