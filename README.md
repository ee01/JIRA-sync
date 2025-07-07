# JIRA-sync
Bidirectional sync data between Google Sheet and JIRA

## Deployment

### Test env

`npm run push` 无需versions管理，即可直接生效

测试sheet: https://docs.google.com/spreadsheets/d/1GNeBIM6Z6cnUz1qnlB9rztQJjv6BebCTZQ-6oEGNmbo/edit?gid=0#gid=0

### Prod

1. Copy ./Code.js overwrite here: https://script.google.com/u/1/home/projects/18Os5bP8YSpMWjxQS8HDB8Q5neD7A9gk_IIBvMyxFFFj11344D_nPZmSe/edit

2. Goto script Manage Deploy, edit Versions and select a new Version.
https://script.google.com/home/projects/18Os5bP8YSpMWjxQS8HDB8Q5neD7A9gk_IIBvMyxFFFj11344D_nPZmSe/edit

3. Go to cloud console then edit the version number for Sheets Add-on script version as the new Version number.
https://console.cloud.google.com/apis/api/appsmarket-component.googleapis.com/googleapps_sdk?project=jira-sync-418403&invt=AbuzdQ
