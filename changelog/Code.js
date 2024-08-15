/**
 * 定时轮询 JIRA webhook 的数据，清除无关联的，并生成有关联的到 changelog 等待执行数据推送
 * Install: 添加 filterChangesNoRelatedToSheets() 每分钟执行一次
 * 
 * 
 * Version: 2024-8-15
 * 
 * Author: Esone
 *  */

const syncbackSheetURL = 'https://docs.google.com/spreadsheets/d/107ER5MfUeWfTZAKmvMvYwGTaHxI8zcaw3AnlEIAzKNk/edit#gid=0'

const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
const changelogSheet = activeSpreadsheet.getSheetByName('changelog')
const colSheetURL = getHeaderCol('sheet URL', changelogSheet)
const jiraWebhookSheet = activeSpreadsheet.getSheetByName('jira_webhook_data')
const colFrom = getHeaderCol('from', jiraWebhookSheet)
const colEditor = getHeaderCol('Editor', jiraWebhookSheet)
const colJIRAKey = getHeaderCol('JIRA key', jiraWebhookSheet)
const colJIRAFieldName = getHeaderCol('JIRA field name', jiraWebhookSheet)
const colOldValue = getHeaderCol('old value', jiraWebhookSheet)
const colNewValue = getHeaderCol('new value', jiraWebhookSheet)
const colIsSync = getHeaderCol('isSync', jiraWebhookSheet)
const colTime = getHeaderCol('time', jiraWebhookSheet)
const colSyncTime = getHeaderCol('sync time', jiraWebhookSheet)
const colTookSeconds = getHeaderCol('took seconds', jiraWebhookSheet)
const colFailReason = getHeaderCol('fail reason', jiraWebhookSheet)

function filterChangesNoRelatedToSheets() {
  let logs = jiraWebhookSheet.getRange(2, 1, 10000, 21).getValues()
  Logger.log(logs.length)
  logs = logs.filter((log, i) => {
    log['rowIndex'] = 2 + i
    if (!log[0]) return false
    if (log[colFrom-1] != 'jira') return false
    if (log[colIsSync-1] == 'Done' || log[colIsSync-1] == 'Failed') return false
    if (!log[colJIRAKey-1]) return false
    return true
  })
  Logger.log(logs.length)
  logs = logs.forEach((log, i) => {
    if (log[colEditor-1] == 'sync.service@ringcentral.com') {
      jiraWebhookSheet.getRange(log['rowIndex'], colIsSync).setValue('Failed')
      jiraWebhookSheet.getRange(log['rowIndex'], colSyncTime).setValue(new Date())
      jiraWebhookSheet.getRange(log['rowIndex'], colTookSeconds).setValue(Math.ceil((new Date().getTime() - new Date(log[colTime-1]).getTime()) / 1000))
      jiraWebhookSheet.getRange(log['rowIndex'], colFailReason).setValue('Ignore sync.service to avoid running into loop!')
      return
    }
    locationInSheets = getLocationInSheets(log[colJIRAKey-1], log[colJIRAFieldName-1])
    if (!locationInSheets || locationInSheets.length <= 0) {
      jiraWebhookSheet.getRange(log['rowIndex'], colIsSync).setValue('Failed')
      jiraWebhookSheet.getRange(log['rowIndex'], colSyncTime).setValue(new Date())
      jiraWebhookSheet.getRange(log['rowIndex'], colTookSeconds).setValue(Math.ceil((new Date().getTime() - new Date(log[colTime-1]).getTime()) / 1000))
      jiraWebhookSheet.getRange(log['rowIndex'], colFailReason).setValue('No mapping tickets in syncback index sheet!')
    } else {
      locationInSheets.forEach(location => {
        changelogSheet.appendRow([log[colEditor-1], log[colFrom-1], "replace", log[colOldValue-1], log[colNewValue-1], location['sheet name'], location['sheet URL'], location['sheet tab'], location['sheet tab gid'], location['sheet row'], location['sheet column'], location['sheet key header'], log[colJIRAKey-1], location['JIRA field desc'], log[colJIRAFieldName-1], location['JIRA field type'], new Date()])
      })
      jiraWebhookSheet.getRange(log['rowIndex'], colIsSync).setValue('Done')
      jiraWebhookSheet.getRange(log['rowIndex'], colSyncTime).setValue(new Date())
      jiraWebhookSheet.getRange(log['rowIndex'], colTookSeconds).setValue(Math.ceil((new Date().getTime() - new Date(log[colTime-1]).getTime()) / 1000))
      Logger.log('Append to changelog as found changed ticket in sheets. Sheet locations:')
      Logger.log(locationInSheets)
    }
  })
}
/* Deprecated: webhook in changelog sheet
function filterChangesNoRelatedToSheets() {
  let logs = changelogSheet.getRange(2, 1, 10000, 21).getValues()
  Logger.log(logs.length)
  logs = logs.filter((log, i) => {
    log['rowIndex'] = 2 + i
    if (!log[0]) return false
    if (log[colFrom-1] != 'jira') return false
    if (log[colIsSync-1] == 'Done' || log[colIsSync-1] == 'Failed') return false
    if (log[colSheetURL-1]) return false
    if (!log[colJIRAKey-1]) return false
    return true
  })
  Logger.log(logs.length)
  let countDeleted = 0
  for (var i in logs) {
    locationInSheets = getLocationInSheets(logs[i][colJIRAKey-1], logs[i][colJIRAFieldName-1])
    if (logs[i][colEditor-1] == 'sync.service@ringcentral.com' || !locationInSheets || locationInSheets.length <= 0) {
      // 只删除jira同步过来不在syncback表的raw行，以及过滤循环同步
      if (deleteRow(changelogSheet, logs[i]['rowIndex'] - countDeleted, colJIRAKey, logs[i][colJIRAKey-1])) countDeleted++
      // if(countDeleted > 3) return  // 有一条终止，防止删除影响rowIndex
    }
  }
} */

function getLocationInSheets(key, fieldName) {
  const syncbackSS = SpreadsheetApp.openByUrl(syncbackSheetURL)
  syncbackSheets = syncbackSS.getSheets()
  syncbackIndexes = getSyncbackIndexes()
  let whereArr = [
    {name: 'JIRA key', value: key},
    {name: 'JIRA field name', value: fieldName},
  ]
  for (var i in whereArr) {
    if (!whereArr[i]['name']) delete whereArr[i]
    let col = getHeaderCol(whereArr[i]['name'], syncbackSheets[0])
    if (!col) delete whereArr[i]
    whereArr[i]['col'] = col
  }

  let indexes = syncbackIndexes.filter(row => {
    return whereArr.reduce((aggr, cur, i) => {
      if (!aggr) return false
      if (row[ cur['col']-1 ] != cur['value']) return false
      return true
    }, true)
  })
  // if (indexes.length) Logger.log(whereArr)

  /* Deprecated: search from 1 sheet by 1 sheet
  const syncbackSS = SpreadsheetApp.openByUrl(syncbackSheetURL)
  syncbackSheets = syncbackSS.getSheets()
  let indexes = []
  syncbackSheets.forEach(sheet => {
    Logger.log(sheet.getName())
    indexes = [...indexes, ...getRowByValues([
      {name: 'JIRA key', value: key},
      {name: 'JIRA field name', value: fieldName},
    ]), sheet]
  }) */

  // format
  let headers = syncbackSheets[0].getRange(1, 1, 1, 100).getValues()[0]
  return indexes.map(item => {
    return item.reduce((aggr, cur, i) => {
      if (headers[i]) aggr[ headers[i] ] = cur
      return aggr
    }, {})
  })
}

let syncbackIndexes = []
function getSyncbackIndexes() {
  if (syncbackIndexes.length) return syncbackIndexes
  const syncbackSS = SpreadsheetApp.openByUrl(syncbackSheetURL)
  syncbackSheets = syncbackSS.getSheets()
  syncbackSheets.forEach(sheet => {
    let sheetValues = sheet.getRange(2, 1, 10000, 100).getValues()
    sheetValues = sheetValues.filter(v => v[0])
    syncbackIndexes = [...syncbackIndexes, ...sheetValues]
  })
  return syncbackIndexes
}

function deleteRow(sheet, rowIndex, validateCol, validateValue) {
  if (!rowIndex) return
  if (!validateCol) return
  let values = sheet.getRange(rowIndex, 1, 1, 100).getValues()
  let valueToValidate = values[0][validateCol-1]
  if (validateValue != valueToValidate) {Logger.log(rowIndex + ' no match deletion! col:'+validateCol+', value:'+validateValue); Logger.log(values[0]); return false}
  sheet.deleteRow(rowIndex)
  Logger.log('Row '+rowIndex+' has been deleted! Row data:')
  Logger.log(values[0])
  return true
}


/* Utils */
function getHeaderCol(columnName, sheet = SpreadsheetApp.getActiveSheet()) {
  let headerValues = sheet.getRange(1, 1, 1, 100).getValues()
  let headerColByName = headerValues[0].reduce((pre,cur,i) => {pre[cur] = i+1; return pre}, {})
  return headerColByName[columnName]
}

function getRowByValue(colValue, colName, sheet = SpreadsheetApp.getActiveSheet()) {
  if (!colName) return null
  let col = getHeaderCol(colName, sheet)
  if (!col) return null
  let colValues = sheet.getRange(1, col, 10000, 1).getValues()
  let rowByColValue = colValues.reduce((pre,cur,i) => {pre[cur[0]] = i+1; return pre}, {})
  return rowByColValue[colValue]
}
function getRowByValues(whereArr, sheet = SpreadsheetApp.getActiveSheet()) {
  if (!whereArr.length) return null
  for (var i in whereArr) {
    if (!whereArr[i]['name']) delete whereArr[i]
    let col = getHeaderCol(whereArr[i]['name'], sheet)
    if (!col) delete whereArr[i]
    whereArr[i]['col'] = col
  }
  Logger.log(whereArr)
  let rows = sheet.getRange(1, 1, 10000, 100).getValues()
  return rows.filter(row => {
    return whereArr.reduce((aggr, cur, i) => {
      if (!aggr) return false
      if (row[ cur['col']-1 ] != cur['value']) return false
      return true
    }, true)
  })
}
