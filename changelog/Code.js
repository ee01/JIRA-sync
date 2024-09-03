/**
 * 定时轮询 JIRA webhook 的数据，清除无关联的，并生成有关联的到 changelog 等待执行数据推送
 * Install: 添加 filterChangesNoRelatedToSheets() 每分钟执行一次
 * 
 * 
 * Version: 2024-8-30
 * 
 * Author: Esone
 *  */

const syncbackSheetURL = 'https://docs.google.com/spreadsheets/d/107ER5MfUeWfTZAKmvMvYwGTaHxI8zcaw3AnlEIAzKNk/edit#gid=0'
const backupSheetURL = 'https://docs.google.com/spreadsheets/d/1qYptY_YHEw-Pr75UkH4dvm4g_Z0cwubxwXH2CzSus0Q/edit?gid=0#gid=0'

const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
const changelogSheetName = 'changelog'
const changelogSheet = activeSpreadsheet.getSheetByName(changelogSheetName)
const colSheetURL = getHeaderCol('sheet URL', changelogSheet)
const colIsSync_log = getHeaderCol('isSync', changelogSheet)
const jiraWebhookSheetName = 'jira_webhook_data'
const jiraWebhookSheet = activeSpreadsheet.getSheetByName(jiraWebhookSheetName)
const colFrom = getHeaderCol('from', jiraWebhookSheet)
const colAction = getHeaderCol('action', jiraWebhookSheet)
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
    /* Deprecated: Keep sync.service in case to sync changes to other sheets
    if (log[colEditor-1] == 'sync.service@ringcentral.com') {
      jiraWebhookSheet.getRange(log['rowIndex'], colIsSync).setValue('Failed')
      jiraWebhookSheet.getRange(log['rowIndex'], colSyncTime).setValue(new Date())
      jiraWebhookSheet.getRange(log['rowIndex'], colTookSeconds).setValue(Math.ceil((new Date().getTime() - new Date(log[colTime-1]).getTime()) / 1000))
      jiraWebhookSheet.getRange(log['rowIndex'], colFailReason).setValue('Ignore sync.service to avoid running into loop!')
      return
    } */
    locationInSheets = getLocationInSheets(log[colJIRAKey-1], log[colJIRAFieldName-1])
    if (!locationInSheets || locationInSheets.length <= 0) {
      jiraWebhookSheet.getRange(log['rowIndex'], colIsSync).setValue('Failed')
      jiraWebhookSheet.getRange(log['rowIndex'], colSyncTime).setValue(new Date())
      jiraWebhookSheet.getRange(log['rowIndex'], colTookSeconds).setValue(Math.ceil((new Date().getTime() - new Date(log[colTime-1]).getTime()) / 1000))
      jiraWebhookSheet.getRange(log['rowIndex'], colFailReason).setValue('No mapping tickets in syncback index sheet!')
    } else {
      locationInSheets.forEach(location => {
        let newValue = log[colNewValue-1].replace(location['remove prefix'], '').replace(location['remove suffix'], '')
        let backFormatFuc = function(value) {
          if (!location['back format func']) return value
          try{
            return eval(location['back format func'].replaceAll('{value}', '"'+value+'"'))
          }catch{
            return value
          }
        }
        newValue = backFormatFuc ? backFormatFuc(newValue) : newValue
        if (newValue == log[colOldValue-1]) return
        changelogSheet.appendRow([log[colEditor-1], log[colFrom-1], log[colAction-1], log[colOldValue-1], newValue, location['sheet name'], location['sheet URL'], location['sheet tab'], location['sheet tab gid'], location['sheet row'], location['sheet column'], location['sheet key header'], `=HYPERLINK("https://jira.ringcentral.com/browse/${log[colJIRAKey-1]}", "${log[colJIRAKey-1]}")`, location['JIRA field desc'], log[colJIRAFieldName-1], location['JIRA field type'], new Date()])
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


/* Clean and backup old data */
function cleanData() {
  _cleanData(changelogSheetName, 1000, 500, ["editor", "from", "action", "old value", "new value", "sheet name", "sheet URL", "sheet tab", "sheet tab gid", "sheet row", "sheet column", "sheet key header", "JIRA key", "JIRA field desc", "JIRA field name", "JIRA field type", "time", "isSync", "sync time", "took seconds", "fail reason"], colIsSync_log)
  _cleanData(jiraWebhookSheetName, 2000, 1000, ["editor", "from", "action", "old value", "new value", "JIRA key", "JIRA field desc", "JIRA field name", "JIRA field type", "time", "isSync", "sync time", "took seconds", "fail reason"], colIsSync)
}
function _cleanData(backupSheetName, overCapacity = 1000, cleanRows = 500, headers, isSyncColumn) {
  // 检测备份条件
  const cleanSheet = activeSpreadsheet.getSheetByName(backupSheetName)
  let logs1000_1010 = cleanSheet.getRange(overCapacity, 1, 10, 21).getValues()
  logs1000_1010 = logs1000_1010.filter(log => {
    if (!log[0]) return false
    if (!log[isSyncColumn-1]) return false
    return true
  })
  if (logs1000_1010.length < 5) return

  // 读取备份表
  const backupSS = SpreadsheetApp.openByUrl(backupSheetURL)
  let backupSheet = backupSS.getSheetByName(backupSheetName)
  if (!backupSheet) {
    backupSheet = backupSS.insertSheet(backupSheetName)
    backupSheet.appendRow(headers);
  }
  let backupDataLength = backupSheet.getDataRange().getValues().length

  // 备份
  let logs500 = cleanSheet.getRange(2, 1, cleanRows, 21).getValues()
  logs500 = logs500.filter(log => log[0])
  backupSheet.getRange(backupDataLength+1, 1, logs500.length, 21).setValues(logs500)
  Logger.log('Backup '+logs500.length+' rows in ['+backupSheetName+'] to '+backupSheetURL)

  // 清除
  cleanSheet.deleteRows(2, cleanRows)
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
