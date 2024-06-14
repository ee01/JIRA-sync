/**
 * 记录每次jira相关的修改推送到log表，之后可以用jenkins job去定时刷log表任务来update jira ticket
 * 创建配置表推送到syncback缓存表，当jira接收到修改webhook到时候，可以通过缓存表把jira变动推送回来
 * 
 * Install the test deployment with this script: https://script.google.com/u/0/home/projects/1Fozil1svOmiFilRgNIi0O3iTonXTVnCA4hZtJZuGmJErb2LnJnSi-8Oa/edit
 * 
 * Version: 2024-6-14 （版本更新勿替换Configurations区域）
 * 
 * Author: Esone
 *  */

/* Initial Configurations */
let env = 'test'  // production|test
try { // auto detech env
  env = ScriptApp.getScriptId() == '18Os5bP8YSpMWjxQS8HDB8Q5neD7A9gk_IIBvMyxFFFj11344D_nPZmSe' ? 'production' : 'test'
} catch { env = 'production' }
Logger.log(env + ' environment')
const logSheetURL = env == 'production' ? 'https://docs.google.com/spreadsheets/d/1tWJ9Mr8TdCPaXAawrRT1Opvza-UFvm4TMZ6RO_fyFcA/edit#gid=0' : 'https://docs.google.com/spreadsheets/d/1_V3PkFgh2lVRX2JB5H3oS1cl1FyoKH-mm0PHkcIdooc/edit#gid=937199117'
const syncbackSheetURL = env == 'production' ? 'https://docs.google.com/spreadsheets/d/107ER5MfUeWfTZAKmvMvYwGTaHxI8zcaw3AnlEIAzKNk/edit#gid=0' : 'https://docs.google.com/spreadsheets/d/1Niy_BEmx57TAirrP5wH5CRiXYeflTDIImAUvaimCqY4/edit#gid=0'

/* Configurations End */

/* ------------------------- Replace following code to upgrade ------------------------------------------ */


const ui = SpreadsheetApp.getUi()
const jiraFields = ['summary', 'priority', 'description', 'assignee', 'reporter', 'labels', 'components', 'issuetype', 'duedate', 'fixversions', 'status']
const installProperty = 'install-onEdit_recordChanges'
const userEmail = Session.getActiveUser().getEmail()
const userName =  userEmail.split('@')[0].replace('.', ' ')
const countConfigColumns = 8
let primaryJiraKeyCol = 0
let secondaryJiraKeyCol = null
let primaryJiraFieldMap = null
let secondaryJiraFieldMap = null
let installInfo = null
try {
const Properties = PropertiesService.getDocumentProperties()
installInfo = JSON.parse(Properties.getProperty(installProperty) || null)
if (!installInfo) { // To be remove, 兼容旧版
  installInfo = Properties.getProperty('trigger-onEdit_recordChanges')
  try { installInfo = JSON.parse(installInfo || null) }
  catch { installInfo = {triggerId: installInfo, creator: 'Someone'} }
}
} catch (err) { console.log('Get Properties failed with error: %s', err.message) }
let isInstalled = !!installInfo
let menu = env == 'production' ? ui.createAddonMenu() : ui.createMenu('JIRA sync test')

/* Sync and config */

// Header class
function testHeader() {
  initHeaders()
  Logger.log(primaryJiraKeyCol)
  Logger.log(primaryJiraFieldMap)
  Logger.log(secondaryJiraKeyCol)
  Logger.log(secondaryJiraFieldMap)
}
function initHeaders(dataSheet = null, noCache = false) {
  if (noCache) { primaryJiraFieldMap = null; secondaryJiraFieldMap = null }
  primaryJiraFieldMap = primaryJiraFieldMap || getHeaders(dataSheet)
  secondaryJiraFieldMap = secondaryJiraFieldMap || getHeaders(dataSheet, countConfigColumns + 2)
}
function getHeaders(dataSheet = null, startCol = 1) {
  if (!dataSheet) dataSheet = SpreadsheetApp.getActiveSheet()
  const sheetHeaders = dataSheet.getRange(1, 1, 1, 50).getValues()
  let headers = [...sheetHeaders[0]]

  const configFiledsByColum = getFieldsConfig(dataSheet, headers, startCol)
  for (var col in headers) {
    headers[col] = {name: headers[col], row: 1, col}
    if (headers[col].name == 'JIRA')  {primaryJiraKeyCol = parseInt(col) + 1; continue} // 存储默认JIRA key字段
    if (configFiledsByColum[ headers[col].name ]) {
      headers[col] = configFiledsByColum[ headers[col].name ]
      headers[col].row = 1
      headers[col].col = col
      // 存储JIRA key字段
      if (headers[col].name == 'JIRA') {
        if (startCol == 1) primaryJiraKeyCol = parseInt(col) + 1
        else secondaryJiraKeyCol = parseInt(col) + 1
      }
      // 第二个jira key列需要同步的字段需在第一个同步字段列表中排除
      if (startCol != 1 && headers[col].name != 'JIRA') delete primaryJiraFieldMap[parseInt(col)+1]
      continue
    }
    if (new Set(jiraFields).has(headers[col].name.toLowerCase()) && startCol == 1) continue
    delete headers[col]
  }
  headers.unshift(null)
  // Logger.log(headers)
  return headers
}
function getFieldsConfig(dataSheet = null, headers = [], startCol = 1) {
  if (!dataSheet) dataSheet = SpreadsheetApp.getActiveSheet()
  const configSheetName = dataSheet.getName() + '_config'
  let configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(configSheetName)
  if (!configSheet) {
    if (/.*_config$/.test(dataSheet.getName())) return {}
    // 只检查整张表有创建过config表一次，就不再创建
    const Properties = PropertiesService.getDocumentProperties()
    let didCreateConfigSheet = Properties.getProperty('did-create-config') || false
    if (didCreateConfigSheet) return {}
    // 针对带有 JIRA 表头的第一次修改，创建一张config表给用户配置
    if (!new Set(headers).has('JIRA') && !new Set(headers).has('JIRA ID')) return {}
    _createConfigSheet(configSheetName)
    Logger.log(configSheetName)
    let configSheetNames = Properties.getProperty('did-create-config') || []
    configSheetNames.push(configSheetName)
    Properties.setProperty('did-create-config', configSheetNames);
    return {}
  }

  let fieldsConfig = configSheet.getRange(2, startCol, 100, countConfigColumns).getValues()
  // Logger.log(fieldsConfig)
  let fieldsConfigBySheetcolumn = fieldsConfig.reduce((keyby, x) => {
    keyby[x[0]] = {
      name: x[1],
      desc: x[0],
      syncMode: x[2],
      type: x[3],
      isChangeAsAdding: !!x[4] && x[4] != 0 && x[4] != 'No',
      prefix: x[5],
      suffix: x[6],
      formatFuc: function(value) {
        if (!x[7]) return value
        try{
          return eval(x[7].replace('{value}', '"'+value+'"'))
        }catch{
          return value
        }
      },
    };
    return keyby;
  }, {})
  // Logger.log(fieldsConfigBySheetcolumn)
  return fieldsConfigBySheetcolumn
}

function createConfigSheet() {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const activeSheet = SpreadsheetApp.getActiveSheet()
  const configSheetName = activeSheet.getName() + '_config'
  let configSheet = activeSpreadsheet.getSheetByName(configSheetName)
  if (configSheet) { ui.alert('Config sheet already exist!'); return }
  if (/.*_config$/.test(activeSheet.getName())) { ui.alert('You now stay on a config sheet!'); return }
  _createConfigSheet(configSheetName)
}
function _createConfigSheet(configSheetName) {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  configSheet = activeSpreadsheet.insertSheet(configSheetName)
  configSheet.appendRow(["Sheet Column", "JIRA Field", "Sync mode", "Field type", "Change as adding?", "Prefix", "Suffix", "Format function", "", "Sheet Column for another ticket in same row", "JIRA Field", "Sync mode", "Field type", "Change as adding?", "Prefix", "Suffix", "Format function"]);  // 如修改，请同步修改 countConfigColumns, fieldsConfigBySheetcolumn
  configSheet.getRange(1, 1, 1, 50).setFontWeight("bold")
  configSheet.appendRow(["JIRA ID", "JIRA"]);
  configSheet.appendRow(["Title", "summary", "2-ways", "text"]);
  configSheet.appendRow(["Label", "labels", "To", "list", "Yes"]);
  configSheet.appendRow(["Release", "fixVersions", "2-ways", "list", "No", "mThor "]);
  configSheet.appendRow(["Affect versions", "versions", "2-ways", "list", "No", "mThor "]);
  configSheet.appendRow(["Due date", "duedate", "2-ways", "date"]);
  configSheet.appendRow(["BV", "customfield_10423", "2-ways", "text"]);
  configSheet.appendRow(["Sprint", "customfield_10652", "2-ways", "list", "No"]);
  configSheet.appendRow(["Team", "customfield_17553", "2-ways", "list", "No"]);
  configSheet.appendRow(["Story Point", "customfield_10422", "2-ways", "text"]);
  configSheet.appendRow(["SDK Story Point", "customfield_24666", "2-ways", "text"]);
  configSheet.appendRow(["Vertical Track", "customfield_24174", "2-ways", "list"]);
  configSheet.appendRow(["Assignee", "assignee", "2-ways", "list", "", "", "", "{value}.toLowerCase().replace(' ', '.')"]);
  configSheet.appendRow(["Reporter", "reporter", "2-ways", "list", "", "", "", "{value}.toLowerCase().replace(' ', '.')"]);
  configSheet.appendRow(["Local PM", "customfield_24893", "2-ways", "list", "", "", "", "{value}.toLowerCase().replace(' ', '.')"]);
  configSheet.appendRow(["Dev manday", "customfield_25757", "2-ways", "text"]);
  configSheet.appendRow(["QA manday", "customfield_25958", "2-ways", "text"]);
  configSheet.appendRow(["Target start", "customfield_18350", "2-ways", "date"]);
  configSheet.appendRow(["Target end", "customfield_18351", "2-ways", "date"]);
  configSheet.appendRow(["Exist on Production", "customfield_10570", "2-ways", "text"]);
  configSheet.appendRow(["Affect customers", "customfield_13250", "2-ways", "text"]);
  configSheet.appendRow(["DEA", "customfield_26055", "2-ways", "list", "No"]);
  configSheet.appendRow(["UX Ticket", "links:depends on", "2-ways", "link"]);
  configSheet.appendRow(["Status", "status", "Back", "text"]);
}

function insertJIRAColumn() {
  const activeSheet = SpreadsheetApp.getActiveSheet()
  let JIRAcolumn = activeSheet.insertColumnBefore(1)
  JIRAcolumn.getRange(1, 1).setValue('JIRA')
}


// 主页面
function onHomepage(e) {
  Logger.log('installInfo:')
  Logger.log(installInfo)
  // Logger.log('trigger property:')
  // const Properties = PropertiesService.getDocumentProperties()
  // Logger.log(Properties.getProperty('trigger-onEdit_recordChanges'))
  let myInstallTrigger = getMyInstallTrigger()
  isInstalled = isInstalled || !!myInstallTrigger
  // test环境不会执行 onOpen
  if (env == 'test') if (!isInstalled) menu.addItem('Sync this sheet', 'createSpreadsheetEditTrigger').addToUi()
  else menu.addItem('Stop sync this sheet', 'removeSpreadsheetEditTrigger').addToUi()

  const card = CardService.newCardBuilder()
  const section = CardService.newCardSection()
  section.addWidget(CardService.newTextParagraph().setText("Welcome to JIRA sync!"))
  // if (!isInstalled) section.addWidget(CardService.newDivider()).addWidget(CardService.newDecoratedText().setText('Let\'s make it sync!').setWrapText(true)
  //   .setButton(CardService.newTextButton().setText('Sync').setOnClickAction(CardService.newAction().setFunctionName("createSpreadsheetEditTrigger"))))
  if (!isInstalled) section.addWidget(CardService.newDivider()).addWidget(CardService.newDecoratedText().setText('Start sync from menu: Extensions -> JIRA sync -> Sync this sheet!').setWrapText(true))
  if (isInstalled) section.addWidget(CardService.newDivider()).addWidget(CardService.newDecoratedText().setText('Create config sheet').setWrapText(true)
    .setButton(CardService.newTextButton().setText('Create').setOnClickAction(CardService.newAction().setFunctionName("createConfigSheet"))))
  if (isInstalled) section.addWidget(CardService.newDivider()).addWidget(CardService.newDecoratedText().setText('This sheet already synced. Start with adding a column named "JIRA" with issue id!').setWrapText(true)
    .setButton(CardService.newTextButton().setText('Insert').setOnClickAction(CardService.newAction().setFunctionName("insertJIRAColumn"))))
  if (isInstalled) section.addWidget(CardService.newDecoratedText().setText('Fetch the latest data from JIRA!').setWrapText(true)
    .setButton(CardService.newTextButton().setText('Fetch').setOnClickAction(CardService.newAction().setFunctionName("getIssues"))))
  if (isInstalled) section.addWidget(CardService.newDecoratedText().setText('Expand the sub issues to the Epic!').setWrapText(true)
    .setButton(CardService.newTextButton().setText('Expand').setOnClickAction(CardService.newAction().setFunctionName("expandSubIssues"))))
  if (isInstalled) section.addWidget(CardService.newDivider()).addWidget(CardService.newDecoratedText().setText('If you change the config sheet, please refresh the catch for data sync back!').setWrapText(true)
    .setButton(CardService.newTextButton().setText('Refresh').setOnClickAction(CardService.newAction().setFunctionName("homepage_maintainSyncback"))))
  if (isInstalled) section.addWidget(CardService.newDivider()).addWidget(CardService.newDecoratedText().setText('You can pause sync for a while!').setWrapText(true)
    .setButton(CardService.newTextButton().setText('Pause').setOnClickAction(CardService.newAction().setFunctionName("comingSoon"))))

  let fixedButton = CardService.newTextButton()
  if (isInstalled) {
    fixedButton.setText("Stop sync this sheet")
    if (!myInstallTrigger && installInfo) fixedButton.setDisabled(true).setAltText(`The sync for this sheet is managed by ${installInfo.creator}!`)
    fixedButton.setOnClickAction(CardService.newAction().setFunctionName("removeSpreadsheetEditTrigger"))
  } else fixedButton.setText("Sync this sheet")
            .setOnClickAction(CardService.newAction().setFunctionName("createSpreadsheetEditTrigger"))
  let fixedFooter = CardService.newFixedFooter()
  if (isInstalled && installInfo && !myInstallTrigger) fixedFooter.setSecondaryButton(CardService.newTextButton().setText("i").setOnClickAction(CardService.newAction().setFunctionName("alertInstalledByOther")))
  // 从toolbar安装后的 edit trigger 别人触发不了。menu安装的可以
  if (!isInstalled) {
    fixedButton.setDisabled(true)
    fixedFooter.setSecondaryButton(CardService.newTextButton().setText("i").setOnClickAction(CardService.newAction().setFunctionName("alertInstallTips")))
  }
  fixedFooter.setPrimaryButton(fixedButton)
  card.setFixedFooter(fixedFooter)
  card.addSection(section)

  return card.build()
}
function alertInstalledByOther() {
  if (!installInfo) return
  ui.alert(`This sheet has been enabled the sync by ${installInfo.creator} at ${installInfo.date}.
    To stop the sync, please reach out to ${installInfo.creator}!`)
}
function alertInstallTips() {
  ui.alert(`Please start sync from the menu: 
    Extensions -> JIRA sync -> Sync this sheet!`)
}
function comingSoon() {
  ui.alert(`Coming soon!`)
}


// Triggers
function onInstall(e) {
  menu.addItem('Click to start!', 'openHomepage').addToUi()
  createSpreadsheetEditTrigger();
}
function onOpen(e) {
  Logger.log(userEmail)
  Logger.log('authMode:')
  Logger.log(e.authMode)
  try {
    Logger.log({'sheet name:': SpreadsheetApp.getActiveSpreadsheet().getName(), 'sheet tab': SpreadsheetApp.getActiveSheet().getName(), 'sheet url:': SpreadsheetApp.getActiveSpreadsheet().getUrl()})
  } catch {}
  if (e && e.authMode != ScriptApp.AuthMode.NONE) {
    if (!isInstalled) {
      menu.addItem('Sync this sheet', 'createSpreadsheetEditTrigger').addToUi()
    } else {
      menu.addItem('Stop sync this sheet', 'removeSpreadsheetEditTrigger').addToUi()
      menu.addItem('Create config sheet', 'createConfigSheet').addToUi()
      menu.addItem('View the sync logs', 'openTheLogSheet').addToUi()
    }
  } else menu.addItem('Sync this sheet!', 'createSpreadsheetEditTrigger').addToUi();

  menu.addSeparator()
  menu.addItem('More..', 'openHomepage').addToUi()
  // ui.createAddonMenu().addItem('test menu', 'openHomepage').addToUi() // To be removed
}
function onFileScopeGrantedSheets(e) {
}
function openTheLogSheet() {
  let html = '<a href="' + logSheetURL + '" target="_blank">Click to open!</a>';
  let userInterface = HtmlService.createHtmlOutput(html)
                               .setWidth(300)
                               .setHeight(100);
  ui.showModalDialog(userInterface, 'Log sheet');
}
function openHomepage() {
  // var html = HtmlService.createHtmlOutput().setContent('Hello, world! <input type="button" value="Close" onclick="google.script.host.close()" />')
  //     .setTitle('My custom sidebar');
  // ui.showSidebar(html)
  // CardService.newActionResponseBuilder()
  //       .setNavigation(CardService.newNavigation().popToRoot())
  //       .build();
  // var card = onHomepage(); // 假设 onHomepage() 函数返回一个 CardService.newCardBuilder().build()
  // var userInterface = CardService.newCardService()
  //                      .createCardFromBuilder(card)
  //                      .build();
  // ui.showModalDialog(userInterface, 'Add-on Homepage');
  return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().updateCard(onHomepage()))
        .build();
}

// Trigger class
function getMyInstallTrigger() {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  Logger.log('User triggers:')
  Logger.log(ScriptApp.getUserTriggers(activeSpreadsheet).map(v => ({sourece: v.getTriggerSource(), sid: v.getTriggerSourceId(), func:v.getHandlerFunction(), type: v.getEventType(), id: v.getUniqueId()})))
  return ScriptApp.getUserTriggers(activeSpreadsheet).find(t => t.getHandlerFunction() == 'onEdit_recordChanges') || null
}
function createSpreadsheetEditTrigger() {
  if (installInfo && installInfo.creatorEmail && installInfo.creatorEmail != userEmail) {
    ui.alert(`This sheet has been already enabled the sync by ${installInfo.creator} at ${installInfo.date}.`)
    return
  }

  const Properties = PropertiesService.getDocumentProperties()
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  let myInstallTrigger = ScriptApp.newTrigger('onEdit_recordChanges')
      .forSpreadsheet(activeSpreadsheet)
      .onEdit()
      .create()
  installInfo = {
    creator: userName,
    creatorEmail: userEmail,
    triggerId: myInstallTrigger.getUniqueId(),
    date: new Date().toLocaleString(),
  }
  Properties.setProperty(installProperty, JSON.stringify(installInfo))
  isInstalled = !!myInstallTrigger

  let result = ui.alert(
     'Congratulations!',
     'This sheet has been sync to JIRA! Please make sure you have [JIRA] column to present JIRA ticket id.',
      ui.ButtonSet.OK);
  if (result == ui.Button.OK) {
    // Process the user's response.
    menu.addItem('Stop sync this sheet', 'removeSpreadsheetEditTrigger')
    .addToUi();
  }

  return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().updateCard(onHomepage()))
        .build();
}
function removeSpreadsheetEditTrigger() {
  let myInstallTrigger = getMyInstallTrigger()
  if (myInstallTrigger) ScriptApp.deleteTrigger(myInstallTrigger)
  else if (installInfo && installInfo.creatorEmail && installInfo.creatorEmail != userEmail) {
    ui.alert(`This sheet was synced by ${installInfo.creator}.
      Please reach out to ${installInfo.creator} to manage the sync!`)
    return
  }

  const Properties = PropertiesService.getDocumentProperties()
  if (installInfo && installInfo.creatorEmail == userEmail) Properties.deleteProperty(installProperty)
  if (Properties.getProperty('trigger-onEdit_recordChanges')) Properties.deleteProperty('trigger-onEdit_recordChanges') // To be removed 兼容旧版
  isInstalled = false
  menu.addItem('Sync this sheet', 'createSpreadsheetEditTrigger').addToUi()
  ui.alert('This sheet now will not sync with JIRA ever.')

  return CardService.newNavigation().updateCard(onHomepage())
}

// function onEdit(e) {
  // onEdit_recordChanges(e)
// }

function onEdit_recordChanges(e) {
  Logger.log({user: userEmail, sheet: SpreadsheetApp.getActiveSpreadsheet().getName(), tab: SpreadsheetApp.getActiveSheet().getName(), col: e.range.getColumn(), row: e.range.getRow()})
  recordChanges(e)
  onEdit_maintainSyncback(e)
  runOnEditQueue_maintainSyncback()
}

// Queue class
function runOnEditQueue_maintainSyncback() {
  const Properties = PropertiesService.getDocumentProperties()
  let queue = []
  try { queue = JSON.parse(Properties.getProperty('queue-to-run')) || [] } catch {}
  Logger.log('queue:')
  Logger.log(queue)
  paramsToQueueFunc = {}
  queue.forEach((job,i) => {
    if (!job.time) return
    Logger.log(new Date(job.time).toLocaleString())
    if (new Date(job.time).getTime() > new Date().getTime()) return
    if (!job.functionName) return
    if (job.params) paramsToQueueFunc[job.functionName] = job.params || []
    try {
      eval(job.functionName + '(...paramsToQueueFunc.'+job.functionName+')')
      queue.splice(i, 1)
      Properties.setProperty('queue-to-run', JSON.stringify(queue))
    }
    catch { Logger.log('Run queue function failed: ' + job.functionName) }
  })
}
function setQueue(functionName, time, ...params) {
  const Properties = PropertiesService.getDocumentProperties()
  let queue = []
  try { queue = JSON.parse(Properties.getProperty('queue-to-run')) || [] } catch {}
  let job = queue.find(j => j.functionName == functionName)
  if (job) {
    if (new Date(job.time).getTime() > new Date().getTime()) job.time = time
    job.params = params
    for (var i in queue) if (queue[i].functionName == functionName) queue[i] = job
  } else queue.push({
    functionName,
    params,
    time,
  })
  Properties.setProperty('queue-to-run', JSON.stringify(queue))
}


/* Fetch JIRA data */

function getIssues() {
  const dataSheet = SpreadsheetApp.getActiveSheet()
  if (/.*_config$/.test(dataSheet.getName())) {ui.alert('You are not at config sheet, please swtich to data sheet!'); return}
  initHeaders()
  if (!primaryJiraKeyCol) {ui.alert('Here is not data sheet, or config one mapping to "JIRA"!'); return}


  let range = dataSheet.getActiveRange()
  for (var row = range.getRow(); row < range.getRow() + range.getNumRows(); row++) {
    let primaryJiraValue = dataSheet.getRange(row, primaryJiraKeyCol).getValue()
    if (!primaryJiraValue) continue
    for (var column in primaryJiraFieldMap) {
      if (column == primaryJiraKeyCol) continue
      if (!primaryJiraFieldMap[column]) continue
      Logger.log({row, column, primaryJiraValue})
      _syncDataToLogSheet({
        row,
        column,
        id: primaryJiraValue,
        field: primaryJiraFieldMap[column].name,
        fieldDesc: primaryJiraFieldMap[column].desc,
        type: primaryJiraFieldMap[column].type,
        oldValue: dataSheet.getRange(row, column).getValue(),
        newValue: "",
        isChangeAsAdding: primaryJiraFieldMap[column].isChangeAsAdding,
        prefix: '',
        suffix: '',
        time: new Date().getTime(),
      }, "sheet", "get")
    }
  }
  ui.alert("Please wait a min for data fetching!")
}
function expandSubIssues() {
  const dataSheet = SpreadsheetApp.getActiveSheet()
  if (/.*_config$/.test(dataSheet.getName())) {ui.alert('You are not at config sheet, please swtich to data sheet!'); return}
  initHeaders()
  if (!primaryJiraKeyCol) {ui.alert('Here is not data sheet, or config one mapping to "JIRA"!'); return}

  let range = dataSheet.getActiveRange()
  let lastRow = range.getRow() + range.getNumRows() - 1
  for (var row = range.getRow(); row <= lastRow; row++) {
    let primaryJiraValue = dataSheet.getRange(row, primaryJiraKeyCol).getValue()
    Logger.log({row, lastRow, primaryJiraValue})
    if (!primaryJiraValue) continue
    dataSheet.insertRowAfter(row)
    row++
    lastRow++
    let newRowAfterPrimaryJiraFirstCell = dataSheet.getRange(row, primaryJiraKeyCol)
    newRowAfterPrimaryJiraFirstCell.shiftRowGroupDepth(1)
    _syncDataToLogSheet({
      row: row,
      column: primaryJiraKeyCol,
      id: primaryJiraValue,
      field: "issuesInEpic",
      fieldDesc: "Issues in Epic",
      type: "",
      oldValue: "",
      newValue: "",
      prefix: "",
      suffix: "",
      time: new Date().getTime(),
    }, "sheet", "getSubissuesInsert")
  }
  ui.alert("Empty row placeholder added. Please wait a min for data fetching!")
}


/* Log to changelog sheet */

// 记录每次jira相关的修改推送到log表
function recordChanges(e) {
  const activeSheet = SpreadsheetApp.getActiveSheet()
  const range = e.range;
  const column = range.getColumn();
  const row = range.getRow();
  const isMultiple = !!(range.getNumRows() || range.getNumColumns())
  let firstValue = e.value || (e.oldValue?null:range.getValue()) // 粘帖和撤销异常值 - 
    // Copy:{e.value:null, e.oldValue:null, range.value:"test"} - range为准，可能多列
    // Delete:{e.value:null, e.oldValue:"test", range.value:null} - 可能多列
    // Undo:{e.value:null, e.oldValue:null, range.value:"test"} - range为准，可能多列
    // 快速Delete&Undo:{e.value:null, e.oldValue:"test", range.value:"test"} - range为准，可能多列
    // 拖拽自动填充:{e.value:null, e.oldValue:null, range.value:"test"} - range为准，可能多列
  if (/.*_config$/.test(activeSheet.getName())) {Logger.log('Config sheet, no sync!'); return}
  if (row == 1) {Logger.log('Header change, no sync!'); return}
  initHeaders()
  if (!primaryJiraKeyCol) {Logger.log('No specific column defaultly named "JIRA", or config one mapping to "JIRA"!'); return}
  
  // 推送第一字段列表的变化给JIRA
  !function(){
    if (!isMultiple) {
      if (!e.value) return;
      // if (!e.oldValue) return;
      if (!primaryJiraFieldMap[column]) {Logger.log('No JIRA mapping field change!'); return}
      if (column == primaryJiraKeyCol) {Logger.log('No sync on changing JIRA key column!'); return}
      const _jiraKey = range.getSheet().getRange(row, primaryJiraKeyCol);
      if (!_jiraKey.getValue()) {Logger.log('No specific JIRA key!'); return}

      let data = getData(_jiraKey.getValue(), e.oldValue, e.value, primaryJiraFieldMap)
      Logger.log(data)
      _syncDataToLogSheet(data)
    } else {
      let values = range.getValues()
      for (var row = range.getRow(); row < range.getRow() + range.getNumRows(); row++) {
        for (var column = range.getColumn(); column < range.getColumn() + range.getNumColumns(); column++) {
          let value = values[row-range.getRow()][column-range.getColumn()]
          Logger.log({value_range: value, row, column})
          if (!primaryJiraFieldMap[column]) {Logger.log('Row:'+row+' Column:'+column + '. No JIRA mapping field change!'); continue}
          if (column == primaryJiraKeyCol) {Logger.log('Row:'+row+' Column:'+column + '. No sync on changing JIRA key column!'); continue}
          if (!value) continue
          const _jiraKey = range.getSheet().getRange(row, primaryJiraKeyCol);
          if (!_jiraKey.getValue()) {Logger.log('Row:'+row+' Column:'+column + '. No specific JIRA key!'); continue}

          let data = getData(_jiraKey.getValue(), row==range.getRow()&&column==range.getColumn()?e.oldValue:'', value, primaryJiraFieldMap, row, column)
          Logger.log(data)
          _syncDataToLogSheet(data)
        }
      }
    }
  }()

  // 推送第二字段列表的变化给JIRA
  !function(){
    if (!isMultiple) {
      if (!e.value) return;
      // if (!e.oldValue) return;
      if (secondaryJiraKeyCol === null) return;
      if (!secondaryJiraFieldMap[column]) return;
      if (column == secondaryJiraKeyCol) return;
      const _jiraKey = range.getSheet().getRange(row, secondaryJiraKeyCol);
      if (!_jiraKey.getValue()) return;

      let data = getData(_jiraKey.getValue(), e.oldValue, e.value, secondaryJiraFieldMap)
      Logger.log(data)
      _syncDataToLogSheet(data)
    } else {
      let values = range.getValues()
      for (var row = range.getRow(); row < range.getRow() + range.getNumRows(); row++) {
        for (var column = range.getColumn(); column < range.getColumn() + range.getNumColumns(); column++) {
          let value = values[row-range.getRow()][column-range.getColumn()]
          if (!secondaryJiraFieldMap[column]) continue
          if (column == secondaryJiraKeyCol) continue
          if (!value) continue
          const _jiraKey = range.getSheet().getRange(row, secondaryJiraKeyCol);
          if (!_jiraKey.getValue()) continue

          let data = getData(_jiraKey.getValue(), row==range.getRow()&&column==range.getColumn()?e.oldValue:'', value, secondaryJiraFieldMap, row, column)
          Logger.log(data)
          _syncDataToLogSheet(data)
        }
      }
    }
  }()

  function getData(id, oldValue, newValue, JiraFieldMap = primaryJiraFieldMap, row = row, column = column) {
    // Format with custom function
    newValue = JiraFieldMap[column].formatFuc ? JiraFieldMap[column].formatFuc(newValue) : newValue
    // Format date
    newValue = JiraFieldMap[column].type == 'date' ? formatDate(newValue) : newValue

    return {
      row,
      column,
      id: id,
      field: JiraFieldMap[column].name,
      fieldDesc: JiraFieldMap[column].desc,
      type: JiraFieldMap[column].type,
      oldValue: oldValue,
      newValue: newValue,
      isChangeAsAdding: JiraFieldMap[column].isChangeAsAdding,
      prefix: JiraFieldMap[column].prefix || '',
      suffix: JiraFieldMap[column].suffix || '',
      time: new Date().getTime(),
    }
  }
}
function _syncDataToLogSheet(data, from = "sheet", action = "") {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const activeSheet = SpreadsheetApp.getActiveSheet()
  const logSS = SpreadsheetApp.openByUrl(logSheetURL)
  const logSheetName = activeSpreadsheet.getName() + ': ' + activeSheet.getName()
  let logSheet = logSS.getSheetByName(logSheetName)
  if (!logSheet) {
    logSheet = logSS.insertSheet(logSheetName)
    logSheet.appendRow(["editor", "from", "action", "old value", "new value", "sheet name", "sheet URL", "sheet tab", "sheet tab gid", "sheet row", "sheet column", "JIRA key", "JIRA field desc", "JIRA field name", "JIRA field type", "time", "isSync", "sync time", "took seconds", "fail reason"]);
  }

  logSheet.appendRow([userEmail, from, action||(data.isChangeAsAdding?'add':'replace'), data.oldValue, data.prefix+data.newValue+data.suffix, activeSpreadsheet.getName(), activeSpreadsheet.getUrl()+'#gid='+activeSheet.getSheetId(), activeSheet.getName(), activeSheet.getSheetId(), data.row, data.column, `=HYPERLINK("https://jira.ringcentral.com/browse/${data.id}", "${data.id}")`, data.fieldDesc, data.field, data.type, new Date().toLocaleString()]);
  Logger.log('Sync log successfully!\n' + logSheetURL)
}

/* Maintain sync back index sheet */

// 修改配置表的时候，维护一张 sync back 表 供 JIRA webhook 调用的时候索引 Tickets
function onEdit_maintainSyncback(e) {
  const range = e.range;
  const column = range.getColumn();
  const row = range.getRow();
  if (row == 1) {Logger.log('Header change, quit!'); return}

  // Config sheet changed
  !function(){
    if (!e.value) return;
    const configSheet = SpreadsheetApp.getActiveSheet()
    const dataSheetName = configSheet.getName().replace(/_config$/, "")
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheetName)
    if (!dataSheet) {Logger.log('No found data sheet, quit!'); return}
    if (!/.*_config$/.test(configSheet.getName())) return
    if (column != 1 && column != 1 + 1
      && column != 2 + countConfigColumns && column != 2 + countConfigColumns + 1) {Logger.log('Auto update syncback config is only trigger by field name changes!'); return}

    // Throttle
    setQueue('maintainSyncback', new Date().getTime() + 2 * 1000, dataSheet.getName())
    // maintainSyncback(dataSheet)
  }()

  // Data sheet new JIRA key added
  !function(){
    if (!e.value) return;
    if (e.oldValue) return; // only add JIRA key
    const dataSheet = SpreadsheetApp.getActiveSheet()
    if (/.*_config$/.test(dataSheet.getName())) return
    if (column != primaryJiraKeyCol && column != secondaryJiraKeyCol) {Logger.log('Auto update syncback config is only trigger by JIRA key added!'); return}

    // Throttle
    setQueue('maintainSyncback', new Date().getTime() + 2 * 1000, dataSheet.getName())
  }()
}
function homepage_maintainSyncback() {
  if (maintainSyncback()) ui.alert('Refresh catch successfully!')
}
function maintainSyncback(dataSheet = null) {
  if (!dataSheet) {
    const activeSheet = SpreadsheetApp.getActiveSheet()
    if (!/.*_config$/.test(activeSheet.getName())) {ui.alert('You are currently not in the config sheet!'); return}
    const dataSheetName = activeSheet.getName().replace(/_config$/, "")
    dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheetName)
  } else if (typeof dataSheet == 'string') dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet)

  initHeaders(dataSheet, true)
  // Logger.log(primaryJiraFieldMap)
  let dataValues = dataSheet.getRange(2, 1, 500, primaryJiraFieldMap.length).getValues()
  let tickets = []
  for (var r in dataValues) {
    if (!dataValues[r][primaryJiraKeyCol-1]) continue
    // Logger.log(dataValues[r])
    for (var c in dataValues[r]) {
      const primaryJiraField = primaryJiraFieldMap[parseInt(c)+1]
      if (!primaryJiraField) continue
      if (parseInt(c)+1 == primaryJiraKeyCol) continue
      tickets.push({
        jiraKey: dataValues[r][primaryJiraKeyCol-1],
        row: parseInt(r) + 1 + 1,
        column: parseInt(c) + 1,
        fieldName: primaryJiraField.name,
        fieldDesc: primaryJiraField.desc,
        fieldType: primaryJiraField.type,
        isListAllValue: primaryJiraField.isChangeAsAdding ? 'editing' : 'all',  // Todo: 'max'
        removePrefix: primaryJiraField.prefix,
        removeSuffix: primaryJiraField.suffix,
        // backFormatFunc: primaryJiraField.formatFuc, // Todo
      })
    }
  }
  // Logger.log(JSON.stringify(tickets))
  return _syncTicketsToSyncbackSheet(tickets, dataSheet)
}
function _syncTicketsToSyncbackSheet(tickets = [], dataSheet) {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const syncbackSS = SpreadsheetApp.openByUrl(syncbackSheetURL)
  const syncbackSheetName = activeSpreadsheet.getName() + ': ' + dataSheet.getName()
  let syncbackSheet = syncbackSS.getSheetByName(syncbackSheetName)
  if (!syncbackSheet) syncbackSheet = syncbackSS.insertSheet(syncbackSheetName)

  syncbackSheet.clear()
  syncbackSheet.appendRow(["owner", "JIRA key", "JIRA field desc", "JIRA field name", "JIRA field type", "sheet name", "sheet URL", "sheet tab", "sheet tab gid", "sheet row", "sheet column", "list all value", "remove prefix", "remove suffix", "back format func", "last edit back time", "fail reason"]);
  if (!tickets.length) { Logger.log('No data with config to sync!'); return }

  // tickets.forEach(ticket => {
  //   syncbackSheet.appendRow([userEmail, `=HYPERLINK("https://jira.ringcentral.com/browse/${ticket.jiraKey}", "${ticket.jiraKey}")`, ticket.fieldDesc, ticket.fieldName, ticket.fieldType, activeSpreadsheet.getName(), activeSpreadsheet.getUrl()+'#gid='+dataSheet.getSheetId(), dataSheet.getName(), dataSheet.getSheetId(), ticket.row, ticket.column, ticket.isListAllValue, ticket.prefix, ticket.suffix, ticket.backFormatFunc]);
  // })
  let rangeValues = tickets.map(ticket => [userEmail, `=HYPERLINK("https://jira.ringcentral.com/browse/${ticket.jiraKey}", "${ticket.jiraKey}")`, ticket.fieldDesc, ticket.fieldName, ticket.fieldType, activeSpreadsheet.getName(), activeSpreadsheet.getUrl()+'#gid='+dataSheet.getSheetId(), dataSheet.getName(), dataSheet.getSheetId(), ticket.row, ticket.column, ticket.isListAllValue, ticket.prefix, ticket.suffix, ticket.backFormatFunc])
  syncbackSheet.getRange(2, 1, rangeValues.length, rangeValues[0].length).setValues(rangeValues)
  Logger.log('Sync syncback config successfully!\n' + syncbackSheetURL)
  return true
}

/* Utils */
function formatDate(value) {
  let date = convertToDateWithCurrentYearIfNoYear(value)
  if (value.length && value.length == parseInt(value).toString().length) {
    return Intl.DateTimeFormat('en-CA', {dateStyle:'short'}).format(valueToDate(value))
  } else if (date) return Intl.DateTimeFormat('en-CA', {dateStyle:'short'}).format(date)
  else return value

  function valueToDate(GoogleDateValue) {
    return new Date(new Date(1899,11,30+Math.floor(GoogleDateValue),0,0,0,0).getTime()+(GoogleDateValue%1)*86400000)
  }
  function convertToDateWithCurrentYearIfNoYear(dateStr) {
    const currentYear = new Date().getFullYear();
    if (!/\d{4}/.test(dateStr)) { // 检查是否包含年份
      dateStr = dateStr.replace(/^(\d{1,2})([^\d]+)(\d{1,2})/, `$1/$3/${currentYear}`); // 没有年份则添加当前年份
        console.log(dateStr)
    }
    const date = new Date(dateStr);
    return date.toString() === 'Invalid Date' ? null : date; // 如果日期无效，返回null
  }
}