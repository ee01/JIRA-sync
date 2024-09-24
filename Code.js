/**
 * 记录每次jira相关的修改推送到log表，之后可以用jenkins job去定时刷log表任务来update jira ticket
 * 创建配置表推送到syncback缓存表，当jira接收到修改webhook到时候，可以通过缓存表把jira变动推送回来
 * 
 * Install the test deployment with this script: https://script.google.com/u/0/home/projects/1Fozil1svOmiFilRgNIi0O3iTonXTVnCA4hZtJZuGmJErb2LnJnSi-8Oa/edit
 * 
 * Version: 2024-9-13 （版本更新勿替换Configurations区域）
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
const jiraGetDataWebhook = env == 'production' ? 'https://jira.ringcentral.com/rest/cb-automation/latest/hooks/9edd6e55ec9b7da28206ab927562da913f5532bf' : 'https://jira.ringcentral.com/rest/cb-automation/latest/hooks/9edd6e55ec9b7da28206ab927562da913f5532bfhttps://jira.ringcentral.com/rest/cb-automation/latest/hooks/5697bd573396a29a190ffe78e6eb88d74a5cf252'
const editorEmail = 'jirasheetsyncer@jirasheetsyncer.iam.gserviceaccount.com'
const editorBackupEmail = 'sync.service@ringcentral.com'

/* Configurations End */

/* ------------------------- Replace following code to upgrade ------------------------------------------ */


const logSheetName = 'changelog'
const dataGettingSheetName = 'get_jira_data'
const jiraFields = ['summary', 'priority', 'description', 'assignee', 'reporter', 'labels', 'components', 'issuetype', 'duedate', 'fixversions', 'status']  // Deprecated: Use config sheet to manage
const installProperty = 'install-onEdit_recordChanges'
const userEmail = Session.getActiveUser().getEmail()
const userName =  userEmail.split('@')[0].replace('.', ' ')
const countConfigColumns = 9
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
let menu = null

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
    if (!headers[col]) continue
    headers[col] = {name: headers[col], row: 1, col: col+1}
    if (headers[col].name == 'JIRA key')  {primaryJiraKeyCol = parseInt(col) + 1; continue} // 存储默认JIRA key字段
    if (configFiledsByColum[ headers[col].name ]) {
      headers[col] = configFiledsByColum[ headers[col].name ]
      headers[col].row = 1
      headers[col].col = parseInt(col) + 1
      // 存储JIRA key字段
      if (headers[col].name == 'JIRA key') primaryJiraKeyCol = parseInt(col) + 1
      else if (headers[col].name == 'link key') secondaryJiraKeyCol = parseInt(col) + 1
      continue
    }
    // if (new Set(jiraFields).has(headers[col].name.toLowerCase()) && startCol == 1) continue // 通过预设字段自动同步
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
    if (!new Set(headers).has('JIRA') && !new Set(headers).has('JIRA key')) return {}
    _createConfigSheet(configSheetName)
    Logger.log(configSheetName)
    let configSheetNames = JSON.parse(Properties.getProperty('did-create-config')) || []
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
      format: x[7],
      formatFuc: function(value) {
        if (!x[7]) return value
        try{
          return eval(x[7].replaceAll('{value}', '"'+value+'"'))
        }catch{
          return value
        }
      },
      backFormat: x[8],
      backFormatFuc: function(value) {
        if (!x[8]) return value
        try{
          return eval(x[8].replaceAll('{value}', '"'+value+'"'))
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
  const ui = SpreadsheetApp.getUi()
  const Properties = PropertiesService.getDocumentProperties()
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const activeSheet = SpreadsheetApp.getActiveSheet()
  const configSheetName = activeSheet.getName() + '_config'
  let configSheet = activeSpreadsheet.getSheetByName(configSheetName)
  if (configSheet) { ui.alert('Config sheet already exist!'); return }
  if (/.*_config$/.test(activeSheet.getName())) { ui.alert('You now stay on a config sheet!'); return }
  _createConfigSheet(configSheetName)
  let configSheetNames = JSON.parse(Properties.getProperty('did-create-config')) || []
  configSheetNames.push(configSheetName)
  Properties.setProperty('did-create-config', configSheetNames);
}
function _createConfigSheet(configSheetName) {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  configSheet = activeSpreadsheet.insertSheet(configSheetName)
  configSheet.appendRow(["Sheet Column", "JIRA Field", "Sync mode", "Field type", "Change as adding?", "Prefix", "Suffix", "Format function", "Back format", "", "Sheet Column - link ticket", "JIRA Field", "Sync mode", "Field type", "Change as adding?", "Prefix", "Suffix", "Format function", "Back format"]);  // 如修改，请同步修改 countConfigColumns, fieldsConfigBySheetcolumn
  configSheet.getRange(1, 1, 1, 50).setFontWeight("bold")
  configSheet.appendRow(["JIRA", "JIRA key", "", "", "", "", "", "", "", "", "UX Ticket", "link key"]);
  configSheet.appendRow(["Title", "summary", "Back", "text"]);
  configSheet.appendRow(["Type", "issuetype", "Back", "text"]);
  configSheet.appendRow(["Label", "labels", "To", "list", "Yes"]);
  configSheet.appendRow(["Component", "components", "2-ways", "list", "No"]);
  configSheet.appendRow(["Release", "fixVersions", "2-ways", "list", "No", "", "", "{value}=='Video Wishlist'?{value}:'mThor '+{value}", '{value}.split(",").map(re => re.replace(/[^\\d]*/, "").replace(/.*\\W(\\d+\\.\\d+\\.\\d+)/, "$1")).reduce((aggr, cur) => aggr>cur?aggr:cur, -Infinity)']);
  configSheet.appendRow(["Affect versions", "versions", "2-ways", "list", "No", "mThor "]);
  configSheet.appendRow(["Due date", "duedate", "2-ways", "date"]);
  configSheet.appendRow(["BV", "customfield_10423", "2-ways", "text"]);
  configSheet.appendRow(["Priority", "Priority", "2-ways", "text"]);
  configSheet.appendRow(["Sprint", "customfield_10652", "2-ways", "list", "No"]);
  configSheet.appendRow(["Team", "customfield_17553", "2-ways", "list", "No"]);
  configSheet.appendRow(["Story Point", "customfield_10422", "2-ways", "text"]);
  configSheet.appendRow(["SDK Story Point", "customfield_24666", "2-ways", "text"]);
  configSheet.appendRow(["Vertical Track", "customfield_24174", "2-ways", "list"]);
  configSheet.appendRow(["Assignee", "assignee", "2-ways", "list", "", "", "", "{value}.toLowerCase().replace(' ', '.')"]);
  configSheet.appendRow(["Reporter", "reporter", "2-ways", "list", "", "", "", "{value}.toLowerCase().replace(' ', '.')"]);
  configSheet.appendRow(["Local PM", "customfield_24893", "2-ways", "list", "", "", "", "{value}.toLowerCase().replace(' ', '.')"]);
  configSheet.appendRow(["Dev estimate", "customfield_25757", "2-ways", "text"]);
  configSheet.appendRow(["QA estimate", "customfield_25958", "2-ways", "text"]);
  configSheet.appendRow(["Target start", "customfield_18350", "2-ways", "date"]);
  configSheet.appendRow(["Target end", "customfield_18351", "2-ways", "date"]);
  configSheet.appendRow(["Exist on Production", "customfield_10570", "2-ways", "text"]);
  configSheet.appendRow(["Affect customers", "customfield_13250", "2-ways", "text"]);
  configSheet.appendRow(["DEA", "customfield_26055", "2-ways", "list", "No"]);
  configSheet.appendRow(["UX Ticket", "depends on", "2-ways", "link"]);
  configSheet.appendRow(["Status", "status", "Back", "text"]);
}

function onEdit_markSyncHeaders(e) {
  const range = e.range;
  const column = range.getColumn();
  const row = range.getRow();
  if (row == 1) {Logger.log('Header change, quit!'); return}
  if (!e.value) return;
  const configSheet = SpreadsheetApp.getActiveSheet()
  const dataSheetName = configSheet.getName().replace(/_config$/, "")
  const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheetName)
  if (!dataSheet) {Logger.log('No found data sheet, quit!'); return}
  if (!/.*_config$/.test(configSheet.getName())) return
  if (column != 1 && column != 1 + 1
    && column != 2 + countConfigColumns && column != 2 + countConfigColumns + 1) {Logger.log('Mark synced headers is only trigger by field name changes!'); return}

  // Throttle
  setQueue('markSyncHeaders', new Date().getTime() + queueIntevalSeconds * 1000, dataSheet.getName())
}
function markSyncHeaders(dataSheetName = null) {
  const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheetName)
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheetName + '_config')
  if (!dataSheet) return
  if (!configSheet) return

  let columnValues = configSheet.getRange(2, 1, 100, 3).getValues()
  let columnNames_2ways = columnValues.filter(v => v[2]=='2-ways').map(v => v[0])
  let columnNames_back = columnValues.filter(v => v[2]=='Back').map(v => v[0])
  let columnNames_to = columnValues.filter(v => v[2]!='Back'&&v[2]!='2-ways').map(v => v[0])
  for (var col = 1; col <= 100; col++) {
    let headerCell = dataSheet.getRange(1, col)
    let headerValue = headerCell.getValue()
    if (!headerValue) continue
    if (columnNames_2ways.includes(headerValue)) headerCell.setNote('This column is synced with JIRA')
    else if (columnNames_back.includes(headerValue)) headerCell.setNote('This column will get changes from JIRA')
    else if (columnNames_to.includes(headerValue)) headerCell.setNote('This column is updating to JIRA')
    else headerCell.clearNote()
  }
}

function insertJIRAColumn() {
  const activeSheet = SpreadsheetApp.getActiveSheet()
  let JIRAcolumn = activeSheet.insertColumnBefore(1)
  JIRAcolumn.getRange(1, 1).setValue('JIRA key')
}


// 主页面
function onHomepage(e) {
  Logger.log('installInfo:')
  Logger.log(installInfo)
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const ui = SpreadsheetApp.getUi()
  menu = menu || (env == 'production' ? ui.createAddonMenu() : ui.createMenu('JIRA sync test'))
  let myInstallTriggers = getMyInstallTriggers()
  isInstalled = isInstalled || myInstallTriggers.length > 0
  let isInstalledHourTrigger = !!myInstallTriggers.find(t => t.getHandlerFunction() == '_runEveryHour')
  // test环境不会执行 onOpen
  if (env == 'test') if (!isInstalled) menu.addItem('Sync this sheet', 'createSpreadsheetEditTrigger').addToUi()
  else menu.addItem('Stop sync this sheet', 'removeSpreadsheetEditTrigger').addToUi()

  const card = CardService.newCardBuilder()
  const section = CardService.newCardSection()
  section.addWidget(CardService.newTextParagraph().setText("Welcome to JIRA sync!"))
  /* Deprecated: 用 homepage 安装会导致只有安装者有权限，其他用户无法触发 onEdit
  if (!isInstalled) section.addWidget(CardService.newDivider()).addWidget(CardService.newDecoratedText().setText('Let\'s make it sync!').setWrapText(true)
    .setButton(CardService.newTextButton().setText('Sync').setOnClickAction(CardService.newAction().setFunctionName("homepage_createSpreadsheetEditTrigger")))) */
  if (!isInstalled) {
    section.addWidget(CardService.newDivider()).addWidget(CardService.newDecoratedText()
      .setText('Start sync from menu: Extensions -> JIRA sync -> Sync this sheet!\n\nThen refresh here.').setWrapText(true))
  } else {
    let editorEmails = activeSpreadsheet.getEditors().map(editor => editor.getEmail())
    const canEdit = editorEmails.includes(editorEmail)
    const didCreatedConfig = PropertiesService.getDocumentProperties().getProperty('did-create-config') || false
    section.addWidget(CardService.newDivider())
    section.addWidget(CardService.newTextParagraph().setText("Step 1."))
    const buttonCreateConfig = CardService.newTextButton().setText('Create').setOnClickAction(CardService.newAction().setFunctionName("createConfigSheet"))
    if (!didCreatedConfig) buttonCreateConfig.setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    section.addWidget(CardService.newDecoratedText().setText('Create config sheet').setWrapText(true)
      .setButton(buttonCreateConfig))
    const buttonRefreshIndex = CardService.newTextButton().setText('Refresh').setOnClickAction(CardService.newAction().setFunctionName("homepage_indexSyncback"))
    if (canEdit) buttonRefreshIndex.setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    if (didCreatedConfig) section.addWidget(CardService.newDecoratedText().setText('If you change the config sheet, please refresh the catch for data sync back!').setWrapText(true)
      .setButton(buttonRefreshIndex))
    section.addWidget(CardService.newDivider())
    section.addWidget(CardService.newTextParagraph().setText("Step 2."))
    const buttonGrant = CardService.newTextButton().setText('Grant').setOnClickAction(CardService.newAction().setFunctionName("homepage_grantAccessToEditAccount"))
    section.addWidget(CardService.newDecoratedText().setText('Grant access for JIRA changes').setWrapText(true)
      .setButton(!canEdit ? buttonGrant.setTextButtonStyle(CardService.TextButtonStyle.FILLED) : buttonGrant.setText('Granted').setDisabled(true)))
    section.addWidget(CardService.newDivider())
    section.addWidget(CardService.newTextParagraph().setText("Step 3."))
    section.addWidget(CardService.newTextParagraph().setText("You've done the settings! Edit sheet data to sync."))
    section.addWidget(CardService.newDivider())
    section.addWidget(CardService.newTextParagraph().setText("Other tools"))
    section.addWidget(CardService.newDecoratedText().setText('This sheet already synced. Start with adding a column named "JIRA" with issue id!').setWrapText(true)
      .setButton(CardService.newTextButton().setText('Insert').setOnClickAction(CardService.newAction().setFunctionName("insertJIRAColumn"))))
    section.addWidget(CardService.newDecoratedText().setText('Fetch the latest data from JIRA!').setWrapText(true)
      .setButton(CardService.newTextButton().setText('Fetch').setOnClickAction(CardService.newAction().setFunctionName("getIssues"))))
    section.addWidget(CardService.newDecoratedText().setText('Expand the sub issues to the Epic!').setWrapText(true)
      .setButton(CardService.newTextButton().setText('Expand').setOnClickAction(CardService.newAction().setFunctionName("expandSubIssues"))))
    section.addWidget(CardService.newDivider()).addWidget(CardService.newDecoratedText().setText('You can pause sync for a while!').setWrapText(true)
      .setButton(CardService.newTextButton().setText('Pause').setOnClickAction(CardService.newAction().setFunctionName("comingSoon"))))
    if (!isInstalledHourTrigger) section.addWidget(CardService.newDivider()).addWidget(CardService.newDecoratedText().setText('Re-do sync made you apply the new feature: Bidirectional-sync!').setWrapText(true)
      .setButton(CardService.newTextButton().setText('Stop Sync').setOnClickAction(CardService.newAction().setFunctionName("homepage_removeSpreadsheetEditTrigger"))))
  }

  let fixedButton = CardService.newTextButton()
  if (isInstalled) {
    fixedButton.setText("Stop sync this sheet")
    if (!myInstallTriggers.length && installInfo) fixedButton.setDisabled(true).setAltText(`The sync for this sheet is managed by ${installInfo.creator}!`)
    fixedButton.setOnClickAction(CardService.newAction().setFunctionName("homepage_removeSpreadsheetEditTrigger"))
  } else fixedButton.setText("Sync this sheet")
            .setOnClickAction(CardService.newAction().setFunctionName("homepage_createSpreadsheetEditTrigger"))
  let fixedFooter = CardService.newFixedFooter()
  if (isInstalled && installInfo && !myInstallTriggers.length) fixedFooter.setSecondaryButton(CardService.newTextButton().setText("i").setOnClickAction(CardService.newAction().setFunctionName("alertInstalledByOther")))
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
  const ui = SpreadsheetApp.getUi()
  ui.alert(`This sheet has been enabled the sync by ${installInfo.creator} at ${installInfo.date}.
    To stop the sync, please reach out to ${installInfo.creator}!`)
}
function alertInstallTips() {
  const ui = SpreadsheetApp.getUi()
  ui.alert(`Please start sync from the menu: 
    Extensions -> JIRA sync -> Sync this sheet!`)
}
function comingSoon() {
  const ui = SpreadsheetApp.getUi()
  ui.alert(`Coming soon!`)
}

// Grant access
function homepage_grantAccessToEditAccount() {
  const ui = SpreadsheetApp.getUi()
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  activeSpreadsheet.addEditor(editorEmail)
  activeSpreadsheet.addEditor(editorBackupEmail)
  ui.alert(`Granted! Now the JIRA changes can be synced back to this sheet!`)
  return CardService.newNavigation().updateCard(onHomepage())
}


// Triggers
function onInstall(e) {
  const ui = SpreadsheetApp.getUi()
  menu = menu || (env == 'production' ? ui.createAddonMenu() : ui.createMenu('JIRA sync test'))
  menu.addItem('Click to start!', 'openHomepage').addToUi()
  createSpreadsheetEditTrigger();
}
function onOpen(e) {
  Logger.log(userEmail)
  Logger.log('authMode:')
  Logger.log(e.authMode)
  const ui = SpreadsheetApp.getUi()
  menu = menu || (env == 'production' ? ui.createAddonMenu() : ui.createMenu('JIRA sync test'))
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
  // SpreadsheetApp.getUi().createAddonMenu().addItem('test menu', 'openHomepage').addToUi() // To be removed
}
function onFileScopeGrantedSheets(e) {
}
function openTheLogSheet() {
  const ui = SpreadsheetApp.getUi()
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
function getMyInstallTriggers() {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  Logger.log('User triggers:')
  Logger.log(ScriptApp.getUserTriggers(activeSpreadsheet).map(v => ({sourece: v.getTriggerSource(), sid: v.getTriggerSourceId(), func:v.getHandlerFunction(), type: v.getEventType(), id: v.getUniqueId()})))
  // Logger.log('Project triggers:')  // 当前用户用此脚本包括在其他表的trigger，以及此脚本(非standalone)绑定那张表手动创建的trigger
  // Logger.log(ScriptApp.getProjectTriggers().map(v => ({sourece: v.getTriggerSource(), sid: v.getTriggerSourceId(), func:v.getHandlerFunction(), type: v.getEventType(), id: v.getUniqueId()})))
  return ScriptApp.getUserTriggers(activeSpreadsheet)
}
function homepage_createSpreadsheetEditTrigger() {
  createSpreadsheetEditTrigger()
  return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().updateCard(onHomepage()))
        .build();
}
function createSpreadsheetEditTrigger() {
  const ui = SpreadsheetApp.getUi()
  if (installInfo && installInfo.creatorEmail && installInfo.creatorEmail != userEmail) {
    ui.alert(`This sheet has been already enabled the sync by ${installInfo.creator} at ${installInfo.date}.`)
    return
  }

  const Properties = PropertiesService.getDocumentProperties()
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  let myInstallTrigger = ScriptApp.newTrigger('_onEdit')
      .forSpreadsheet(activeSpreadsheet)
      .onEdit()
      .create()
  let myChangeTrigger = ScriptApp.newTrigger('_onChange')
      .forSpreadsheet(activeSpreadsheet)
      .onChange()
      .create()
  let myFetchTrigger = ScriptApp.newTrigger('_runEveryHour')
      .timeBased()
      .everyHours(1)
      // .everyMinutes(1) // Add-on support 1 hour at least
      .create();
  installInfo = {
    creator: userName,
    creatorEmail: userEmail,
    triggerId: myInstallTrigger.getUniqueId(),
    triggerIdChange: myChangeTrigger.getUniqueId(),
    triggerIdFetch: myFetchTrigger.getUniqueId(),
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
    menu = menu || (env == 'production' ? ui.createAddonMenu() : ui.createMenu('JIRA sync test'))
    menu.addItem('Stop sync this sheet', 'removeSpreadsheetEditTrigger')
    .addToUi();
  }
}
function homepage_removeSpreadsheetEditTrigger() {
  removeSpreadsheetEditTrigger()
  return CardService.newNavigation().updateCard(onHomepage())
}
function removeSpreadsheetEditTrigger() {
  const ui = SpreadsheetApp.getUi()
  menu = menu || (env == 'production' ? ui.createAddonMenu() : ui.createMenu('JIRA sync test'))
  let myInstallTriggers = getMyInstallTriggers()
  if (myInstallTriggers.length > 0) {
    myInstallTriggers.forEach(trigger => ScriptApp.deleteTrigger(trigger))
  } else if (installInfo && installInfo.creatorEmail && installInfo.creatorEmail != userEmail) {
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
}

// function onEdit(e) {
  // onEdit_recordChanges(e)
// }

function _onEdit(e) {
  Logger.log({user: userEmail, sheet: SpreadsheetApp.getActiveSpreadsheet().getName(), tab: SpreadsheetApp.getActiveSheet().getName(), col: e.range.getColumn(), row: e.range.getRow()})
  recordChanges(e)
  onEdit_markSyncHeaders(e)
  onEdit_indexSyncback(e)
  // runQueue()  // Run queue to sync syncback index data
}
function onEdit_recordChanges(e) {  // Deprecated: but keep it for the compatibility
  _onEdit(e)
}

function _onChange(e) {
  Logger.log({user: userEmail, sheet: SpreadsheetApp.getActiveSpreadsheet().getName(), tab: SpreadsheetApp.getActiveSheet().getName(), type: e.changeType})
  onChange_indexSyncback(e)
  runQueue()  // Run queue to sync syncback index data
}

function _runEveryHour() {
  fetchJIRADataFromLogSheet()
  runQueue()  // Run queue to sync syncback index data
}

// Queue class
const queueIntevalSeconds = 15
function runQueue() {
  const Properties = PropertiesService.getDocumentProperties()
  let isBusy = Properties.getProperty('is-queue-instant-busy')
  if (isBusy && new Date(isBusy).getTime() > new Date().getTime() - 2 * queueIntevalSeconds * 1000) {Logger.log('Other queue instant is running now!'); return}
  Properties.setProperty('is-queue-instant-busy', new Date())
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
    if (!job.status == 'running') return
    queue[i].status = 'running'
    Properties.setProperty('queue-to-run', JSON.stringify(queue))
    try {
      eval(job.functionName + '(...paramsToQueueFunc.'+job.functionName+')')
      queue.splice(i, 1)
      Properties.setProperty('queue-to-run', JSON.stringify(queue))
    } catch {
      queue[i].status = 'failed'
      Properties.setProperty('queue-to-run', JSON.stringify(queue))
      Logger.log('Run queue function failed: ' + job.functionName)
    }
  })
  Properties.deleteProperty('is-queue-instant-busy')
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
  const ui = SpreadsheetApp.getUi()
  const dataSheet = SpreadsheetApp.getActiveSheet()
  if (/.*_config$/.test(dataSheet.getName())) {ui.alert('You are not at config sheet, please swtich to data sheet!'); return}
  initHeaders()
  if (!primaryJiraKeyCol) {ui.alert('Here is not data sheet, or config one mapping to "JIRA"!'); return}

  let issueKeys = []
  let range = dataSheet.getActiveRange()
  for (var row = range.getRow(); row < range.getRow() + range.getNumRows(); row++) {
    let primaryJiraValue = dataSheet.getRange(row, primaryJiraKeyCol).getValue()
    if (!primaryJiraValue) continue

    issueKeys.push({key: primaryJiraValue, keyHeader: primaryJiraFieldMap[primaryJiraKeyCol].name, row})

    /* Deprecated: RC server is not available outside
    try {
      var response = UrlFetchApp.fetch(jiraGetDataWebhook, {
        method: 'post',
        contentType: 'application/json',
        muteHttpExceptions: 'true', // 异步请求
        payload: JSON.stringify({
            "data": {
                "emailAddress": userEmail
            },
            "issues": [ primaryJiraValue ]
        })
      });
      var responseData = JSON.parse(response.getContentText());
      
      Logger.log("响应数据: " + JSON.stringify(responseData));
      
      CardService.newActionResponseBuilder()
        .setNotification(
          CardService.newNotification()
            .setText("POST 请求成功，响应 ID: " + responseData.id)
            .setType(CardService.NotificationType.INFO)
        )
        .build();
    } catch (error) {
      Logger.log("请求失败: " + error);
      
      CardService.newActionResponseBuilder()
        .setNotification(
          CardService.newNotification()
            .setText("请求失败: " + error.message)
            .setType(CardService.NotificationType.ERROR)
        )
        .build();
    } */

    /* Deprecated: Need python script to fetch every column data
    for (var column in primaryJiraFieldMap) {
      if (column == primaryJiraKeyCol) continue
      if (!primaryJiraFieldMap[column]) continue
      Logger.log({row, column, primaryJiraValue})
      _syncDataToLogSheet({
        row,
        column,
        id: primaryJiraValue,
        idName: primaryJiraFieldMap[primaryJiraKeyCol].name,
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
    } */
  }

  _syncGettingListToLogSheet(issueKeys)
  ui.alert("Please wait a min for data fetching!")
}
function expandSubIssues() {
  const ui = SpreadsheetApp.getUi()
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
      idName: primaryJiraFieldMap[primaryJiraKeyCol].name,
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
function _syncGettingListToLogSheet(issues) {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const activeSheet = SpreadsheetApp.getActiveSheet()
  const activeSheetId = activeSheet.getSheetId()
  const activeSheetName = activeSheet.getName()
  const activeSheetUrl = activeSpreadsheet.getUrl() + '#gid=' + activeSheetId
  const activeSpreadsheetName = activeSpreadsheet.getName()
  const logSS = SpreadsheetApp.openByUrl(logSheetURL)
  let logSheet = logSS.getSheetByName(dataGettingSheetName)
  if (!logSheet) {
    logSheet = logSS.insertSheet(dataGettingSheetName)
    logSheet.appendRow(["editor", "JIRA key", "sheet key header", "sheet name", "sheet URL", "sheet tab", "sheet tab gid", "sheet row", "time", "isSync", "sync time", "took seconds", "fail reason"]);
    // 有调整，同步修改 getIssuesPendingData (in changelog/Code.js)
  }

  issues.forEach(issue => {
    logSheet.appendRow([userEmail, `=HYPERLINK("https://jira.ringcentral.com/browse/${issue.key}", "${issue.key}")`, issue.keyHeader, activeSpreadsheetName, activeSheetUrl, activeSheetName, activeSheetId, issue.row, new Date().toLocaleString()]);
  })
  Logger.log('Sync log successfully!\n' + logSheetURL)
}

function fetchJIRADataFromLogSheet() {
  const dataSS = env == 'production' ? SpreadsheetApp.getActiveSpreadsheet() : SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1GNeBIM6Z6cnUz1qnlB9rztQJjv6BebCTZQ-6oEGNmbo/edit?gid=0#gid=0")
  const logSS = SpreadsheetApp.openByUrl(logSheetURL)
  const dataSSId = dataSS.getId()

  const dataSheets = dataSS.getSheets()
  dataSheets.forEach(function(sheet) {
    if (!/_config$/.test(sheet.getSheetName())) return
    const dataSheet = dataSS.getSheetByName(sheet.getSheetName().replace('_config', ''))
    if (!dataSheet) return
    // logSheetName = dataSS.getName() + ': ' + dataSheet.getName()
    let logSheet = logSS.getSheetByName(logSheetName)
    if (!logSheet) {Logger.log(dataSheet.getName() + ': No log sheet found!'); return}
  
    let logs = logSheet.getDataRange().getValues()
    let colSheetUrl = getHeaderCol('sheet URL', logSheet)
    let colSheetTabGid = getHeaderCol('sheet tab gid', logSheet)
    let colTime = getHeaderCol('time', logSheet)
    let colStatus = getHeaderCol('isSync', logSheet)
    let colSyncTime = getHeaderCol('sync time', logSheet)
    let colTookSeconds = getHeaderCol('took seconds', logSheet)
    let colFrom = getHeaderCol('from', logSheet)
    let colAction = getHeaderCol('action', logSheet)
    let colNewValue = getHeaderCol('new value', logSheet)
    let colOldValue = getHeaderCol('old value', logSheet)
    let colFieldDesc = getHeaderCol('JIRA field desc', logSheet)
    let colSheetRow = getHeaderCol('sheet row', logSheet)
    let colKeyHeader = getHeaderCol('sheet key header', logSheet)
    let colKey = getHeaderCol('JIRA key', logSheet)
    let logsFetched = logs.filter(function(log, i) {
      if (!RegExp(dataSSId).test(log[colSheetUrl-1])) return
      if (log[colSheetTabGid-1] != dataSheet.getSheetId()) return
      if (log[colStatus-1] == 'Done') return
      if (log[colStatus-1] == 'Failed') return
      if (log[colFrom-1] == 'sheet' && log[colAction-1] == 'get') {
        if (!log[colNewValue-1]) return
        return _copyDataFromChangelog()
      } else if (log[colFrom-1] == 'sheet' && log[colAction-1] == 'getSubissuesInsert') {
        // Todo
        if (!log[colNewValue-1]) return
      } else if (log[colFrom-1] == 'jira') {
        if (log[colNewValue-1] == log[colOldValue-1]) return false
        return _copyDataFromChangelog()
      }

      function _copyDataFromChangelog() {
        let colDataSheetJIRA = getHeaderCol(log[colKeyHeader-1], dataSheet)
        if (!colDataSheetJIRA) return false
        let colDataSheetField = getHeaderCol(log[colFieldDesc-1], dataSheet)
        if (!colDataSheetField) return false
        let dataRow = log[colSheetRow-1]
        if (dataSheet.getRange(dataRow, colDataSheetJIRA).getValue() != log[colKey-1]) {
          // dataSheet row/column 发生错位。检索 jira key 对应行进行修改
          dataRow = getRowByValue(log[colKey-1], log[colKeyHeader-1], dataSheet)
          if (!dataRow) return false
        }
        dataSheet.getRange(dataRow, colDataSheetField).setValue(log[colNewValue-1])
        logSheet.getRange(i+1, colStatus).setValue('Done')
        logSheet.getRange(i+1, colSyncTime).setValue(new Date())
        logSheet.getRange(i+1, colTookSeconds).setValue(Math.ceil((new Date().getTime() - new Date(log[colTime-1]).getTime()) / 1000))
        Logger.log({logSheetName, colDataSheetField, colSheetRow, newValue: log[colNewValue-1]})
        return true
      }
    })
    Logger.log('Fetch ' + dataSheet.getName() + ' data from log done!\nData sheet:' + dataSS.getUrl() + '\nData sheet owner:' + dataSS.getOwner().getEmail() + '\nLog sheet:' + logSheetURL)
    Logger.log(logsFetched)
  })
}


/* Log to changelog sheet */

// 记录每次jira相关的修改推送到log表
function recordChanges(e) {
  const activeSheet = SpreadsheetApp.getActiveSheet()
  const range = e.range
  const column = range.getColumn()
  const row = range.getRow()
  const isMultiple = !!(range.getNumRows() > 1 || range.getNumColumns() > 1)
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
      if (e.value === '') return;
      if (e.value === undefined) return;
      // if (!e.oldValue) return;
      if (column == primaryJiraKeyCol) {Logger.log('Skip as it is JIRA key column!'); return}
      if (!primaryJiraFieldMap[column]) {Logger.log('Skip as change is no mapping to config JIRA fields!'); return}
      if (primaryJiraFieldMap[column].syncMode == 'Back') {Logger.log('Skip per to sync mode config!'); return}
      const jiraKey = range.getSheet().getRange(row, primaryJiraKeyCol).getValue();
      if (!jiraKey) {Logger.log('No specific JIRA key!'); return}
      
      const jiraKeyName = range.getSheet().getRange(1, primaryJiraKeyCol).getValue();
      let data = getData(jiraKey, jiraKeyName, e.oldValue, e.value, primaryJiraFieldMap)
      Logger.log(data)
      _syncDataToLogSheet(data)
    } else {
      let values = range.getValues()
      for (var r = row; r < row + range.getNumRows(); r++) {
        for (var c = column; c < column + range.getNumColumns(); c++) {
          let value = values[r-row][c-column]
          // Logger.log({value_range: value, r, c})
          if (c == primaryJiraKeyCol) {Logger.log('Row:'+r+' Column:'+c + '. Skip as it is JIRA key column!'); continue}
          if (!primaryJiraFieldMap[c]) {Logger.log('Row:'+r+' Column:'+c + '. Skip as change is no mapping to config JIRA fields!'); continue}
          if (primaryJiraFieldMap[c].syncMode == 'Back') {Logger.log('Row:'+r+' Column:'+c + '. Skip per to sync mode config!'); continue}
          if (!value) {Logger.log('Row:'+r+' Column:'+c + '. Skip by no value, probably it is dragging row!'); continue}
          const jiraKey = range.getSheet().getRange(r, primaryJiraKeyCol).getValue();
          if (!jiraKey) {Logger.log('Row:'+r+' Column:'+c + '. No specific JIRA key!'); continue}
          
          const jiraKeyName = range.getSheet().getRange(1, primaryJiraKeyCol).getValue();
          let data = getData(jiraKey, jiraKeyName, r==row&&c==column?e.oldValue:'', value, primaryJiraFieldMap, r, c)
          Logger.log(data)
          _syncDataToLogSheet(data)
        }
      }
    }
  }()

  // 推送第二字段列表的变化给JIRA
  !function(){
    if (!isMultiple) {
      if (e.value === '') return;
      if (e.value === undefined) return;
      // if (!e.oldValue) return;
      if (secondaryJiraKeyCol === null) return;
      if (column == secondaryJiraKeyCol) return;
      if (!secondaryJiraFieldMap[column]) return;
      if (secondaryJiraFieldMap[column].syncMode == 'Back') return;
      const jiraKey = range.getSheet().getRange(row, secondaryJiraKeyCol).getValue();
      if (!jiraKey) return;
      
      const jiraKeyName = range.getSheet().getRange(1, secondaryJiraKeyCol).getValue();
      let data = getData(jiraKey, jiraKeyName, e.oldValue, e.value, secondaryJiraFieldMap)
      Logger.log(data)
      _syncDataToLogSheet(data)
    } else {
      let values = range.getValues()
      for (var r = row; r < row + range.getNumRows(); r++) {
        for (var c = column; c < column + range.getNumColumns(); c++) {
          let value = values[r-row][c-column]
          if (c == secondaryJiraKeyCol) continue
          if (!secondaryJiraFieldMap[c]) continue
          if (secondaryJiraFieldMap[c].syncMode == 'Back') continue
          if (!value) continue
          const jiraKey = range.getSheet().getRange(r, secondaryJiraKeyCol).getValue();
          if (!jiraKey) continue
          
          const jiraKeyName = range.getSheet().getRange(1, secondaryJiraKeyCol).getValue();
          let data = getData(jiraKey, jiraKeyName, r==row&&c==column?e.oldValue:'', value, secondaryJiraFieldMap, r, c)
          Logger.log(data)
          _syncDataToLogSheet(data)
        }
      }
    }
  }()

  function getData(id, idName, oldValue, newValue, JiraFieldMap = primaryJiraFieldMap, row = range.getRow(), column = range.getColumn()) {
    // Format with custom function
    newValue = JiraFieldMap[column].formatFuc ? JiraFieldMap[column].formatFuc(newValue) : newValue
    // Format date
    newValue = JiraFieldMap[column].type == 'date' ? formatDate(newValue) : newValue

    return {
      row,
      column,
      id,
      idName,
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
  // logSheetName = activeSpreadsheet.getName() + ': ' + activeSheet.getName()
  let logSheet = logSS.getSheetByName(logSheetName)
  if (!logSheet) {
    logSheet = logSS.insertSheet(logSheetName)
    logSheet.appendRow(["editor", "from", "action", "old value", "new value", "sheet name", "sheet URL", "sheet tab", "sheet tab gid", "sheet row", "sheet column", "sheet key header", "JIRA key", "JIRA field desc", "JIRA field name", "JIRA field type", "time", "isSync", "sync time", "took seconds", "fail reason"]);
  }

  logSheet.appendRow([userEmail, from, action||(data.isChangeAsAdding?'add':'replace'), data.oldValue, data.newValue?data.prefix+data.newValue+data.suffix:'', activeSpreadsheet.getName(), activeSpreadsheet.getUrl()+'#gid='+activeSheet.getSheetId(), activeSheet.getName(), activeSheet.getSheetId(), data.row, data.column, data.idName, `=HYPERLINK("https://jira.ringcentral.com/browse/${data.id}", "${data.id}")`, data.fieldDesc, data.field, data.type, new Date().toLocaleString()]);
  Logger.log('Sync log successfully!\n' + logSheetURL)
}


/* Reindex sync back index sheet */

// 修改配置表的时候，维护一张 sync back 索引表供 JIRA webhook 调用的时候索引 Tickets
function onEdit_indexSyncback(e) {
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
      && column != 2 + countConfigColumns && column != 2 + countConfigColumns + 1) {Logger.log('Auto update syncback index is only trigger by field name changes!'); return}

    // Throttle
    setQueue('indexSyncback', new Date().getTime() + queueIntevalSeconds * 1000, dataSheet.getName())
    // indexSyncback(dataSheet) // Instead of run directly, add to queue above
  }()

  // Data sheet new JIRA key added
  !function(){
    if (!e.value) return;
    if (e.oldValue) return; // only add JIRA key
    // Todo: 支持 copy/paste
    const dataSheet = SpreadsheetApp.getActiveSheet()
    if (/.*_config$/.test(dataSheet.getName())) return
    const configSheetName = dataSheet.getName() + '_config'
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(configSheetName)
    if (!configSheet) {Logger.log('No found config sheet, quit!'); return}
    initHeaders(dataSheet, true)
    if (column != primaryJiraKeyCol && column != secondaryJiraKeyCol) {Logger.log('Auto update syncback index is only trigger by JIRA key added!'); return}

    // Throttle
    setQueue('indexSyncback', new Date().getTime() + queueIntevalSeconds * 1000, dataSheet.getName())
  }()

  /* Deprecated: use onChange_indexSyncback instead
  // Data sheet new row inserted (delete cannot be supported in onEdit trigger)
  !function(){
    if (e.value) return;
    if (e.oldValue) return;
    const isMultiple = !!(range.getNumRows() || range.getNumColumns())
    if (!isMultiple) return;
    const dataSheet = SpreadsheetApp.getActiveSheet()
    if (/.*_config$/.test(dataSheet.getName())) return
    const configSheetName = dataSheet.getName() + '_config'
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(configSheetName)
    if (!configSheet) {Logger.log('No found config sheet, quit!'); return}
    let values = range.getValues()
    if (values.some(rowValues => rowValues.some(cell => cell !== ''))) return // 但其实插入行列和清空操作是一样的，cell全为空

    // Throttle
    setQueue('indexSyncback', new Date().getTime() + queueIntevalSeconds * 1000, dataSheet.getName())
  }() */
}

// 表头变化需要用 onChange
function onChange_indexSyncback(e) {
  const dataSheet = SpreadsheetApp.getActiveSheet()
  if (/.*_config$/.test(dataSheet.getName())) return
  const configSheetName = dataSheet.getName() + '_config'
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(configSheetName)
  if (!configSheet) {Logger.log('No found config sheet, quit!'); return}
  if (!/^INSERT_ROW|INSERT_COLUMN|REMOVE_ROW|REMOVE_COLUMN$/.test(e.changeType)) return
  Logger.log('Data sheet structure changed: ' + e.changeType)

  // Throttle
  setQueue('indexSyncback', new Date().getTime() + queueIntevalSeconds * 1000, dataSheet.getName())
}

function homepage_indexSyncback() {
  const ui = SpreadsheetApp.getUi()
  if (indexSyncback()) ui.alert('Refresh catch successfully!')
    return CardService.newActionResponseBuilder()
}
function indexSyncback(dataSheet = null) {
  const ui = SpreadsheetApp.getUi()
  if (!dataSheet) {
    const activeSheet = SpreadsheetApp.getActiveSheet()
    if (!/.*_config$/.test(activeSheet.getName())) {ui.alert('You are currently not in the config sheet!'); return}
    const dataSheetName = activeSheet.getName().replace(/_config$/, "")
    dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheetName)
  } else if (typeof dataSheet == 'string') dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet)

  initHeaders(dataSheet, true)
  // Logger.log(primaryJiraFieldMap)
  let tickets = []
  // Primary JIRA columns index
  let dataValues = dataSheet.getRange(2, 1, 1000, primaryJiraFieldMap.length).getValues()
  for (var r in dataValues) {
    if (!dataValues[r][primaryJiraKeyCol-1]) continue
    // Logger.log(dataValues[r])
    for (var c in dataValues[r]) {
      const primaryJiraField = primaryJiraFieldMap[parseInt(c)+1]
      if (!primaryJiraField) continue
      if (parseInt(c)+1 == primaryJiraKeyCol) continue
      if (primaryJiraField.syncMode != '2-ways' && primaryJiraField.syncMode != 'Back') continue
      tickets.push({
        jiraKey: dataValues[r][primaryJiraKeyCol-1],
        keyHeader: primaryJiraFieldMap[primaryJiraKeyCol].desc,
        row: parseInt(r) + 1 + 1,
        column: parseInt(c) + 1,
        fieldName: primaryJiraField.name,
        fieldDesc: primaryJiraField.desc,
        fieldType: primaryJiraField.type,
        isListAllValue: primaryJiraField.isChangeAsAdding ? 'editing' : 'all',  // Todo: 'max'
        removePrefix: primaryJiraField.prefix,
        removeSuffix: primaryJiraField.suffix,
        backFormat: primaryJiraField.backFormat,
      })
    }
  }
  // Secondary JIRA columns index
  for (var r in dataValues) {
    if (!dataValues[r][secondaryJiraKeyCol-1]) continue
    // Logger.log(dataValues[r])
    for (var c in dataValues[r]) {
      const secondaryJiraField = secondaryJiraFieldMap[parseInt(c)+1]
      if (!secondaryJiraField) continue
      if (parseInt(c)+1 == secondaryJiraKeyCol) continue
      if (secondaryJiraField.syncMode != '2-ways' && secondaryJiraField.syncMode != 'Back') continue
      tickets.push({
        jiraKey: dataValues[r][secondaryJiraKeyCol-1],
        keyHeader: secondaryJiraFieldMap[secondaryJiraKeyCol].desc,
        row: parseInt(r) + 1 + 1,
        column: parseInt(c) + 1,
        fieldName: secondaryJiraField.name,
        fieldDesc: secondaryJiraField.desc,
        fieldType: secondaryJiraField.type,
        isListAllValue: secondaryJiraField.isChangeAsAdding ? 'editing' : 'all',  // Todo: 'max'
        removePrefix: secondaryJiraField.prefix,
        removeSuffix: secondaryJiraField.suffix,
        backFormat: secondaryJiraField.backFormat,
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
  syncbackSheet.appendRow(["owner", "sheet key header", "JIRA key", "JIRA field desc", "JIRA field name", "JIRA field type", "sheet name", "sheet URL", "sheet tab", "sheet tab gid", "sheet row", "sheet column", "list all value", "remove prefix", "remove suffix", "back format func", "last edit back time", "fail reason"]);
  if (!tickets.length) { Logger.log('No data with config to sync!'); return }

  const spreadSheetName = activeSpreadsheet.getName()
  const sheetURL = activeSpreadsheet.getUrl()+'#gid='+dataSheet.getSheetId()
  const dataSheetName = dataSheet.getName()
  const dataSheetId = dataSheet.getSheetId()
  /* Deprecated: use setValues instead of appendRow
  tickets.forEach(ticket => {
    syncbackSheet.appendRow([userEmail, `=HYPERLINK("https://jira.ringcentral.com/browse/${ticket.jiraKey}", "${ticket.jiraKey}")`, ticket.fieldDesc, ticket.fieldName, ticket.fieldType, spreadSheetName, sheetURL, dataSheetName, dataSheetId, ticket.row, ticket.column, ticket.isListAllValue, ticket.removePrefix, ticket.removeSuffix, ticket.backFormat]);
  }) */
  let rangeValues = tickets.map(ticket => [userEmail, ticket.keyHeader, `=HYPERLINK("https://jira.ringcentral.com/browse/${ticket.jiraKey}", "${ticket.jiraKey}")`, ticket.fieldDesc, ticket.fieldName, ticket.fieldType, spreadSheetName, sheetURL, dataSheetName, dataSheetId, ticket.row, ticket.column, ticket.isListAllValue, ticket.removePrefix, ticket.removeSuffix, ticket.backFormat])
  syncbackSheet.getRange(2, 1, rangeValues.length, rangeValues[0].length).setValues(rangeValues)
  Logger.log('Sync syncback index successfully!\n' + syncbackSheetURL)
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

function getHeaderCol(columnName, sheet = SpreadsheetApp.getActiveSheet()) {
  let headerValues = sheet.getRange(1, 1, 1, 100).getValues()
  let headerColByName = headerValues[0].reduce((pre,cur,i) => {pre[cur] = i+1; return pre}, {})
  return headerColByName[columnName]
}

function getRowByValue(colValue, colName, sheet = SpreadsheetApp.getActiveSheet()) {
  if (!colName) return null
  let col = getHeaderCol(colName, sheet)
  if (!col) return null
  let colValues = sheet.getRange(1, col, 100, 1).getValues()
  let rowByColValue = colValues.reduce((pre,cur,i) => {pre[cur[0]] = i+1; return pre}, {})
  return rowByColValue[colValue]
}


/**
 * Web app API
 * 
 * Configure: 
 *  1. Deploy as web app, set access to "Anyone, even anonymous"
 *  2. Set the URL+'?action=jiraChange' to JIRA webhook
 * 
 *  */
function doGet(e) {
  switch (e.parameter.action) {
    // case 'getIssuesPendingData': // 移植到 Changelog script
    //   return ContentService.createTextOutput(JSON.stringify(getIssuesPendingData(e)))
    default:
      var name = e.parameter.name || "World";
      return ContentService.createTextOutput(JSON.stringify({ message: "Hello, " + name + "!" }))
                           .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  switch (e.parameter.action) {
    case 'sendJIRAIssues':
      var jsonData = JSON.parse(e.postData.issues);
      return ContentService.createTextOutput(populateJIRAIssues(jsonData))
    case 'jiraChange':
      // return jiraChange(e)
    default:
      return ContentService.createTextOutput('No action found!')
      var jsonData = JSON.parse(e.postData.contents);
      return ContentService.createTextOutput("Received: " + jsonData);
  }
}

function populateJIRAIssues(issues) {

}
