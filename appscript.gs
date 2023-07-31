function updateForm(sheetId, formId) {
  var [group, groupAndTopic, studentList] = getDataFromSheet(sheetId);  // group: danh sách nhóm khả dụng, groupAndTopic: topic được đăng ký theo nhóm, studentList: danh sách sinh viên theo nhóm
  // form
  // delte old form
  if (!group.length)
    return [false, '']
  if (!groupAndTopic) {
    return [false, '']
  }
  if (formId) {
    var oldForm = DriveApp.getFileById(formId);
    oldForm.setTrashed(true);
  }
  var formId = duplicateForm();
  var form = FormApp.openById(formId);
  const oldItems = form.getItems(); 
  const selItem = oldItems.find(item => item.getType() === FormApp.ItemType.TEXT);
  if (selItem) {
    form.deleteItem(selItem);
  }
  // First section
  groupSelectedItem = form.addListItem();
  groupSelectedItem.setTitle('Nhóm');
  groupSelectedItem.setRequired(true);

  // create other section
  var pageBreakCompoment = [];
  var navigationPage = []

  for(var i = 0;i < group.length;i++) {
    pagetmp = form.addPageBreakItem();
    pagetmp.setTitle('Nhóm ' + String(group[i]));
    pagetmp.setHelpText('Đề tài: ' + String(groupAndTopic[group[i]])) // set topic
    pageBreakCompoment[i] = pagetmp;
    pageBreakCompoment[i].setGoToPage(FormApp.PageNavigationType.SUBMIT)

    nameItem = form.addListItem();
    nameItem.setTitle('Tên người đánh giá');
    nameItem.setChoiceValues(studentList[group[i]]);

    for (var j = 0; j < studentList[group[i]].length;j++) {
      item = form.addScaleItem();
      item.setTitle('Đánh giá điểm của ' + studentList[group[i]][j] + ':')
      item.setBounds(1, 10);
    }
    navigationPage[i] = groupSelectedItem.createChoice('Nhóm ' + String(group[i]), pageBreakCompoment[i]);
  }
  groupSelectedItem.setChoices(navigationPage);
  
  return [true, formId];
}

function getDataFromSheet(id){
  var app = SpreadsheetApp.openById(id);
  var sheet1 = app.getSheetByName('Danh sách sinh viên');
  var sheet2 = app.getSheetByName('Đăng ký đề tài');
  var sheet3 = app.getSheetByName('Danh sách đề tài');

  // lấy dữ liệu về nhóm và tên các sinh viên lưu trong group và 
  var numRow1 = sheet1.getLastRow()
  if (!numRow1) {
    return [[], [], []]
  }
  var raw_group = sheet1.getSheetValues(2, 7, numRow1 - 1, 1)
  var raw_student = sheet1.getSheetValues(2, 6, numRow1 - 1, 1)
  var group = []
  var name_student = []
  var studentList = {}
  Logger.log(raw_group)
  Logger.log(raw_student)
  for(var i = 0; i < numRow1 - 1;i++)
  {
    group.push(raw_group[i][0])
    name_student.push(raw_student[i][0])
  }
  // group.shift()
  var groups = group
  group = [...new Set(group)]
  group.filter((element) => { return element != undefined})
  name_student.filter((element) => { return element != undefined})
  
  // name_student.shift()
  for (var i = 0; i < group.length; i++){
    list = []
    for (var j = 0;j < groups.length;j++)
      if (group[i] == groups[j])
        list.push(name_student[j])
    studentList[group[i]] = list
  }

  // danh sách đề tài nhóm đã đăng ký
  var numRow2 = sheet2.getLastRow();
  if (!numRow2) {
    return [group, [], studentList]
  }
  var topicRegister = sheet2.getSheetValues(2, 1, numRow2 - 1, 2);
  var numRow3 = sheet3.getLastRow();
  var topicName = sheet3.getSheetValues(2, 1, numRow3 - 1, 2);
  print(topicRegister)
  print(topicName)
  // a = [[1, 2], [3, 4]]
  topicRegister.filter((element) => { return element != undefined});
  topicName.filter((element) => { return element != undefined});
  print(topicRegister)
  print(topicName)
  var topicDict = {}
  topicName.map(val => topicDict[val[0]] = val[1]) // topicDict = {'1': 'topic name',...}
  print(topicDict)
  var groupAndTopic = {} // dictory key=nhóm value= tên đề tài
  topicRegister.map(val => groupAndTopic[val[0]] = topicDict[val[1]])
  print(groupAndTopic)
  return [group, groupAndTopic, studentList]
}

function print(s) {
  Logger.log(s)
}

function getResponseFromForm(formId, sheetId) {
  try {
    var form = FormApp.openById(formId);
  }
  catch (e) {
    return []
  }
  var formResponses = form.getResponses();
  try {
    var spreadsheet = SpreadsheetApp.openById(sheetId);
  }
  catch {
    return []
  }
  formResponses
  var sheet = spreadsheet.getSheetByName('Danh sách sinh viên');
  var row = sheet.getLastRow();
  var auth_email = sheet.getSheetValues(2, 9, row - 1, 1);
  var auth_name = sheet.getSheetValues(2, 6, row - 1, 1);
  var email_list = {};
  for (var i = 0;i < auth_email.length;i++) {
    if (!auth_email[i].length)
      email_list[auth_email[i]] = '';
    else 
      email_list[auth_name[i]] = auth_email[i][0];
  }
  var result_list = [];
  var tmp_list = [];
  
  for(var i = 0;i < formResponses.length; i++) {
    var email = formResponses[i].getRespondentEmail();
    var itemResponse = formResponses[i].getItemResponses();
    var tmpString = '';
    tmp_list.push(email);
    for(var j = 0; j < itemResponse.length;j++){
      var response = itemResponse[j];
      if (j > 1) {
        tmpString = response.getItem().getTitle();
        tmpString = tmpString.replace('Đánh giá điểm của ', '');
        tmpString = tmpString.replace(':', '');
        tmp_list.push(tmpString);
      }
      if (j == 0) {
        tmpString = response.getResponse();
        tmpString = tmpString.replace('Nhóm', '');
        tmp_list.push(tmpString);
      }
      else
        tmp_list.push(response.getResponse());
    }
    tmp_list.splice(0, 0, email_list[tmp_list[2]])
    if (tmp_list[0] != tmp_list[1]) {
      tmp_list.splice(0, 0, false);
    }
    else 
      tmp_list.splice(0, 0, true);
    result_list.push(tmp_list);
    tmp_list = [];
  }
  print(result_list);
  return result_list;
}

function getAverageMark(formId, sheetId) {
  var raw_result = getResponseFromForm(formId, sheetId);
  if (!raw_result.length)
    return []
  var accessed = []
  for(var i = 0; i < raw_result.length;i++) {
    if(raw_result[i][0])
      accessed.push(raw_result[i].slice(3, raw_result[i].length));
  }
  if (!accessed.length)
    return []
  var [group, groupAndTopic, studentList] = getDataFromSheet(sheetId);
  var sortedResponse = {};
  var tmpList = [];
  for (var i = 0; i < group.length;i++) {
    sortedResponse[group[i]]=[];
  }
  for (var i = 0;i < accessed.length;i++) {
    for(var j = 0;j < group.length;j++) {
      if (accessed[i][0] == group[j]) {
        sortedResponse[group[j]].push(accessed[i]);
        break;
      }
    }
  }
  print(sortedResponse);
  var averageMark = []
  var tmp = [];
  var mark = [];
  var num = 0;
  var tmpmark =0;
  for (var i = 0;i < group.length;i++) {
    mark = [];
    print(averageMark)
    tmp = sortedResponse[group[i]];
    if (!tmp.length)
      continue;
    num = studentList[group[i]].length;
    if (!num) 
      continue;
    for (var j = 0;j < num;j++) {
      tmpmark = 0;
      for (var k = 0;k < tmp.length;k++) {
        
        tmpmark += Number(tmp[k][3+2*j]);
      }
      mark[j] =tmpmark / num;

    }
    tmp = [];
    for (var j = 0;j < num;j++) {
      var sec0 = group[i];
      var sec1 = sortedResponse[group[i]][0][2*j + 2];
      var sec2 = mark[j];
      averageMark.push([sec0, sec1, sec2]);
      tmp = []
    }
    
  }
  print(averageMark)
  return averageMark;
}

function test() {
  return 'hello'
}
function test2(){
  var sheet = '1OKK0OLLbj-Mqk0ahh0zk4sgBl_FGYpWRX34f4U9loeE'
  var form = '1gjtEZ5ui-e0EvWHTJdqWGWe3ZqQJwZXvi4IRQ-rmaNo'
  // getResponseFromForm(form, sheet);
  getAverageMark(form, sheet);
}

// create Sheet 
function createSheet() {
  var app = SpreadsheetApp.create('Danh sách đăng ký');
  app.renameActiveSheet('Danh sách sinh viên');
  app.insertSheet('Danh sách đề tài');
  app.insertSheet('Đăng ký đề tài');
  app.getSheetByName('Đăng ký đề tài').hideSheet()
  var drive = DriveApp.getFileById(app.getId())
  drive.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT)
  return app.getId();
}

function updateSheet1(sheetId, studentList) {
  var app = SpreadsheetApp.openById(sheetId);
  print(studentList);
  var sheet1 = app.getSheetByName('Danh sách sinh viên');
  sheet1.clearContents();
  sheet1.getRange(1, 1, studentList.length, studentList[0].length).setValues(studentList);
  
}
function updateSheet2(sheetId, topicList) {
  var app = SpreadsheetApp.openById(sheetId);
  print(topicList);
  print(typeof(topicList));
  print(topicList.length);
  var sheet2 = app.getSheetByName('Danh sách đề tài');
  sheet2.clearContents();
  sheet2.getRange(1, 1, topicList.length, topicList[0].length).setValues(topicList);
}


function addRegisterSheet(sheetId) {
  var app = SpreadsheetApp.openById(sheetId);
  var sheet1 = app.getSheetByName('Danh sách sinh viên');
  var sheet3 = app.getSheetByName('Đăng ký đề tài');
  sheet3.showSheet();
  var n_row = sheet1.getLastRow();
  raw_data = sheet1.getSheetValues(2, 7, n_row - 1, 1);
  var data = [];
  for (var i = 0;i < raw_data.length;i++) {
    data.push(raw_data[i][0]);
  }
  var data = [...new Set(data)];
  
  u_data = [['Nhóm', 'Đề tài']]
  for (var i = 0; i < data.length;i++) {
    u_data.push([data[i], ''])
  }
  Logger.log(u_data)
  sheet3.clearContents();
  sheet3.getRange(1, 1, u_data.length, u_data[0].length).setValues(u_data);
}

function test3(){
  // var sheetId = '1-XTyZRpBTdaPtDE7NiU1BYOPKNXXnIAS45T5bAePr_A'
  // var formId ='1gOpUM8PpW-9-hyJTveYGAFh6WCsgZQvcro5d9IVgapM'
  // var data = updateForm(sheetId, formId)
  var id = '1aOHEWZq87rFYjbnxYR4qNLRPRGXxz0Jg_XdyelDPLaU'
  var form = FormApp.openById(id);
  form.get
  // Get the email field.
  var emailField = form.getEmailFieldByName("EMAIL");

  // Enable the Verified checkbox.
  emailField.setVerified(true);

}

function checkFormId(id) {
  try {
    FormApp.openById(id);
  }
  catch(e){
    print(e)
    return false
  }
  print('hello')
  return true
}

function checkSheetId(id) {
  try {
    SpreadsheetApp.openById(id);
  }
  catch (e) {
    print(e);
    return false
  }
  print('hello')
  return true
}

function duplicateForm() {
  var templateformId = '1D0XNFhXZm-QA9EPwyg5XCnDBPWYWU9LErsPe6SAGuLk';
  var file = DriveApp.getFileById(templateformId).makeCopy('Form đánh giá bài tập lớn');
  var fileId = file.getId();
  return fileId;
}
















