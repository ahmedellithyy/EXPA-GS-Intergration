function dataExtraction(graphql)
{
  
  var requestOptions = {
    'method': 'post',
    'payload': graphql,
    'contentType':'application/json',
    'headers':{
      'access_token': "ACCESS_TOKEN"
    }
  };
  var response = UrlFetchApp.fetch(`https://gis-api.aiesec.org/graphql?access_token=${access_token}`, requestOptions);
  var recievedDate = JSON.parse(response.getContentText());
  return recievedDate.data.allOpportunityApplication.data;
}

//Set the extracted data into 2D array
function dataManipulation(dataSet)
{
 
 var rows = [];
 var standards = [];
 for(var i = 0; i < dataSet.length; i++)
 {
      
  rows.push([
    dataSet[i].person.id+"_"+dataSet[i].opportunity.id, dataSet[i].person.full_name,dataSet[i].opportunity.programme.short_name_display, dataSet[i].status,""/*remote status*/, "",
    
    dataSet[i].person.email,  "" /*phone number*/, dataSet[i].person.home_mc.name, dataSet[i].person.home_lc.name, "",
    
    dataSet[i].opportunity.home_mc.name,dataSet[i].opportunity.host_lc.name,"https://aiesec.org/opportunity/"+dataSet[i].opportunity.id,dataSet[i].opportunity.programme.short_name_display,""/*duration_type*/,""/*Project Fees*/,""/*salary Fees*/,"",
    
    ""/*APD Date*/,
    ""/*slot start date*/,
    ""/*slot end date*/,
    ""/*RE date*/,
    ""/*Fi date*/,
    ""/*remote date*/
 ]);

 standards.push([
   dataSet[i].person.id+"_"+dataSet[i].opportunity.id, 
   dataSet[i].person.full_name, 
   dataSet[i].opportunity.programme.short_name_display, 
   dataSet[i].status,
   ""/*remote status*/,
   "",
   ""/*APD Date*/,
   ""/*RE date*/,
   ""/*Fi date*/,
   ""/*remote date*/,
   "",""/*standard 1*/,
   "",""/*standard 2*/,
   "",""/*standard 3*/,
   "",""/*standard 4*/,
   "",""/*standard 5*/,
   "",""/*standard 6*/,
   "",""/*standard 7*/,
   "",""/*standard 8*/,
   "",""/*standard 9*/,
   "",""/*standard 10*/,
   "",""/*standard 11*/,
   "",""/*standard 12*/,
   "",""/*standard 13*/,
   "",""/*standard 14*/,
   "",""/*standard 15*/,
   "",""/*standard 16*/
 ])

  if( dataSet[i].opportunity.remote_opportunity == true){
        rows[i][4] = "Yes";
        standards[i][4] = "Yes";
  }
  else{
    rows[i][4] = "No"
    standards[i][4] = "No";

  }

  if (dataSet[i].person.contact_detail != null){
      rows[i][7] = dataSet[i].person.contact_detail.phone;

  }

  if(dataSet[i].opportunity.programme.short_name_display == "GV")
  {
      rows[i][16] = dataSet[i].opportunity.project_fee.fee+" "+dataSet[i].opportunity.project_fee.currency;

  }
  else{
    if(dataSet[i].opportunity.specifics_info.salary == null){
        dataSet[i].opportunity.specifics_info.salary = 0;
        rows[i][17] = dataSet[i].opportunity.specifics_info.salary+" "+dataSet[i].opportunity.specifics_info.salary_currency.alphabetic_code;
    }
    else{
      rows[i][17] = dataSet[i].opportunity.specifics_info.salary+" "+dataSet[i].opportunity.specifics_info.salary_currency.alphabetic_code;

    }
  }
  if(dataSet[i].slot!=null){
    rows[i][20] = dataSet[i].slot.start_date;
    rows[i][21] = dataSet[i].slot.end_date;
  }

  if(dataSet[i].opportunity.opportunity_duration_type != null){
      rows[i][15]=dataSet[i].opportunity.opportunity_duration_type.duration_type;

  }

  // set approval, realization and experience end dates
    if(dataSet[i].date_approved != null){
      rows[i][19]=  dataSet[i].date_approved.toString().substring(0,10);
      standards[i][6]=  dataSet[i].date_approved.toString().substring(0,10);

    }
    else{
          rows[i][19]=  dataSet[i].updated_at.toString().substring(0,10);
          standards[i][6]=  dataSet[i].updated_at.toString().substring(0,10);

    }

    if(dataSet[i].status == "remote_realized" ){
          rows[i][24]=  dataSet[i].updated_at.toString().substring(0,10);
          standards[i][9]=  dataSet[i].updated_at.toString().substring(0,10);
         
    }
    if(dataSet[i].status == "realized" )
    { 
      if(dataSet[i].date_realized != null){
        rows[i][22]=  dataSet[i].date_realized.toString().substring(0,10);
        standards[i][7]=  dataSet[i].date_realized.toString().substring(0,10);
      }
      if(dataSet[i].experience_end_date != null){
              rows[i][23]=  dataSet[i].experience_end_date.toString().substring(0,10);
              standards[i][8]=  dataSet[i].experience_end_date.toString().substring(0,10);
        }
      
    }
    else if (dataSet[i].status == "finished" || dataSet[i].status == "completed"){
      if(dataSet[i].date_realized != null){
        rows[i][22]=  dataSet[i].date_realized.toString().substring(0,10);
        standards[i][7]=  dataSet[i].date_realized.toString().substring(0,10);

      }
        if(dataSet[i].experience_end_date != null){
              rows[i][23]=  dataSet[i].experience_end_date.toString().substring(0,10);
              standards[i][8]=  dataSet[i].experience_end_date.toString().substring(0,10);
        }
        
    }
    for(var k =0; k <16;k++)
    {
      if(dataSet[i].standards[k].standard_option != null ){
        r = k*2;
       standards[i][11+r] = dataSet[i].standards[k].standard_option.meta.option
     }
    }
 }

    return [rows, standards]; 
}
// Take the raw data recieved from the HTTP response and arrange it into the corresponding sheet
function dataUpdating(sheet,standardSheet, rows, standards)
{
  var lr = getLastRow(sheet);
  var oldRows = sheet.getRange(5,1,lr,25).getValues();
  var newRows = [];
  for(var i =0; i < rows.length;i++)
  {
      var duplicated = false;
      for(var j = 0; j <oldRows.length ;j++)
      {
            if(rows[i][0] == oldRows[j][0])
            {
                duplicated = true;
                oldRows[j][3] = rows[i][3];
                if(rows[i][3] == "realized" || rows[i][3] == "finished" || rows[i][3] == "completed"){
                  oldRows[j].splice(22,1,rows[i][22]);
                  oldRows[j].splice(23,1,rows[i][23]);
                }
                else if(rows[i][2] == "remote_realized"){
                  oldRows[j].splice(24,1,rows[i][24]);
                }
            }
            
      }
      if(!duplicated){
        newRows.push(rows[i]); 
      }
  }

  afterEdits = sheet.getRange(5,1,oldRows.length,25).setValues(oldRows);

  if(newRows.length > 0){
    dataRange = sheet.getRange(lr+5,1,newRows.length,25).setValues(newRows);
  }


  lr = getLastRow(standardSheet);
  oldRows = standardSheet.getRange(5,1,lr,42).getValues();
  newRows = [];
  for(var i =0; i < standards.length;i++)
  {
      var duplicated = false;
      for(var j = 0; j <oldRows.length ;j++)
      {
            if(standards[i][0] == oldRows[j][0])
            {
                duplicated = true;
                oldRows[j][3] = standards[i][3];
                for(var k =0; k < 16;k++){
                    r = k*2;
                    oldRows[j].splice(11+r,1,standards[i][10+r])
                }
                if(rows[i][2] == "realized" || rows[i][2] == "finished" || rows[i][2] == "completed"){
                  oldRows[j].splice(7,1,rows[i][19]);
                  oldRows[j].splice(8,1,rows[i][20]);
                }
                else if(rows[i][2] == "remote_realized"){
                  oldRows[j].splice(9,1,rows[i][21]);
                }
            }
            
      }
      if(!duplicated){
        newRows.push(standards[i]); 
      }
  }

  afterEdits = standardSheet.getRange(5,1,oldRows.length,42).setValues(oldRows);

  if(newRows.length > 0){
    dataRange = standardSheet.getRange(lr+5,1,newRows.length,42).setValues(newRows);
  }


}
 

function main()
{
  var sheetInterface = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interface"); // write sheet name 
  var startDate = Utilities.formatDate(sheetInterface.getRange(8,3).getValue(), "GMT+2", "dd/MM/yyyy");
  var lcCode = sheetInterface.getRange(8,2).getValue();
  var lc = ["opportunity_home_lc","person_home_lc"]

  for(var i = 0; i < 2;i++){

  if(i == 0){
    // Write LC_NAME as it's existed in const file
      if(lc_codes["LC_NAME"][1] == -1) continue
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ICX Database/Auditing"); // write sheet name 
      var standardSheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ICX Standards Tracker"); 
  }
  else{
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("OGX Database/Auditing"); // write sheet name 
      var standardSheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("OGX Standards Tracker"); 

  }

  var queryAPDs= `query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n      ${lc[i]}:${lcCode}\n      date_approved:{from:"${startDate}"}\n      \n\t\t}\n    \n    page:1\n    per_page:3000\n\t)\n\t{\n    paging\n    {\n      total_items\n    }\n\t\tdata\n    {\n      person\n      {\n        id\n        full_name\n        email\n        contact_detail{\n          phone\n        }\n        home_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n      }\n      opportunity\n      {\n        id\n        programme\n        { short_name_display }\n        host_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n        remote_opportunity\n        project_fee\n        earliest_start_date\n        latest_end_date\n        specifics_info{\n          salary\n          salary_currency{\n            alphabetic_code\n          }\n        }\n        opportunity_duration_type{\n          duration_type\n          salary\n        }\n        \n      }\n      slot{\n        start_date\n        end_date\n      }\n      status\n      updated_at\n      date_approved\n      date_realized\n      experience_end_date\n    standards{\n        \n standard_option{\n          \n          meta\n        }\n \n   } }\n  }\n}`
    
  var graphql_APDs = JSON.stringify({query: queryAPDs})
  var dataSet_APDs = dataExtraction(graphql_APDs);
  var APDs = dataManipulation(dataSet_APDs);
  dataUpdating(sheet,standardSheet,APDs[0],APDs[1]);

  //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 var queryREs = `query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n      ${lc[i]}:${lcCode}\n      date_realized:{from:"${startDate}"}\n      \n\t\t}\n    \n    page:1\n    per_page:3000\n\t)\n\t{\n    paging\n    {\n      total_items\n    }\n\t\tdata\n    {\n      person\n      {\n        id\n        full_name\n        email\n        contact_detail{\n          phone\n        }\n        home_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n      }\n      opportunity\n      {\n        id\n        programme\n        { short_name_display }\n        host_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n        remote_opportunity\n        project_fee\n        earliest_start_date\n        latest_end_date\n        specifics_info{\n          salary\n          salary_currency{\n            alphabetic_code\n          }\n        }\n        opportunity_duration_type{\n          duration_type\n          salary\n        }\n        \n      }\n      slot{\n        start_date\n        end_date\n      }\n      status\n      updated_at\n      date_approved\n      date_realized\n      experience_end_date\n    standards{\n        \n standard_option{\n          \n          meta\n        }\n \n   } }\n  }\n}`
    
  var graphql_REs = JSON.stringify({query: queryREs})
  var dataSet_REs = dataExtraction(graphql_REs);
  var REs = dataManipulation(dataSet_REs);
  dataUpdating(sheet,standardSheet,REs[0],REs[1]);
  //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

  var queryRemote = `query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n      ${lc[i]}:${lcCode}\n      date_remote_realized:{from:"${startDate}"}\n      \n\t\t}\n    \n    page:1\n    per_page:3000\n\t)\n\t{\n    paging\n    {\n      total_items\n    }\n\t\tdata\n    {\n      person\n      {\n        id\n        full_name\n        email\n        contact_detail{\n          phone\n        }\n        home_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n      }\n      opportunity\n      {\n        id\n        programme\n        { short_name_display }\n        host_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n        remote_opportunity\n        project_fee\n        earliest_start_date\n        latest_end_date\n        specifics_info{\n          salary\n          salary_currency{\n            alphabetic_code\n          }\n        }\n        opportunity_duration_type{\n          duration_type\n          salary\n        }\n        \n      }\n      slot{\n        start_date\n        end_date\n      }\n      status\n      updated_at\n      date_approved\n      date_realized\n      experience_end_date\n    standards{\n        \n standard_option{\n          \n          meta\n        }\n \n   } }\n  }\n}`
    
  var graphql_Remote_REs = JSON.stringify({query: queryRemote})
  var dataSet_Remote_REs = dataExtraction(graphql_Remote_REs);
  var Remote_REs = dataManipulation(dataSet_Remote_REs);
  dataUpdating(sheet,standardSheet,Remote_REs[0],Remote_REs[1]);
  //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  var queryFIs = `query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n      ${lc[i]}:${lcCode}\n      experience_end_date:{from:"${startDate}"}\n statuses:[\"finished\",\"completed\"]     \n\t\t}\n    \n    page:1\n    per_page:3000\n\t)\n\t{\n    paging\n    {\n      total_items\n    }\n\t\tdata\n    {\n      person\n      {\n        id\n        full_name\n        email\n        contact_detail{\n          phone\n        }\n        home_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n      }\n      opportunity\n      {\n        id\n        programme\n        { short_name_display }\n        host_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n        remote_opportunity\n        project_fee\n        earliest_start_date\n        latest_end_date\n        specifics_info{\n          salary\n          salary_currency{\n            alphabetic_code\n          }\n        }\n        opportunity_duration_type{\n          duration_type\n          salary\n        }\n        \n      }\n      slot{\n        start_date\n        end_date\n      }\n      status\n      updated_at\n      date_approved\n      date_realized\n      experience_end_date\n standards{\n        \n        standard_option{\n          \n          meta\n        }\n \n   }\n  }\n}}`
  var graphql_FIs = JSON.stringify({query: queryFIs})
  var dataSet_FIs = dataExtraction(graphql_FIs);
  var FIs = dataManipulation(dataSet_FIs);
  dataUpdating(sheet,standardSheet,FIs[0],FIs[1]);
  //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

  var queryBreaks = `query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n      ${lc[i]}:${lcCode}\n      statuses:[\"approval_broken\",\"realization_broken\"]\n\t\t}\n    \n    page:1\n    per_page:5000\n\t)\n\t{\n    data\n    {\n      person\n      {\n        id\n      }\n      opportunity\n      {\n        id\n        }\n      \n      status      \n     }\n  }\n}`
  var graphql_Breaks = JSON.stringify({query: queryBreaks})
  var dataSet_Breaks = dataExtraction(graphql_Breaks);
  updateBreaks(sheet,standardSheet,dataSet_Breaks)

 }
  var now = new Date();
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interface"); // write sheet name 
  updateDate = sheet.getRange(8,5).setValue(now);
  updateDate = sheet.getRange(8,4).setValue("Succeed");

}

function updateBreaks(sheet,standardSheet,breaks){
  for(var i = 0; i < breaks.length; i++){
    if(breaks[i].person != null && breaks[i].opportunity != null){
      var rowIndexInSheet = sheet.createTextFinder(breaks[i].person.id+"_"+breaks[i].opportunity.id).matchEntireCell(false).findAll().map(x => x.getRow())
      var rowIndexInStandardSheet = standardSheet.createTextFinder(breaks[i].person.id+"_"+breaks[i].opportunity.id).matchEntireCell(false).findAll().map(x => x.getRow())
      if(rowIndexInSheet.length>0){
        sheet.getRange(rowIndexInSheet,4).setValue(breaks[i].status)
        standardSheet.getRange(rowIndexInStandardSheet,4).setValue(breaks[i].status) 
      }
    }
  }
}


function getLastRow(sheet)
{ 
  var lr = sheet.getLastRow()
  var range = sheet.getRange(5,1,lr,1).getValues()
  var lastRow = 0;
  for(var i =0; i < lr; i++){
    if(range[i]!= "")
    {
      lastRow++;
    }
  }
  return lastRow;
}


function onOpen() {  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Run')
      .addItem("Run the main code", 'main')
      .addItem("Run the permissions code", 'updatePermissions')
      .addToUi();
}

function getLastRowOfCol(sheet){
    let data = sheet.getRange(2,8,30,1).getDisplayValues()
    let i = 0
    while(data[i][0]!=""){
      i++
    }
    return i+1
  }

function updatePermissions(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("References")
  const lastRowForECBTeamColumn = getLastRowOfCol(sheet)
  const ecbTeam = sheet.getRange(2,8,lastRowForECBTeamColumn-1,1).getDisplayValues()
  const icxranges = sheet.getRange(2,9,19,2).getDisplayValues()
  const ogxranges = sheet.getRange(2,11,18,2).getDisplayValues()
  const standardRangesEBEdits = sheet.getRange(2,13,16,1).getDisplayValues()
  const standardRangesECBTeam = sheet.getRange(19,13).getDisplayValue()
  const documentsRange = sheet.getRange(2,14).getDisplayValue()
  var eb = sheet.getRange(35,4,sheet.getLastRow()-34,1).getDisplayValues()

  eb.forEach((item,index)=>{
        eb[index][0] = item[0].replace(" ","")              
    })
  
  
  //Deleting the ones that have accesses
    var editors = SpreadsheetApp.getActiveSpreadsheet().getEditors().toString().split(",")
    for(let i = 0;i<editors.length;i++){
      SpreadsheetApp.getActiveSpreadsheet().removeEditor(editors[i])
    }

    var viewers = SpreadsheetApp.getActiveSpreadsheet().getViewers().toString().split(",")
    for(let i = 0;i<viewers.length;i++){
      SpreadsheetApp.getActiveSpreadsheet().removeViewer(viewers[i])
    }

    for(let i = 0;i<eb.length;i++){
      SpreadsheetApp.getActiveSpreadsheet().addViewer(eb[i][0])
    }


  
  
  var icx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ICX Database/Auditing")
  var protections = icx.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    var protection = protections[i];
    protection.remove(); 
  }

  var ogx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("OGX Database/Auditing")
  var protections = ogx.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    var protection = protections[i];
    protection.remove();
  }

  var ogx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("OGX Standards Tracker")
  var protections = ogx.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    var protection = protections[i];
    protection.remove();
  }
  var icx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ICX Standards Tracker")
  var protections = icx.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    var protection = protections[i];
    protection.remove();
  }
  
  Logger.log("ICX Ranges")
  for(let j = 0; j < icxranges.length;j++){
    let range = SpreadsheetApp.getActive().getRange(`ICX Database/Auditing!${icxranges[j][1]}`)
    let protection = range.protect()
    for(let i = 0; i < ecbTeam.length;i++){
      SpreadsheetApp.getActiveSpreadsheet().addEditor(ecbTeam[i][0])
      protection.addEditor(ecbTeam[i][0])
    }
  }
  Logger.log("OGX Ranges")
  for(let j = 0; j < ogxranges.length;j++){
      let range = SpreadsheetApp.getActive().getRange(`OGX Database/Auditing!${ogxranges[j][1]}`)
      let protection = range.protect()
      for(let i = 0; i < ecbTeam.length;i++){
        protection.addEditor(ecbTeam[i][0])
      }
  }

  Logger.log("Standards EB Ranges")
  
  for(let j = 0; j < standardRangesEBEdits.length;j++){
    let rangeOGX = SpreadsheetApp.getActive().getRange(`OGX Standards Tracker!${standardRangesEBEdits[j][0]}`)
    let protectionOGX = rangeOGX.protect()
    let rangeICX = SpreadsheetApp.getActive().getRange(`ICX Standards Tracker!${standardRangesEBEdits[j][0]}`)
    let protectionICX = rangeICX.protect()
    for(let i = 0; i < eb.length;i++){
      protectionOGX.addEditor(eb[i][0])
      protectionICX.addEditor(eb[i][0])
    }
  }

  Logger.log("Standards ECB Ranges")
  var documents = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("IPS & OPS & Accommdation")
  var protections = documents.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    var protection = protections[i];
    protection.remove();
  }

  var rangeICX = SpreadsheetApp.getActive().getRange(`ICX Standards Tracker!${standardRangesECBTeam}`)
  var protectionICX = rangeICX.protect()
  var rangeOGX = SpreadsheetApp.getActive().getRange(`OGX Standards Tracker!${standardRangesECBTeam}`)
  var protectionOGX = rangeOGX.protect()
  var rangeDocuments = SpreadsheetApp.getActive().getRange(`IPS & OPS & Accommdation!${documentsRange}`)
  var protectionDocuments = rangeDocuments.protect()
  for(let i = 0; i < ecbTeam.length;i++){
    protectionICX.addEditor(ecbTeam[i][0])
    protectionOGX.addEditor(ecbTeam[i][0])
    protectionDocuments.addEditor(ecbTeam[i][0])
  }

}
 
