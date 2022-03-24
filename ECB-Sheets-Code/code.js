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
    dataSet[i].opportunity !=null?dataSet[i].person.id+"_"+dataSet[i].opportunity.id:dataSet[i].person.id+"_", dataSet[i].person.full_name,dataSet[i].opportunity !=null?dataSet[i].opportunity.programme.short_name_display:"", dataSet[i].status,
    dataSet[i].opportunity !=null?dataSet[i].opportunity.remote_opportunity == true?"Yes":"No":"", "",
    
    dataSet[i].person.email,  dataSet[i].person.contact_detail != null?dataSet[i].person.contact_detail.phone: "", dataSet[i].person.home_mc.name, dataSet[i].person.home_lc.name, "",
    
    dataSet[i].opportunity !=null? dataSet[i].opportunity.home_mc.name:"",
    dataSet[i].opportunity !=null?dataSet[i].opportunity.host_lc.name:"",
    dataSet[i].opportunity !=null?"https://aiesec.org/opportunity/"+dataSet[i].opportunity.id:"https://aiesec.org/opportunity/",
    dataSet[i].opportunity !=null?dataSet[i].opportunity.programme.short_name_display:"",
    dataSet[i].opportunity !=null?dataSet[i].opportunity.opportunity_duration_type != null?dataSet[i].opportunity.opportunity_duration_type.duration_type:"":""/*duration_type*/,
    dataSet[i].opportunity !=null?dataSet[i].opportunity.programme.short_name_display == "GV"?dataSet[i].opportunity.project_fee.fee+" "+dataSet[i].opportunity.project_fee.currency:"-":""/*Project Fees*/,
    dataSet[i].opportunity !=null?dataSet[i].opportunity.specifics_info.salary == null?0:dataSet[i].opportunity.specifics_info.salary+" "+dataSet[i].opportunity.specifics_info.salary_currency.alphabetic_code:""/*salary Fees*/,"",
    
    dataSet[i].date_approved != null?dataSet[i].date_approved.toString().substring(0,10):dataSet[i].updated_at.toString().substring(0,10)/*APD Date*/,
    dataSet[i].slot!=null?dataSet[i].slot.start_date:""/*slot start date*/,
    dataSet[i].slot!=null?dataSet[i].slot.end_date:""/*slot end date*/,
    ""/*RE date*/,
    ""/*Fi date*/,
    dataSet[i].status == "remote_realized"?dataSet[i].updated_at.toString().substring(0,10):""/*remote date*/
 ]);

 standards.push([
   dataSet[i].opportunity !=null?dataSet[i].person.id+"_"+dataSet[i].opportunity.id:dataSet[i].person.id+"_", 
   dataSet[i].person.full_name, 
   dataSet[i].opportunity !=null?dataSet[i].opportunity.programme.short_name_display:"", 
   dataSet[i].status,
   dataSet[i].opportunity !=null?dataSet[i].opportunity.remote_opportunity == true?"Yes":"No":""/*remote status*/,
   "",
   dataSet[i].date_approved != null?dataSet[i].date_approved.toString().substring(0,10):dataSet[i].updated_at.toString().substring(0,10)/*APD Date*/,
   ""/*RE date*/,
   ""/*Fi date*/,
   dataSet[i].status == "remote_realized"?dataSet[i].updated_at.toString().substring(0,10):""/*remote date*/,
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
  Logger.log(sheet.getName())
  var newRows = [];
  for(var i =0; i < rows.length ;i++){
      var rowIndexInSheet = sheet.createTextFinder(rows[i][0]).matchEntireCell(true).findAll().map(x => x.getRow())
      if(rowIndexInSheet.length > 0)
      {
        var value = sheet.getRange(rowIndexInSheet[0],4).getValue()
        if(value != rows[i][3])
        {
          sheet.getRange(rowIndexInSheet[0],4).setValue(rows[i][3])
          if(rows[i][3] == "realized" || rows[i][3] == "finished" || rows[i][3] == "completed")
          {
            sheet.getRange(rowIndexInSheet[0],23).setValue(rows[i][22])
            sheet.getRange(rowIndexInSheet[0],24).setValue(rows[i][23])
          }
          else if(rows[i][3] == "remote_realized")
          {
            sheet.getRange(rowIndexInSheet[0],25).setValue(rows[i][24])
          }
        }
        else
        {
          continue
        } 
      }
      else{
        newRows.push(rows[i])
      }
    
  }
  if(newRows.length > 0){
    sheet.getRange(getLastRow(sheet)+5,1,newRows.length,newRows[0].length).setValues(newRows);
  }
  
  var newRows = [];
  for(var i =0; i < standards.length;i++)
  {
    Logger.log(i)
      var rowIndexInStandardSheet = standardSheet.createTextFinder(standards[i][0]).matchEntireCell(true).findAll().map(x => x.getRow())
      if(rowIndexInStandardSheet.length>0){
          if((standards[i][3] == "realized" || standards[i][3] == "finished" || standards[i][3] == "completed") && standardSheet.getRange(rowIndexInStandardSheet,4).getValue()!="completed"){
            standardSheet.getRange(rowIndexInStandardSheet,4).setValue(standards[i][3])
            standardSheet.getRange(rowIndexInStandardSheet,8).setValue(standards[i][7])
            standardSheet.getRange(rowIndexInStandardSheet,9).setValue(standards[i][8])   
            for(var k =0; k < 16;k++){
              r = k*2;
              standardSheet.getRange(rowIndexInStandardSheet,12+r).setValue(standards[i][11+r])
            }     
          }
          else if(standards[i][3] == "remote_realized"){
            standardSheet.getRange(rowIndexInStandardSheet,4).setValue(standards[i][3])
            standardSheet.getRange(rowIndexInStandardSheet,9).setValue(standards[i][8])   
          }
      }
      else{
        newRows.push(standards[i])
      }
          
  }
  if(newRows.length > 0){
    standardSheet.getRange(getLastRow(standardSheet)+5,1,newRows.length,newRows[0].length).setValues(newRows);
  }


}
 

function main()
{
  try{
  var sheetInterface = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interface"); // write sheet name 
  var startDate = Utilities.formatDate(sheetInterface.getRange(8,3).getValue(), "GMT+2", "dd/MM/yyyy");
  var lcCode = sheetInterface.getRange(8,2).getValue();
  var lc = ["opportunity_home_lc","person_home_lc"]

  for(var i = 0; i < 2;i++){

  if(i == 0){
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
  if(dataSet_APDs.length>0){
    Logger.log(`${lc[i]} APDs`)
    var APDs = dataManipulation(dataSet_APDs);
    dataUpdating(sheet,standardSheet,APDs[0],APDs[1]);
  }

  //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 var queryREs = `query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n      ${lc[i]}:${lcCode}\n      date_realized:{from:"${startDate}"}\n      \n\t\t}\n    \n    page:1\n    per_page:3000\n\t)\n\t{\n    paging\n    {\n      total_items\n    }\n\t\tdata\n    {\n      person\n      {\n        id\n        full_name\n        email\n        contact_detail{\n          phone\n        }\n        home_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n      }\n      opportunity\n      {\n        id\n        programme\n        { short_name_display }\n        host_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n        remote_opportunity\n        project_fee\n        earliest_start_date\n        latest_end_date\n        specifics_info{\n          salary\n          salary_currency{\n            alphabetic_code\n          }\n        }\n        opportunity_duration_type{\n          duration_type\n          salary\n        }\n        \n      }\n      slot{\n        start_date\n        end_date\n      }\n      status\n      updated_at\n      date_approved\n      date_realized\n      experience_end_date\n    standards{\n        \n standard_option{\n          \n          meta\n        }\n \n   } }\n  }\n}`
    
  var graphql_REs = JSON.stringify({query: queryREs})
  var dataSet_REs = dataExtraction(graphql_REs);
  if(dataSet_REs.length>0 ){
        Logger.log(`${lc[i]} REs`)

    var REs = dataManipulation(dataSet_REs);
    dataUpdating(sheet,standardSheet,REs[0],REs[1]);
  }
  //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

  var queryRemote = `query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n      ${lc[i]}:${lcCode}\n      date_remote_realized:{from:"${startDate}"}\n      \n\t\t}\n    \n    page:1\n    per_page:3000\n\t)\n\t{\n    paging\n    {\n      total_items\n    }\n\t\tdata\n    {\n      person\n      {\n        id\n        full_name\n        email\n        contact_detail{\n          phone\n        }\n        home_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n      }\n      opportunity\n      {\n        id\n        programme\n        { short_name_display }\n        host_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n        remote_opportunity\n        project_fee\n        earliest_start_date\n        latest_end_date\n        specifics_info{\n          salary\n          salary_currency{\n            alphabetic_code\n          }\n        }\n        opportunity_duration_type{\n          duration_type\n          salary\n        }\n        \n      }\n      slot{\n        start_date\n        end_date\n      }\n      status\n      updated_at\n      date_approved\n      date_realized\n      experience_end_date\n    standards{\n        \n standard_option{\n          \n          meta\n        }\n \n   } }\n  }\n}`
    
  var graphql_Remote_REs = JSON.stringify({query: queryRemote})
  var dataSet_Remote_REs = dataExtraction(graphql_Remote_REs);
  if(dataSet_Remote_REs.length>0)
  {
        Logger.log(`${lc[i]} REMOTES`)

    var Remote_REs = dataManipulation(dataSet_Remote_REs);
    dataUpdating(sheet,standardSheet,Remote_REs[0],Remote_REs[1]);
  }

  
  //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  var queryFIs = `query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n      ${lc[i]}:${lcCode}\n      experience_end_date:{from:"${startDate}"}\n statuses:[\"finished\",\"completed\"]     \n\t\t}\n    \n    page:1\n    per_page:3000\n\t)\n\t{\n    paging\n    {\n      total_items\n    }\n\t\tdata\n    {\n      person\n      {\n        id\n        full_name\n        email\n        contact_detail{\n          phone\n        }\n        home_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n      }\n      opportunity\n      {\n        id\n        programme\n        { short_name_display }\n        host_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n        remote_opportunity\n        project_fee\n        earliest_start_date\n        latest_end_date\n        specifics_info{\n          salary\n          salary_currency{\n            alphabetic_code\n          }\n        }\n        opportunity_duration_type{\n          duration_type\n          salary\n        }\n        \n      }\n      slot{\n        start_date\n        end_date\n      }\n      status\n      updated_at\n      date_approved\n      date_realized\n      experience_end_date\n standards{\n        \n        standard_option{\n          \n          meta\n        }\n \n   }\n  }\n}}`
  var graphql_FIs = JSON.stringify({query: queryFIs})
  var dataSet_FIs = dataExtraction(graphql_FIs);
  if(dataSet_FIs.length>0){
        Logger.log(`${lc[i]} FIs`)

    var FIs = dataManipulation(dataSet_FIs);
    dataUpdating(sheet,standardSheet,FIs[0],FIs[1]);
  }
  //-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

  var queryBreaks = `query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n      ${lc[i]}:${lcCode}\n last_interaction:{from:\"${startDate}\"}   \n\t\t\t       statuses:[\"approval_broken\",\"realization_broken\"]\n\t\t}\n    \n    page:1\n    per_page:1000\n\t)\n\t{\n    data\n    {\n      person\n      {\n        id\n      }\n      opportunity\n      {\n        id\n        }\n      \n      status      \n     }\n  }\n}`
  var graphql_Breaks = JSON.stringify({query: queryBreaks})
  var dataSet_Breaks = dataExtraction(graphql_Breaks);
  updateBreaks(sheet,standardSheet,dataSet_Breaks)

 }
  var now = new Date();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interface"); // write sheet name 
  updateDate = sheet.getRange(8,5).setValue(now);
  updateDate = sheet.getRange(8,4).setValue("Succeed");

  }
  catch(e){
    Logger.log(e.toString())
    var now = new Date();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interface"); // write sheet name 
    updateDate = sheet.getRange(8,5).setValue(now);
    updateDate = sheet.getRange(8,4).setValue("Failed");
  }
}






function updateBreaks(sheet,standardSheet,breaks){
  for(var i = 0; i < breaks.length; i++){
    if(breaks[i].person != null && breaks[i].opportunity != null){
      var rowIndexInSheet = sheet.createTextFinder(breaks[i].person.id+"_"+breaks[i].opportunity.id).matchEntireCell(true).findAll().map(x => x.getRow())
      var rowIndexInStandardSheet = standardSheet.createTextFinder(breaks[i].person.id+"_"+breaks[i].opportunity.id).matchEntireCell(true).findAll().map(x => x.getRow())
      if(rowIndexInSheet.length>0 && sheet.getRange(rowIndexInSheet,4).getValue()!=breaks[i].status){
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
      SpreadsheetApp.getActiveSpreadsheet().addEditor(eb[i][0])
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

  Logger.log("ICX Ranges")
  var icxSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`ICX Database/Auditing`)
  var protections = icxSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for(let j = 0; j < protections.length;j++){
    for(let i = 0; i < eb.length;i++){
      protections[j].removeEditor(eb[i][0])
    }
  }


  Logger.log("OGX Ranges")
  var ogxSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`OGX Database/Auditing`)
  var protections = ogxSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for(let j = 0; j < protections.length;j++){
      for(let i = 0; i < eb.length;i++){
        protections[j].removeEditor(eb[i][0])
      }
  }

}
