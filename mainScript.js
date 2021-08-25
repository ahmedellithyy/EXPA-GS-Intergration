function dataExtraction(graphql)
{
  
  var requestOptions = {
    'method': 'post',
    'payload': graphql,
    'contentType':'application/json',
    'headers':{
      'access_token': "" //write the access token here
    }
  };
  var response = UrlFetchApp.fetch("https://gis-api.aiesec.org/graphql?access_token=", requestOptions);  //write the access token here
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
    dataSet[i].person.id+"_"+dataSet[i].opportunity.id, dataSet[i].person.full_name, dataSet[i].status,""/*remote status*/, "",
    
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
   dataSet[i].person.id+"_"+dataSet[i].opportunity.id, dataSet[i].person.full_name, dataSet[i].opportunity.programme.short_name_display, dataSet[i].status,""/*remote status*/,"",
   dataSet[i].person.email,  "" /*phone number*/, dataSet[i].person.home_mc.name, dataSet[i].person.home_lc.name, "",
   dataSet[i].opportunity.home_mc.name, dataSet[i].opportunity.host_lc.name,"https://aiesec.org/opportunity/"+dataSet[i].opportunity.id,""/*duration_type*/,""/*Project Fees*/,""/*salary Fees*/,"",
   ""/*APD Date*/,
   ""/*RE date*/,
   ""/*Fi date*/,
   ""/*remote date*/,"","",
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
        rows[i][3] = "Yes";
        standards[i][4] = "Yes";
  }
  else{
    rows[i][3] = "No"
    standards[i][4] = "No";

  }

  if (dataSet[i].person.contact_detail != null){
      rows[i][6] = dataSet[i].person.contact_detail.phone;
      standards[i][7] = dataSet[i].person.contact_detail.phone;

  }

  if(dataSet[i].opportunity.programme.short_name_display == "GV")
  {
      rows[i][15] = dataSet[i].opportunity.project_fee.fee+" "+dataSet[i].opportunity.project_fee.currency;
      standards[i][15] = dataSet[i].opportunity.project_fee.fee+" "+dataSet[i].opportunity.project_fee.currency;

  }
  else{
    if(dataSet[i].opportunity.specifics_info.salary == null){
        dataSet[i].opportunity.specifics_info.salary = 0;
        rows[i][16] = dataSet[i].opportunity.specifics_info.salary+" "+dataSet[i].opportunity.specifics_info.salary_currency.alphabetic_code;
        standards[i][16] = dataSet[i].opportunity.specifics_info.salary+" "+dataSet[i].opportunity.specifics_info.salary_currency.alphabetic_code;
    }
    else{
      rows[i][16] = dataSet[i].opportunity.specifics_info.salary+" "+dataSet[i].opportunity.specifics_info.salary_currency.alphabetic_code;
      standards[i][16] = dataSet[i].opportunity.specifics_info.salary+" "+dataSet[i].opportunity.specifics_info.salary_currency.alphabetic_code;

    }
  }
  if(dataSet[i].slot!=null){
    rows[i][19] = dataSet[i].slot.start_date;
    rows[i][20] = dataSet[i].slot.end_date;
  }

  if(dataSet[i].opportunity.opportunity_duration_type != null){
      rows[i][14]=dataSet[i].opportunity.opportunity_duration_type.duration_type;
      standards[i][14]=dataSet[i].opportunity.opportunity_duration_type.duration_type;

  }

  // set approval, realization and experience end dates
    if(dataSet[i].date_approved != null){
      rows[i][18]=  dataSet[i].date_approved.toString().substring(0,10);
      standards[i][18]=  dataSet[i].date_approved.toString().substring(0,10);

    }
    else{
          rows[i][18]=  dataSet[i].updated_at.toString().substring(0,10);
          standards[i][18]=  dataSet[i].updated_at.toString().substring(0,10);

    }

    if(dataSet[i].status == "remote_realized" ){
          rows[i][23]=  dataSet[i].updated_at.toString().substring(0,10);
          standards[i][21]=  dataSet[i].updated_at.toString().substring(0,10);
         
    }
    if(dataSet[i].status == "realized" )
    { 
      if(dataSet[i].date_realized != null){
        rows[i][21]=  dataSet[i].date_realized.toString().substring(0,10);
        standards[i][19]=  dataSet[i].date_realized.toString().substring(0,10);
      }
      if(dataSet[i].experience_end_date != null){
              rows[i][22]=  dataSet[i].experience_end_date.toString().substring(0,10);
              standards[i][20]=  dataSet[i].experience_end_date.toString().substring(0,10);
        }
      
    }
    else if (dataSet[i].status == "finished" || dataSet[i].status == "completed"){
      if(dataSet[i].date_realized != null){
        rows[i][21]=  dataSet[i].date_realized.toString().substring(0,10);
        standards[i][19]=  dataSet[i].date_realized.toString().substring(0,10);

      }
        if(dataSet[i].experience_end_date != null){
              rows[i][22]=  dataSet[i].experience_end_date.toString().substring(0,10);
              standards[i][20]=  dataSet[i].experience_end_date.toString().substring(0,10);
        }
        
    }
    for(var k =0; k <16;k++)
    {
      if(dataSet[i].standards[k].standard_option != null ){
        r = k*2;
       standards[i][25+r] = dataSet[i].standards[k].standard_option.meta.option
     }
    }
 }

    return [rows, standards]; 
}
// Take the raw data recieved from the HTTP response and arrange it into the corresponding sheet
function dataUpdating(sheet,standardSheet, rows, standards)
{
  var lr = getLastRow(sheet);
  var oldRows = sheet.getRange(5,1,lr,24).getValues();
  var newRows = [];
  for(var i =0; i < rows.length;i++)
  {
      var duplicated = false;
      for(var j = 0; j <oldRows.length ;j++)
      {
            if(rows[i][0] == oldRows[j][0])
            {
                duplicated = true;
                oldRows[j][2] = rows[i][2];
                if(rows[i][2] == "realized" || rows[i][2] == "finished" || rows[i][2] == "completed"){
                  oldRows[j].splice(21,1,rows[i][21]);
                  oldRows[j].splice(22,1,rows[i][22]);
                }
                else if(rows[i][2] == "remote_realized"){
                  oldRows[j].splice(23,1,rows[i][23]);
                }
            }
            
      }
      if(!duplicated){
        newRows.push(rows[i]); 
      }
  }

  afterEdits = sheet.getRange(5,1,oldRows.length,24).setValues(oldRows);

  if(newRows.length > 0){
    dataRange = sheet.getRange(lr+5,1,newRows.length,24).setValues(newRows);
  }


  lr = getLastRow(standardSheet);
  oldRows = standardSheet.getRange(5,1,lr,56).getValues();
  newRows = [];
  for(var i =0; i < rows.length;i++)
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
                    oldRows[j].splice(25+r,1,standards[i][25+r])
                }
                if(rows[i][2] == "realized" || rows[i][2] == "finished" || rows[i][2] == "completed"){
                  oldRows[j].splice(19,1,rows[i][19]);
                  oldRows[j].splice(20,1,rows[i][20]);
                }
                else if(rows[i][2] == "remote_realized"){
                  oldRows[j].splice(21,1,rows[i][21]);
                }
            }
            
      }
      if(!duplicated){
        newRows.push(standards[i]); 
      }
  }

  afterEdits = standardSheet.getRange(5,1,oldRows.length,56).setValues(oldRows);

  if(newRows.length > 0){
    dataRange = standardSheet.getRange(lr+5,1,newRows.length,56).setValues(newRows);
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

  var queryApplications = `query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n      person_home_lc:${lcCode}\n      created_at:{from:\"01/08/2021\"}\n      \n\t\t}\n    \n    page:1\n    per_page:5000\n\t)\n\t{\n    data\n    {\n   id\n   person\n      {\n        id\n        full_name\n      managers{full_name}  \n      }\n      created_at\n      \n      status\n     }\n  }\n}`
  var graphql_Applications = JSON.stringify({query: queryApplications});
  var dataSet_Applications = dataExtraction(graphql_Applications);
  updateApplications(dataSet_Applications)

  var queryOpen = `query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n      opportunity_home_lc:${lcCode}\n      created_at:{from:\"01/08/2021\"}\n      status:"open"\n\t\t}\n    \n    page:1\n    per_page:5000\n\t)\n\t{\n    data\n    {\n id\n      person\n      {\n        id\n        full_name\n      managers{full_name}  \n      }\n      created_at\n      \n      status\n     }\n  }\n}`
  var graphql_Open = JSON.stringify({query:queryOpen});
  var dataSet_Open = dataExtraction(graphql_Open)
  udapteOpen(dataSet_Open)

 }
  var now = new Date();
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interface"); // write sheet name 
  updateDate = sheet.getRange(8,5).setValue(now);
  updateDate = sheet.getRange(8,4).setValue("Succeed");

}

function updateBreaks(sheet,standardSheet,breaks){
 var rows = [];
 for(var i = 0; i < breaks.length; i++){
    rows.push([
    breaks[i].person.id+"_"+breaks[i].opportunity.id,breaks[i].status])
  }

  var lr = getLastRow(sheet);
  var oldRows = sheet.getRange(5,1,lr,4).getValues();
  for(var i =0; i < rows.length;i++)
  {
      for(var j = 0; j <oldRows.length ;j++)
      {
            if(rows[i][0] == oldRows[j][0])
            {
                oldRows[j][2] = rows[i][1];
            } 
      }
  }
    afterEdits = sheet.getRange(5,1,oldRows.length,4).setValues(oldRows);

  var lr = getLastRow(standardSheet);
  var oldRows = standardSheet.getRange(5,1,lr,4).getValues();
  for(var i =0; i < rows.length;i++)
  {
      for(var j = 0; j <oldRows.length ;j++)
      {
            if(rows[i][0] == oldRows[j][0])
            {
                oldRows[j][3] = rows[i][1];
            } 
      }
  }
    afterEdits = standardSheet.getRange(5,1,oldRows.length,4).setValues(oldRows);
}

function updateApplications(applications){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("OGX Applicants"); // write sheet name 
  var rows = [];
  for(var i = 0; i < applications.length; i++){
    rows.push([applications[i].id, applications[i].person.full_name, applications[i].status, applications[i].created_at, "", ""])
    if(applications[i].person.managers != null){
      rows[i][4] = "Yes"
      var managers = []
      for(var j = 0; j < applications[i].person.managers.length; j++){
        managers.push(applications[i].person.managers[j].full_name+",")
      }
      rows[i][5] = managers;
    }
    else{
      rows[i][4] = "No"
    }
    
  }

  var lr = getLastRow(sheet);
  var oldRows = sheet.getRange(3,1,lr,6).getValues();
  var newRows =[]
  for(var i =0; i < rows.length;i++)
  {       
      var duplicated = false;
      for(var j = 0; j <oldRows.length ;j++)
      {
            if(rows[i][0] == oldRows[j][0])
            {
                duplicated = true
                oldRows[j][2] = rows[i][2];
            }
      }
      if(!duplicated){
        newRows.push(rows[i])
      }
  }
  afterEdits = sheet.getRange(3,1,oldRows.length,6).setValues(oldRows);
  if(newRows.length > 0){
    dataRange = sheet.getRange(lr+3,1,newRows.length,6).setValues(newRows);
  }

}

function udapteOpen(open){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Open Applications"); // write sheet name 
  var newRows = [];
  for(var i = 0; i < open.length; i++){
      newRows.push([open[i].id, open[i].person.full_name, open[i].status, open[i].created_at])
  }
  var lr = getLastRow(sheet);
  var range = sheet.getRange(3,1,lr,4)
  range.clearContent()
  range = sheet.getRange(3,1,newRows.length,4).setValues(newRows);

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
