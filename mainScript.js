//Extract OGX Data
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
  var response = UrlFetchApp.fetch("https://gis-api.aiesec.org/graphql?access_token=ACCESS_TOKEN", requestOptions);
  var recievedDate = JSON.parse(response.getContentText());

  return recievedDate.data.allOpportunityApplication.data;
}

//Set the extracted data into 2D array
function dataManipulation(dataSet)
{
 var rows = [];
 for (var i = 0; i < dataSet.length; i++){
      
  rows.push([
  dataSet[i].person.id+"_"+dataSet[i].opportunity.id, dataSet[i].person.full_name, dataSet[i].status, dataSet[i].opportunity.remote_opportunity, "",
  
  dataSet[i].person.email,  dataSet[i].person.phone, dataSet[i].person.home_mc.name, dataSet[i].person.home_lc.name, "",
  
  dataSet[i].opportunity.home_mc.name,dataSet[i].opportunity.host_lc.name,"https://aiesec.org/opportunity/"+dataSet[i].opportunity.id,dataSet[i].opportunity.programme. short_name_display,""/*duration_type*/,""/*fees*/,"",
  
  ""/*APD Date*/,
  ""/*slot start date*/,
  ""/*slot end date*/,
  ""/*RE date*/,
  ""/*Fi date*/,
  ""/*remote date*/,
  "",
  "",
  "",
  "",
  "" /*standard #1*/,
  "" /*standard #2*/,
  "" /*standard #3*/,
  "" /*standard #4*/,
  "" /*standard #5*/,
  "" /*standard #6*/,
  "" /*standard #7*/,
  "" /*standard #8*/,
  "" /*standard #9*/,
  "" /*standard #10*/,
  "" /*standard #11*/,
  "" /*standard #12*/,
  "" /*standard #13*/,
  "" /*standard #14*/,
  "" /*standard #15*/,
  "" /*standard #16*/,


 ]);
  // Value setting for Salary/Project Fees column 
  if(dataSet[i].opportunity.programme.short_name_display == "GV")
  {
            rows[i][15] = dataSet[i].opportunity.project_fee.fee+" "+dataSet[i].opportunity.project_fee.currency;
  }
  else{
            if(dataSet[i].opportunity.specifics_info.salary == null){
            dataSet[i].opportunity.specifics_info.salary = 0;
            rows[i][15] = dataSet[i].opportunity.specifics_info.salary+" "+dataSet[i].opportunity.specifics_info.salary_currency.alphabetic_code;}
            else{
              rows[i][15] = dataSet[i].opportunity.specifics_info.salary+" "+dataSet[i].opportunity.specifics_info.salary_currency.alphabetic_code;
            }
  }
  if(dataSet[i].slot!=null){
    rows[i][18] = dataSet[i].slot.start_date;
    rows[i][19] = dataSet[i].slot.end_date;
  }

  if(dataSet[i].opportunity.opportunity_duration_type != null){
      rows[i][14]=dataSet[i].opportunity.opportunity_duration_type.duration_type;
  }

  // set approval, realization and experience end dates
  if(dataSet[i].date_approved != null){
     rows[i][17]=  dataSet[i].date_approved.toString().substring(0,10);
  }
  else{
        rows[i][17]=  dataSet[i].updated_at.toString().substring(0,10);
  }

  if(dataSet[i].status == "remote_realized" ){
        rows[i][22]=  dataSet[i].updated_at.toString().substring(0,10);
  }
  else if(dataSet[i].status == "realized" ){
        rows[i][20]=  dataSet[i].date_realized.toString().substring(0,10);
  }
  else if (dataSet[i].status == "finished" || dataSet[i].status == "completed"){
            rows[i][20]=  dataSet[i].date_realized.toString().substring(0,10);
            rows[i][21]=  dataSet[i].experience_end_date.toString().substring(0,10);
 }

  //push standards
  for(var k =0; k <16;k++)
  {
    if(dataSet[i].standards[k].standard_option != null ){
       
       rows[i][26+k] = dataSet[i].standards[k].standard_option.meta.option
    
     }
  
  }
 }



    return rows; 
}
// Take the raw data recieved from the HTTP response and arrange it into the corresponding sheet
function dataUpdating(sheet, rows)
{
  var lr = sheet.getLastRow();
  var oldRows = sheet.getRange(5,1,lr-2,23).getValues();

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
                if(rows[i][2] == "realized"){
                  oldRows[j].splice(20,1,rows[i][20]);
                }
                else if(rows[i][2] == "finished")
                {
                  oldRows[j].splice(20,1,rows[i][20]);
                  oldRows[j].splice(21,1,rows[i][21]);
                }
                else if(rows[i][2] == "remote_realized"){
                  oldRows[j].splice(22,1,rows[i][22]);
                }
            }
            
      }
      if(!duplicated){
        newRows.push(rows[i]); 
      }
  }

  afterEdits = sheet.getRange(5,1,lr-2,23).setValues(oldRows);

  if(newRows.length > 0){
    dataRange = sheet.getRange(lr+1,1,newRows.length,newRows[0].length).setValues(newRows);
  }

 // dataSort = sheet.getRange(5,1,sheet.getLastRow(),sheet.getLastColumn()).sort({column: 18, ascending: false});

}

//main function that use others function to perform whole Data Extraction process
function main_ICX()
{
  var graphql_ICX_APDs = JSON.stringify({
    query: "query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n      opportunity_home_lc:1064\n      date_approved:{from:\"01/07/2021\"}\n      \n\t\t}\n    \n    page:1\n    per_page:3000\n\t)\n\t{\n    paging\n    {\n      total_items\n    }\n\t\tdata\n    {\n      person\n      {\n        id\n        full_name\n        email\n        contact_detail{\n          phone\n        }\n        home_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n      }\n      opportunity\n      {\n        id\n        programme\n        { short_name_display }\n        host_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n        remote_opportunity\n        project_fee\n        earliest_start_date\n        latest_end_date\n        specifics_info{\n          salary\n          salary_currency{\n            alphabetic_code\n          }\n        }\n        opportunity_duration_type{\n          duration_type\n          salary\n        }\n        \n      }\n      slot{\n        start_date\n        end_date\n      }\n      status\n      updated_at\n      date_approved\n      date_realized\n      experience_end_date\n    standards{\n        \n standard_option{\n          \n          meta\n        }\n \n   } }\n  }\n}"
    
    ,
    variables: {}
  })
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ICX Database/Auditing"); // write sheet name 
  var dataSet_ICX_APDs = dataExtraction(graphql_ICX_APDs);
  var rows_ICX_APDs = dataManipulation(dataSet_ICX_APDs);
  dataUpdating(sheet,rows_ICX_APDs);

  var graphql_ICX_REs = JSON.stringify({
    query: "query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n      opportunity_home_lc:1064\n      date_realized:{from:\"01/07/2021\"}\n     \n\t\t}\n    \n    page:1\n    per_page:3000\n\t)\n\t{\n    paging\n    {\n      total_items\n    }\n\t\tdata\n    {\n      person\n      {\n        id\n        full_name\n        email\n        contact_detail{\n          phone\n        }\n        home_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n      }\n      opportunity\n      {\n        id\n        programme\n        { short_name_display }\n        host_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n        remote_opportunity\n        project_fee\n        earliest_start_date\n        latest_end_date\n        specifics_info{\n          salary\n          salary_currency{\n            alphabetic_code\n          }\n        }\n        opportunity_duration_type{\n          duration_type\n          salary\n        }\n        \n      }\n      slot{\n        start_date\n        end_date\n      }\n      status\n      updated_at\n      date_approved\n      date_realized\n      experience_end_date\n standards{\n        \n        standard_option{\n          \n          meta\n        }\n \n  } }\n  }\n}"
    
    ,
    variables: {}
  })
  var dataSet_ICX_REs = dataExtraction(graphql_ICX_REs);
  var rows_ICX_REs = dataManipulation(dataSet_ICX_REs);
  dataUpdating(sheet,rows_ICX_REs);

  var graphql_ICX_Remote_REs = JSON.stringify({
    query: "query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n      opportunity_home_lc:1064\n      date_remote_realized:{from:\"01/07/2021\"}\n\n\t\t}\n    \n    page:1\n    per_page:3000\n\t)\n\t{\n    paging\n    {\n      total_items\n    }\n\t\tdata\n    {\n      person\n      {\n        id\n        full_name\n        email\n        contact_detail{\n          phone\n        }\n        home_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n      }\n      opportunity\n      {\n        id\n        programme\n        { short_name_display }\n        host_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n        remote_opportunity\n        project_fee\n        earliest_start_date\n        latest_end_date\n        specifics_info{\n          salary\n          salary_currency{\n            alphabetic_code\n          }\n        }\n        opportunity_duration_type{\n          duration_type\n          salary\n        }\n        \n      }\n      slot{\n        start_date\n        end_date\n      }\n      status\n      updated_at\n      date_approved\n      date_realized\n      experience_end_date\n standards{\n        \n        standard_option{\n          \n          meta\n        }\n \n   }}\n  }\n}"
    
    
    ,
    variables: {}
  })
  var dataSet_ICX_Remote_REs = dataExtraction(graphql_ICX_Remote_REs);
  var rows_ICX_Remote_REs = dataManipulation(dataSet_ICX_Remote_REs);
  dataUpdating(sheet,rows_ICX_Remote_REs);


  var graphql_ICX_FIs = JSON.stringify({
    query: "query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n      opportunity_home_lc:1064\n      experience_end_date:{from:\"01/07/2021\"}\n statuses:[\"finished\",\"completed\"]     \n\t\t}\n    \n    page:1\n    per_page:3000\n\t)\n\t{\n    paging\n    {\n      total_items\n    }\n\t\tdata\n    {\n      person\n      {\n        id\n        full_name\n        email\n        contact_detail{\n          phone\n        }\n        home_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n      }\n      opportunity\n      {\n        id\n        programme\n        { short_name_display }\n        host_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n        remote_opportunity\n        project_fee\n        earliest_start_date\n        latest_end_date\n        specifics_info{\n          salary\n          salary_currency{\n            alphabetic_code\n          }\n        }\n        opportunity_duration_type{\n          duration_type\n          salary\n        }\n        \n      }\n      slot{\n        start_date\n        end_date\n      }\n      status\n      updated_at\n      date_approved\n      date_realized\n      experience_end_date\n standards{\n        \n        standard_option{\n          \n          meta\n        }\n \n   }\n  }\n}}"
    
    
    ,
    variables: {}
  })
  var dataSet_ICX_FIs = dataExtraction(graphql_ICX_FIs);
  var rows_ICX_FIs = dataManipulation(dataSet_ICX_FIs);
  dataUpdating(sheet,rows_ICX_FIs);


}

function main_OGX()
{
  var graphql_OGX_APDs = JSON.stringify({
    query: "query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n      person_home_lc:1064\n      date_approved:{from:\"01/07/2021\"}\n      \n\t\t}\n    \n    page:1\n    per_page:3000\n\t)\n\t{\n    paging\n    {\n      total_items\n    }\n\t\tdata\n    {\n      person\n      {\n        id\n        full_name\n        email\n        contact_detail{\n          phone\n        }\n        home_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n      }\n      opportunity\n      {\n        id\n        programme\n        { short_name_display }\n        host_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n        remote_opportunity\n        project_fee\n        earliest_start_date\n        latest_end_date\n        specifics_info{\n          salary\n          salary_currency{\n            alphabetic_code\n          }\n        }\n        opportunity_duration_type{\n          duration_type\n          salary\n        }\n        \n      }\n      slot{\n        start_date\n        end_date\n      }\n      status\n      updated_at\n      date_approved\n      date_realized\n      experience_end_date\n    standards{\n        \n standard_option{\n          \n          meta\n        }\n \n   } }\n  }\n}"
    
    ,
    variables: {}
  })
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("OGX Database/Auditing"); // write sheet name 
  var dataSet_OGX_APDs = dataExtraction(graphql_OGX_APDs);
  var rows_OGX_APDs = dataManipulation(dataSet_OGX_APDs);
  dataUpdating(sheet,rows_OGX_APDs);

  var graphql_OGX_REs = JSON.stringify({
    query: "query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n      person_home_lc:1064\n      date_realized:{from:\"01/07/2021\"}\n      \n\t\t}\n    \n    page:1\n    per_page:3000\n\t)\n\t{\n    paging\n    {\n      total_items\n    }\n\t\tdata\n    {\n      person\n      {\n        id\n        full_name\n        email\n        contact_detail{\n          phone\n        }\n        home_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n      }\n      opportunity\n      {\n        id\n        programme\n        { short_name_display }\n        host_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n        remote_opportunity\n        project_fee\n        earliest_start_date\n        latest_end_date\n        specifics_info{\n          salary\n          salary_currency{\n            alphabetic_code\n          }\n        }\n        opportunity_duration_type{\n          duration_type\n          salary\n        }\n        \n      }\n      slot{\n        start_date\n        end_date\n      }\n      status\n      updated_at\n      date_approved\n      date_realized\n      experience_end_date\n    standards{\n        \n standard_option{\n          \n          meta\n        }\n \n   } }\n  }\n}"
    
    ,
    variables: {}
  })
  var dataSet_OGX_REs = dataExtraction(graphql_OGX_REs);
  var rows_OGX_REs = dataManipulation(dataSet_OGX_REs);
  dataUpdating(sheet,rows_OGX_REs);

  var graphql_OGX_Remote_REs = JSON.stringify({
    query: "query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n      person_home_lc:1064\n      date_remote_realized:{from:\"01/07/2021\"}\n      \n\t\t}\n    \n    page:1\n    per_page:3000\n\t)\n\t{\n    paging\n    {\n      total_items\n    }\n\t\tdata\n    {\n      person\n      {\n        id\n        full_name\n        email\n        contact_detail{\n          phone\n        }\n        home_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n      }\n      opportunity\n      {\n        id\n        programme\n        { short_name_display }\n        host_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n        remote_opportunity\n        project_fee\n        earliest_start_date\n        latest_end_date\n        specifics_info{\n          salary\n          salary_currency{\n            alphabetic_code\n          }\n        }\n        opportunity_duration_type{\n          duration_type\n          salary\n        }\n        \n      }\n      slot{\n        start_date\n        end_date\n      }\n      status\n      updated_at\n      date_approved\n      date_realized\n      experience_end_date\n    standards{\n        \n standard_option{\n          \n          meta\n        }\n \n   } }\n  }\n}"
    
    ,
    variables: {}
  })
  var dataSet_OGX_Remote_REs = dataExtraction(graphql_OGX_Remote_REs);
  var rows_OGX_Remote_REs = dataManipulation(dataSet_OGX_Remote_REs);
  dataUpdating(sheet,rows_OGX_Remote_REs);


  var graphql_OGX_FIs = JSON.stringify({
    query: "query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n      person_home_lc:1064\n      experience_end_date:{from:\"01/07/2021\"}\n statuses:[\"finished\",\"completed\"]     \n\t\t}\n    \n    page:1\n    per_page:3000\n\t)\n\t{\n    paging\n    {\n      total_items\n    }\n\t\tdata\n    {\n      person\n      {\n        id\n        full_name\n        email\n        contact_detail{\n          phone\n        }\n        home_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n      }\n      opportunity\n      {\n        id\n        programme\n        { short_name_display }\n        host_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n        remote_opportunity\n        project_fee\n        earliest_start_date\n        latest_end_date\n        specifics_info{\n          salary\n          salary_currency{\n            alphabetic_code\n          }\n        }\n        opportunity_duration_type{\n          duration_type\n          salary\n        }\n        \n      }\n      slot{\n        start_date\n        end_date\n      }\n      status\n      updated_at\n      date_approved\n      date_realized\n      experience_end_date\n standards{\n        \n        standard_option{\n          \n          meta\n        }\n \n   }\n  }\n}}"
    
    ,
    variables: {}
  })
  var dataSet_OGX_FIs = dataExtraction(graphql_OGX_FIs);
  var rows_OGX_FIs = dataManipulation(dataSet_OGX_FIs);
  dataUpdating(sheet,rows_OGX_FIs);


}

