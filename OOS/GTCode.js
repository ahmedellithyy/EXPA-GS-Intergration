/*------------------------------------------------------------------------ Main Function ------------------------------------------------------------------------*/
function updateOpportunityData(){

  // Get the sheet of the IGT opportunities
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`${gtSheet}`)
  var oppData = sheet.getRange(gtHeaderRow+1,1,sheet.getLastRow()-gtHeaderRow,sheet.getLastColumn()).getValues()
  var lr = getLastRowForSheet(oppData)

  for(let i =0; i < lr; i++){
    try{
      var response = UrlFetchApp.fetch(`https://gis-api.aiesec.org/v2/opportunities/${oppData[i][opportunity_id_column_GT-1]}?access_token=${access_token}`);  
      var recievedDate = JSON.parse(response.getContentText());
      sheet.getRange(i+gvHeaderRow+1,expa_status_column_GT).setValue(recievedDate.status)

      if(recievedDate.home_lc.parent_id == egypt_expa_id){
    Logger.log(i)
        if(recievedDate.opened_by != null){
          sheet.getRange(i+gvHeaderRow+1,opened_by_name_column_GT).setValue(recievedDate.opened_by.full_name)
        }
        else{
          sheet.getRange(i+gvHeaderRow+1,opened_by_name_column_GT).setValue("-")
        }
        if(recievedDate.status == "open"){

            sheet.getRange(i+gvHeaderRow+1,updated_at_column_GT).setValue(recievedDate.updated_at.toString().substring(0,10))     

          if((oppData[i][closed_by_ecb_column_GT-1]=="Closed" && oppData[i][audited_by_ecb_column_GT-1] == "Audited & Not Passed") || oppData[i][audited_by_mc_column_GT-1] != "Passed"){
            
            UrlFetchApp.fetch(`https://gis-api.aiesec.org/v2/opportunities/${oppData[i][opportunity_id_column_GT-1]}?access_token=${access_token}`,options_delete)
            sheet.getRange(i+gvHeaderRow+1,expa_status_column_GT).setValue("removed as it didn't pass")
            let message = `Dear ${oppData[i][full_name_column_GT-1]},\nYour Opportunity that has this ID ${oppData[i][opportunity_id_column_GT-1]} got closed becuase it's not approved by ECB or MCVP as it doesn't meet opportunity guidelines.`

            GmailApp.sendEmail(oppData[i][email_column_GT-1],`Your Opportunity Got Closed - Opportunity ID: ${oppData[i][opportunity_id_column_GT-1]}`,message,
            { body:message,
              cc: emails_for_the_GT_submissions
              }
            )
            sheet.getRange(i+gvHeaderRow+1,updated_at_column_GT).setValue(recievedDate.updated_at.toString().substring(0,10))     
          }
        }
        else if(recievedDate.status == "removed"){
          sheet.getRange(i+gvHeaderRow+1,updated_at_column_GT).setValue(recievedDate.updated_at.toString().substring(0,10))
        }
        
      }
      else{
        sheet.getRange(i+gvHeaderRow+1,expa_status_column_GT).setValue("Wrong Opportunity ID")
        if(oppData[i][wrong_opp_id_colum_GT-1] != true){
          let message = `Dear ${oppData[i][full_name_column_GT-1]},\nYou have entered a wrong opportunity ID in Opportunities Openning System Form which its role title is ${oppData[i][opportunity_role_title_column_GT-1]}. PLEASE UPDATE it with the right one (don't do re-submission to avoid having duplicate submissions on the system which will cause your opportunity to get automatically closed even if one of the submissions is correct)!`
          GmailApp.sendEmail(oppData[i][email_column_GT-1],`Wrong Submission In Opportunities Openning System Form - ID ${oppData[i][opportunity_id_column_GT-1]}`,message,
            {
              htmlBody:message,
              cc: emails_for_the_GT_submissions
            }
          )
          sheet.getRange(i+gvHeaderRow+1,wrong_opp_id_colum_GT).setValue(true)
        }
      }
    }
    catch(err){ 
      try{
          if(oppData[i][wrong_opp_id_colum_GT-1] != true && err["message"].toString().includes("Couldn't find Opportunity") == true){
            let message = `Dear ${oppData[i][full_name_column_GT-1]},\nYou have entered a wrong opportunity ID in Opportunities Openning System Form which its role title is ${oppData[i][opportunity_role_title_column_GT-1]}. PLEASE UPDATE it with the right one (don't do re-submission to avoid having duplicate submissions on the system which will cause your opportunity to get automatically closed even if one of the submissions is correct)!`
              GmailApp.sendEmail(oppData[i][email_column_GT-1],`Wrong Submission In Opportunities Openning System Form - ID ${oppData[i][opportunity_id_column_GT-1]}`,message,
              {
                htmlBody:message,
                cc: emails_for_the_GT_submissions
                }
              )
              sheet.getRange(i+gvHeaderRow+1,wrong_opp_id_colum_GT).setValue(true)
              sheet.getRange(i+gvHeaderRow+1,expa_status_column_GT).setValue("Wrong Opportunity ID")
          }
          else{
            let message = `Check OOS,\n ${e.toString()}`
              GmailApp.sendEmail(mcvpIM,"ERROR IN OOS",message,
              {
                body:message,
                }
              )
          }
          }
          catch(err){
              sheet.getRange(i+gvHeaderRow+1,expa_status_column_GT).setValue("Wrong Opportunity ID & Wrong Email or No Email")
          }
            
    }
    
  }
}

function checkOpportunities(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`${gtSheet_not_submitted}`)
  var sheetGT = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`${gtSheet}`)

  var response = UrlFetchApp.fetch(`https://gis-api.aiesec.org/v2/opportunities?access_token=${access_token}&only=data&per_page=1000&filters[status]=open&filters[created]%5Bfrom%5D=01%2F12%2F2021`)
  var data = JSON.parse(response.getContentText()).data;
  var newRows = []
  
  for(let i = 0; i< data.length; i++){
    if(data[i].programmes.short_name == "GTa" || data[i].programmes.short_name == "GTe"){
      var row = sheetGT.createTextFinder(data[i].id).matchEntireCell(true).findAll().map(x => x.getRow())
      if(row.length ==0){
        var responseOpp = UrlFetchApp.fetch(`https://gis-api.aiesec.org/v2/opportunities/${data[i].id}?access_token=${access_token}`);  
        var recievedOppDate = JSON.parse(responseOpp.getContentText());
        var today = new Date()
        newRows.push([
          data[i].id, data[i].title,  data[i].programmes.short_name,  data[i].office.full_name, "removed",today ,recievedOppDate.opened_by.full_name, 
          recievedOppDate.managers[0] != null? recievedOppDate.managers[0].full_name:""
        ])
        UrlFetchApp.fetch(`https://gis-api.aiesec.org/v2/opportunities/${data[i].id}?access_token=${access_token}`,options_delete)
        GmailApp.sendEmail(`${emails_for_the_GT_submissions}`,"GTa/e Opportunity Got Closed",`The opportunity with this id (${data[i].id}) got closed because it's not submitted in The opportunity Opening System. So please, contact the LC to submit it.`)
      }
    }
  }
  if(newRows.length > 0){
    sheet.getRange(gtvSheet_not_submitted_header_row+1,1,sheet.getLastRow(),sheet.getLastColumn()).clearContent()
    sheet.getRange(gtvSheet_not_submitted_header_row+1,1,newRows.length,newRows[0].length).setValues(newRows)
  }
}


