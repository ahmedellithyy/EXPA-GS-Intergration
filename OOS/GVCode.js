function updateOpportunityData_GV(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GV Opportunities")
  var oppData = sheet.getRange(4,1,sheet.getLastRow()-3,sheet.getLastColumn()).getValues()
  var lr = getLastRowForSheet(oppData)
  for(let i = 0; i < lr; i++){
    try{
      var response = UrlFetchApp.fetch(`https://gis-api.aiesec.org/v2/opportunities/${oppData[i][14]}?access_token=2f4c006eecfe5fb616839f27a4731c19cd7faecc11e1c97afd0ccc86491e207d`);  
      var recievedDate = JSON.parse(response.getContentText());
      sheet.getRange(i+4,1).setValue(recievedDate.status)
      if(recievedDate.home_lc.parent_id == 1609){
      if(recievedDate.opened_by != null){
        sheet.getRange(i+4,2).setValue(recievedDate.opened_by.full_name)
      }
      else{
        sheet.getRange(i+4,2).setValue("-")
      }
      if(recievedDate.status == "open"){
        if((oppData[i][38]=="Closed" && oppData[i][4] == "Audited & Not Passed") || oppData[i][3] != "Passed"){
          var options = {
            "method":"DELETE"
          }
          UrlFetchApp.fetch(`https://gis-api.aiesec.org/v2/opportunities/${oppData[i][14]}?access_token=2f4c006eecfe5fb616839f27a4731c19cd7faecc11e1c97afd0ccc86491e207d`,options)
          sheet.getRange(i+4,1).setValue("removed")
          let message = `Dear ${oppData[i][9]},\nYour Opportunity that has this ID ${oppData[i][14]} got closed becuase it's not approved by ECB or MCVP as it doesn't meet opportunity guidelines.`
          GmailApp.sendEmail(oppData[i][8],`Your Opportunity Got Closed - Opportunity ID: ${oppData[i][14]}`,message,
          { body:message,
            cc: "ahmed.ellithy4@aiesec.net,adel.mohamed@aiesec.net,c.sidhom@aiesec.org.eg"}

          )
          sheet.getRange(i+4,3).setValue(recievedDate.updated_at.toString().substring(0,10))     
        }
      }
      else if(recievedDate.status == "removed"){
        sheet.getRange(i+4,3).setValue(recievedDate.updated_at.toString().substring(0,10))
      }
      }
      else{
        sheet.getRange(i+4,1).setValue("Wrong Opp ID")
      }
    }
    catch(err){ 
      try{
          if(oppData[i][40] != true && err["message"].toString().includes("Couldn't find Opportunity") == true){
            let message = `Dear ${oppData[i][9]},\nYou have entered a wrong opportunity ID in Opportunities Openning System Form which its role title is ${oppData[i][11]}. PLEASE UPDATE it with the right one!`
              GmailApp.sendEmail(oppData[i][8],"Wrong Submission In Opportunities Openning System Form",message,
              {
                body:message,
                cc: "ahmed.ellithy4@aiesec.net,adel.mohamed@aiesec.net,c.sidhom@aiesec.org.eg"
                }
              )
              sheet.getRange(i+4,41).setValue(true)
              sheet.getRange(i+4,1).setValue("Wrong Opportunity ID")
          }
          }
          catch(err){
              sheet.getRange(i+4,1).setValue("Wrong Opportunity ID & Wrong Email or No Email")
          }
            
    }
    
  }
}

function checkOpportunities_GV(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Opportunities opened without submiting - GV")
  var sheetGT = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GV Opportunities")

  var response = UrlFetchApp.fetch("https://gis-api.aiesec.org/v2/opportunities?access_token=2f4c006eecfe5fb616839f27a4731c19cd7faecc11e1c97afd0ccc86491e207d&only=data&per_page=1000&filters[status]=open&filters[created]%5Bfrom%5D=01%2F12%2F2021")
  var data = JSON.parse(response.getContentText()).data;
  var newRows = []
  var options = { 
    "method":"DELETE"
  }
  for(let i = 0; i< data.length; i++){
    if(data[i].programmes.short_name == "GV"){
      var responseOpp = UrlFetchApp.fetch(`https://gis-api.aiesec.org/v2/opportunities/${data[i].id}?access_token=2f4c006eecfe5fb616839f27a4731c19cd7faecc11e1c97afd0ccc86491e207d`);  
      var recievedOppDate = JSON.parse(responseOpp.getContentText());
      
      var row = sheetGT.createTextFinder(data[i].id).matchEntireCell(true).findAll().map(x => x.getRow())
      if(row.length ==0){
        var today = new Date()
        newRows.push([
          data[i].id, data[i].title,  data[i].programmes.short_name,  data[i].office.full_name, "removed",today ,recievedOppDate.opened_by.full_name, 
          recievedOppDate.managers[0] != null? recievedOppDate.managers[0].full_name:""
        ])
        UrlFetchApp.fetch(`https://gis-api.aiesec.org/v2/opportunities/${data[i].id}?access_token=2f4c006eecfe5fb616839f27a4731c19cd7faecc11e1c97afd0ccc86491e207d`,options)
        GmailApp.sendEmail("c.safwat@aiesec.org.eg,lujain.shaltout@aiesec.net","GV Opportunity Got Closed",`The opportunity with this id (${data[i].id}) got closed because it's not submitted in The opportunity Opening System. So please, contact the LC to submit it.`)
      }
    }
  }
  if(newRows.length > 0){
    sheet.getRange(2,1,sheet.getLastRow(),sheet.getLastColumn()).clearContent()
    sheet.getRange(2,1,newRows.length,newRows[0].length).setValues(newRows)
  }
}

function sendEmailonSubmissions_GV(){
  GmailApp.sendEmail("c.sidhom@aiesec.org.eg","New Submission in Opportunity Opening System","Check the OOS for checking the new submission.") 
}

function fillFormulas_GV()
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GV Opportunities")
  var oppData = sheet.getRange(4,1,sheet.getLastRow()-3,sheet.getLastColumn()).getValues()
  var lr = getLastRowForSheet(oppData)
  
  for(let i = 0; i < lr; i++){
    sheet.getRange(i+4,4).setFormula(`=if(H${i+4}<>"",ifs(AND(Y${i+4}="Correct",Z${i+4}="Correct",AA${i+4}="Correct",AB${i+4}="Correct",AC${i+4}="Correct",AD${i+4}="Correct",AE${i+4}="Correct"),"Passed",AND(Y${i+4}="",Z${i+4}="",AA${i+4}="",AB${i+4}="",AC${i+4}="",AD${i+4}="",AE${i+4}=""),"Not Checked",OR(Y${i+4}="",Z${i+4}="",AA${i+4}="",AB${i+4}="",AC${i+4}="",AD${i+4}="",AE${i+4}=""),"Not Fully Checked",OR(Y${i+4}="Incorrect",Z${i+4}="Incorrect",AA${i+4}="Incorrect",AB${i+4}="Incorrect",AC${i+4}="Incorrect",AD${i+4}="Incorrect",AE${i+4}="Incorrect"),"Not Passed"),"")`)
    
    
    sheet.getRange(i+4,5).setFormula(`=if(H${i+4}<>"",ifs(AND(AH${i+4}="Correct",AK${i+4}="Correct",AJ${i+4}="Correct",AK${i+4}="Correct",AL${i+4}="Correct"),"Audited & Passed",AND(AJ${i+4}="",AK${i+4}="",AL${i+4}="",AM${i+4}="",AN${i+4}=""),"Not Audited",OR(AH${i+4}="",AI${i+4}="",AJ${i+4}="",AK${i+4}="",AL${i+4}=""),"Not Fully Audited",OR(AH${i+4}="Incorrect",AI${i+4}="Incorrect",AJ${i+4}="Incorrect",AK${i+4}="",AL${i+4}="Incorrect"),"Audited & Not Passed"),"")`)
    
    sheet.getRange(i+4,6).setFormula(`=if(H${i+4}<>"",year(H${i+4})&"-"&if(AND(month(H${i+4})<>10,month(H${i+4})<>11,month(H${i+4})<>12),"0"&month(H${i+4}),month(H${i+4})),"")`)
  }
}
