function deleteOpps(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Close Opps")
  var index = sheet.createTextFinder("Closed").matchEntireCell(true).findAll().map(x => x.getRow()-1)
  var index = index[index.length-1]
  var oppData = sheet.getRange(1,1,sheet.getLastRow(),2).getValues()
  var options = {
    "method":"DELETE"
  }
  for(let i = index; i< oppData.length;i++){
    Logger.log(i)
    try{
      UrlFetchApp.fetch(`https://gis-api.aiesec.org/v2/opportunities/${oppData[i][0]}?access_token=${access_token}`,options)
      sheet.getRange(i+1,3).setValue("Closed").setBackground("red")
    }
    catch(e){
      Logger.log(e.toString())
    }
  }
}
