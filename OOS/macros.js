function fillFormulas_GT(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`${gtSheet}`)
  var oppData = sheet.getRange(gtHeaderRow+1,1,sheet.getLastRow()-gtHeaderRow,sheet.getLastColumn()).getValues()
  var lr = getLastRowForSheet(oppData)
  
  for(let i = 0; i < lr; i++){
        // updating status for MCVP
    sheet.getRange(i+gtHeaderRow+1,mcvpStatusColumn_GT).setFormula(`=if(I${i+gtHeaderRow+1}<>"",ifs(AND(AA${i+gtHeaderRow+1}="Correct",AB${i+gtHeaderRow+1}="Correct",AC${i+gtHeaderRow+1}="Correct",AD${i+gtHeaderRow+1}="Correct",AE${i+gtHeaderRow+1}="Correct",AF${i+gtHeaderRow+1}="Correct",AG${i+gtHeaderRow+1}="Correct"),"Passed",AND(AA${i+gtHeaderRow+1}="",AB${i+gtHeaderRow+1}="",AC${i+gtHeaderRow+1}="",AD${i+gtHeaderRow+1}="",AE${i+gtHeaderRow+1}="",AF${i+gtHeaderRow+1}="",AG${i+gtHeaderRow+1}=""),"Not Checked",OR(AA${i+gtHeaderRow+1}="",AB${i+gtHeaderRow+1}="",AC${i+gtHeaderRow+1}="",AD${i+gtHeaderRow+1}="",AE${i+gtHeaderRow+1}="",AF${i+gtHeaderRow+1}="",AG${i+gtHeaderRow+1}=""),"Not Fully Checked",OR(AA${i+gtHeaderRow+1}="Incorrect",AB${i+gtHeaderRow+1}="Incorrect",AC${i+gtHeaderRow+1}="Incorrect",AD${i+gtHeaderRow+1}="Incorrect",AE${i+gtHeaderRow+1}="Incorrect",AF${i+gtHeaderRow+1}="Incorrect",AG${i+gtHeaderRow+1}="Incorrect"),"Not Passed"),"")`)
        // updating status for ECB
    sheet.getRange(i+gtHeaderRow+1,ecbStatusColumn_GT).setFormula(`=if(I${i+gtHeaderRow+1}<>"",ifs(AND(AK${i+gtHeaderRow+1}="Correct",AL${i+gtHeaderRow+1}="Correct",AM${i+gtHeaderRow+1}="Correct",AN${i+gtHeaderRow+1}="Correct",AO${i+gtHeaderRow+1}="Correct"),"Audited & Passed",AND(AK${i+gtHeaderRow+1}="",AL${i+gtHeaderRow+1}="",AM${i+gtHeaderRow+1}="",AN${i+gtHeaderRow+1}="",AO${i+gtHeaderRow+1}=""),"Not Audited",OR(AK${i+gtHeaderRow+1}="",AL${i+gtHeaderRow+1}="",AM${i+gtHeaderRow+1}="",AN${i+gtHeaderRow+1}="",AO${i+gtHeaderRow+1}=""),"Not Fully Audited",OR(AK${i+gtHeaderRow+1}="Incorrect",AL${i+gtHeaderRow+1}="Incorrect",AM${i+gtHeaderRow+1}="Incorrect",AN${i+gtHeaderRow+1}="",AO${i+gtHeaderRow+1}="Incorrect"),"Audited & Not Passed"),"")`)
    
        // updating submission month forumla
    sheet.getRange(i+gtHeaderRow+1,submissionMonthColumn_GT).setFormula(`=if(I${i+gtHeaderRow+1}<>"",year(I${i+gtHeaderRow+1})&"-"&if(AND(month(I${i+gtHeaderRow+1})<>10,month(I${i+gtHeaderRow+1})<>11,month(I${i+gtHeaderRow+1})<>12),"0"&month(I${i+gtHeaderRow+1}),month(I${i+gtHeaderRow+1})),"")`)

        // updating the total number of openings
    sheet.getRange(i+gvHeaderRow+1,totalOpeningsColumn_GT).setFormula(`=S${i+gvHeaderRow+1}*T${i+gvHeaderRow+1}`)

  }
}

function fillFormulas_GV()
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`${gvSheet}`)
  var oppData = sheet.getRange(gvHeaderRow+1,1,sheet.getLastRow()-gvHeaderRow,sheet.getLastColumn()).getValues()
  var lr = getLastRowForSheet(oppData)
  
  for(let i = 0; i < lr; i++){
    // updating status for MCVP
    sheet.getRange(i+gvHeaderRow+1,mcvpStatusColumn_GV).setFormula(
      `=if(H${i+gvHeaderRow+1}<>"",ifs(AND(Y${i+gvHeaderRow+1}="Correct",Z${i+gvHeaderRow+1}="Correct",AA${i+gvHeaderRow+1}="Correct",AB${i+gvHeaderRow+1}="Correct",AC${i+gvHeaderRow+1}="Correct",AD${i+gvHeaderRow+1}="Correct",AE${i+gvHeaderRow+1}="Correct"),"Passed",AND(Y${i+gvHeaderRow+1}="",Z${i+gvHeaderRow+1}="",AA${i+gvHeaderRow+1}="",AB${i+gvHeaderRow+1}="",AC${i+gvHeaderRow+1}="",AD${i+gvHeaderRow+1}="",AE${i+gvHeaderRow+1}=""),"Not Checked",OR(Y${i+gvHeaderRow+1}="",Z${i+gvHeaderRow+1}="",AA${i+gvHeaderRow+1}="",AB${i+gvHeaderRow+1}="",AC${i+gvHeaderRow+1}="",AD${i+gvHeaderRow+1}="",AE${i+gvHeaderRow+1}=""),"Not Fully Checked",OR(Y${i+gvHeaderRow+1}="Incorrect",Z${i+gvHeaderRow+1}="Incorrect",AA${i+gvHeaderRow+1}="Incorrect",AB${i+gvHeaderRow+1}="Incorrect",AC${i+gvHeaderRow+1}="Incorrect",AD${i+gvHeaderRow+1}="Incorrect",AE${i+gvHeaderRow+1}="Incorrect"),"Not Passed"),"")`
      )
    
    // updating status for ECB
    sheet.getRange(i+gvHeaderRow+1,ecbStatusColumn_GV).setFormula(`=if(H${i+gvHeaderRow+1}<>"",ifs(AND(AI${i+gvHeaderRow+1}="Correct",AJ${i+gvHeaderRow+1}="Correct",AK${i+gvHeaderRow+1}="Correct",AL${i+gvHeaderRow+1}="Correct",AM${i+gvHeaderRow+1}="Correct"),"Audited & Passed",AND(AI${i+gvHeaderRow+1}="",AJ${i+gvHeaderRow+1}="",AK${i+gvHeaderRow+1}="",AL${i+gvHeaderRow+1}="",AM${i+gvHeaderRow+1}=""),"Not Audited",OR(AI${i+gvHeaderRow+1}="",AJ${i+gvHeaderRow+1}="",AK${i+gvHeaderRow+1}="",AL${i+gvHeaderRow+1}="",AM${i+gvHeaderRow+1}=""),"Not Fully Audited",OR(AI${i+gvHeaderRow+1}="Incorrect",AJ${i+gvHeaderRow+1}="Incorrect",AK${i+gvHeaderRow+1}="Incorrect",AL${i+gvHeaderRow+1}="",AM${i+gvHeaderRow+1}="Incorrect"),"Audited & Not Passed"),"")`)
    
    // updating submission month forumla
    sheet.getRange(i+gvHeaderRow+1,submissionMonthColumn_GV).setFormula(`=if(I${i+gvHeaderRow+1}<>"",year(I${i+gvHeaderRow+1})&"-"&if(AND(month(I${i+gvHeaderRow+1})<>10,month(I${i+gvHeaderRow+1})<>11,month(I${i+gvHeaderRow+1})<>12),"0"&month(I${i+gvHeaderRow+1}),month(I${i+gvHeaderRow+1})),"")`)
    
    // updating the total number of openings
    sheet.getRange(i+gvHeaderRow+1,totalOpeningsColumn_GV).setFormula(`=S${i+gvHeaderRow+1}*T${i+gvHeaderRow+1}`)

  }
}



