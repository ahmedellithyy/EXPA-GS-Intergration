function reportReleasing(){
  const feedbackSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Common Mistake Reports")
  const complianceMatrixSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Compliance Matrix")
  const ogxSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("OGX Submissions Dashboard")
  const icxSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ICX Submissions Dashboard")
  var folder = DriveApp.getFolderById(driveFolderId)
  var ebTeam = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("References").getRange(35,4,SpreadsheetApp.getActiveSpreadsheet().getSheetByName("References").getLastRow()-34,1).getDisplayValues()
  ebTeam.forEach((item,index)=>{
        ebTeam[index][0] = item[0].replace(" ","")              
  })

  var date = new Date()
  date.setMonth(date.getMonth()-2)
  date = Utilities.formatDate(date,"GMT+2","yyyy-MM")

  var colIndexInComplianceMatrixSheet = complianceMatrixSheet.createTextFinder(date).matchEntireCell(true).findAll().map(x => x.getColumn())
  var rowIndexIn_ogxSheet = ogxSheet.createTextFinder(date).matchEntireCell(true).findAll().map(x => x.getRow())
  var rowIndexIn_icxSheet = icxSheet.createTextFinder(date).matchEntireCell(true).findAll().map(x => x.getRow())

  var presentation = SlidesApp.openById("1Cjt_J9IEnDZ-0pfhvnAezFeiTpQgOTJEdphYX5NOnF0")
  var name = presentation.getName()
  var slides = presentation.getSlides()
  //var slide = SlidesApp.openById("1Cjt_J9IEnDZ-0pfhvnAezFeiTpQgOTJEdphYX5NOnF0").getSlideById("g1185394d3d3_0_0")

  var month = new Date()
  name = name.replace("{{LC_Name}}",`${lcName}`).replace("{{Month}}",`${monthNames[month.getMonth()-1]}`)
  var newPresentation = SlidesApp.create(name)
  for(let i = 0; i <  slides.length;i++){
    newPresentation.appendSlide(slides[i])
  }
  newPresentation.getSlides()[0].remove()
  var file = DriveApp.getFileById(newPresentation.getId())
  file.moveTo(folder)
  
  var slides = newPresentation.getSlides()
  
  // Cover slide
  slides[0].replaceAllText("{{MONTH}}",`${monthNames[month.getMonth()-1]}`)
  slides[0].replaceAllText("{{YEAR}}",`${month.getFullYear()}`)

  // OGX Analysis
  var sl1 = complianceMatrixSheet.getRange(indices["OGX_Severity_Row"],colIndexInComplianceMatrixSheet[0]).getValue()=="-"? 0:complianceMatrixSheet.getRange(indices["OGX_Severity_Row"],colIndexInComplianceMatrixSheet[0]).getValue();
  var sl2 = complianceMatrixSheet.getRange(indices["OGX_Severity_Row"],colIndexInComplianceMatrixSheet[0]+1).getValue()=="-"?0:complianceMatrixSheet.getRange(indices["OGX_Severity_Row"],colIndexInComplianceMatrixSheet[0]+1).getValue()
  var sl3 = complianceMatrixSheet.getRange(indices["OGX_Severity_Row"],colIndexInComplianceMatrixSheet[0]+2).getValue()=="-"?0:complianceMatrixSheet.getRange(indices["OGX_Severity_Row"],colIndexInComplianceMatrixSheet[0]+2).getValue()
  var sl4 = complianceMatrixSheet.getRange(indices["OGX_Severity_Row"],colIndexInComplianceMatrixSheet[0]+3).getValue()=="-"?0:complianceMatrixSheet.getRange(indices["OGX_Severity_Row"],colIndexInComplianceMatrixSheet[0]+3).getValue()
  var sl5 = complianceMatrixSheet.getRange(indices["OGX_Severity_Row"],colIndexInComplianceMatrixSheet[0]+4).getValue()=="-"?0:complianceMatrixSheet.getRange(indices["OGX_Severity_Row"],colIndexInComplianceMatrixSheet[0]+4).getValue()
  var fines = complianceMatrixSheet.getRange(indices["OGX_Fine_Row"],colIndexInComplianceMatrixSheet[0]+3).getValue()
  var ogvPass = ogxSheet.getRange(rowIndexIn_ogxSheet[0],indices["GV_Pass_Col"]).getValue()=="-"?0:ogxSheet.getRange(rowIndexIn_ogxSheet[0],indices["GV_Pass_Col"]).getValue()*100
  var ogtePass = ogxSheet.getRange(rowIndexIn_ogxSheet[1],indices["GTe_Pass_Col"]).getValue()=="-"?0:ogxSheet.getRange(rowIndexIn_ogxSheet[1],indices["GTe_Pass_Col"]).getValue()*100
  var ogtaPass = ogxSheet.getRange(rowIndexIn_ogxSheet[2],indices["GTa_Pass_Col"]).getValue()=="-"?0:ogxSheet.getRange(rowIndexIn_ogxSheet[2],indices["GTa_Pass_Col"]).getValue()*100
  var passed
  sl1<6 && sl2<6 && sl3<6 && sl4==0 && sl5==0 && (ogvPass+ogtePass+ogtaPass)/3<=70?passed= "PASSED": passed = "FAILED" 
  
  slides[2].replaceAllText("{{SL1}}",`${sl1}`)
  slides[2].replaceAllText("{{SL2}}",`${sl2}`)
  slides[2].replaceAllText("{{SL3}}",`${sl3}`)
  slides[2].replaceAllText("{{SL4}}",`${sl4}`)
  slides[2].replaceAllText("{{SL5}}",`${sl5}`)
  slides[2].replaceAllText("{{NUM_FINES}}",`${fines}`)
  slides[2].replaceAllText("{{OGV_PASS}}",ogvPass.toFixed(2)+"%")
  slides[2].replaceAllText("{{OGTA_PASS}}",ogtaPass.toFixed(2)+"%")
  slides[2].replaceAllText("{{OGTE_PASS}}",ogtePass.toFixed(2)+"%")
  slides[2].replaceAllText("{{PASSED?}}",passed)

  // ICX Analysis
  var sl1 = complianceMatrixSheet.getRange(indices["ICX_Severity_Row"],colIndexInComplianceMatrixSheet[0]).getValue()=="-"? 0:complianceMatrixSheet.getRange(indices["ICX_Severity_Row"],colIndexInComplianceMatrixSheet[0]).getValue();
  var sl2 = complianceMatrixSheet.getRange(indices["ICX_Severity_Row"],colIndexInComplianceMatrixSheet[0]+1).getValue()=="-"?0:complianceMatrixSheet.getRange(indices["ICX_Severity_Row"],colIndexInComplianceMatrixSheet[0]+1).getValue()
  var sl3 = complianceMatrixSheet.getRange(indices["ICX_Severity_Row"],colIndexInComplianceMatrixSheet[0]+2).getValue()=="-"?0:complianceMatrixSheet.getRange(indices["ICX_Severity_Row"],colIndexInComplianceMatrixSheet[0]+2).getValue()
  var sl4 = complianceMatrixSheet.getRange(indices["ICX_Severity_Row"],colIndexInComplianceMatrixSheet[0]+3).getValue()=="-"?0:complianceMatrixSheet.getRange(indices["ICX_Severity_Row"],colIndexInComplianceMatrixSheet[0]+3).getValue()
  var sl5 = complianceMatrixSheet.getRange(indices["ICX_Severity_Row"],colIndexInComplianceMatrixSheet[0]+4).getValue()=="-"?0:complianceMatrixSheet.getRange(indices["ICX_Severity_Row"],colIndexInComplianceMatrixSheet[0]+4).getValue()
  var fines = complianceMatrixSheet.getRange(indices["ICX_Fine_Row"],colIndexInComplianceMatrixSheet[0]+3).getValue()
  var igvPass = icxSheet.getRange(rowIndexIn_icxSheet[0],indices["GV_Pass_Col"]).getValue()=="-"?0:icxSheet.getRange(rowIndexIn_icxSheet[0],indices["GV_Pass_Col"]).getValue()*100
  var igtePass = icxSheet.getRange(rowIndexIn_icxSheet[1],indices["GTe_Pass_Col"]).getValue()=="-"?0:icxSheet.getRange(rowIndexIn_icxSheet[1],indices["GTe_Pass_Col"]).getValue()*100
  var igtaPass = icxSheet.getRange(rowIndexIn_icxSheet[2],indices["GTa_Pass_Col"]).getValue()=="-"?0:icxSheet.getRange(rowIndexIn_icxSheet[2],indices["GTa_Pass_Col"]).getValue()*100
  var passed
  sl1<6 && sl2<6 && sl3<6 && sl4==0 && sl5==0 && (ogvPass+ogtePass+ogtaPass)/3<=70?passed= "PASSED": passed = "FAILED" 
  
  slides[3].replaceAllText("{{SL1}}",`${sl1}`)
  slides[3].replaceAllText("{{SL2}}",`${sl2}`)
  slides[3].replaceAllText("{{SL3}}",`${sl3}`)
  slides[3].replaceAllText("{{SL4}}",`${sl4}`)
  slides[3].replaceAllText("{{SL5}}",`${sl5}`)
  slides[3].replaceAllText("{{NUM_FINES}}",`${fines}`)
  slides[3].replaceAllText("{{IGV_PASS}}",igvPass.toFixed(2)+"%")
  slides[3].replaceAllText("{{IGTA_PASS}}",igtePass.toFixed(2)+"%")
  slides[3].replaceAllText("{{IGTE_PASS}}",igtaPass.toFixed(2)+"%")
  slides[3].replaceAllText("{{PASSED?}}",passed)
  

  // OGX Common Mistakes Filling
  slides[4].replaceAllText("{{APD_Mistakes}}",feedbackSheet.getRange(commonMistakesRow["OGXAPDs"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]).getValue())
  slides[4].replaceAllText("{{APD_Advice}}",feedbackSheet.getRange(commonMistakesRow["OGXAPDs"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]+1).getValue())

  slides[5].replaceAllText("{{RE_Mistakes}}",feedbackSheet.getRange(commonMistakesRow["OGXREs"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]).getValue())
  slides[5].replaceAllText("{{RE_Advice}}",feedbackSheet.getRange(commonMistakesRow["OGXREs"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]+1).getValue())

  slides[6].replaceAllText("{{GV_FI_Mistakes}}",feedbackSheet.getRange(commonMistakesRow["OGVFIs"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]).getValue())
  slides[6].replaceAllText("{{GV_FI_Advice}}",feedbackSheet.getRange(commonMistakesRow["OGVFIs"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]+1).getValue())
  slides[6].replaceAllText("{{GT_Monthly_Mistakes}}",feedbackSheet.getRange(commonMistakesRow["OGTMonthly"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]).getValue())
  slides[6].replaceAllText("{{GT_Monthly_Advice}}",feedbackSheet.getRange(commonMistakesRow["OGTMonthly"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]+1).getValue())
  
  slides[7].replaceAllText("{{GT_FI_Mistakes}}",feedbackSheet.getRange(commonMistakesRow["OGTFIs"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]).getValue())
  slides[7].replaceAllText("{{GT_FI_Advice}}",feedbackSheet.getRange(commonMistakesRow["OGTFIs"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]+1).getValue())
  slides[7].replaceAllText("{{OGX_Mistakes}}",feedbackSheet.getRange(commonMistakesRow["OGXStandards"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]).getValue())
  slides[7].replaceAllText("{{OGX_Advice}}",feedbackSheet.getRange(commonMistakesRow["OGXStandards"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]+1).getValue())
  slides[7].replaceAllText("{{APD_TO_BE_BROKEN}}",feedbackSheet.getRange(commonMistakesRow["OGXAPDsToBeBroken"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]).getValue())
  slides[7].replaceAllText("{{APD_TO_BE_BROKEN_Advice}}",feedbackSheet.getRange(commonMistakesRow["OGXAPDsToBeBroken"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]+1).getValue())

  slides[8].replaceAllText("{{OPS_Mistake}}",feedbackSheet.getRange(commonMistakesRow["OPS"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]).getValue())
  slides[8].replaceAllText("{{OPS_Advice}}",feedbackSheet.getRange(commonMistakesRow["OPS"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]+1).getValue())


  // ICX Common Mistakes Filling
  slides[9].replaceAllText("{{APD_Mistakes}}",feedbackSheet.getRange(commonMistakesRow["ICXAPDs"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]).getValue())
  slides[9].replaceAllText("{{APD_Advice}}",feedbackSheet.getRange(commonMistakesRow["ICXAPDs"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]+1).getValue())

  slides[10].replaceAllText("{{RE_Mistakes}}",feedbackSheet.getRange(commonMistakesRow["ICXREs"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]).getValue())
  slides[10].replaceAllText("{{RE_Advice}}",feedbackSheet.getRange(commonMistakesRow["ICXREs"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]+1).getValue())

  slides[11].replaceAllText("{{GV_FI_Mistakes}}",feedbackSheet.getRange(commonMistakesRow["IGVFIs"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]).getValue())
  slides[11].replaceAllText("{{GV_FI_Advice}}",feedbackSheet.getRange(commonMistakesRow["IGVFIs"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]+1).getValue())
  slides[11].replaceAllText("{{GT_Monthly_Mistakes}}",feedbackSheet.getRange(commonMistakesRow["IGTMonthly"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]).getValue())
  slides[11].replaceAllText("{{GT_Monthly_Advice}}",feedbackSheet.getRange(commonMistakesRow["IGTMonthly"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]+1).getValue())
  
  slides[12].replaceAllText("{{GT_FI_Mistakes}}",feedbackSheet.getRange(commonMistakesRow["IGTFIs"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]).getValue())
  slides[12].replaceAllText("{{GT_FI_Advice}}",feedbackSheet.getRange(commonMistakesRow["IGTFIs"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]+1).getValue())
  slides[12].replaceAllText("{{ICX_Mistakes}}",feedbackSheet.getRange(commonMistakesRow["ICXStandards"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]).getValue())
  slides[12].replaceAllText("{{ICX_Advice}}",feedbackSheet.getRange(commonMistakesRow["ICXStandards"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]+1).getValue())
  slides[12].replaceAllText("{{APD_TO_BE_BROKEN}}",feedbackSheet.getRange(commonMistakesRow["ICXAPDsToBeBroken"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]).getValue())
  slides[12].replaceAllText("{{APD_TO_BE_BROKEN_Advice}}",feedbackSheet.getRange(commonMistakesRow["ICXAPDsToBeBroken"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]+1).getValue())

  slides[13].replaceAllText("{{IPS_Mistake}}",feedbackSheet.getRange(commonMistakesRow["IPS"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]).getValue())
  slides[13].replaceAllText("{{IPS_Advice}}",feedbackSheet.getRange(commonMistakesRow["IPS"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]+1).getValue())
  slides[13].replaceAllText("{{ACC_Mistakes}}",feedbackSheet.getRange(commonMistakesRow["Accommodation"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]).getValue())
  slides[13].replaceAllText("{{ACC_Advice}}",feedbackSheet.getRange(commonMistakesRow["Accommodation"],commonMistakesCol[`${monthNames[month.getMonth()-1]}${month.getFullYear()}`]+1).getValue())

  newPresentation.addViewer("ecb.chair@aiesec.net.eg")
  MailApp.sendEmail("ecb.chair@aiesec.net.eg",`${name}`,`This is the link of the final Report ${newPresentation.getUrl()}`)
}
