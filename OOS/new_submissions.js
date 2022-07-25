function sendEmailonSubmissions_GV(){
  GmailApp.sendEmail(`${emails_for_the_GV_submissions}`,"New Submission in OOS","Check the OOS for checking the new submission.") 
}


function sendEmailonSubmissions_GT(){
  GmailApp.sendEmail(`${emails_for_the_GT_submissions}`,"New Submission in Opportunity Opening System","Check the OOS for checking the new submission.") 
}
