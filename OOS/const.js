/*------------------------------------------------------------------------ General Constants ------------------------------------------------------------------------*/
const egypt_expa_id = 000000 // expa id
const access_token = "" // the access token
const mcvpIM = ""

const gvSheet = "iGV Opportunities"
const gvHeaderRow = 3
const gvSheet_not_submitted = "Opportunities opened without submiting - GV"
const gvSheet_not_submitted_header_row = 1


const gtSheet = "iGTa/e Opportunities"
const gtHeaderRow = 3
const gtSheet_not_submitted = "Opportunities opened without submiting - GTa/e"
const gtvSheet_not_submitted_header_row = 1


const options_delete = {
  method: 'DELETE',
  headers:{
    'Accept': 'application/json'
  }          
}


/*------------------------------------------------------------------------ Constants for GVCode ------------------------------------------------------------------------*/
const emails_for_the_GV_submissions = ""
const opportunity_id_column_GV = 8
const expa_status_column_GV = 1
const opened_by_name_column_GV = 2
const updated_at_column_GV = 3
const closed_by_ecb_column_GV = 40
const audited_by_mc_column_GV = 4
const audited_by_ecb_column_GV = 5
const full_name_column_GV = 11
const email_column_GV = 10
const wrong_opp_id_colum_GV = 42
const opportunity_role_title_column_GV = 13


/*------------------------------------------------------------------------ Constants for GTCode ------------------------------------------------------------------------*/
const emails_for_the_GT_submissions = ""
const opportunity_id_column_GT = 8
const expa_status_column_GT = 1
const opened_by_name_column_GT = 2
const updated_at_column_GT = 3
const closed_by_ecb_column_GT = 42
const audited_by_mc_column_GT = 4
const audited_by_ecb_column_GT = 5
const full_name_column_GT = 11
const email_column_GT = 10
const wrong_opp_id_colum_GT = 44
const opportunity_role_title_column_GT = 13



/*------------------------------------------------------------------------ Constants for macros ------------------------------------------------------------------------*/
// Write the number of the column for exmaple A -> 1, B -> 2. You can get the number of the column by writing =COLUMN() forumla in the desired column
const mcvpStatusColumn_GV = 4               
const ecbStatusColumn_GV = 5
const submissionMonthColumn_GV = 6
const totalOpeningsColumn_GV = 44


const mcvpStatusColumn_GT = 4               
const ecbStatusColumn_GT = 5
const submissionMonthColumn_GT = 6
const totalOpeningsColumn_GT = 46






