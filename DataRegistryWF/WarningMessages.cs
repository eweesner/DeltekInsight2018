using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Deltek.Framework.API.Server;
using Deltek.Vision.WorkflowAPI.Server;
using System.Data;

namespace DataRegistryWF
{
    /// <summary>
    /// Class to manage the pop-up warning messages for the Data Registry
    /// </summary>
    public class WarningMessages : WorkflowBaseClass
    {
        public void checkForErrors(string udic_uid, string projectLookup_Old, string question2_Old, string question3_Old, string question4_Old, string question5a_Old, string question5b_Old)
        {
            //get all of the questions that fields to be checked and 
            //only create variables for the questions that need to be checked
            string sql = $@"select * from UDIC_PIIDataRegistry where UDIC_UID = '{udic_uid}'";
            DataTable dt = QueryData(sql);
            if (dt.Rows.Count == 1)
            {
                string projectErrorMessage = "";
                string warningMessages = "";
                string question2, question3, question4, question5a, question5b, question10, question13, question14, question15, question18, projectLookup, fieldsAutoPopulated;
                question2 = dt.Rows[0]["CustWillProjectUseData"].ToString();
                question3 = dt.Rows[0]["CustDoesthedatacontaininformationaboutpeople"].ToString();
                question4 = dt.Rows[0]["CustDoesthedatacontaininformationaboutanindividualpersonorpersons"].ToString();
                question5a = dt.Rows[0]["CustDoesDataIdentifyOrSufficientToIdentifyPersons"].ToString();
                question5b = dt.Rows[0]["CustDoesDataIdentifyOrSufficientToIdentifyPersonsbyVendor"].ToString();
                question10 = dt.Rows[0]["CustTypeofPersonalData"].ToString();
                question13 = dt.Rows[0]["CustWillRTIHSStaffAccessOnlyAAData"].ToString();
                question14 = dt.Rows[0]["CustWillDataOtherThanAABeSentOrAccessByUSA"].ToString();
                question15 = dt.Rows[0]["CustAuthorizationthatpermitsexporttoUSA"].ToString();
                question18 = dt.Rows[0]["CustIsanynonanonymizednonaggregatedDataheldbyanRTIHSVendor"].ToString();
                projectLookup = dt.Rows[0]["CustProjectLookup"].ToString();
                fieldsAutoPopulated = dt.Rows[0]["CustFieldsAutoPopulated"].ToString();

                //The project number can't be changed.
                if (projectLookup_Old != projectLookup)
                    projectErrorMessage = "ERROR:  The project number can't be changed. Please delete the form and start over.";
                //Form clearing
                if ((!question2_Old.Equals(question2) || !question3_Old.Equals(question3) || !question4_Old.Equals(question4) || !question5a_Old.Equals(question5a) || !question5b_Old.Equals(question5b))
                        && fieldsAutoPopulated.Equals("Y"))
                {
                    warningMessages += (messageEmpty(warningMessages) + @"** Clearing Form: Based on your new answers for Questions 2, 3, 4, 5a, or 5b the answers below will update.");
                }                
                //q2 is No and q3 is yes
                if (question2.Equals("No") && question3.Equals("Yes"))
                {
                    warningMessages += (messageEmpty(warningMessages) + @"** If there is data on people, then Q2 (""Will Project Use Data?"") should be marked ""Yes"".");
                }
                //q2=Y, q3=Y, q4=Y, 5a=N, 5b=y or n, 
                if (question2.Equals("Yes") && question3.Equals("Yes") && question4.Equals("Yes") && question5a.Equals("No") && (question5b.Equals("No") || question5b.Equals("Yes"))
                        && !string.IsNullOrEmpty(question10) && !string.IsNullOrEmpty(question13) && !string.IsNullOrEmpty(question14))
                {
                    if (!question10.Equals("Aggregated") && !question10.Equals("Anonymized"))
                    {
                        //warningMessages += (messageEmpty(warningMessages) + @"** The answers to 5a and 10 appear inconsistent. Please review the help box for 5a and 10 and determine if your answer should change for 5a, 10, or both. Update the answer(s), click Save.");
                    }
                    //q10 is either anonymized or aggregated, q13 should always = Yes and q14 = No
                    else
                    {
                        if (!question13.Equals("Yes") && !question14.Equals("No"))
                        {
                            warningMessages += (messageEmpty(warningMessages) + @"** Your answers to Q10, Q13, and Q14 are not consistent.  Please review the help box definitions and determine if your answer should change for Q10, Q13, Q14, or all.  Update the answer(s), click Save.");
                        }
                        else if (question13.Equals("Yes") && !question14.Equals("No"))
                        {
                            warningMessages += (messageEmpty(warningMessages) + @"** Your answers to Q10 and Q14 are not consistent.  Please review the help box definitions and determine if your answer should change for Q10, Q14, or both.  Update the answer(s), click Save.");
                        }
                        else if (!question13.Equals("Yes") && question14.Equals("No"))
                        {
                            warningMessages += (messageEmpty(warningMessages) + @"** Your answers to Q10 and Q13 are not consistent.  Please review the help box definitions and determine if your answer should change for Q10, Q13, or both.  Update the answer(s), click Save.");
                        }
                    }
                }
                else
                {
                    //q13 and q14 should be oppposite answers
                    if ((question13.Equals("No") || question13.Equals("Yes")) && question13.Equals(question14))
                    {
                        warningMessages += (messageEmpty(warningMessages) + @"** The answers to 13 and 14 are not consistent. Please review the help box and determine if your answer should change for 13, 14, or both. Update the answer(s), click Save.");
                    }
                }

                //q5b and q18 should be opposite answwers
                if ((question5b.Equals("No") || question5b.Equals("Yes")) && question5b.Equals(question18)) 
                {
                    warningMessages += (messageEmpty(warningMessages) + @"** The answers to 5b and 18 are not consistent. Please review the help box and determine if your answer should change for 5b, 18, or both. Update the answer(s), click Save.");
                }

                //creating the warning popup box
                if (!string.IsNullOrEmpty(warningMessages) || !string.IsNullOrEmpty(projectErrorMessage))
                {
                    if(!string.IsNullOrEmpty(warningMessages))
                        warningMessages = "-----Warning Messages-----\n\n" + warningMessages;
                    AddWarning(projectErrorMessage + messageEmpty(projectErrorMessage) + warningMessages);
                }                
            }
            else
            {
                AddFatal("Please contact the Vision System Administrator");
            }  
     
        }
        private string messageEmpty(string s)
        {
            return string.IsNullOrEmpty(s) ? "" : "\n\n";
        }
    }
}
