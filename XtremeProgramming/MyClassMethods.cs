using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Deltek.Vision.WorkflowAPI.Server;
using System.Data;

namespace XtremeProgramming
{
    public class MyClassMethods : WorkflowBaseClass
    {
        /// <summary>
        /// Basic information message. This one doesn't have any parameters.
        /// </summary>
        public void myInformationMessage()
        {
            AddInformation("The light is green.");
        }

        /// <summary>
        /// Warninng message with one parameter.
        /// </summary>
        /// <param name="myWord">Value passed/set in Workflow Configuration</param>
        public void myWarningMessage(string myWord)
        {
            AddWarning("The light is yellow " + myWord);
        }

        /// <summary>
        /// Error Message with two parameters
        /// </summary>
        /// <param name="myFirstWord"></param>
        /// <param name="mySecondWord"></param>
        public void myErrorMessage(string myFirstWord, string mySecondWord)
        {
            AddError($@"The light is red {myFirstWord}! {mySecondWord}!");
        }

        /// <summary>
        /// Using the AddFatal method with some extras.
        /// </summary>
        /// <param name="myWord"></param>
        public void myFatalMessage(string myWord)
        {
            string whoHitTheLight;
            string sql = @"select 
	                    Name = e.FirstName + ' ' + e.LastName
                    from EMMain e
                    inner join SEUser
                    on e.Employee = SEUser.Employee
                    and SEUser.Username = dbo.GetVisionAuditUsername()";
            //You should probably wrap this in a try-catch block
            DataTable dt = QueryData(sql);
            whoHitTheLight = dt.Rows[0]["Name"].ToString();
            AddFatal($@"{whoHitTheLight} has hit the light! Please notify your {myWord}!");
        }

        public void setDefaultApplicationForAll(string defaultApplication)
        {
            // Examples are Dashboard, Blank, or Timesheet
            string startPage, startPageDesc;
            string sql;
            if (defaultApplication.ToUpper().Equals("BLANK"))
            {

                startPage = @"'None'";
                startPageDesc = "NULL";
            }
            else if (defaultApplication.ToUpper().Equals("TIMESHEET"))
            {
                startPage = @"'Timekeeper'";
                startPageDesc = @"'Timesheet'";
            }
            else
            {
                startPage = "NULL";
                startPageDesc = "NULL";
            }
            sql = $@"update SEUser set StartPage = {startPage} 
                , StartPageDesc = {startPageDesc}";
            int rows = ExecuteSQL(sql);
            AddInformation("The default application has been "
                + $@"set to {defaultApplication} for all users. {rows} records updated");
        }

        public void getMyInformation()
        {
            string FirstName, LastName, Email, EmployeeID, message;
            string sql = $@"select e.* from EMAllCompany e 
                        inner join SEUser 
                        on SEUser.Employee = e.Employee 
                        and SEUser.Username = dbo.GetVisionAuditUsername()";
            DataTable dt = QueryData(sql);
            if (dt.Rows.Count == 1)
            {
                FirstName = dt.Rows[0]["FirstName"].ToString();
                LastName = dt.Rows[0]["LastName"].ToString();
                Email = dt.Rows[0]["Email"].ToString();
                EmployeeID = dt.Rows[0]["Employee"].ToString();
            }
            else
            {
                AddFatal("Bad programming. Please contact your Vision System Administrator.");
                return;
            }
            message = "Here is your information:\n\nEmployeeID: _EmployeeID"
                + "\nFirstName: _FirstName\nLastName: _LastName\nEmail: _Email";
            message = message.Replace("_EmployeeID", EmployeeID)
                .Replace("_FirstName", FirstName)
                .Replace("_LastName", LastName)
                .Replace("_Email", Email);
            AddInformation(message);
        }
    }
}
