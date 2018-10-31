using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Deltek.Vision.WorkflowAPI.Server;
using System.Data;

namespace ExpenseAlerts
{
    public class ExpenseLineItemAlerts : WorkflowBaseClass
    {
        
        private ReportDetail _currentLine;
        private List<ReportDetail> _alertList;
        private static string _CONTACTSYSADMIN = "\r\nPlease click email to contact the System Administrator.";
        
        public void SendEmailAlertOnSubmit(
            string Employee,
            DateTime ReportDate,
            string ReportName,
            string WBS1,
            string WBS2,
            string WBS3,
            string SubmittedBy,
            string ProjMgr,
            string ProjLead)
        {
            _currentLine = new ReportDetail(
                Employee,
                ReportDate,
                ReportName,
                WBS1,
                WBS2,
                WBS3,
                SubmittedBy,
                ProjMgr,
                ProjLead
                );
            InsertIntoAlertTable(_currentLine);
            //this.AddInformation("Before Last Line Item: " + ReportDate);
            if (IsLastLineItem())
            {
                //this.AddInformation("Last line submitted");
                //this.AddInformation("New Alerts Sent");
                getAlertList(); //1
                try
                {
                    emailAlertOnSubmit(); //2
                }
                catch (Exception e)
                {
                    this.AddFatal(e.ToString() + _CONTACTSYSADMIN);
                }
                CleanAlertTable(); //3
            }
            
        }

        //this method inserts the data into the alert table
        private void InsertIntoAlertTable(ReportDetail r)
        {
            string sql = "insert into ekAlerts_RTIHS ";
            sql += @"values('_Employee','_ReportDate','_ReportName','_WBS1','_WBS2','_WBS3','_SubmittedBy','_ProjMgr','_ProjLead');";
            sql = sql.Replace("_Employee", r.Employee);
            sql = sql.Replace("_ReportDate", r.ReportDate.ToString("yyyy-MM-dd"));
            sql = sql.Replace("_ReportName",r.ReportName);
            sql = sql.Replace("_WBS1",r.wbs1);
            sql = sql.Replace("_WBS2",r.wbs2);
            sql = sql.Replace("_WBS3",r.wbs3);
            sql = sql.Replace("_SubmittedBy",r.SubmittedBy);
            sql = sql.Replace("_ProjMgr",r.ProjMgr);
            sql = sql.Replace("_ProjLead", r.ProjLead);
            try
            {
                this.ExecuteSQL(sql);
            }
            catch (Exception e)
            {
                this.AddFatal(e.ToString() + "\r\n" + sql);
            }
            /*finally
            {
                this.AddInformation(r.wbs1 + ":" + r.wbs2 + ":" + r.wbs3);
            }*/
        }

        //check if the expense report workflow is done
        //if number of rows in ApprovalItem table = 0 then set to true; this is the last item
        //EWW 7-18-17 added inner join ekDetail to ensure that the expense report item list was accurate
        private bool IsLastLineItem()
        {
            bool value = false;
            string sqlDetail; 
            // string sqlAlerts;
            sqlDetail = @"select 'x' from ApprovalItem a "
                    + @"inner join ekDetail ekd on ekd.Employee = N'_Employee' "
                    + @"and ekd.ReportDate = '_ReportDate' and ekd.ReportName = '_ReportName' "
                    + @"and ekd.Seq = cast(left(a.ApplicationKey,4) as int) "
                    + @"where ApplicationID = 'EXPENSE' "
                    + @"and ApplicationKey like '%|_Employee|_ReportDate|_ReportName' "
                    + @"and ItemType_UID = 'DefaultExpenseReportandLine' "
                    + @"and (WorkflowStep = 0 OR (WorkflowStep <> 0 AND IsUpdated = 'Y')) "
                    + @"and IsSuspended = 'N'";
            sqlDetail = sqlDetail.Replace("_Employee", _currentLine.Employee);
            sqlDetail = sqlDetail.Replace("_ReportDate", _currentLine.ReportDate.ToString("yyyy-MM-dd"));
            sqlDetail = sqlDetail.Replace("_ReportName", _currentLine.ReportName);
            int ekDetailRows = this.QueryData(sqlDetail).Rows.Count;
            if (ekDetailRows == 0)
                value = true;
            return value;
        }

        //email message
        private void emailAlertOnSubmit()
        {
            foreach (ReportDetail rd in _alertList)
            {
                ExpenseEmailTemplate alert = new ExpenseEmailTemplate();
                alert.addSubject(rd.wbs1,rd.wbs2,rd.wbs3);
                alert.addBody(
                    getEmployeeName(rd.SubmittedBy),
                    getEmployeeName(rd.Employee),
                    rd.wbs1,
                    rd.wbs2,
                    rd.wbs3
                    );
                alert.addTo(getEmailAddress(rd.ProjMgr));
                /*if (rd.ProjMgr == rd.SubmittedBy)
                    alert.addCC(getEmailAddress(rd.ProjLead));*/
                //alert.addBcc("eweesner@rti.org");
                alert.sendMail();
            }
        }

        //clean out expense report information for ekAlert_RTIHS table
        private void CleanAlertTable()
        {
            string sql;
            /*ReportDetail rd = new ReportDetail();
            rd = (from row in _alertList
                         select row).First();*/
            sql = @"delete from ekAlerts_RTIHS where Employee = '_Employee' and ReportDate = '_ReportDate' and ReportName = '_ReportName'";
            sql = sql.Replace("_Employee", _currentLine.Employee);
            sql = sql.Replace("_ReportName", _currentLine.ReportName);
            sql = sql.Replace("_ReportDate", _currentLine.ReportDate.ToString("yyyy-MM-dd"));
            try
            {
                this.ExecuteSQL(sql);
            }
            catch (Exception e)
            {
                this.AddFatal(e.ToString() + _CONTACTSYSADMIN);
            }
        }

        //get list of unique values from ekAlerts_RTIHS table
        //EWW 7-18-17 added where clause for sql statement to ensure accuracy
        private void getAlertList()
        { 
            DataTable dt = new DataTable();
            _alertList = new List<ReportDetail>();
            string sql;
            sql = @"select Employee, ReportDate, ReportName, WBS1, WBS2, WBS3, SubmittedBy = MAX(SubmittedBy), ProjMgr = MAX(ProjMgr), 
                ProjLead = MAX(ProjLead) from ekAlerts_RTIHS where Employee = '_Employee' and ReportDate = '_ReportDate' and ReportName = '_ReportName' 
                group by Employee, ReportDate, ReportName, WBS1, WBS2, WBS3";
            sql = sql.Replace("_Employee", _currentLine.Employee);
            sql = sql.Replace("_ReportName", _currentLine.ReportName);
            sql = sql.Replace("_ReportDate", _currentLine.ReportDate.ToString("yyyy-MM-dd"));
            try
            {
                dt = this.QueryData(sql);
            }
            catch(Exception e)
            {
                this.AddFatal(e.ToString() + _CONTACTSYSADMIN);
            }
            foreach (DataRow r in dt.Rows)
            {
                _alertList.Add(new ReportDetail(
                    r["Employee"].ToString(),
                    Convert.ToDateTime(r["ReportDate"]),
                    r["ReportName"].ToString(),
                    r["WBS1"].ToString(),
                    r["WBS2"].ToString(),
                    r["WBS3"].ToString(),
                    r["SubmittedBy"].ToString(),
                    r["ProjMgr"].ToString(),
                    r["ProjLead"].ToString()
                    ));
            }            
        }
        
        //get employee name; use this for submitter and employee expense report
        private string getEmployeeName(string employee)
        {
            string employeeName = getEmployeeData("FullName", @"LastName + ', ' + FirstName", employee);
            return employeeName;
        }

        //get email address; use this for PM and PL email addresses
        private string getEmailAddress(string employee)
        {
            string emailAddress = getEmployeeData("Email", "Email", employee);
            return emailAddress;
        }

        private string getEmployeeData(string colName, string colValues, string employee)
        {
            string name = "";
            DataTable dt = new DataTable();
            string sql;
            sql = "select " + colName + " = " +  colValues + " from EMMain where Employee = '" + employee + "'";
            try
            {
                dt = this.QueryData(sql);
            }
            catch(Exception e)
            {
                this.AddFatal(e.ToString() + _CONTACTSYSADMIN);
            }
            name = dt.Rows[0][colName].ToString();
            return name;
        }
    }
}
