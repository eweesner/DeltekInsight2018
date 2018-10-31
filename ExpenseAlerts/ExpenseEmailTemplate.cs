using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExpenseAlerts
{
    internal class ExpenseEmailTemplate : VisionEmail
    {
        private string _body;
        private string _subject;
        private string _database;

        public ExpenseEmailTemplate() : base()
        {}

        public void addSubject(string wbs1, string wbs2, string wbs3)
        { 
            _subject = "New expense lines submitted on " + wbs1 + "." + wbs2 + "." + wbs3;
            this.subject(_subject);
        }
        
        public void addBody(string submitter, string reportEmployeeName,string wbs1, string wbs2, string wbs3)
        {
            string newLine = @"<br \>";
            _body = submitter + " has submitted an expense report for " + reportEmployeeName + "."
                + newLine + "Charge Code: \t" + wbs1 + "." + wbs2 + "." + wbs3
                + newLine + "Employee: \t" + reportEmployeeName
                + newLine
                + newLine + "Click "+ getRecordLink() + " to approve or reject the records.";
            this.body(_body);
        }


        private string getRecordLink()
        {
            string link;
            string url = @"DeltekVision.application?navtreeID=ExpenseLineItemApproval&databaseDescription=RTI-HS Production";
            if (this.getAPPURL().Substring(this.getAPPURL().Length - 1).Equals(@"/"))
                url = getAPPURL() + url;
            else
                url = getAPPURL() + @"/" + url;
            link = @"<a href = """ + url + @""">here</a>";
            return link;
        }



        


    }
}