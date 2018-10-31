using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExpenseAlerts
{
    internal class ReportDetail
    {
        public string Employee { set; get; }
        public string ReportName { set; get; }
        public DateTime ReportDate { set; get; }
        public string wbs1 { set; get; }
        public string wbs2 { set; get; }
        public string wbs3 { set; get; }
        public string SubmittedBy { set; get; }
        public string ProjMgr { set; get; }
        public string ProjLead { set; get; }

        public ReportDetail(){}
        public ReportDetail(
            string newEmployee,
            DateTime newReportDate,
            string newReportName,
            string newWBS1,
            string newWBS2,
            string newWBS3,
            string newSubmittedBy,
            string newProjMgr,
            string newProjLead
            )
        {
            Employee = newEmployee;
            ReportName = newReportName;
            ReportDate = newReportDate;
            wbs1 = newWBS1;
            wbs2 = newWBS2;
            wbs3 = newWBS3;
            SubmittedBy = newSubmittedBy;
            ProjMgr = newProjMgr;
            ProjLead = newProjLead;

        }

    }
}
