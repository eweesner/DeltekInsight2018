using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Deltek.Framework.API.Server;
using Deltek.Vision.WorkflowAPI.Server;
using System.Data;
using System.IO;
using System.Globalization;

namespace HyperionWF
{
    /// <summary>
    /// Generate a file of Hyperion data and load into Vision as an attachment
    /// </summary>
    public class GenerateHyperionFile : WorkflowBaseClass
    {
        public void QueryToFile(string Period, string username, string dbServerName, string filedbname)
        {
            /// Test Period against what's in CFGDates
            /// 
            string cfgdates = $@"select Period from CFGDates";
            DataTable dtcfgdates = QueryData(cfgdates);
            bool exists = dtcfgdates.Select().ToList().Exists(row => row["Period"].ToString() == $@"{Period}");
            if (!(exists))
            {
                this.AddError("Please enter a valid period in the format YYYYNN");
                return;
            }           

            /// Define variables
            /// 
            string periodYear = Period.Substring(0, 4);
            string periodNumber = Period.Substring(4, 2);
            string periodStart = periodYear + "00";
            string filename = $@"Hyperion{Period}_" + DateTime.Now.ToString("MMddyyyy-HHmmss") + ".txt";
            string path = $@"\\{dbServerName}\c$\Users\Public";
            string localpath = $@"C:\Users\Public";
            string fullfilename = $@"{path}\{filename}";
            string fulllocalfilename = $@"{localpath}\{filename}";
            string UDIC_UID = "39FF7DA96D2E46B9AB7B41315D50BF0D";

            /// Define query as a string
            /// 
            string query = $@"select 'A' as SCENARIO, '2' as COMPANY_CODE, 'RTI Health Solutions' as COMPANY_DESCRIPTION,
CASE WHEN group1 in ('231.00', '212.00') and group2 = '' THEN '05020942' ELSE  
CASE WHEN group1 ='311.00' and group2 = '' THEN '01020835' ELSE replace(group2, ':', '')  END END as FM_AREA, '{periodYear}' as YEAR, '{periodNumber}' as PERIOD, 
replace(group1, '.', '') as GL_ACCOUNT, groupDesc1 as GL_ACCOUNT_DESCRIPTION, '' as COMMITMENT_ITEM,
sum(Amount) as PERIOD_BALANCE, (sum(openingBalance)-sum(creditAmount) + sum(debitAmount)) as YTD_PERIOD_BALANCE,  'USD' as CURRENCY, 'US Dollars' as CURRENCY_DESCRIPTION, '' as VERSION
from (
Select IsNull(CA.Account, '') As group1, Min(CA.Name) As groupDesc1, IsNull(groupOrganization.Org, '') As group2,  IsNull(Sum(LedgerMisc.Amount ),0) As amount,   
Sum((CASE WHEN LedgerMisc.Amount < 0 THEN Abs(LedgerMisc.Amount) ELSE 0 END) ) As creditAmount, Sum((CASE WHEN LedgerMisc.Amount > 0 THEN LedgerMisc.Amount ELSE 0 END) ) As debitAmount, 0 As openingBalance  
FROM   CA INNER JOIN LedgerMisc ON (CA.Account = LedgerMisc.Account) LEFT JOIN AccountCustomTabFields ON (CA.Account = AccountCustomTabFields.Account) LEFT JOIN Organization AS groupOrganization ON LedgerMisc.Org = groupOrganization.Org Left Join Organization On LedgerMisc.Org = Organization.Org Left Join CFGPostControl On LedgerMisc.Period = CFGPostControl.Period And LedgerMisc.PostSeq = CFGPostControl.PostSeq LEFT JOIN FW_CFGCurrency AS CFGCurrencyTrans ON LedgerMisc.TransactionCurrencyCode = CFGCurrencyTrans.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes1 ON colCFGOrgCodes1.orgLevel = 1 AND substring(LedgerMisc.Org, 1, 2) = colCFGOrgCodes1.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes2 ON colCFGOrgCodes2.orgLevel = 2 AND substring(LedgerMisc.Org, 4, 4) = colCFGOrgCodes2.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes3 ON colCFGOrgCodes3.orgLevel = 3 AND substring(LedgerMisc.Org, 9, 2) = colCFGOrgCodes3.Code , FW_CFGCurrency AS CFGCurrencyFunct, CFGMainData AS CFGMainDataFunct Where (LedgerMisc.Period >= {Period} And LedgerMisc.Period <= {Period}) AND '' = CFGMainDataFunct.Company AND CFGMainDataFunct.FunctionalCurrencyCode = CFGCurrencyFunct.Code And LedgerMisc.Period <= {Period} And LedgerMisc.SkipGL = 'N' /***EXTEND WHERE CLAUSE***/ /***orgWhere***/ GROUP BY IsNull(CA.Account, ''), IsNull(groupOrganization.Org, '') Having Sum(Abs(LedgerMisc.Amount )) <> 0 
UNION ALL
 Select IsNull(CA.Account, '') As group1, Min(CA.Name) As groupDesc1, IsNull(groupOrganization.Org, '') As group2,  0 As amount,   0 As creditAmount, 0 As debitAmount, IsNull(Sum(LedgerMisc.Amount ),0) As openingBalance  FROM   
CA INNER JOIN LedgerMisc ON (CA.Account = LedgerMisc.Account) LEFT JOIN AccountCustomTabFields ON (CA.Account = AccountCustomTabFields.Account) LEFT JOIN Organization AS groupOrganization ON LedgerMisc.Org = groupOrganization.Org Left Join Organization On LedgerMisc.Org = Organization.Org Left Join CFGPostControl On LedgerMisc.Period = CFGPostControl.Period And LedgerMisc.PostSeq = CFGPostControl.PostSeq LEFT JOIN FW_CFGCurrency AS CFGCurrencyTrans ON LedgerMisc.TransactionCurrencyCode = CFGCurrencyTrans.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes1 ON colCFGOrgCodes1.orgLevel = 1 AND substring(LedgerMisc.Org, 1, 2) = colCFGOrgCodes1.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes2 ON colCFGOrgCodes2.orgLevel = 2 AND substring(LedgerMisc.Org, 4, 4) = colCFGOrgCodes2.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes3 ON colCFGOrgCodes3.orgLevel = 3 AND substring(LedgerMisc.Org, 9, 2) = colCFGOrgCodes3.Code , FW_CFGCurrency AS CFGCurrencyFunct, CFGMainData AS CFGMainDataFunct Where (LedgerMisc.Period < {Period}) And ((CA.Type < 4) Or (CA.Type >= 4 And LedgerMisc.Period >= {periodStart})) AND '' = CFGMainDataFunct.Company AND CFGMainDataFunct.FunctionalCurrencyCode = CFGCurrencyFunct.Code And LedgerMisc.Period <= {Period} And LedgerMisc.SkipGL = 'N' /***EXTEND WHERE CLAUSE***/ /***orgWhere***/ GROUP BY IsNull(CA.Account, ''), IsNull(groupOrganization.Org, '') Having Sum(Abs(LedgerMisc.Amount )) <> 0 
UNION ALL
 Select IsNull(CA.Account, '') As group1, Min(CA.Name) As groupDesc1, IsNull(groupOrganization.Org, '') As group2,  0 As amount,    0 As creditAmount, 0 As debitAmount, IsNull(Sum(LedgerMisc.Amount ),0) As openingBalance  FROM   
LedgerMisc INNER JOIN CA AS AccountType ON (AccountType.Account = LedgerMisc.Account) INNER JOIN CA ON (CA.Account = '311.00') LEFT JOIN AccountCustomTabFields ON (CA.Account = AccountCustomTabFields.Account) LEFT JOIN Organization AS groupOrganization ON LedgerMisc.Org = groupOrganization.Org Left Join Organization On LedgerMisc.Org = Organization.Org Left Join CFGPostControl On LedgerMisc.Period = CFGPostControl.Period And LedgerMisc.PostSeq = CFGPostControl.PostSeq LEFT JOIN FW_CFGCurrency AS CFGCurrencyTrans ON LedgerMisc.TransactionCurrencyCode = CFGCurrencyTrans.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes1 ON colCFGOrgCodes1.orgLevel = 1 AND substring(LedgerMisc.Org, 1, 2) = colCFGOrgCodes1.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes2 ON colCFGOrgCodes2.orgLevel = 2 AND substring(LedgerMisc.Org, 4, 4) = colCFGOrgCodes2.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes3 ON colCFGOrgCodes3.orgLevel = 3 AND substring(LedgerMisc.Org, 9, 2) = colCFGOrgCodes3.Code , FW_CFGCurrency AS CFGCurrencyFunct, CFGMainData AS CFGMainDataFunct Where (AccountType.Type >= 4) And (LedgerMisc.Period < {periodStart}) AND '' = CFGMainDataFunct.Company AND CFGMainDataFunct.FunctionalCurrencyCode = CFGCurrencyFunct.Code And LedgerMisc.Period <= {Period} And LedgerMisc.SkipGL = 'N' And AccountType.Account = LedgerMisc.Account /***EXTEND WHERE CLAUSE***/ /***orgWhere***/ And Abs(LedgerMisc.Amount) <> 0 And AccountType.Type >= 4 GROUP BY IsNull(CA.Account, ''), IsNull(groupOrganization.Org, '') Having Sum(Abs(LedgerMisc.Amount )) <> 0 
UNION ALL
 Select IsNull(CA.Account, '') As group1, Min(CA.Name) As groupDesc1, IsNull(groupOrganization.Org, '') As group2,  IsNull(Sum(LedgerEX.Amount ),0) As amount,    
Sum((CASE WHEN LedgerEX.Amount < 0 THEN Abs(LedgerEX.Amount) ELSE 0 END) ) As creditAmount, Sum((CASE WHEN LedgerEX.Amount > 0 THEN LedgerEX.Amount ELSE 0 END) ) As debitAmount, 0 As openingBalance  FROM   
CA INNER JOIN LedgerEX ON (CA.Account = LedgerEX.Account) LEFT JOIN AccountCustomTabFields ON (CA.Account = AccountCustomTabFields.Account) LEFT JOIN Organization AS groupOrganization ON LedgerEX.Org = groupOrganization.Org Left Join Organization On LedgerEX.Org = Organization.Org Left Join CFGPostControl On LedgerEX.Period = CFGPostControl.Period And LedgerEX.PostSeq = CFGPostControl.PostSeq LEFT JOIN FW_CFGCurrency AS CFGCurrencyTrans ON LedgerEX.TransactionCurrencyCode = CFGCurrencyTrans.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes1 ON colCFGOrgCodes1.orgLevel = 1 AND substring(LedgerEX.Org, 1, 2) = colCFGOrgCodes1.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes2 ON colCFGOrgCodes2.orgLevel = 2 AND substring(LedgerEX.Org, 4, 4) = colCFGOrgCodes2.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes3 ON colCFGOrgCodes3.orgLevel = 3 AND substring(LedgerEX.Org, 9, 2) = colCFGOrgCodes3.Code , FW_CFGCurrency AS CFGCurrencyFunct, CFGMainData AS CFGMainDataFunct Where (LedgerEX.Period >= {Period} And LedgerEX.Period <= {Period}) AND '' = CFGMainDataFunct.Company AND CFGMainDataFunct.FunctionalCurrencyCode = CFGCurrencyFunct.Code And LedgerEX.Period <= {Period} And LedgerEX.SkipGL = 'N' /***EXTEND WHERE CLAUSE***/ /***orgWhere***/ GROUP BY IsNull(CA.Account, ''), IsNull(groupOrganization.Org, '') Having Sum(Abs(LedgerEX.Amount )) <> 0 
UNION ALL
 Select IsNull(CA.Account, '') As group1, Min(CA.Name) As groupDesc1, IsNull(groupOrganization.Org, '') As group2,  0 As amount,   0 As creditAmount, 0 As debitAmount, IsNull(Sum(LedgerEX.Amount ),0) As openingBalance  
FROM   CA INNER JOIN LedgerEX ON (CA.Account = LedgerEX.Account) LEFT JOIN AccountCustomTabFields ON (CA.Account = AccountCustomTabFields.Account) LEFT JOIN Organization AS groupOrganization ON LedgerEX.Org = groupOrganization.Org Left Join Organization On LedgerEX.Org = Organization.Org Left Join CFGPostControl On LedgerEX.Period = CFGPostControl.Period And LedgerEX.PostSeq = CFGPostControl.PostSeq LEFT JOIN FW_CFGCurrency AS CFGCurrencyTrans ON LedgerEX.TransactionCurrencyCode = CFGCurrencyTrans.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes1 ON colCFGOrgCodes1.orgLevel = 1 AND substring(LedgerEX.Org, 1, 2) = colCFGOrgCodes1.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes2 ON colCFGOrgCodes2.orgLevel = 2 AND substring(LedgerEX.Org, 4, 4) = colCFGOrgCodes2.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes3 ON colCFGOrgCodes3.orgLevel = 3 AND substring(LedgerEX.Org, 9, 2) = colCFGOrgCodes3.Code , FW_CFGCurrency AS CFGCurrencyFunct, CFGMainData AS CFGMainDataFunct Where (LedgerEX.Period < {Period}) And ((CA.Type < 4) Or (CA.Type >= 4 And LedgerEX.Period >= {periodStart})) AND '' = CFGMainDataFunct.Company AND CFGMainDataFunct.FunctionalCurrencyCode = CFGCurrencyFunct.Code And LedgerEX.Period <= {Period} And LedgerEX.SkipGL = 'N' /***EXTEND WHERE CLAUSE***/ /***orgWhere***/ GROUP BY IsNull(CA.Account, ''), IsNull(groupOrganization.Org, '') Having Sum(Abs(LedgerEX.Amount )) <> 0 
UNION ALL
 Select IsNull(CA.Account, '') As group1, Min(CA.Name) As groupDesc1, IsNull(groupOrganization.Org, '') As group2,  0 As amount,   0 As creditAmount, 0 As debitAmount, IsNull(Sum(LedgerEX.Amount ),0) As openingBalance  
 FROM   LedgerEX INNER JOIN CA AS AccountType ON (AccountType.Account = LedgerEX.Account) INNER JOIN CA ON (CA.Account = '311.00') LEFT JOIN AccountCustomTabFields ON (CA.Account = AccountCustomTabFields.Account) LEFT JOIN Organization AS groupOrganization ON LedgerEX.Org = groupOrganization.Org Left Join Organization On LedgerEX.Org = Organization.Org Left Join CFGPostControl On LedgerEX.Period = CFGPostControl.Period And LedgerEX.PostSeq = CFGPostControl.PostSeq LEFT JOIN FW_CFGCurrency AS CFGCurrencyTrans ON LedgerEX.TransactionCurrencyCode = CFGCurrencyTrans.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes1 ON colCFGOrgCodes1.orgLevel = 1 AND substring(LedgerEX.Org, 1, 2) = colCFGOrgCodes1.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes2 ON colCFGOrgCodes2.orgLevel = 2 AND substring(LedgerEX.Org, 4, 4) = colCFGOrgCodes2.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes3 ON colCFGOrgCodes3.orgLevel = 3 AND substring(LedgerEX.Org, 9, 2) = colCFGOrgCodes3.Code , FW_CFGCurrency AS CFGCurrencyFunct, CFGMainData AS CFGMainDataFunct Where (AccountType.Type >= 4) And (LedgerEX.Period < {periodStart}) AND '' = CFGMainDataFunct.Company AND CFGMainDataFunct.FunctionalCurrencyCode = CFGCurrencyFunct.Code And LedgerEX.Period <= {Period} And LedgerEX.SkipGL = 'N' And AccountType.Account = LedgerEX.Account /***EXTEND WHERE CLAUSE***/ /***orgWhere***/ And Abs(LedgerEX.Amount) <> 0 And AccountType.Type >= 4 GROUP BY IsNull(CA.Account, ''), IsNull(groupOrganization.Org, '') Having Sum(Abs(LedgerEX.Amount )) <> 0 
UNION ALL
 Select IsNull(CA.Account, '') As group1, Min(CA.Name) As groupDesc1, IsNull(groupOrganization.Org, '') As group2, IsNull(Sum(LedgerAP.Amount ),0) As amount,   
Sum((CASE WHEN LedgerAP.Amount < 0 THEN Abs(LedgerAP.Amount) ELSE 0 END) ) As creditAmount, Sum((CASE WHEN LedgerAP.Amount > 0 THEN LedgerAP.Amount ELSE 0 END) ) As debitAmount, 0 As openingBalance  
FROM   CA INNER JOIN LedgerAP ON (CA.Account = LedgerAP.Account) LEFT JOIN AccountCustomTabFields ON (CA.Account = AccountCustomTabFields.Account) LEFT JOIN Organization AS groupOrganization ON LedgerAP.Org = groupOrganization.Org Left Join Organization On LedgerAP.Org = Organization.Org Left Join CFGPostControl On LedgerAP.Period = CFGPostControl.Period And LedgerAP.PostSeq = CFGPostControl.PostSeq LEFT JOIN FW_CFGCurrency AS CFGCurrencyTrans ON LedgerAP.TransactionCurrencyCode = CFGCurrencyTrans.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes1 ON colCFGOrgCodes1.orgLevel = 1 AND substring(LedgerAP.Org, 1, 2) = colCFGOrgCodes1.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes2 ON colCFGOrgCodes2.orgLevel = 2 AND substring(LedgerAP.Org, 4, 4) = colCFGOrgCodes2.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes3 ON colCFGOrgCodes3.orgLevel = 3 AND substring(LedgerAP.Org, 9, 2) = colCFGOrgCodes3.Code , FW_CFGCurrency AS CFGCurrencyFunct, CFGMainData AS CFGMainDataFunct Where (LedgerAP.Period >= {Period} And LedgerAP.Period <= {Period}) AND '' = CFGMainDataFunct.Company AND CFGMainDataFunct.FunctionalCurrencyCode = CFGCurrencyFunct.Code And LedgerAP.Period <= {Period} And LedgerAP.SkipGL = 'N' /***EXTEND WHERE CLAUSE***/ /***orgWhere***/ GROUP BY IsNull(CA.Account, ''), IsNull(groupOrganization.Org, '') Having Sum(Abs(LedgerAP.Amount )) <> 0 
UNION ALL
 Select IsNull(CA.Account, '') As group1, Min(CA.Name) As groupDesc1, IsNull(groupOrganization.Org, '') As group2, 0 As amount,  0 As creditAmount, 0 As debitAmount, IsNull(Sum(LedgerAP.Amount ),0) As openingBalance  
FROM   CA INNER JOIN LedgerAP ON (CA.Account = LedgerAP.Account) LEFT JOIN AccountCustomTabFields ON (CA.Account = AccountCustomTabFields.Account) LEFT JOIN Organization AS groupOrganization ON LedgerAP.Org = groupOrganization.Org Left Join Organization On LedgerAP.Org = Organization.Org Left Join CFGPostControl On LedgerAP.Period = CFGPostControl.Period And LedgerAP.PostSeq = CFGPostControl.PostSeq LEFT JOIN FW_CFGCurrency AS CFGCurrencyTrans ON LedgerAP.TransactionCurrencyCode = CFGCurrencyTrans.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes1 ON colCFGOrgCodes1.orgLevel = 1 AND substring(LedgerAP.Org, 1, 2) = colCFGOrgCodes1.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes2 ON colCFGOrgCodes2.orgLevel = 2 AND substring(LedgerAP.Org, 4, 4) = colCFGOrgCodes2.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes3 ON colCFGOrgCodes3.orgLevel = 3 AND substring(LedgerAP.Org, 9, 2) = colCFGOrgCodes3.Code , FW_CFGCurrency AS CFGCurrencyFunct, CFGMainData AS CFGMainDataFunct Where (LedgerAP.Period < {Period}) And ((CA.Type < 4) Or (CA.Type >= 4 And LedgerAP.Period >= {periodStart})) AND '' = CFGMainDataFunct.Company AND CFGMainDataFunct.FunctionalCurrencyCode = CFGCurrencyFunct.Code And LedgerAP.Period <= {Period} And LedgerAP.SkipGL = 'N' /***EXTEND WHERE CLAUSE***/ /***orgWhere***/ GROUP BY IsNull(CA.Account, ''), IsNull(groupOrganization.Org, '') Having Sum(Abs(LedgerAP.Amount )) <> 0 
UNION ALL
 Select IsNull(CA.Account, '') As group1, Min(CA.Name) As groupDesc1, IsNull(groupOrganization.Org, '') As group2,   0 As amount,    0 As creditAmount, 0 As debitAmount, IsNull(Sum(LedgerAP.Amount ),0) As openingBalance  
 FROM   LedgerAP INNER JOIN CA AS AccountType ON (AccountType.Account = LedgerAP.Account) INNER JOIN CA ON (CA.Account = '311.00') LEFT JOIN AccountCustomTabFields ON (CA.Account = AccountCustomTabFields.Account) LEFT JOIN Organization AS groupOrganization ON LedgerAP.Org = groupOrganization.Org Left Join Organization On LedgerAP.Org = Organization.Org Left Join CFGPostControl On LedgerAP.Period = CFGPostControl.Period And LedgerAP.PostSeq = CFGPostControl.PostSeq LEFT JOIN FW_CFGCurrency AS CFGCurrencyTrans ON LedgerAP.TransactionCurrencyCode = CFGCurrencyTrans.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes1 ON colCFGOrgCodes1.orgLevel = 1 AND substring(LedgerAP.Org, 1, 2) = colCFGOrgCodes1.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes2 ON colCFGOrgCodes2.orgLevel = 2 AND substring(LedgerAP.Org, 4, 4) = colCFGOrgCodes2.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes3 ON colCFGOrgCodes3.orgLevel = 3 AND substring(LedgerAP.Org, 9, 2) = colCFGOrgCodes3.Code , FW_CFGCurrency AS CFGCurrencyFunct, CFGMainData AS CFGMainDataFunct Where (AccountType.Type >= 4) And (LedgerAP.Period < {periodStart}) AND '' = CFGMainDataFunct.Company AND CFGMainDataFunct.FunctionalCurrencyCode = CFGCurrencyFunct.Code And LedgerAP.Period <= {Period} And LedgerAP.SkipGL = 'N' And AccountType.Account = LedgerAP.Account /***EXTEND WHERE CLAUSE***/ /***orgWhere***/ And Abs(LedgerAP.Amount) <> 0 And AccountType.Type >= 4 GROUP BY IsNull(CA.Account, ''), IsNull(groupOrganization.Org, '') Having Sum(Abs(LedgerAP.Amount )) <> 0 
UNION ALL
 Select IsNull(CA.Account, '') As group1, Min(CA.Name) As groupDesc1, IsNull(groupOrganization.Org, '') As group2,   IsNull(Sum(LedgerAR.Amount ),0) As amount,     
Sum((CASE WHEN LedgerAR.Amount < 0 THEN Abs(LedgerAR.Amount) ELSE 0 END) ) As creditAmount, Sum((CASE WHEN LedgerAR.Amount > 0 THEN LedgerAR.Amount ELSE 0 END) ) As debitAmount, 0 As openingBalance  
FROM   CA INNER JOIN LedgerAR ON (CA.Account = LedgerAR.Account) LEFT JOIN AccountCustomTabFields ON (CA.Account = AccountCustomTabFields.Account) LEFT JOIN Organization AS groupOrganization ON LedgerAR.Org = groupOrganization.Org Left Join Organization On LedgerAR.Org = Organization.Org Left Join CFGPostControl On LedgerAR.Period = CFGPostControl.Period And LedgerAR.PostSeq = CFGPostControl.PostSeq LEFT JOIN FW_CFGCurrency AS CFGCurrencyTrans ON LedgerAR.TransactionCurrencyCode = CFGCurrencyTrans.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes1 ON colCFGOrgCodes1.orgLevel = 1 AND substring(LedgerAR.Org, 1, 2) = colCFGOrgCodes1.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes2 ON colCFGOrgCodes2.orgLevel = 2 AND substring(LedgerAR.Org, 4, 4) = colCFGOrgCodes2.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes3 ON colCFGOrgCodes3.orgLevel = 3 AND substring(LedgerAR.Org, 9, 2) = colCFGOrgCodes3.Code , FW_CFGCurrency AS CFGCurrencyFunct, CFGMainData AS CFGMainDataFunct Where (LedgerAR.Period >= {Period} And LedgerAR.Period <= {Period}) AND '' = CFGMainDataFunct.Company AND CFGMainDataFunct.FunctionalCurrencyCode = CFGCurrencyFunct.Code And LedgerAR.Period <= {Period} And LedgerAR.SkipGL = 'N' /***EXTEND WHERE CLAUSE***/ /***orgWhere***/ GROUP BY IsNull(CA.Account, ''), IsNull(groupOrganization.Org, '') Having Sum(Abs(LedgerAR.Amount )) <> 0 
UNION ALL
 Select IsNull(CA.Account, '') As group1, Min(CA.Name) As groupDesc1, IsNull(groupOrganization.Org, '') As group2,  0 As amount,    0 As creditAmount, 0 As debitAmount, IsNull(Sum(LedgerAR.Amount ),0) As openingBalance  
FROM   CA INNER JOIN LedgerAR ON (CA.Account = LedgerAR.Account) LEFT JOIN AccountCustomTabFields ON (CA.Account = AccountCustomTabFields.Account) LEFT JOIN Organization AS groupOrganization ON LedgerAR.Org = groupOrganization.Org Left Join Organization On LedgerAR.Org = Organization.Org Left Join CFGPostControl On LedgerAR.Period = CFGPostControl.Period And LedgerAR.PostSeq = CFGPostControl.PostSeq LEFT JOIN FW_CFGCurrency AS CFGCurrencyTrans ON LedgerAR.TransactionCurrencyCode = CFGCurrencyTrans.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes1 ON colCFGOrgCodes1.orgLevel = 1 AND substring(LedgerAR.Org, 1, 2) = colCFGOrgCodes1.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes2 ON colCFGOrgCodes2.orgLevel = 2 AND substring(LedgerAR.Org, 4, 4) = colCFGOrgCodes2.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes3 ON colCFGOrgCodes3.orgLevel = 3 AND substring(LedgerAR.Org, 9, 2) = colCFGOrgCodes3.Code , FW_CFGCurrency AS CFGCurrencyFunct, CFGMainData AS CFGMainDataFunct Where (LedgerAR.Period < {Period}) And ((CA.Type < 4) Or (CA.Type >= 4 And LedgerAR.Period >= {periodStart})) AND '' = CFGMainDataFunct.Company AND CFGMainDataFunct.FunctionalCurrencyCode = CFGCurrencyFunct.Code And LedgerAR.Period <= {Period} And LedgerAR.SkipGL = 'N' /***EXTEND WHERE CLAUSE***/ /***orgWhere***/ GROUP BY IsNull(CA.Account, ''), IsNull(groupOrganization.Org, '') Having Sum(Abs(LedgerAR.Amount )) <> 0 
UNION ALL
 Select IsNull(CA.Account, '') As group1, Min(CA.Name) As groupDesc1, IsNull(groupOrganization.Org, '') As group2,  0 As amount,    0 As creditAmount, 0 As debitAmount, IsNull(Sum(LedgerAR.Amount ),0) As openingBalance  
 FROM   LedgerAR INNER JOIN CA AS AccountType ON (AccountType.Account = LedgerAR.Account) INNER JOIN CA ON (CA.Account = '311.00') LEFT JOIN AccountCustomTabFields ON (CA.Account = AccountCustomTabFields.Account) LEFT JOIN Organization AS groupOrganization ON LedgerAR.Org = groupOrganization.Org Left Join Organization On LedgerAR.Org = Organization.Org Left Join CFGPostControl On LedgerAR.Period = CFGPostControl.Period And LedgerAR.PostSeq = CFGPostControl.PostSeq LEFT JOIN FW_CFGCurrency AS CFGCurrencyTrans ON LedgerAR.TransactionCurrencyCode = CFGCurrencyTrans.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes1 ON colCFGOrgCodes1.orgLevel = 1 AND substring(LedgerAR.Org, 1, 2) = colCFGOrgCodes1.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes2 ON colCFGOrgCodes2.orgLevel = 2 AND substring(LedgerAR.Org, 4, 4) = colCFGOrgCodes2.Code LEFT JOIN cfgOrgCodes As colCFGOrgCodes3 ON colCFGOrgCodes3.orgLevel = 3 AND substring(LedgerAR.Org, 9, 2) = colCFGOrgCodes3.Code , FW_CFGCurrency AS CFGCurrencyFunct, CFGMainData AS CFGMainDataFunct Where (AccountType.Type >= 4) And (LedgerAR.Period < {periodStart}) AND '' = CFGMainDataFunct.Company AND CFGMainDataFunct.FunctionalCurrencyCode = CFGCurrencyFunct.Code And LedgerAR.Period <= {Period} And LedgerAR.SkipGL = 'N' And AccountType.Account = LedgerAR.Account /***EXTEND WHERE CLAUSE***/ /***orgWhere***/ And Abs(LedgerAR.Amount) <> 0 And AccountType.Type >= 4 GROUP BY IsNull(CA.Account, ''), IsNull(groupOrganization.Org, '') Having Sum(Abs(LedgerAR.Amount )) <> 0 
) as Q1 
group by group1, groupDesc1, group2
ORDER BY 1 ASC, 3 ASC";

            /// Run query and send results to a DataTable
           DataTable dt = QueryData(query);

            /// Write DataTable to semicolon-delimited text file
            try
            {
                StreamWriter sw = new StreamWriter(fullfilename, false, Encoding.UTF8);
                int columnCount = dt.Columns.Count;

                for (int i = 0; i < columnCount; i++)
                {
                    sw.Write(dt.Columns[i]);

                    if (i < columnCount - 1)
                    {
                        sw.Write(";");
                    }
                }

                sw.Write(sw.NewLine);

                foreach (DataRow dr in dt.Rows)
                {
                    for (int i = 0; i < columnCount; i++)
                    {
                        if (!Convert.IsDBNull(dr[i]))
                        {
                            sw.Write(dr[i].ToString());
                        }

                        if (i < columnCount - 1)
                        {
                            sw.Write(";");
                        }
                    }

                    sw.Write(sw.NewLine);
                }

                sw.Close();
            }
            catch (Exception ex)
            {
                this.AddFatal("Cannot create file" + ex.ToString());
                //throw ex;
            }

            /// Insert file into Vision database
            /// 
            string fileid = Guid.NewGuid().ToString().ToUpper();
            string insertfile = $@"INSERT INTO {filedbname}.dbo.FW_Files(FileID, FileData) SELECT '{fileid}', BulkColumn FROM OPENROWSET(BULK N'{fulllocalfilename}', SINGLE_BLOB) as f;";
            ExecuteSQL(insertfile);
        
            /// Get file size
            /// 
            long fileSizeInBytes = new FileInfo(fullfilename).Length;

            /// Insert file metadata
            /// 
            string insertmetadata = $@"INSERT INTO FW_Files(FileID, FileName, ContentType, FileApplicationType, FileDescription, FileSize, CreateUser, CreateDate) 
                VALUES ('{fileid}', '{filename}', 'text/plain', 'UDIC_SystemWorkflows', '{filename}', '{fileSizeInBytes}', '{username}', getutcdate())";
            ExecuteSQL(insertmetadata);

            /// Insert data into FW_Attachments
            /// 
            string insertattachmentdata = $@"INSERT INTO FW_Attachments(PKey, Key1, FileID, Application, CreateUser, CreateDate, ModUser, ModDate)
                VALUES (replace(newid(),'-',''), '{UDIC_UID}', '{fileid}', 'UDIC_SystemWorkflows', '{username}', getutcdate(), '{username}', getutcdate())";
            ExecuteSQL(insertattachmentdata);

            File.Delete(fullfilename);

        }
    }
}
