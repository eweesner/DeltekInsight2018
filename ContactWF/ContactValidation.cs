using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Deltek.Framework.API.Server;
using Deltek.Vision.WorkflowAPI.Server;
using System.Data;

namespace ContactsWF
{
    /// <summary>
    /// Contact Validation for the first and last name.
    /// </summary>
    public class ContactValidation : WorkflowBaseClass
    {

        // query table to see if there are any matches
        // if any rows exist
        public void validateName(string contactID, string firstName, string lastName)
        {
            string sql = $@"select FirstName, LastName, Client = Isnull((select Name from CL where CL.ClientID = Contacts.ClientID) ,''), Vendor = isnull((select Name from VE where VE.Vendor = Contacts.Vendor),'') from Contacts where FirstName = N'{firstName}' and LastName = N'{lastName}' and ContactID <> '{contactID}'";
            DataTable dt = this.QueryData(sql);
            int numberOfRows = dt.Rows.Count;
            if (numberOfRows > 0)
            {
                string str = "";
                foreach (DataRow r in dt.Rows)
                {
                    str += ("\n" + $@"{r["FirstName"].ToString()} {r["LastName"].ToString()}, ");
                    if (!string.IsNullOrWhiteSpace(r["Client"].ToString()))
                    {
                        str += $@"Client: {r["Client"].ToString()}";
                    }
                    else if (!string.IsNullOrWhiteSpace(r["Vendor"].ToString()))
                    {
                        str += $@"Vendor: {r["Vendor"].ToString()}";
                    }
                    else
                        str += "No Client or Vendor listed";
                }
                this.AddWarning("Possible Duplicate(s) Found.\n" + str);
            }
            
        }

    }
}
