using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sap.HelloSign.AddOn.Helpers
{
    public static class CommonHelper
    {
        public static bool IsBranchEnabled()
        {
            string query = $@"SELECT T0.""MltpBrnchs"" FROM OADM T0";
            resetRec();
            rec.DoQuery(query);
            return rec.Fields.Item("MltpBrnchs").Value.ToString() == "Y" ? true : false;
        }

        private static SAPbobsCOM.Recordset rec;
        private static void resetRec()
        {
            if (rec != null)
                rec = null;
            if (rec == null)
                rec = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        }
    }
}
