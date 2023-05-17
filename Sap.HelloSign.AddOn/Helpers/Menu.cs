using Core.SAPB1;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;

namespace Sap.HelloSign.AddOn.Helpers
{
    public class Menu
    {
        public static void addMenu()
        {
            B1Helper.addMenuItem("43520", "Payment Voucher", "Payment Voucher");
            //B1Helper.AddSubMenu(MenuID_UD.MODULE, "Gate Pass master", "Gate Pass", -1, string.Concat(System.Windows.Forms.Application.StartupPath, @"\Images\Icon.png"));
        }
        public static void RemoveMenu()
        {
            //B1Helper.RemoveMenuItem("3072", "Auto Item Master Code Setup", "Auto Item Master Code Setup");
        }
    }
}
