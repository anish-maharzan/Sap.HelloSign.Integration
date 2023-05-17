﻿using Sap.HelloSign.AddOn.Helpers;
using SAPbobsCOM;
using SAPbouiCOM.Framework;
using System;
using System.Configuration;

namespace Sap.HelloSign.AddOn
{
    class Program
    {
        #region Variables

        public static SAPbobsCOM.Company oCompany;
        public static SAPbouiCOM.Application SBO_Application;
        public static SAPbouiCOM.Form oForm { get; set; }

        #endregion

        /// <summary_>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                if (args.Length < 1)
                    oApp = new Application();
                else
                    oApp = new Application(args[0]);

                SBO_Application = Application.SBO_Application;
                oCompany = (Company)SBO_Application.Company.GetDICompany();
                Menu.addMenu();

                if (ConfigurationManager.AppSettings["UDF"].ToString() == "N")
                {
                    AddonInfoInfo.InstallUDOs();
                }
                //AddonInfoInfo.GetCommonSettings();
                var applicationHandler = new ApplicationHandlers();
                Application.SBO_Application.StatusBar.SetSystemMessage("Add-on installed successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                oApp.Run();

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
    }
}
