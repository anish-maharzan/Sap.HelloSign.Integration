using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;
using GlobalVariable;
using SAPbobsCOM;
using Core.SAPB1;
using Core.Utilities;

namespace Sap.HelloSign.AddOn.Helpers
{
    public class AddonInfoInfo
    {
        public AddonInfoInfo()
        {
        }

        public int Index { get; set; }
        public bool isHana { get; set; }
        private static int RetCode = 0;
        private static string ErrMsg = null;

        #region Methods
        public static bool InstallUDOs()
        {
            try
            {
                bool UDOAdded = true;

                string[] ChildTable = new string[0];
                string[] FindColumn = new string[0];
                string[] FormColumn = new string[0];

                //B1Helper.DiCompany.StartTransaction();

                #region Payment Voucher

                B1Helper.AddTable("SPC_HS_OPMV", "Payment Voucher for HelloSign", BoUTBTableType.bott_Document);
                B1Helper.AddTable("SPC_HS_PMV1", "Payment Voucher - Content", BoUTBTableType.bott_DocumentLines);
                B1Helper.AddTable("SPC_HS_PMV2", "Payment Voucher - Attachments", BoUTBTableType.bott_DocumentLines);
                B1Helper.AddTable("SPC_HS_PMV3", "Payment Voucher - Approver", BoUTBTableType.bott_DocumentLines);
                B1Helper.AddTable("SPC_HS_PMV4", "Payment Voucher - Bank Receipt", BoUTBTableType.bott_DocumentLines);

                B1Helper.AddField("CODE", "Code", "SPC_HS_OPMV", BoFieldTypes.db_Alpha, 30, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("NAME", "Name", "SPC_HS_OPMV", BoFieldTypes.db_Alpha, 50, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("BNKPYDT", "Bank Pay Date", "SPC_HS_OPMV", BoFieldTypes.db_Date, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true);
                B1Helper.AddField("STATUS", "Status", "SPC_HS_OPMV", BoFieldTypes.db_Alpha, 30, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("PAYTO", "Pay To", "SPC_HS_OPMV", BoFieldTypes.db_Alpha, 100, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("ADDRESS", "Address", "SPC_HS_OPMV", BoFieldTypes.db_Alpha, 150, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("INTEGRATION", "Is Ready For Integration", "SPC_HS_OPMV", BoFieldTypes.db_Alpha, 1, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("SAPUSER", "SAP User Name", "SPC_HS_OPMV", BoFieldTypes.db_Alpha, 30, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("HELOSIGNKEY", "HelloSign Key", "SPC_HS_OPMV", BoFieldTypes.db_Alpha, 100, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("REMARKS", "Remarks", "SPC_HS_OPMV", BoFieldTypes.db_Alpha, 100, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("CURR", "Doc Curr", "SPC_HS_OPMV", BoFieldTypes.db_Alpha, 15, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("EXRATE", "Exchange Rate", "SPC_HS_OPMV", BoFieldTypes.db_Float, 15, BoYesNoEnum.tNO, BoFldSubTypes.st_Rate, true, "");
                B1Helper.AddField("AMOUNT", "Total Pay Amount", "SPC_HS_OPMV", BoFieldTypes.db_Float, 15, BoYesNoEnum.tNO, BoFldSubTypes.st_Sum, true, "");


                B1Helper.AddField("DOCNUM", "AP DocNum", "SPC_HS_PMV1", BoFieldTypes.db_Alpha, 10, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("DOCENTY", "AP DocEntry", "SPC_HS_PMV1", BoFieldTypes.db_Alpha, 10, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("POSTDATE", "AP Posting Date", "SPC_HS_PMV1", BoFieldTypes.db_Date, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true);
                B1Helper.AddField("DUEATE", "AP Doc Due Date", "SPC_HS_PMV1", BoFieldTypes.db_Date, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true);
                B1Helper.AddField("OVERDAYS", "AP Overdays", "SPC_HS_PMV1", BoFieldTypes.db_Numeric, 10, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("AMOUNT", "AP Total Amount", "SPC_HS_PMV1", BoFieldTypes.db_Float, 15, BoYesNoEnum.tNO, BoFldSubTypes.st_Sum, true, "");
                B1Helper.AddField("CURR", "Currency", "SPC_HS_PMV1", BoFieldTypes.db_Alpha, 15, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("PAYAMOUNT", "Pay Amount", "SPC_HS_PMV1", BoFieldTypes.db_Float, 15, BoYesNoEnum.tNO, BoFldSubTypes.st_Sum, true, "");

                B1Helper.AddField("SELECT", "Select", "SPC_HS_PMV2", BoFieldTypes.db_Alpha, 1, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("DOCNUM", "AP DocNum", "SPC_HS_PMV2", BoFieldTypes.db_Alpha, 10, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("DOCENTY", "AP DocEntry", "SPC_HS_PMV2", BoFieldTypes.db_Alpha, 10, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("NAME", "Attached Doc Name", "SPC_HS_PMV2", BoFieldTypes.db_Alpha, 50, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                //B1Helper.AddField("SAVE", "Save", "SPC_HS_PMV2", BoFieldTypes.db_Alpha, 50, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                //B1Helper.AddField("HELLOSIGN", "Integrated to HelloSign", "SPC_HS_PMV2", BoFieldTypes.db_Alpha, 50, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("DATE", "Attach Date", "SPC_HS_PMV2", BoFieldTypes.db_Date, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true);
                B1Helper.AddField("FILEPATH", "File Path", "SPC_HS_PMV2", BoFieldTypes.db_Memo, 100, BoYesNoEnum.tNO, BoFldSubTypes.st_Link, true);

                B1Helper.AddField("USER", "User", "SPC_HS_PMV3", BoFieldTypes.db_Alpha, 1, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("EMAIL", "Email", "SPC_HS_PMV3", BoFieldTypes.db_Alpha, 30, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("ORDER", "Order of Approver", "SPC_HS_PMV3", BoFieldTypes.db_Alpha, 30, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");

                B1Helper.AddField("BNKRECPT", "Bank Receipt", "SPC_HS_PMV4", BoFieldTypes.db_Alpha, 30, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true, "");
                B1Helper.AddField("FILEPATH", "File Path", "SPC_HS_PMV4", BoFieldTypes.db_Memo, 100, BoYesNoEnum.tNO, BoFldSubTypes.st_Link, true);
                B1Helper.AddField("DATE", "Approval Date", "SPC_HS_PMV4", BoFieldTypes.db_Date, BoYesNoEnum.tNO, BoFldSubTypes.st_None, true);

                Array.Resize(ref ChildTable, 4);
                ChildTable[0] = "SPC_HS_PMV1";
                ChildTable[1] = "SPC_HS_PMV2";
                ChildTable[2] = "SPC_HS_PMV3";
                ChildTable[3] = "SPC_HS_PMV4";

                B1Helper.CreateUdo("SPC_HS_OPMV", "Payment Voucher for HelloSign", "SPC_HS_OPMV", "D", "N", FormColumn, ChildTable);

                #endregion

                //B1Helper.DiCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                //Utility.LogException("Ending Transaction: UDOs Creation Process");

                return UDOAdded;
            }
            catch (Exception ex)
            {
                //Utility.LogException(ex);
                //B1Helper.DiCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                return false;
            }
        }

        private static bool CreateUDO(string CodeID, string Name, string TableName, string[] FormColoums, SAPbobsCOM.BoUDOObjType ObjectType, string ManageSeries)
        {
            SAPbobsCOM.UserObjectsMD oUserObjectMD = default(SAPbobsCOM.UserObjectsMD);
            try
            {
                oUserObjectMD = ((SAPbobsCOM.UserObjectsMD)(Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)));
                if (oUserObjectMD.GetByKey(CodeID) == true)
                {
                    return true;
                }
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO;

                oUserObjectMD.Code = CodeID;
                oUserObjectMD.Name = Name;
                oUserObjectMD.TableName = TableName;
                oUserObjectMD.ObjectType = ObjectType;


                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.MenuItem = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.MenuCaption = Name;
                oUserObjectMD.FatherMenuID = 47616;
                oUserObjectMD.Position = 0;
                oUserObjectMD.MenuUID = CodeID;

                if (FormColoums != null)
                {
                    for (int i = 0; i <= FormColoums.Length - 1; i++)
                    {
                        if (FormColoums[i].Trim() != "U_RUNDB")
                        {
                            oUserObjectMD.FormColumns.FormColumnAlias = FormColoums[i];
                            oUserObjectMD.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO;
                            oUserObjectMD.FormColumns.Add();
                        }
                        else
                        {
                            oUserObjectMD.FormColumns.FormColumnAlias = FormColoums[i];
                            oUserObjectMD.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                            oUserObjectMD.FormColumns.Add();
                        }
                    }
                }
                // check for errors in the process
                RetCode = oUserObjectMD.Add();

                if (RetCode != 0)
                {
                    if (RetCode != -1)
                    {
                        Program.oCompany.GetLastError(out RetCode, out ErrMsg);
                        Program.SBO_Application.StatusBar.SetText("Object Failed : " + ErrMsg + "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                else
                {
                    Program.SBO_Application.StatusBar.SetText("Object Registered : " + Name + "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public static bool GetCommonSettings()
        {
            string query = "SELECT T0.\"U_A_Email\", T0.\"U_S_Email\", T0.\"U_J_Email\" , \"U_ExcessDay\" , \"U_N_Email\" FROM OADM T0";
            SAPbobsCOM.Recordset rsQry = (SAPbobsCOM.Recordset)B1Helper.DiCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rsQry.DoQuery(query);
            if (rsQry.RecordCount > 0)
            {
                Globals.SetsAEmail(rsQry.Fields.Item(0).Value.ToString());
                Globals.SetsSEmail(rsQry.Fields.Item(1).Value.ToString());
                Globals.SetsJournal(rsQry.Fields.Item(2).Value.ToString());
                Globals.SetsExcessDay(Convert.ToDouble(rsQry.Fields.Item(3).Value.ToString()));
                Globals.SetsNEmail(rsQry.Fields.Item(4).Value.ToString());
            }

            query = "SELECT T0.\"U_BillProcees\", T0.\"U_Account\" FROM \"@Z_SCGL\"  T0";
            rsQry = (SAPbobsCOM.Recordset)B1Helper.DiCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rsQry.DoQuery(query);
            if (rsQry.RecordCount > 0)
            {
                while (rsQry.EoF == false)
                {
                    if (rsQry.Fields.Item(0).Value.ToString() == "A")
                    { Globals.SetsSAdvance(rsQry.Fields.Item(1).Value.ToString()); }
                    else if (rsQry.Fields.Item(0).Value.ToString() == "C") { Globals.SetsSCredit(rsQry.Fields.Item(1).Value.ToString()); }
                    rsQry.MoveNext();
                }
            }
            rsQry = null;
            return true;

        }

        public static void SetFormFilter()
        {
            try
            {
                //SAPbouiCOM.EventFilters objFilters = new SAPbouiCOM.EventFilters();
                //SAPbouiCOM.EventFilter objFilter;

                //objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);
                //objFilter.AddEx("frm_TransferItems");


                //objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE);
                //objFilter.AddEx("frm_TransferItems");



                //objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
                //objFilter.AddEx("frm_TransferItems");


                //objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);
                //objFilter.AddEx("frm_TransferItems");



                //objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);
                //objFilter.AddEx("frm_TransferItems");

                //objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST);
                //objFilter.AddEx("frm_TransferItems");


                //objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN);
                //objFilter.AddEx("frm_TransferItems");


                //objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS);
                //objFilter.AddEx("frm_TransferItems");


                //objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE);
                //objFilter.AddEx("frm_TransferItems");


                //objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD);
                //objFilter.AddEx("frm_TransferItems");



                //objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
                //objFilter.AddEx("frm_TransferItems");


                //objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK);
                //objFilter.AddEx("frm_TransferItems");


                //objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK);
                //objFilter.AddEx("frm_TransferItems");

                //objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED);
                //objFilter.AddEx("frm_TransferItems");


                //SetFilter(objFilters);
            }
            catch (Exception ex)
            {
                Utility.LogException(ex);
                // Log.LogException(LogLevel.Error, ex);
            }
        }

        public static void RemoveMenu(string menuId)
        {
            Application.SBO_Application.Menus.RemoveEx(menuId);
        }

        public static string GetNextEntryIndex(string tableName)
        {
            try
            {
                var result = B1Helper.GetNextEntryIndex(tableName);
                if (result.Equals(string.Empty))
                    result = "0";
                else
                    if (result.Equals("0"))
                {
                    result = "1";
                }

                return result;
            }
            catch (Exception ex)
            {
                Utility.LogException(ex);
                // Log.LogException(LogLevel.Error, ex);
                return null;
            }

        }

        protected static void SetFilter(SAPbouiCOM.EventFilters Filters)
        {
            Application.SBO_Application.SetFilter(Filters);
        }
        #endregion
    }
}

