using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using Core.SAPB1;
using SAPbobsCOM;
using System.Windows.Forms;
using Core.Utilities;
using System.IO;
using System.Data;
using System.Collections;

namespace Sap.HelloSign.AddOn.Forms
{
    [FormAttribute("Sap.HelloSign.AddOn.Forms.Payment_Voucher", "Forms/Payment Voucher.b1f")]
    class Payment_Voucher : UserFormBase
    {
        public Payment_Voucher()
        {
        }

        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText txtCode;
        private SAPbouiCOM.EditText txtName;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.EditText txtPayTo;
        private SAPbouiCOM.EditText txtNo;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.EditText txtDate;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.EditText txtBPD;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.Folder Folder1;
        private SAPbouiCOM.Folder Folder2;
        private SAPbouiCOM.Folder Folder3;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Matrix Matrix1;
        private SAPbouiCOM.Matrix Matrix3;
        private SAPbouiCOM.Matrix Matrix4;
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.EditText txtSUN;
        private SAPbouiCOM.StaticText StaticText8;
        private SAPbouiCOM.EditText txtDC;
        private SAPbouiCOM.StaticText StaticText9;
        private SAPbouiCOM.EditText txtER;
        private SAPbouiCOM.StaticText StaticText10;
        private SAPbouiCOM.EditText txtTPA;
        private SAPbouiCOM.StaticText StaticText11;
        private SAPbouiCOM.EditText txtHSK;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Button Button3;
        private SAPbouiCOM.Button Button4;
        private SAPbouiCOM.CheckBox CheckBox0;
        private SAPbouiCOM.ComboBox ComboBox2;
        private SAPbouiCOM.StaticText StaticText12;
        private SAPbouiCOM.EditText txtRemarks;
        private string FileName;
        private string stFilePathAndName = "";
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.LinkedButton LinkedButton0;
        List<string> before = new List<string>();
        List<string> after = new List<string>();
        private Recordset Rec;

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.txtCode = ((SAPbouiCOM.EditText)(this.GetItem("txtCode").Specific));
            this.txtCode.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.txtCode_ChooseFromListBefore);
            this.txtCode.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.txtCode_ChooseFromListAfter);
            this.txtName = ((SAPbouiCOM.EditText)(this.GetItem("txtName").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_6").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.txtPayTo = ((SAPbouiCOM.EditText)(this.GetItem("txtPayTo").Specific));
            this.txtNo = ((SAPbouiCOM.EditText)(this.GetItem("txtNo").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_9").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_10").Specific));
            this.txtDate = ((SAPbouiCOM.EditText)(this.GetItem("txtDate").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_12").Specific));
            this.txtBPD = ((SAPbouiCOM.EditText)(this.GetItem("txtBPD").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_14").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("Item_17").Specific));
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("Item_18").Specific));
            this.Folder2 = ((SAPbouiCOM.Folder)(this.GetItem("Item_19").Specific));
            this.Folder3 = ((SAPbouiCOM.Folder)(this.GetItem("Item_20").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_0").Specific));
            this.Matrix0.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix0_ValidateAfter);
            this.Matrix0.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.Matrix0_ChooseFromListBefore);
            this.Matrix0.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix0_ChooseFromListAfter);
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_2").Specific));
            this.Matrix3 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_7").Specific));
            this.Matrix4 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_8").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_11").Specific));
            this.txtSUN = ((SAPbouiCOM.EditText)(this.GetItem("txtSUN").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_15").Specific));
            this.txtDC = ((SAPbouiCOM.EditText)(this.GetItem("txtDC").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_22").Specific));
            this.txtER = ((SAPbouiCOM.EditText)(this.GetItem("txtER").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_24").Specific));
            this.txtTPA = ((SAPbouiCOM.EditText)(this.GetItem("txtTPA").Specific));
            this.txtTPA.ClickAfter += new SAPbouiCOM._IEditTextEvents_ClickAfterEventHandler(this.txtTPA_ClickAfter);
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_26").Specific));
            this.txtHSK = ((SAPbouiCOM.EditText)(this.GetItem("txtHSK").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_30").Specific));
            this.oForm = ((SAPbouiCOM.Form)(this.UIAPIRawForm));
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("Item_34").Specific));
            this.Button3.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button3_ClickAfter);
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("Item_35").Specific));
            this.Button4.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button4_ClickAfter);
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("Item_13").Specific));
            this.ComboBox2 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_21").Specific));
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_23").Specific));
            this.txtRemarks = ((SAPbouiCOM.EditText)(this.GetItem("txtRemarks").Specific));
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_3").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.RightClickBefore += new SAPbouiCOM.Framework.FormBase.RightClickBeforeHandler(this.Form_RightClickBefore);

        }

        private void OnCustomInitialize()
        {
            //oForm.EnableMenu("1284", true);
            //oForm.EnableMenu("1286", true);
            //oForm.EnableMenu("1292", true);
            //oForm.EnableMenu("1293", true);
            //UIAPIRawForm.EnableMenu("773", false);
            //UIAPIRawForm.EnableMenu("5377", false);
            UIAPIRawForm.EnableMenu("774", false);


            BasicBinding();

            Program.SBO_Application.MenuEvent += this.SBO_Application_MenuEvent;
        }

        private void BasicBinding()
        {
            txtNo.Value = B1Helper.GetNextDocNum("@SPC_HS_OPMV").ToString();
            string usercode = Program.SBO_Application.Company.UserName;
            string Query = "Select U_NAME FROM OUSR WHERE USER_CODE = '" + usercode + "'";
            Recordset rs = (Recordset)B1Helper.DiCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rs.DoQuery(Query);
            DateTime date = DateTime.Now;
            txtDate.Value = date.ToString("yyyyMMdd");
            txtBPD.Value = date.AddDays(3).ToString("yyyyMMdd");
            txtSUN.Value = rs.Fields.Item("U_NAME").Value.ToString();
            Matrix0.AddLine();
            Extentions.SetLineId(Matrix0);
            Matrix0.AutoResizeColumns();

            Matrix1.AddLine();
            Extentions.SetLineId(Matrix1);
            Matrix1.AutoResizeColumns();

            Matrix4.AddLine();
            Extentions.SetLineId(Matrix4);
            Matrix4.AutoResizeColumns();

            Matrix3.AddLine();
            Extentions.SetLineId(Matrix3);
            Matrix3.AutoResizeColumns();

            //Matrix4.Columns.Item("13").Editable = true;
            Matrix4.AutoResizeColumns();
            Extentions.SetLineId(Matrix4);
            Folder0.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            ComboBox2.Select("Open", SAPbouiCOM.BoSearchKey.psk_ByValue);
        }

        private string showOpenFileDialog()
        {

            System.Threading.Thread ShowFolderBrowserThread;
            try
            {
                ShowFolderBrowserThread = new System.Threading.Thread(ShowFolderBrowser);
                if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted)
                {
                    ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA);
                    ShowFolderBrowserThread.Start();
                    ShowFolderBrowserThread.Join();
                }
                else if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Stopped)
                {
                    ShowFolderBrowserThread.Start();
                    ShowFolderBrowserThread.Join();
                }
                while (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Running)
                    System.Windows.Forms.Application.DoEvents();
                if (FileName != "")
                    return FileName;
            }
            catch (Exception ex)
            {
                // SBO_Application.MessageBox("FileFile" & ex.Message)
                MessageBox.Show(ex.ToString());
            }

            return "";
        }

        private void ShowFolderBrowser()
        {
            System.Diagnostics.Process[] MyProcs;
            FileName = "";
            stFilePathAndName = "";
            OpenFileDialog OpenFile = new OpenFileDialog();


            try
            {
                OpenFile.Multiselect = false;

                OpenFile.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"; // "All files(*.CSV)|*.CSV"
                int filterindex = 0;
                try
                {
                    filterindex = 0;
                }
                catch (Exception ex)
                {
                }

                OpenFile.FilterIndex = filterindex;
                OpenFile.InitialDirectory = @"C:\"; // orset.Fields.Item("C:\").Value
                OpenFile.RestoreDirectory = false;
                MyProcs = System.Diagnostics.Process.GetProcessesByName("SAP Business One");

                // If MyProcs.Length = 1 Then
                // For i As Integer = 0 To MyProcs.Length - 1

                WindowWrapper MyWindow = new WindowWrapper(MyProcs[0].MainWindowHandle);
                DialogResult ret = OpenFile.ShowDialog(MyWindow);
                if (ret == DialogResult.OK)
                {
                    stFilePathAndName = OpenFile.FileName;
                    System.IO.FileInfo MyFile = new System.IO.FileInfo(stFilePathAndName);
                    FileName = MyFile.Name;

                    // 'FileName = OpenFile.FileName
                    OpenFile.Dispose();
                }
                else
                {
                    FileName = "";
                    System.Windows.Forms.Application.ExitThread();
                }
            }
            // Next
            // End If
            catch (Exception ex)
            {
                // SBO_Application.StatusBar.SetText(ex.Message)
                MessageBox.Show(ex.ToString());
                FileName = "";
            }
            finally
            {
                OpenFile.Dispose();
            }
        }

        private Recordset ExecuteQuery(string query)
        {
            Recordset rec = null;
            try
            {
                rec = (Recordset)B1Helper.DiCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                rec.DoQuery(query);
                return rec;
            }
            catch (Exception) { }
            return rec;
        }

        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                oForm = Program.SBO_Application.Forms.ActiveForm;
                if (oForm.Title.Trim() == "Payment Voucher for Hello Sign")
                {
                    if (!pVal.BeforeAction)
                    {
                        if (pVal.MenuUID == "1282")
                        {
                            BasicBinding();
                        }
                        else if (pVal.MenuUID == "1281")
                        {
                            txtNo.Item.Enabled = true;
                            ComboBox2.Item.Enabled = true;
                            txtName.Item.Enabled = true;
                        }

                        if (pVal.MenuUID == "1284")
                        {

                        }
                        else if (pVal.MenuUID == "1286")
                        {
                        }

                        if (pVal.MenuUID == "1290" || pVal.MenuUID == "1288" || pVal.MenuUID == "1289" || pVal.MenuUID == "1291")
                        {
                            UIAPIRawForm.EnableMenu("1284", true);
                            UIAPIRawForm.EnableMenu("1286", true);
                        }


                        if (pVal.MenuUID == "1293")
                        {
                            after = new List<string>();
                            for (int i = 1; i <= Matrix0.VisualRowCount; i++)
                            {
                                var docentry = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Selected").Cells.Item(i).Specific).Value;
                                after.Add(docentry);
                            }
                            Extentions.SetLineId(Matrix0);
                            List<string> deleted = before.Except(after).ToList();
                            RemoveAttachment(deleted.First());
                            Matrix0.DeleteRow(Matrix0.RowCount-1);
                            Matrix0.AddLine();
                            Extentions.SetLineId(Matrix0);

                        }
                    }
                    if (pVal.BeforeAction)
                    {
                        if (pVal.MenuUID == "1292")
                        {
                            Matrix0.AddLine();
                            Extentions.SetLineId(Matrix0);
                            BubbleEvent = false;
                        }
                        if (pVal.MenuUID == "1293")
                        {
                            before = new List<string>();
                            for (int i = 1; i <= Matrix0.VisualRowCount; i++)
                            {
                                var docentry = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Selected").Cells.Item(i).Specific).Value;
                                before.Add(docentry);
                            }
                            Extentions.SetLineId(Matrix0);
                        }
                    }
                }
            }
            catch (Exception ex) { }
        }

        private void Button3_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (UIAPIRawForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    showOpenFileDialog();
                    string query = "select \"AttachPath\" from \"OADP\"";
                    Recordset rs = (Recordset)B1Helper.DiCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    rs.DoQuery(query);

                    if (rs.RecordCount == 0)
                    {
                        Program.SBO_Application.StatusBar.SetText("Please Provide the Attachment Path... !", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                    }
                    else
                    {
                        if (FileName != "")
                        {

                            var file = new FileInfo(stFilePathAndName);
                            string test = rs.Fields.Item("AttachPath").Value.ToString();
                            file.CopyTo(Path.Combine(rs.Fields.Item("AttachPath").Value.ToString(), file.Name), true);
                            if (Matrix3.RowCount == 0)
                            {
                                Matrix3.AddRow();
                            }
                            else
                            {
                                if (((SAPbouiCOM.EditText)Matrix3.GetCellSpecific("Path", Matrix3.RowCount)).Value.ToString() != "" /*&&
                                    ((SAPbouiCOM.EditText)Matrix3.GetCellSpecific("11", Matrix3.RowCount)).Value.ToString() != ""*/)
                                    Matrix3.AddRow();
                            }
                            ((SAPbouiCOM.EditText)Matrix3.GetCellSpecific("#", Matrix3.RowCount)).Value = Matrix3.RowCount.ToString();
                            ((SAPbouiCOM.EditText)Matrix3.GetCellSpecific("Path", Matrix3.RowCount)).Value = rs.Fields.Item("AttachPath").Value.ToString() + file.Name;
                            //((SAPbouiCOM.EditText)Matrix3.GetCellSpecific("11", Matrix3.RowCount)).Value = FileName;
                            var DocDate = DateTime.Parse(DateTime.Now.ToString());
                            var Date = DocDate.Year + DocDate.Month.ToString("00") + DocDate.Day.ToString("00");
                            ((SAPbouiCOM.EditText)Matrix3.GetCellSpecific("AttchDate", Matrix3.RowCount)).Value = Date;
                            stFilePathAndName = "";

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_VIEW_MODE)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }

                            Matrix3.AutoResizeColumns();
                        }
                    }

                }
            }
            catch (Exception ex)
            {
            }
        }

        private void Button4_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            for (int i = 1; i <= Matrix3.RowCount; i++)
            {
                if (Matrix3.IsRowSelected(i))
                {
                    Matrix3.DeleteRow(i);
                    if (UIAPIRawForm.Mode == SAPbouiCOM.BoFormMode.fm_VIEW_MODE || UIAPIRawForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    {
                        UIAPIRawForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }
                    break;
                }
            }

        }

        private void txtCode_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.ChooseFromList oCFLEvento = default(SAPbouiCOM.ChooseFromList);
                SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                oCFLEvento = this.UIAPIRawForm.ChooseFromLists.Item("Supplier");
                oCFLEvento.SetConditions(null);
                oCons = oCFLEvento.GetConditions();
                oCon = oCons.Add();
                oCon.Alias = "CardType";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "S";
                oCFLEvento.SetConditions(oCons);
            }
            catch (Exception) { }
        }

        private void txtCode_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg pCFL = ((SAPbouiCOM.ISBOChooseFromListEventArg)pVal);
                if (pCFL.SelectedObjects != null)
                {
                    SAPbouiCOM.DataTable oTable = pCFL.SelectedObjects;
                    try
                    {
                        txtCode.Value = oTable.GetValue("CardCode", 0).ToString();
                    }
                    catch { }
                    try
                    {
                        txtName.Value = oTable.GetValue("CardName", 0).ToString();
                    }
                    catch { }
                }
                string query = $@"SELECT T0.""Address"", T0.""Address"" FROM CRD1 T0 WHERE T0.""CardCode"" ='{txtCode.Value}'";
                B1Helper.SAPFillComboValues(this.ComboBox0, query);
            }
            catch (Exception) { }
        }

        private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            string query = $@"SELECT IFNULL(T1.""Street"", '') || ', ' || IFNULL(T1.""City"", '') || ', ' || IFNULL(T1.""Country"", '') AS ""AddressName"" FROM OCRD T0 INNER JOIN CRD1 T1 ON T0.""CardCode"" = T1.""CardCode"" WHERE T0.""CardCode"" = '{txtCode.Value}' AND T1.""Address"" = '{ComboBox0.Value}'";
            Recordset rec = ExecuteQuery(query);
            if (rec != null)
                txtPayTo.Value = rec.Fields.Item("AddressName").Value.ToString();
        }

        private void CreateOutgoingPayment()
        {
            B1Helper.DiCompany.GetBusinessObject(BoObjectTypes.oVendorPayments);
        }

        private void Matrix0_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID == "Selected")
                {
                    SAPbouiCOM.ISBOChooseFromListEventArg pCFL = ((SAPbouiCOM.ISBOChooseFromListEventArg)pVal);
                    if (pCFL.SelectedObjects != null)
                    {
                        // Selected invoices
                        List<string> selected = new List<string>();
                        for (int i = 1; i <= Matrix0.VisualRowCount; i++)
                        {
                            string data = Matrix0.GetCellValue("Selected", i).ToString();
                            if (!string.IsNullOrEmpty(data))
                                selected.Add(data);
                        }

                        // CFL datatable
                        SAPbouiCOM.DataTable oTable = pCFL.SelectedObjects;
                        string docEntry = "";
                        int matrixRow = pVal.Row;

                        // CFL selected rows
                        for (int j = 0; j < oTable.Rows.Count; j++)
                        {
                            string current = oTable.GetValue("DocEntry", j).ToString();
                            if (selected.Contains(current))
                            {
                                Program.SBO_Application.StatusBar.SetText($"Invoice: {oTable.GetValue("DocNum", j).ToString()} is already selected", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                continue;
                            }
                            try
                            {
                                docEntry = oTable.GetValue("DocEntry", j).ToString();
                                Matrix0.SetCellValue("Selected", matrixRow, docEntry);
                            }
                            catch (Exception ex) { }

                            Matrix0.SetCellValue("DocNum", matrixRow, oTable.GetValue("DocNum", j).ToString());
                            Matrix0.SetCellValue("PDate", matrixRow, ((DateTime)oTable.GetValue("DocDate", j)).ToString("yyyyMMdd"));
                            Matrix0.SetCellValue("DueDate", matrixRow, ((DateTime)oTable.GetValue("DocDueDate", j)).ToString("yyyyMMdd"));
                            Matrix0.SetCellValue("Overdays", matrixRow, (DateTime.Now - (DateTime)oTable.GetValue("DocDueDate", j)).Days);
                            string curr = oTable.GetValue("DocCur", j).ToString();
                            Matrix0.SetCellValue("Curr", matrixRow, curr);

                            double dueBalance = Convert.ToDouble(oTable.GetValue("DocTotal", j).ToString()) - Convert.ToDouble(oTable.GetValue("PaidToDate", j).ToString());

                            Matrix0.SetCellValue("Total", matrixRow, dueBalance);
                            Matrix0.SetCellValue("PayAmount", matrixRow, dueBalance);

                            if (!string.IsNullOrEmpty(Matrix0.GetCellValue("Selected", Matrix0.RowCount).ToString()))
                            {
                                Matrix0.AddLine();
                                Extentions.SetLineId(Matrix0);
                            }
                            LoadAttachment(docEntry);
                            matrixRow++;
                        }
                    }
                }
            }
            catch (Exception) { }
        }

        private void ResetRec()
        {
            if (Rec != null)
                Rec = null;
            if (Rec == null)
                Rec = (Recordset)Program.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
        }

        private void Matrix0_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (string.IsNullOrEmpty(txtCode.Value))
                {
                    Program.SBO_Application.SetStatusBarMessage("Please select supplier", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    BubbleEvent = false;
                    return;
                }

                SAPbouiCOM.ChooseFromList oCFLEvento = default(SAPbouiCOM.ChooseFromList);
                SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                oCFLEvento = this.UIAPIRawForm.ChooseFromLists.Item("Invoice");
                oCFLEvento.SetConditions(null);
                oCons = oCFLEvento.GetConditions();
                oCon = oCons.Add();
                oCon.Alias = "CardCode";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = txtCode.Value;
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCon = oCons.Add();
                oCon.Alias = "DocStatus";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "O";

                //HashSet<string> DocEntry = new HashSet<string>();
                //for (int i = 1; i <= Matrix0.VisualRowCount; i++)
                //{
                //    string data = Matrix0.GetCellValue("Selected", i).ToString();
                //    if (!string.IsNullOrEmpty(data))
                //        DocEntry.Add(data);
                //}

                //foreach (var item in DocEntry)
                //{
                //    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                //    oCon = oCons.Add();
                //    oCon.Alias = "DocEntry";
                //    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                //    oCon.CondVal = item;
                //}
                oCFLEvento.SetConditions(oCons);
            }
            catch (Exception) { }
        }

        private void LoadAttachment(string docEntry)
        {
            string query = $@"SELECT T0.""DocEntry"",T0.""DocNum"", T1.""trgtPath"", T1.""FreeText"", TO_VARCHAR(T1.""Date"",'yyyyMMdd') ""Date"" FROM OPCH T0 INNER JOIN ATC1 T1 ON  T0.""AtcEntry"" = T1.""AbsEntry"" WHERE T0.""DocEntry"" = {docEntry}";
            ResetRec();
            Rec.DoQuery(query);
            while (!Rec.EoF)
            {
                int row = Matrix1.VisualRowCount;
                Matrix1.SetCellValue("DocEntry", row, Rec.Fields.Item("DocEntry").Value.ToString());
                Matrix1.SetCellValue("DocNum", row, Rec.Fields.Item("DocNum").Value.ToString());
                Matrix1.SetCellValue("Name", row, Rec.Fields.Item("FreeText").Value.ToString());
                var date = Rec.Fields.Item("Date").Value.ToString();
                Matrix1.SetCellValue("Date", row, Rec.Fields.Item("Date").Value.ToString());
                Matrix1.SetCellValue("FilePath", row, Rec.Fields.Item("trgtPath").Value.ToString());
                Matrix1.AddLine();
                Extentions.SetLineId(Matrix1);
                Rec.MoveNext();
            }
        }

        private void RemoveAttachment(string docentry)
        {
            for (int i = 1; i <= Matrix1.VisualRowCount; i++)
            {
                var apdocentry = ((SAPbouiCOM.EditText)Matrix1.Columns.Item("DocEntry").Cells.Item(i).Specific).Value;
                if (docentry == apdocentry)
                {
                    Matrix1.DeleteRow(i);
                    i--;
                }
            }
            for (int i = 1; i <= Matrix1.VisualRowCount; i++)
            {
                ((SAPbouiCOM.EditText)Matrix1.Columns.Item("#").Cells.Item(i).Specific).Value = Convert.ToString(i);
            }
        }

        private void Matrix0_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {

            if (pVal.ColUID == "PayAmount" && pVal.ItemChanged == true)
            {
                CalcualteTotal();
            }
        }

        private void CalcualteTotal()
        {
            double sum = 0;
            for (int i = 1; i <= Matrix0.VisualRowCount; i++)
            {
                double total = Convert.ToDouble(((SAPbouiCOM.EditText)Matrix0.Columns.Item("PayAmount").Cells.Item(i).Specific).Value);
                sum += total;
            }
            txtTPA.Value = sum.ToString();
        }

        private void txtTPA_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
        }

        private void Form_RightClickBefore(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (eventInfo.ItemUID == "Item_0")
            {

                UIAPIRawForm.EnableMenu("1292", true);
                UIAPIRawForm.EnableMenu("1293", true);
                //oForm.EnableMenu("5377", false);
                //oForm.EnableMenu("1284", false);
                //oForm.EnableMenu("1286", false);

            }
            else
            {
                UIAPIRawForm.EnableMenu("1292", false);
                UIAPIRawForm.EnableMenu("1293", false);
            }
            UIAPIRawForm.EnableMenu("1283", false);



        }
    }
}
