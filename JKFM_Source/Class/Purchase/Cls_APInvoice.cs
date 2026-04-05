using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JKFM_Source
{
    public class Cls_APInvoice
    {
        #region Declaration
        SAPbouiCOM.Form objform, oForm;
        SAPbouiCOM.Grid oGrid;
        SAPbouiCOM.DataTable oDt;
        SAPbouiCOM.ComboBox oCombo, oCombo1, oCombo2;
        SAPbouiCOM.Matrix oMatrix;
        SAPbobsCOM.Documents oMemo;
        string DocEntry = "";
        string DraftNum = "";
        string BaseType = "";
        string Entry = "";
        string DocType = "";
        string DocStatus = "";
        string CardCode = "";
        public string FormType = "";
        string sErrMsg;
        int lErrCode;
        int lRetCode;
        ClsSBO objSBOAPI;

        //Transport Details
        string Success = "";
        //string Cancel = "N";

        #endregion

        #region Constructor
        public Cls_APInvoice(ClsSBO objSBO)
        {
            objSBOAPI = objSBO;
        }
        #endregion

        #region Item Event
        public void itemevent(string formuid, ref SAPbouiCOM.ItemEvent pval, ref bool bubbleevent)
        {
            try
            {
                if (pval.BeforeAction == true)
                {
                    objform = objSBOAPI.SBO_Appln.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount);
                    switch (pval.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            switch (pval.ItemUID)
                            {
                                case "1":
                                    if (pval.FormMode == Convert.ToInt32(SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                                    {
                                        CardCode = objform.Items.Item("4").Specific.Value;
                                        oCombo1 = objform.Items.Item("3").Specific;
                                        if (oCombo1.Selected.Value == "I")
                                        {
                                            oMatrix = objform.Items.Item("38").Specific;
                                            BaseType = oMatrix.Columns.Item("43").Cells.Item(1).Specific.Value.ToString();
                                            Entry = oMatrix.Columns.Item("45").Cells.Item(1).Specific.Value.ToString();
                                        }
                                        else if (oCombo1.Selected.Value == "S")
                                        {
                                            oMatrix = objform.Items.Item("39").Specific;
                                            BaseType = oMatrix.Columns.Item("23").Cells.Item(1).Specific.Value.ToString();
                                            Entry = oMatrix.Columns.Item("25").Cells.Item(1).Specific.Value.ToString();
                                        }
                                        Add_Freight();
                                        //Transport Details
                                        //if (Cancel != "Y")
                                        //{
                                        //    Validation(ref bubbleevent);
                                        //    if (Success == "Y")
                                        //    {
                                        //        oCombo = objform.Items.Item("Cb_FCon").Specific;
                                        //        if (oCombo.Selected.Value != "DIRECT")
                                        //        {
                                        //            Add_Freight();
                                        //        }
                                        //    }
                                        //}
                                    }
                                    break;
                            }
                            break;

                        case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:

                            Adding_Items();
                            objform.Freeze(true);
                            if (pval.FormMode == Convert.ToInt32(SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                            {
                                SAPbouiCOM.Matrix oMatrix;
                                oCombo1 = objform.Items.Item("3").Specific;
                                if (oCombo1.Selected.Value == "I")
                                {
                                    oMatrix = objform.Items.Item("38").Specific;
                                    string Entry = oMatrix.Columns.Item("45").Cells.Item(1).Specific.Value.ToString();
                                    if (Entry == "")
                                    {
                                        DraftNum = objSBOAPI.Query_Execute("Select \"DocEntry\" from \"ODRF\" where \"DocStatus\" = 'O' AND \"ObjType\" = '18' AND \"U_AVA_PURCHASETOKENNO\" = (Select TOP 1 IFNULL(\"U_AVA_PURCHASETOKENNO\",0) from \"OPDN\" where \"DocEntry\" = NULL)");
                                    }
                                    else
                                    {
                                        DraftNum = objSBOAPI.Query_Execute("Select \"DocEntry\" from \"ODRF\" where \"DocStatus\" = 'O' AND \"ObjType\" = '18' AND \"U_AVA_PURCHASETOKENNO\" = (Select TOP 1 IFNULL(\"U_AVA_PURCHASETOKENNO\",0) from \"OPDN\" where \"DocEntry\" = '" + Entry + "')");
                                    }

                                    if (DraftNum != "")
                                    {
                                        string TokNo = objSBOAPI.Query_Execute("Select \"U_AVA_PURCHASETOKENNO\" from \"ODRF\" where \"DocEntry\" = '" + DraftNum + "'");
                                        Lab_Grid_Loading(TokNo, "", DraftNum);
                                    }
                                }
                                else if (oCombo1.Selected.Value == "S")
                                {
                                    oMatrix = objform.Items.Item("39").Specific;
                                    string Entry = oMatrix.Columns.Item("25").Cells.Item(1).Specific.Value.ToString();
                                    Deduction_Grid_Loading("", "", "");
                                }

                                //Transport Details
                                //oCombo1 = objform.Items.Item("Et_TrNme").Specific;
                                //SAPbobsCOM.Recordset oRec = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                //string Transporter = "Select DISTINCT \"U_AVA_TRANSPORTER\" from \"@AVA_OSTM\"";
                                //oRec.DoQuery(Transporter);
                                //if (oRec.RecordCount > 0)
                                //{
                                //    for (int i = 0; i < oRec.RecordCount; i++)
                                //    {
                                //        oCombo1.ValidValues.Add(oRec.Fields.Item("U_AVA_TRANSPORTER").Value, oRec.Fields.Item("U_AVA_TRANSPORTER").Value);
                                //        oRec.MoveNext();
                                //    }
                                //    oCombo1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                                //}
                                objform.Freeze(false);
                            }
                            break;

                        //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:

                        //    Cancel = "N";
                        //    break;

                    }
                }
                else
                {
                    switch (pval.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            switch (pval.ItemUID)
                            {
                                case "Tb_LD":
                                    objform.PaneLevel = 50;
                                    if (pval.FormMode == Convert.ToInt32(SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
                                    {
                                        objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    }
                                    break;

                                case "Tb_DD":
                                    objform.PaneLevel = 51;
                                    objform.Items.Item("St_Total").Visible = true;
                                    objform.Items.Item("Et_Total").Visible = true;
                                    objform.Items.Item("Et_Total").Enabled = false;
                                    if (pval.FormMode == Convert.ToInt32(SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
                                    {
                                        objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                        objform.Items.Item("Et_Total").Enabled = false;
                                    }
                                    break;

                                case "GRD2":
                                    if (pval.ColUID == "CK_Select")
                                    {
                                        Total_Deduction();
                                    }
                                    break;

                                case "1":
                                    if (pval.FormMode == Convert.ToInt32(SAPbouiCOM.BoFormMode.fm_ADD_MODE) && pval.Action_Success == true)
                                    {
                                        //DocStatus = objSBOAPI.Query_Execute("Select \"DocStatus\" from \"OPCH\" where \"DocEntry\" = '" + DocEntry + "'");
                                        //if (DocStatus == "O")
                                        //{
                                            if (BaseType == "20")
                                            {
                                                DocType = objSBOAPI.Query_Execute("Select \"DocType\" from \"OPCH\" where \"DocEntry\" = '" + DocEntry + "'");
                                                DraftNum = objSBOAPI.Query_Execute("Select \"DocEntry\" from \"ODRF\" where \"DocStatus\" = 'O' AND \"ObjType\" = '18' AND \"U_AVA_PURCHASETOKENNO\" = (Select TOP 1 IFNULL(\"U_AVA_PURCHASETOKENNO\",0) from \"OPDN\" where \"DocEntry\" = '" + Entry + "')");
                                                if (DraftNum == "")
                                                {
                                                    string Count = objSBOAPI.Query_Execute("Select \"DocEntry\" from \"OPCH\" where \"DocEntry\" = '" + DocEntry + "' AND\"U_AVA_PURCHASETOKENNO\" = (Select TOP 1 IFNULL(\"U_AVA_PURCHASETOKENNO\",0) from \"OPDN\" where \"DocEntry\" = '" + Entry + "')");
                                                    if (Count != "")
                                                    {
                                                        Insert_Into_Table(DocEntry);
                                                        APCreditMemo_Creation(DocEntry);
                                                    }
                                                }
                                                else
                                                {
                                                    if (DocEntry != "")
                                                    {
                                                        Insert_Into_Temp(DocEntry);
                                                    }
                                                    else
                                                    {
                                                        Insert_Into_Temp(DraftNum);
                                                    }
                                                }
                                            }
                                            else if (BaseType == "22")
                                            {
                                                DraftNum = objSBOAPI.Query_Execute("Select \"DocEntry\" from \"ODRF\" where \"DocStatus\" = 'O' AND \"ObjType\" = '18' AND \"U_AVA_PURCHASETOKENNO\" = (Select TOP 1 IFNULL(\"U_AVA_PURCHASETOKENNO\",0) from \"OPDN\" where \"DocEntry\" = '" + Entry + "')");
                                                if (DraftNum == "")
                                                {
                                                    DocType = objSBOAPI.Query_Execute("Select \"DocType\" from \"OPCH\" where \"DocEntry\" = '" + DocEntry + "'");
                                                    if (DocType == "S")
                                                    {
                                                        oGrid = objform.Items.Item("GRD2").Specific;
                                                        for (int i = 0; i < oGrid.DataTable.Rows.Count; i++)
                                                        {
                                                            if (oGrid.DataTable.Columns.Item("CK_Select").Cells.Item(i).Value == "Y")
                                                            {
                                                                string Name = oGrid.DataTable.Columns.Item("NE_Deduction").Cells.Item(i).Value.ToString();
                                                                double Debit = Convert.ToDouble(oGrid.DataTable.Columns.Item("E_Debit").Cells.Item(i).Value);
                                                                double Credit = Convert.ToDouble(oGrid.DataTable.Columns.Item("E_Credit").Cells.Item(i).Value);
                                                                string Ledger = oGrid.DataTable.Columns.Item("INV_Ledger").Cells.Item(i).Value.ToString();
                                                                double Total = Convert.ToDouble(objform.Items.Item("Et_Total").Specific.Value);

                                                                string Code = objSBOAPI.Query_Execute("Select IFNULL(MAX(\"Code\"),0)+1 from \"@AVA_DED2\"");
                                                                objSBOAPI.Query_Execute("INSERT INTO \"@AVA_DED2\" (\"Code\",\"Name\",\"U_AVA_DedName\",\"U_AVA_DocEntry\",\"U_AVA_Debit\",\"U_AVA_Credit\",\"U_AVA_Ledger\",\"U_AVA_Total\") VALUES('" + Code + "','" + objSBOAPI.oCompany.UserName + "','" + Name + "','" + DocEntry + "','" + Debit + "','" + Credit + "','" + Ledger + "','" + Total + "')");

                                                            }
                                                        }
                                                        string GCode = objSBOAPI.Query_Execute("Select \"GroupCode\" from \"OCRD\" where \"CardCode\" = '" + CardCode + "'");
                                                        SAPbobsCOM.Recordset oRec = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                        string Group = objSBOAPI.Query_Execute("SELECT \"QString\" FROM OUQR WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM OQCN WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='APCreditMemo'");
                                                        if (Group != "")
                                                        {
                                                            Group = Group.Replace("[%1]", "GroupCode");
                                                            oRec.DoQuery(Group);
                                                            if (oRec.RecordCount > 0)
                                                            {
                                                                if (GCode == oRec.Fields.Item("GroupCode").Value.ToString())
                                                                {
                                                                    string Qty = objSBOAPI.Query_Execute("SELECT \"QString\" FROM OUQR WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM OQCN WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='APCreditMemo'");
                                                                    Qty = Qty.Replace("[%1]", "Quantity");
                                                                    Qty = Qty.Replace("[%2]", DocEntry);
                                                                    oRec.DoQuery(Qty);
                                                                    if (oRec.RecordCount > 0)
                                                                    {
                                                                        APCreditMemo_Creation(DocEntry);
                                                                    }
                                                                    else
                                                                    {
                                                                        objSBOAPI.SBO_Appln.StatusBar.SetText("AP Credit Memo is not created for : " + DocEntry + " because the deduction Amount is zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    objSBOAPI.SBO_Appln.StatusBar.SetText("AP Credit Memo is not created for : " + DocEntry + " because the Vendor is not a transporter Group", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                objSBOAPI.SBO_Appln.StatusBar.SetText("GroupCode not found for Vendor of Transporter Type", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                                            }
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (DocEntry != "")
                                                    {
                                                        Insert_Into_Temp(DocEntry);
                                                    }
                                                    else
                                                    {
                                                        Insert_Into_Temp(DraftNum);
                                                    }
                                                }
                                            }
                                        //}
                                        
                                    }
                                    break;

                                //Transport Details
                                case "Lk_InsNo":

                                    objSBOAPI.SBO_Appln.ActivateMenuItem("AV_INSDTM");
                                    SAPbouiCOM.Form oForm1 = objSBOAPI.SBO_Appln.Forms.ActiveForm;
                                    oForm1.Freeze(true);
                                    oForm1.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                    SAPbouiCOM.ComboBox oCombo2 = oForm1.Items.Item("Cb_Code").Specific;
                                    oForm1.Items.Item("Cb_Code").Enabled = true;
                                    oCombo2.Select(objform.Items.Item("Et_InsNo").Specific.Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                    oForm1.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    oForm1.Items.Item("Cb_Code").Enabled = false;
                                    oForm1.Freeze(false);

                                    break;
                            }
                            break;

                        case SAPbouiCOM.BoEventTypes.et_VALIDATE:

                            if (pval.ItemUID == "GRD2" && (pval.ColUID == "E_Debit" || pval.ColUID == "E_Credit"))
                            {
                                Total_Deduction();
                            }
                            break;

                        case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:

                            if (pval.ItemUID == "10000330")
                            {
                                oCombo = objform.Items.Item("10000330").Specific;
                                if (oCombo.Selected.Value == "Goods Receipt PO")
                                {
                                    FormType = objform.TypeEx;
                                }
                                else if (oCombo.Selected.Value == "Purchase Orders")
                                {
                                    Deduction_Grid_Loading("", "", "");
                                }
                            }
                            //if (pval.ItemUID == "Et_TrNme" && pval.ItemChanged == true)
                            //{
                            //    oCombo1 = objform.Items.Item("Et_TrNme").Specific;
                            //    oCombo2 = objform.Items.Item("Cb_VehNo").Specific;
                            //    int count = oCombo2.ValidValues.Count;
                            //    for (int i = 1; i <= count; i++)
                            //    {
                            //        oCombo2.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            //    }
                            //    if (oCombo2.ValidValues.Count == 0)
                            //    {
                            //        string VehNo = "Select \"U_AVA_VEHICLENO\" from \"@AVA_OSTM\" where \"U_AVA_TRANSPORTER\" = '" + oCombo1.Selected.Value + "'";
                            //        SAPbobsCOM.Recordset oRec = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            //        oRec.DoQuery(VehNo);
                            //        if (oRec.RecordCount > 0)
                            //        {
                            //            for (int i = 0; i < oRec.RecordCount; i++)
                            //            {
                            //                oCombo2.ValidValues.Add(oRec.Fields.Item("U_AVA_VEHICLENO").Value, oRec.Fields.Item("U_AVA_VEHICLENO").Value);
                            //                oRec.MoveNext();
                            //            }
                            //            oCombo2.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                            //        }
                            //    }
                            //}
                            //else if (pval.ItemUID == "Cb_VehNo" && pval.ItemChanged == true)
                            //{
                            //    oCombo2 = objform.Items.Item("Cb_VehNo").Specific;
                            //    string VehNo = "Select \"U_AVA_DRIVERMOBILENO\",\"U_AVA_DRIVERLNO\",IFNULL(\"U_AVA_OWNTRUCK\",'N') AS \"U_AVA_OWNTRUCK\" from \"@AVA_OSTM\" where \"U_AVA_VEHICLENO\" = '" + oCombo2.Selected.Value + "'";
                            //    SAPbobsCOM.Recordset oRec = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            //    oRec.DoQuery(VehNo);
                            //    if (oRec.RecordCount > 0)
                            //    {
                            //        objform.Items.Item("Et_MobNo").Specific.Value = oRec.Fields.Item("U_AVA_DRIVERMOBILENO").Value;
                            //        objform.Items.Item("Et_LicNo").Specific.Value = oRec.Fields.Item("U_AVA_DRIVERLNO").Value;
                            //        oCombo1 = objform.Items.Item("Cb_Trck").Specific;
                            //        oCombo1.Select(oRec.Fields.Item("U_AVA_OWNTRUCK").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                            //    }
                            //}
                            break;

                        //Transport Details
                        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                            if (pval.ItemUID == "Et_InsNo")
                            {
                                SAPbouiCOM.DataTable dt;
                                SAPbouiCOM.ChooseFromListEvent cfl;
                                cfl = (SAPbouiCOM.ChooseFromListEvent)pval;
                                dt = cfl.SelectedObjects;
                                if (dt != null)
                                {
                                    try
                                    {
                                        objform.Items.Item("Et_InsNo").Specific.value = dt.GetValue("Code", 0).ToString();
                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                    objform.Items.Item("Et_InsBal").Specific.value = dt.GetValue("U_AVA_BalAmt", 0).ToString();
                                }
                            }
                            break;

                        //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:

                        //    switch (pval.ItemUID)
                        //    {
                        //        case "Et_FAmt":

                        //            if (objform.Items.Item("Et_FAmt").Specific.Value != "")
                        //            {
                        //                if (Convert.ToDouble(objform.Items.Item("Et_FAmt").Specific.Value) != 0)
                        //                {
                        //                    SAPbouiCOM.Matrix oMatrix1 = objform.Items.Item("38").Specific;
                        //                    double TotQty = 0.0;
                        //                    for (int i = 1; i <= oMatrix1.VisualRowCount; i++)
                        //                    {
                        //                        if (oMatrix1.Columns.Item("1").Cells.Item(i).Specific.Value != "")
                        //                        {
                        //                            string Qty = oMatrix1.Columns.Item("11").Cells.Item(i).Specific.Value;
                        //                            TotQty = TotQty + Convert.ToDouble(Qty);
                        //                        }
                        //                    }
                        //                    double Rate = Convert.ToDouble(objform.Items.Item("Et_FAmt").Specific.Value) / TotQty;
                        //                    objform.Items.Item("Et_FRate").Specific.Value = Rate;
                        //                }
                        //            }
                        //            break;

                        //        case "Et_FRate":

                        //            if (objform.Items.Item("Et_FRate").Specific.Value != "")
                        //            {
                        //                if (Convert.ToDouble(objform.Items.Item("Et_FRate").Specific.Value) != 0)
                        //                {
                        //                    SAPbouiCOM.Matrix oMatrix1 = objform.Items.Item("38").Specific;
                        //                    double TotQty = 0.0;
                        //                    for (int i = 1; i <= oMatrix1.VisualRowCount; i++)
                        //                    {
                        //                        if (oMatrix1.Columns.Item("1").Cells.Item(i).Specific.Value != "")
                        //                        {
                        //                            string Qty = oMatrix1.Columns.Item("11").Cells.Item(i).Specific.Value;
                        //                            TotQty = TotQty + Convert.ToDouble(Qty);
                        //                        }
                        //                    }
                        //                    double Amount = Convert.ToDouble(objform.Items.Item("Et_FRate").Specific.Value) * TotQty;
                        //                    objform.Items.Item("Et_FAmt").Specific.Value = Amount;
                        //                }
                        //            }
                        //            break;

                        //        case "Et_FAdv":

                        //            if (objform.Items.Item("Et_FAdv").Specific.Value != "")
                        //            {
                        //                objform.Items.Item("Et_FBal").Specific.value = Convert.ToDouble(objform.Items.Item("Et_FAmt").Specific.Value) - Convert.ToDouble(objform.Items.Item("Et_FAdv").Specific.Value);
                        //            }
                        //            break;
                        //    }
                        //    break;

                    }
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Menu Event
        public void MenuEvent(ref SAPbouiCOM.MenuEvent pval, ref bool BubbleEvent)
        {
            try
            {
                if (pval.BeforeAction == true)
                {
                    switch (pval.MenuUID)
                    {
                        //case "1284":
                        //    Cancel = "Y";
                        //    break;
                    }
                }
                else if (pval.BeforeAction == false)
                {
                    switch (pval.MenuUID)
                    {
                        case "1282":

                            objform.Items.Item("Et_Total").Specific.Value = 0;
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region FORM DATA EVENT
        public void FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            if (BusinessObjectInfo.ActionSuccess == true)
                            {
                                string MyXml = BusinessObjectInfo.ObjectKey;
                                System.Xml.XmlDocument MyDoc = new System.Xml.XmlDocument();
                                MyDoc.LoadXml(MyXml);
                                DocEntry = MyDoc.SelectSingleNode("//DocumentParams/DocEntry").InnerText;
                                //DocStatus = objSBOAPI.Query_Execute("Select \"DocStatus\" from \"OPCH\" where \"DocEntry\" = '" + DocEntry + "'");
                                //if (DocStatus == "O")
                                //{
                                    //oCombo = objform.Items.Item("Cb_FCon").Specific;
                                    //if (oCombo.Selected.Value != "DIRECT")
                                    //{
                                    //    if (Cancel != "Y")
                                    //    {
                                    //        PO_Creation(DocEntry);
                                    //    }
                                    Insurance_Update(DocEntry);
                                    //}
                                    //Cancel = "N";
                                //}
                            }
                            break;

                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:

                            DocEntry = objform.DataSources.DBDataSources.Item("OPCH").GetValue("DocEntry", 0).ToString();
                            Lab_Grid_Loading("", DocEntry, "");
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                objSBOAPI.SBO_Appln.MessageBox(ex.Message);
            }
        }

        #endregion

        #region Adding Items
        public void Adding_Items()
        {
            try
            {
                objform.Freeze(false);

                SAPbouiCOM.Item oExtItem, oExtItem_GRD, oEdit, oStatic;

                oExtItem = objform.Items.Item("1320002137");
                objform.DataSources.UserDataSources.Add("UD_LD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                objSBOAPI.Adding_Items_Folder("Tb_LD", oExtItem.Width, oExtItem.Top, oExtItem.Left + oExtItem.Width, oExtItem.Height, 50, "112", "Lab Deduction", "UD_LD", objform.UniqueID);

                objform.DataSources.UserDataSources.Add("UD_DD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                objSBOAPI.Adding_Items_Folder("Tb_DD", oExtItem.Width, oExtItem.Top, oExtItem.Left + oExtItem.Width, oExtItem.Height, 51, "Tb_LD", "Deduction Details", "UD_DD", objform.UniqueID);

                objform.DataSources.DataTables.Add("DT_0");
                objform.DataSources.DataTables.Add("DT_1");

                oExtItem_GRD = objform.Items.Item("38");
                objSBOAPI.Adding_Items_Grid("GRD1", oExtItem_GRD.Width, oExtItem_GRD.Top, oExtItem_GRD.Left, oExtItem_GRD.Height, objform.UniqueID, 50, 50);
                objSBOAPI.Adding_Items_Grid("GRD2", oExtItem_GRD.Width, oExtItem_GRD.Top, oExtItem_GRD.Left, oExtItem_GRD.Height, objform.UniqueID, 51, 51);

                oStatic = objform.Items.Item("254000067");
                objSBOAPI.Adding_Items_Static("St_Total", oStatic.Width, oStatic.Top + 20, oStatic.Left + 5, oStatic.Height, "Total Deduction", objform.UniqueID, 51, 51, false);

                oEdit = objform.Items.Item("254000068");
                objform.DataSources.UserDataSources.Add("UD_Total", SAPbouiCOM.BoDataType.dt_PRICE);
                objSBOAPI.Adding_Items_Edit("Et_Total", oEdit.Width, oEdit.Top + 20, oEdit.Left + 5, oEdit.Height, "", "UD_Total", objform.UniqueID, 51, 51, false, false);

                //Transport Details
                SAPbouiCOM.Item oEdit2, oText2, oLink;
                //SAPbouiCOM.Item oEdits, oText, oEdit1, oText1, oEdit2, oText2, oLink, oEdit3, oText3, oChk;
                //oEdits = objform.Items.Item("46");
                //oText = objform.Items.Item("86");
                //oEdit1 = objform.Items.Item("254000015");
                //oText1 = objform.Items.Item("254000012");
                oEdit2 = objform.Items.Item("222");
                oText2 = objform.Items.Item("230");
                //oEdit3 = objform.Items.Item("33");
                //oText3 = objform.Items.Item("34");
                oLink = objform.Items.Item("51");
                //oChk = objform.Items.Item("1320002111");

                //objSBOAPI.Adding_Items_Static("St_FRate", oText.Width, oText.Top + 15, oText.Left, oText.Height, "Freight Rate", objform.UniqueID);
                //objSBOAPI.Adding_Items_Edit("Et_FRate", oEdits.Width, oEdits.Top + 15, oEdits.Left, oEdits.Height, "OPCH", "U_AVA_FREIGHTRATE", objform.UniqueID);

                //objSBOAPI.Adding_Items_Static("St_FAmt", oText.Width, oText.Top + 30, oText.Left, oText.Height, "Freight Amount", objform.UniqueID);
                //objSBOAPI.Adding_Items_Edit("Et_FAmt", oEdits.Width, oEdits.Top + 30, oEdits.Left, oEdits.Height, "OPCH", "U_AVA_FREIGHTAMT", objform.UniqueID);

                //objSBOAPI.Adding_Items_Static("St_FAdv", oText.Width, oText.Top + 45, oText.Left, oText.Height, "Freight Advance", objform.UniqueID);
                //objSBOAPI.Adding_Items_Edit("Et_FAdv", oEdits.Width, oEdits.Top + 45, oEdits.Left, oEdits.Height, "OPCH", "U_AVA_FREIGHTADVANCE", objform.UniqueID);

                //objSBOAPI.Adding_Items_Static("St_FBal", oText.Width, oText.Top + 60, oText.Left, oText.Height, "Freight Balance", objform.UniqueID);
                //objSBOAPI.Adding_Items_Edit("Et_FBal", oEdits.Width, oEdits.Top + 60, oEdits.Left, oEdits.Height, "OPCH", "U_AVA_FREIGHTBALANCE", objform.UniqueID);

                //objSBOAPI.Adding_Items_Static("St_TrNme", oText1.Width, oText1.Top + 15, oText1.Left, oText1.Height, "Transporter Name", objform.UniqueID);
                //objSBOAPI.Adding_Items_Combo("Et_TrNme", oEdit1.Width, oEdit1.Top + 15, oEdit1.Left, oEdit1.Height, "OPCH", "U_AVA_TRANSNAME", objform.UniqueID);

                //objSBOAPI.Adding_Items_Static("St_FCon", oText1.Width, oText1.Top + 30, oText1.Left, oText1.Height, "Freight Condition", objform.UniqueID);
                //objSBOAPI.Adding_Items_Combo("Cb_FCon", oEdit1.Width, oEdit1.Top + 30, oEdit1.Left, oEdit1.Height, "OPCH", "U_AVA_FRTCONDITION", objform.UniqueID);

                objSBOAPI.Adding_Items_Static("St_InsNo", oText2.Width, oText2.Top + 15, oText2.Left, oText2.Height, "Insurance Policy No", objform.UniqueID);
                objSBOAPI.Adding_Items_Edit("Et_InsNo", oEdit2.Width, oEdit2.Top + 15, oEdit2.Left, oEdit2.Height, "OPCH", "U_AVA_INSPNO", objform.UniqueID);
                objSBOAPI.Adding_Items_Link_Button("Lk_InsNo", oLink.Width, oEdit2.Top + 17, oLink.Left, oLink.Height, objform.UniqueID, "Et_InsNo");

                objSBOAPI.Adding_Items_Static("St_InsBal", oText2.Width, oText2.Top + 30, oText2.Left, oText2.Height, "Insurance Balance", objform.UniqueID);
                objSBOAPI.Adding_Items_Edit("Et_InsBal", oEdit2.Width, oEdit2.Top + 30, oEdit2.Left, oEdit2.Height, "OPCH", "U_AVA_INSBAL", objform.UniqueID);

                //objSBOAPI.Adding_Items_Static("St_VehNo", oText2.Width, oText2.Top + 45, oText2.Left, oText2.Height, "Vehicle No", objform.UniqueID);
                //objSBOAPI.Adding_Items_Combo("Cb_VehNo", oEdit2.Width, oEdit2.Top + 45, oEdit2.Left, oEdit2.Height, "OPCH", "U_AVA_VEHICLENO", objform.UniqueID);

                //objSBOAPI.Adding_Items_Static("St_MobNo", oText2.Width, oText2.Top + 60, oText2.Left, oText2.Height, "Driver Mobile No", objform.UniqueID);
                //objSBOAPI.Adding_Items_Edit("Et_MobNo", oEdit2.Width, oEdit2.Top + 60, oEdit2.Left, oEdit2.Height, "OPCH", "U_AVA_DRIVERMOBILENO", objform.UniqueID);

                //objSBOAPI.Adding_Items_Static("St_LicNo", oText2.Width, oText2.Top + 75, oText2.Left, oText2.Height, "Driver License No", objform.UniqueID);
                //objSBOAPI.Adding_Items_Edit("Et_LicNo", oEdit2.Width, oEdit2.Top + 75, oEdit2.Left, oEdit2.Height, "OPCH", "U_AVA_DRIVERLNO", objform.UniqueID);

                //objSBOAPI.Adding_Items_Static("St_Trck", oText2.Width, oText2.Top + 90, oText2.Left, oText2.Height, "Own Truck", objform.UniqueID);
                //objSBOAPI.Adding_Items_Combo("Cb_Trck", oEdit2.Width, oEdit2.Top + 90, oEdit2.Left, oEdit2.Height, "OPCH", "U_AVA_OWNTRUCK", objform.UniqueID);

                //objSBOAPI.Adding_Items_Static("St_Spvr", oText3.Width, oText2.Top + 161, oText3.Left, oText3.Height, "Supervisor", objform.UniqueID);
                //objSBOAPI.Adding_Items_Edit("Et_Spvr", oEdit3.Width, oEdit2.Top + 161, oEdit3.Left, oEdit3.Height, "OPCH", "U_AVA_SUPERVISOR", objform.UniqueID);

                //objSBOAPI.Adding_Items_Static("St_Lpsn", oText3.Width, oText2.Top + 146, oText3.Left, oText3.Height, "Loading Person", objform.UniqueID);
                //objSBOAPI.Adding_Items_Edit("Et_Lpsn", oEdit3.Width, oEdit2.Top + 146, oEdit3.Left, oEdit3.Height, "OPCH", "U_AVA_LOADINGPERSON", objform.UniqueID);

                //oChk.Left = oEdit2.Left + oEdit2.Width + 5;

                SAPbouiCOM.ChooseFromListCollection oCFLs;
                oCFLs = objform.ChooseFromLists;
                SAPbouiCOM.ChooseFromList oCFL;
                SAPbouiCOM.ChooseFromListCreationParams oCFLCreate;
                oCFLCreate = objSBOAPI.SBO_Appln.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                oCFLCreate.MultiSelection = false;
                oCFLCreate.UniqueID = "CFL1";
                oCFLCreate.ObjectType = "AVA_INSURANCE";
                oCFL = oCFLs.Add(oCFLCreate);
                SAPbouiCOM.EditText oEdit4;
                oEdit4 = objform.Items.Item("Et_InsNo").Specific;
                oEdit4.ChooseFromListUID = "CFL1";
                oEdit4.ChooseFromListAlias = "Code";

                objform.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                objform.Freeze(true);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Lab Grid Loading
        public void Lab_Grid_Loading(string TokenNo, string DocEntry, string DraftNum)
        {
            try
            {
                objform.Freeze(true);
                oGrid = objform.Items.Item("GRD1").Specific;
                oDt = oGrid.DataTable;
                oGrid.DataTable = objform.DataSources.DataTables.Item("DT_0");
                oGrid.DataTable.Clear();
                string Load = objSBOAPI.Query_Execute("SELECT \"QString\" FROM OUQR WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM OQCN WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='Lab_Deduction'");
                if (Load != "")
                {
                    if (DocEntry != "")
                    {
                        Load = Load.Replace("[%3]", DocEntry);
                        Load = Load.Replace("[%2]", "Load");
                    }
                    else if (TokenNo != "")
                    {
                        Load = Load.Replace("[%1]", TokenNo);
                        Load = Load.Replace("[%2]", "Add");
                    }
                }
                else
                {
                    objform.Freeze(false);
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Query is Missing, Please Update Query Manager.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }

                oGrid.DataTable.ExecuteQuery(Load);
                if (oGrid.DataTable.IsEmpty == false)
                {
                    try
                    {
                        oGrid.Columns.Item("RowsHeader").Visible = false;

                        for (int i = 0; i <= oGrid.DataTable.Columns.Count - 1; i++)
                        {
                            string columncheck = "";
                            columncheck = oGrid.Columns.Item(i).TitleObject.Caption;

                            columncheck = oGrid.DataTable.Columns.Item(i).Name;

                            switch (columncheck.ToString().Split('_')[0])
                            {
                                case "NE":
                                    oGrid.Columns.Item(columncheck).Editable = false;
                                    oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("NE_", "");
                                    break;

                                case "E":
                                    oGrid.Columns.Item(columncheck).Editable = true;
                                    oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("E_", "");
                                    break;

                                case "EC":
                                    oGrid.Columns.Item(columncheck).Editable = true;
                                    oGrid.Columns.Item(columncheck).BackColor = -1;
                                    oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("EC_", "");
                                    break;

                                case "NEC":
                                    oGrid.Columns.Item(columncheck).Editable = false;
                                    oGrid.Columns.Item(columncheck).BackColor = -1;
                                    oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("NEC_", "");
                                    break;

                                case "INV":
                                    oGrid.Columns.Item(columncheck).Visible = false;
                                    oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("INV_", "");
                                    break;

                                case "CK":
                                    oGrid.Columns.Item(columncheck).Editable = true;
                                    oGrid.Columns.Item(columncheck).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                                    oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("CK_", "");
                                    break;

                                case "ECMB":
                                    oGrid.Columns.Item(columncheck).Editable = true;
                                    oGrid.Columns.Item(columncheck).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                                    oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("ECMB_", "");
                                    break;
                            }
                        }

                        SAPbouiCOM.EditTextColumn ColQty;
                        ColQty = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("NE_Deduction Amount");
                        ColQty.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                        objform.Freeze(false);

                    }
                    catch (Exception ex)
                    {
                        objform.Freeze(false);
                        objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                    oGrid.AutoResizeColumns();
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Lab Deduction Data are Loaded Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                    if (DraftNum != "")
                    {
                        Deduction_Grid_Loading("", "", DraftNum);
                    }
                    else if (TokenNo != "")
                    {
                        Deduction_Grid_Loading(TokenNo, "", "");
                    }
                    else if (DocEntry != "")
                    {
                        Deduction_Grid_Loading("", DocEntry, "");
                    }

                    if (TokenNo != "")
                    {
                        Total_Deduction();
                    }
                }
                else
                {
                    oGrid.DataTable.Clear();
                    objform.Freeze(false);
                    objSBOAPI.SBO_Appln.StatusBar.SetText("No Records Found for Lab Decuction", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                    if (DraftNum != "")
                    {
                        Deduction_Grid_Loading("", "", DraftNum);
                    }
                    else if (TokenNo != "")
                    {
                        Deduction_Grid_Loading(TokenNo, "", "");
                    }
                    else if (DocEntry != "")
                    {
                        Deduction_Grid_Loading("", DocEntry, "");
                    }

                    if (TokenNo != "")
                    {
                        Total_Deduction();
                    }
                }

                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Deduction Grid Loading
        public void Deduction_Grid_Loading(string TokenNo, string DocEntry, string DraftNum)
        {
            try
            {
                objform.Freeze(true);
                oGrid = objform.Items.Item("GRD2").Specific;
                oGrid.DataTable = objform.DataSources.DataTables.Item("DT_1");
                oDt = oGrid.DataTable;
                oGrid.DataTable.Clear();
                string Load = objSBOAPI.Query_Execute("SELECT \"QString\" FROM OUQR WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM OQCN WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='Deduction_Details'");
                if (Load != "")
                {
                    if (DocEntry != "")
                    {
                        Load = Load.Replace("[%3]", DocEntry);
                        Load = Load.Replace("[%2]", "Load");
                    }
                    else if (TokenNo != "")
                    {
                        Load = Load.Replace("[%1]", TokenNo);
                        Load = Load.Replace("[%2]", "Add");
                    }
                    else if (DraftNum != "")
                    {
                        Load = Load.Replace("[%3]", DraftNum);
                        Load = Load.Replace("[%2]", "Draft");
                    }
                    else
                    {
                        Load = Load.Replace("[%1]", "");
                        Load = Load.Replace("[%2]", "Add");
                    }
                }
                else
                {
                    objform.Freeze(false);
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Query is Missing, Please Update Query Manager.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }

                oGrid.DataTable.ExecuteQuery(Load);
                if (oGrid.DataTable.IsEmpty == false)
                {
                    try
                    {
                        oGrid.Columns.Item("RowsHeader").Visible = false;


                        for (int i = 0; i <= oGrid.DataTable.Columns.Count - 1; i++)
                        {
                            string columncheck = "";
                            columncheck = oGrid.Columns.Item(i).TitleObject.Caption;

                            columncheck = oGrid.DataTable.Columns.Item(i).Name;

                            switch (columncheck.ToString().Split('_')[0])
                            {
                                case "NE":
                                    oGrid.Columns.Item(columncheck).Editable = false;
                                    oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("NE_", "");
                                    break;

                                case "E":
                                    oGrid.Columns.Item(columncheck).Editable = true;
                                    oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("E_", "");
                                    break;

                                case "EC":
                                    oGrid.Columns.Item(columncheck).Editable = true;
                                    oGrid.Columns.Item(columncheck).BackColor = -1;
                                    oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("EC_", "");
                                    break;

                                case "NEC":
                                    oGrid.Columns.Item(columncheck).Editable = false;
                                    oGrid.Columns.Item(columncheck).BackColor = -1;
                                    oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("NEC_", "");
                                    break;

                                case "INV":
                                    oGrid.Columns.Item(columncheck).Visible = false;
                                    oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("INV_", "");
                                    break;

                                case "CK":
                                    oGrid.Columns.Item(columncheck).Editable = true;
                                    oGrid.Columns.Item(columncheck).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                                    oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("CK_", "");
                                    break;

                                case "ECMB":
                                    oGrid.Columns.Item(columncheck).Editable = true;
                                    oGrid.Columns.Item(columncheck).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                                    oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("ECMB_", "");
                                    break;
                            }
                        }

                        if (DocEntry != "")
                        {
                            SAPbouiCOM.EditTextColumn ColQty;
                            ColQty = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("NE_Debit");
                            ColQty.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

                            SAPbouiCOM.EditTextColumn ColQty1;
                            ColQty1 = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("NE_Credit");
                            ColQty1.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

                            objform.Items.Item("Et_Total").Enabled = false;
                            objform.Items.Item("Et_Total").Specific.Value = Convert.ToDouble(oGrid.DataTable.Columns.Item("INV_Total").Cells.Item(0).Value);
                        }
                        else
                        {
                            SAPbouiCOM.EditTextColumn ColQty;
                            ColQty = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("E_Debit");
                            ColQty.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

                            SAPbouiCOM.EditTextColumn ColQty1;
                            ColQty1 = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("E_Credit");
                            ColQty1.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                        }

                        objform.Freeze(false);
                    }
                    catch (Exception ex)
                    {
                        objform.Freeze(false);
                        objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                    oGrid.AutoResizeColumns();
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Deduction Details Data are Loaded Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                else
                {
                    oGrid.DataTable.Clear();
                    objform.Freeze(false);
                    objSBOAPI.SBO_Appln.StatusBar.SetText("No Records Found for Deduction Details", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }

                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Total Deduction
        public void Total_Deduction()
        {
            try
            {
                objform.Freeze(true);
                oGrid = objform.Items.Item("GRD2").Specific;
                double debit = 0.0;
                double credit = 0.0;
                double Total = 0.0;
                for (int i = 0; i < oGrid.DataTable.Rows.Count; i++)
                {
                    if (oGrid.DataTable.Columns.Item("CK_Select").Cells.Item(i).Value == "Y")
                    {
                        debit = debit + Convert.ToDouble(oGrid.DataTable.Columns.Item("E_Debit").Cells.Item(i).Value);
                        credit = credit + Convert.ToDouble(oGrid.DataTable.Columns.Item("E_Credit").Cells.Item(i).Value);
                    }
                }
                Total = debit - credit;
                objform.Items.Item("Et_Total").Specific.Value = Convert.ToDouble(Total);

                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Insert Into Temp
        public void Insert_Into_Temp(string DocEntry)
        {
            try
            {
                objform.Freeze(true);
                oGrid = objform.Items.Item("GRD2").Specific;
                for (int i = 0; i < oGrid.DataTable.Rows.Count; i++)
                {
                    string Name = oGrid.DataTable.Columns.Item("NE_Deduction").Cells.Item(i).Value.ToString();
                    string Select = oGrid.DataTable.Columns.Item("CK_Select").Cells.Item(i).Value.ToString();
                    double Debit = Convert.ToDouble(oGrid.DataTable.Columns.Item("E_Debit").Cells.Item(i).Value);
                    double Credit = Convert.ToDouble(oGrid.DataTable.Columns.Item("E_Credit").Cells.Item(i).Value);
                    string Ledger = oGrid.DataTable.Columns.Item("INV_Ledger").Cells.Item(i).Value.ToString();
                    double Total = Convert.ToDouble(objform.Items.Item("Et_Total").Specific.Value);

                    string Count = objSBOAPI.Query_Execute("Select COUNT(*) AS \"Count\" from \"@AVA_TEMP2\" where \"U_AVA_DocEntry\" = '" + DocEntry + "' and \"U_AVA_DedName\" = '" + Name + "'");
                    if (Count == "0")
                    {
                        string Code = objSBOAPI.Query_Execute("Select IFNULL(MAX(\"Code\"),0)+1 from \"@AVA_TEMP2\"");
                        objSBOAPI.Query_Execute("INSERT INTO \"@AVA_TEMP2\" (\"Code\",\"Name\",\"U_AVA_Select\",\"U_AVA_DedName\",\"U_AVA_DocEntry\",\"U_AVA_Debit\",\"U_AVA_Credit\",\"U_AVA_Ledger\",\"U_AVA_Total\") VALUES('" + Code + "','" + objSBOAPI.oCompany.UserName + "','" + Select + "','" + Name + "','" + DocEntry + "','" + Debit + "','" + Credit + "','" + Ledger + "','" + Total + "')");
                    }
                    else
                    {
                        objSBOAPI.Query_Execute("UPDATE \"@AVA_TEMP2\" SET \"Name\" = '" + objSBOAPI.oCompany.UserName + "',\"U_AVA_Select\" = '" + Select + "',\"U_AVA_Debit\" = '" + Debit + "',\"U_AVA_Credit\" = '" + Credit + "',\"U_AVA_Ledger\" = '" + Ledger + "',\"U_AVA_Total\" = '" + Total + "' where \"U_AVA_DocEntry\" = '" + DocEntry + "' and \"U_AVA_DedName\" = '" + Name + "'");
                    }
                }
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Insert Into Table
        public void Insert_Into_Table(string DocEntry)
        {
            try
            {
                objform.Freeze(false);
                oGrid = objform.Items.Item("GRD1").Specific;
                for (int i = 0; i < oGrid.DataTable.Rows.Count; i++)
                {
                    string Analysis = oGrid.DataTable.Columns.Item("NE_Analysis").Cells.Item(i).Value.ToString();
                    double ActualVal = Convert.ToDouble(oGrid.DataTable.Columns.Item("NE_Actual Value").Cells.Item(i).Value);
                    double DedAmount = Convert.ToDouble(oGrid.DataTable.Columns.Item("NE_Deduction Amount").Cells.Item(i).Value);
                    string Unit = oGrid.DataTable.Columns.Item("NE_Unit").Cells.Item(i).Value.ToString();
                    double Quantity = Convert.ToDouble(oGrid.DataTable.Columns.Item("NE_Quantity").Cells.Item(i).Value);
                    double Parameter = Convert.ToDouble(oGrid.DataTable.Columns.Item("NE_Standard Parameter").Cells.Item(i).Value);
                    string Remarks = oGrid.DataTable.Columns.Item("NE_Remarks").Cells.Item(i).Value.ToString();
                    int Sample = Convert.ToInt32(oGrid.DataTable.Columns.Item("NE_Sample").Cells.Item(i).Value);
                    int QCLevel = Convert.ToInt32(oGrid.DataTable.Columns.Item("NE_QC Level").Cells.Item(i).Value);

                    string Code = objSBOAPI.Query_Execute("Select IFNULL(MAX(\"Code\"),0)+1 from \"@AVA_DED1\"");
                    objSBOAPI.Query_Execute("INSERT INTO \"@AVA_DED1\" (\"Code\",\"Name\",\"U_AVA_AnlysNme\",\"U_AVA_DocEntry\",\"U_AVA_ActVal\",\"U_AVA_Unit\",\"U_AVA_DedAmt\",\"U_AVA_Qty\",\"U_AVA_StdPar\",\"U_AVA_Remarks\",\"U_AVA_Sample\",\"U_AVA_QCLevel\") VALUES('" + Code + "','" + objSBOAPI.oCompany.UserName + "','" + Analysis + "','" + DocEntry + "','" + ActualVal + "','" + Unit + "','" + DedAmount + "','" + Quantity + "','" + Parameter + "','" + Remarks + "','" + Sample + "','" + QCLevel + "')");

                }

                oGrid = objform.Items.Item("GRD2").Specific;
                for (int i = 0; i < oGrid.DataTable.Rows.Count; i++)
                {
                    if (oGrid.DataTable.Columns.Item("CK_Select").Cells.Item(i).Value == "Y")
                    {
                        string Name = oGrid.DataTable.Columns.Item("NE_Deduction").Cells.Item(i).Value.ToString();
                        double Debit = Convert.ToDouble(oGrid.DataTable.Columns.Item("E_Debit").Cells.Item(i).Value);
                        double Credit = Convert.ToDouble(oGrid.DataTable.Columns.Item("E_Credit").Cells.Item(i).Value);
                        string Ledger = oGrid.DataTable.Columns.Item("INV_Ledger").Cells.Item(i).Value.ToString();
                        double Total = Convert.ToDouble(objform.Items.Item("Et_Total").Specific.Value);

                        string Code = objSBOAPI.Query_Execute("Select IFNULL(MAX(\"Code\"),0)+1 from \"@AVA_DED2\"");
                        objSBOAPI.Query_Execute("INSERT INTO \"@AVA_DED2\" (\"Code\",\"Name\",\"U_AVA_DedName\",\"U_AVA_DocEntry\",\"U_AVA_Debit\",\"U_AVA_Credit\",\"U_AVA_Ledger\",\"U_AVA_Total\") VALUES('" + Code + "','" + objSBOAPI.oCompany.UserName + "','" + Name + "','" + DocEntry + "','" + Debit + "','" + Credit + "','" + Ledger + "','" + Total + "')");

                    }
                }

                objform.Freeze(true);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion  

        #region AP CreditMemo Creation
        public void APCreditMemo_Creation(string DocEntry)
        {
            try
            {
                objform.Freeze(true);
                double Amount;
                int i = 0;
                oMemo = (SAPbobsCOM.Documents)objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes);
                SAPbobsCOM.Recordset oRec = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRec1 = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRec2 = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string Header = objSBOAPI.Query_Execute("SELECT \"QString\" FROM OUQR WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM OQCN WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='APCreditMemo'");
                if (Header != "")
                {
                    Header = Header.Replace("[%1]", "Header");
                    Header = Header.Replace("[%2]", DocEntry);
                    oRec.DoQuery(Header);
                    if (oRec.RecordCount > 0)
                    {
                        oMemo.CardCode = oRec.Fields.Item("CardCode").Value.ToString();
                        oMemo.DocDate = objSBOAPI.GetDateTimeValue(objSBOAPI.oCompany.GetDBServerDate().ToString());
                        oMemo.DocDueDate = objSBOAPI.GetDateTimeValue(objSBOAPI.oCompany.GetDBServerDate().ToString());
                        oMemo.TaxDate = objSBOAPI.GetDateTimeValue(objSBOAPI.oCompany.GetDBServerDate().ToString());

                        oMemo.OriginalRefNo = oRec.Fields.Item("DocNum").Value.ToString();
                        oMemo.OriginalRefDate = Convert.ToDateTime(oRec.Fields.Item("DocDate").Value.ToString());

                        string Line = objSBOAPI.Query_Execute("SELECT \"QString\" FROM OUQR WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM OQCN WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='APCreditMemo'");
                        Line = Line.Replace("[%1]", "Line");
                        Line = Line.Replace("[%2]", DocEntry);
                        oRec1.DoQuery(Line);
                        if (oRec1.RecordCount > 0)
                        {
                            while (!oRec1.EoF)
                            {
                                oMemo.Lines.BaseType = 18;
                                oMemo.Lines.BaseEntry = oRec1.Fields.Item("DocEntry").Value;
                                oMemo.Lines.BaseLine = i;

                                if (DocType == "I")
                                {
                                    oMemo.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                                    oMemo.Lines.ItemCode = oRec1.Fields.Item("ItemCode").Value;
                                    oMemo.Lines.Quantity = oRec1.Fields.Item("Quantity").Value;
                                    oMemo.Lines.UnitPrice = 0;
                                    oMemo.Lines.WarehouseCode = oRec1.Fields.Item("WhsCode").Value;
                                    oMemo.Lines.TaxCode = oRec1.Fields.Item("TaxCode").Value;
                                    oMemo.Lines.WithoutInventoryMovement = SAPbobsCOM.BoYesNoEnum.tYES;
                                }
                                else if (DocType == "S")
                                {
                                    oMemo.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;
                                    oMemo.Lines.AccountCode = oRec1.Fields.Item("AcctCode").Value;
                                    oMemo.Lines.LineTotal = 0;
                                    oMemo.Lines.LocationCode = oRec1.Fields.Item("LocCode").Value;
                                    oMemo.Lines.TaxCode = oRec1.Fields.Item("TaxCode").Value;
                                }

                                oMemo.Lines.Add();
                                i++;
                                oRec1.MoveNext();
                            }
                        }

                        string Expense = objSBOAPI.Query_Execute("SELECT \"QString\" FROM OUQR WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM OQCN WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='APCreditMemo'");
                        Expense = Expense.Replace("[%1]", "Expense");
                        Expense = Expense.Replace("[%2]", DocEntry);
                        oRec2.DoQuery(Expense);
                        if (oRec2.RecordCount > 0)
                        {
                            while (!oRec2.EoF)
                            {
                                oMemo.Expenses.ExpenseCode = Convert.ToInt32(oRec2.Fields.Item("ExpenseCode").Value.ToString());
                                oMemo.Expenses.TaxCode = oRec2.Fields.Item("TaxCode").Value.ToString();
                                Amount = Convert.ToDouble(oRec2.Fields.Item("Amount").Value.ToString());
                                if (Amount < 0)
                                {
                                    Amount = -(Amount);
                                }
                                oMemo.Expenses.LineTotal = Amount;

                                oMemo.Expenses.Add();
                                oRec2.MoveNext();
                            }
                        }

                        lRetCode = oMemo.Add();

                        if (lRetCode != 0)
                        {
                            objSBOAPI.oCompany.GetLastError(out lErrCode, out sErrMsg);
                            objSBOAPI.SBO_Appln.StatusBar.SetText(lErrCode.ToString() + " - Error in While Creating AP Credit Memo - " + sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                        else
                        {
                            objSBOAPI.SBO_Appln.StatusBar.SetText("AP Credit Memo is created successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }
                    }
                    else
                    {
                        objform.Freeze(false);
                        objSBOAPI.SBO_Appln.StatusBar.SetText("No Records Found.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                else
                {
                    objform.Freeze(false);
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Query is Missing, Please Update Query Manager.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }

                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oMemo);
                GC.Collect();
            }
        }
        #endregion

        //Transport Details

        #region Validation 
        public void Validation(ref bool Bubbleevent)
        {
            try
            {
                //Success = "Y";
                //oCombo = objform.Items.Item("Cb_FCon").Specific;
                //if (objform.Items.Item("Et_TrNme").Specific.Value == "")
                //{
                //    objSBOAPI.SBO_Appln.StatusBar.SetText("Transporter Name is Mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                //    Success = "N";
                //    Bubbleevent = false;
                //    return;
                //}
                //string Str = objSBOAPI.Query_Execute("Select \"CardCode\" from \"OCRD\" where \"CardName\" = '" + objform.Items.Item("Et_TrNme").Specific.Value + "' and \"CardType\" = 'S'");
                //if (Str == "")
                //{
                //    objSBOAPI.SBO_Appln.StatusBar.SetText("Transporter is not a vendor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                //    Success = "N";
                //    Bubbleevent = false;
                //    return;
                //}
                //if (oCombo.Selected.Value == "FOR" || oCombo.Selected.Value == "EX-MILL")
                //{
                    if (objform.Items.Item("Et_InsNo").Specific.Value == "")
                    {
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Insurance Policy No is Mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        Success = "N";
                        Bubbleevent = false;
                        return;
                    }
                    else if (Convert.ToDouble(objform.Items.Item("Et_InsBal").Specific.Value) <= 0)
                    {
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Insurance Balance Amount should be greater than Zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        Success = "N";
                        Bubbleevent = false;
                        return;
                    }
                    else if (Convert.ToDouble(objform.Items.Item("Et_InsBal").Specific.Value) > 0)
                    {
                        SAPbouiCOM.Matrix oMatrix1 = objform.Items.Item("38").Specific;
                        double Amount = 0.0;
                        for (int i = 1; i < oMatrix1.VisualRowCount; i++)
                        {
                            if (oMatrix1.Columns.Item("1").Cells.Item(i).Specific.Value != "")
                            {
                                string Total = System.Text.RegularExpressions.Regex.Replace(oMatrix1.Columns.Item("21").Cells.Item(i).Specific.Value, "[^0-9.]", "");
                                Amount = Amount + Convert.ToDouble(Total);
                            }
                        }
                        if (Amount > Convert.ToDouble(objform.Items.Item("Et_InsBal").Specific.Value))
                        {
                            objSBOAPI.SBO_Appln.StatusBar.SetText("Insurance Balance is low...Please Update the Insurance Balance Amount for the Insurance Policy No : " + objform.Items.Item("Et_InsNo").Specific.Value, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            Success = "N";
                            Bubbleevent = false;
                            return;
                        }
                        else
                        {
                            string DocDate = objSBOAPI.GetDateFromField(objform.Items.Item("46").Specific.Value);
                            string TDate = objSBOAPI.GetDateToInsert(objSBOAPI.Query_Execute("Select \"U_AVA_ToDate\" from \"@AVA_INSURANCEH\" where \"Code\" = '" + objform.Items.Item("Et_InsNo").Specific.Value + "'"));
                            string ToDate = objSBOAPI.GetDateFromField(TDate);
                            int result = string.Compare(DocDate, ToDate);
                            if (result > 0)
                            {
                                objSBOAPI.SBO_Appln.StatusBar.SetText("Insurance Policy Date has been expired...Please update the Insurance Policy To Date for the Policy No : " + objform.Items.Item("Et_InsNo").Specific.Value, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                Success = "N";
                                Bubbleevent = false;
                                return;
                            }
                            string FDate = objSBOAPI.GetDateToInsert(objSBOAPI.Query_Execute("Select \"U_AVA_FromDate\" from \"@AVA_INSURANCEH\" where \"Code\" = '" + objform.Items.Item("Et_InsNo").Specific.Value + "'"));
                            string FromDate = objSBOAPI.GetDateFromField(FDate);
                            int result1 = string.Compare(FromDate, DocDate);
                            if (result1 > 0)
                            {
                                objSBOAPI.SBO_Appln.StatusBar.SetText("The Insurance Policy is not active for this Document Date...Please update the Insurance Policy From Date for the Policy No : " + objform.Items.Item("Et_InsNo").Specific.Value, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                Success = "N";
                                Bubbleevent = false;
                                return;
                            }
                        }
                        string Balance = objSBOAPI.Query_Execute("Select \"U_AVA_BalAmt\" from \"@AVA_INSURANCEH\" where \"Code\" = '" + objform.Items.Item("Et_InsNo").Specific.Value + "'");
                        objform.Items.Item("Et_InsBal").Specific.Value = Convert.ToDouble(Balance) - Amount;
                        objform.Items.Item("Et_FBal").Specific.value = Convert.ToDouble(objform.Items.Item("Et_FAmt").Specific.Value) - Convert.ToDouble(objform.Items.Item("Et_FAdv").Specific.Value);
                    }
                //}
            }
            catch (Exception ex)
            {
                Bubbleevent = false;
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Add Freight
        public void Add_Freight()
        {
            try
            {
                objform.Freeze(true);
                //oCombo = objform.Items.Item("Cb_FCon").Specific;
                objform.Items.Item("91").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                oForm = objSBOAPI.SBO_Appln.Forms.ActiveForm;
                oMatrix = oForm.Items.Item("3").Specific;
                SAPbobsCOM.Recordset oRec = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRec1 = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string Freight = objSBOAPI.Query_Execute("SELECT \"QString\" FROM OUQR WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM OQCN WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='APInvoice'");
                if (Freight != "")
                {
                    for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                    {
                        int ExpCode = Convert.ToInt32(oMatrix.Columns.Item("1").Cells.Item(i).Specific.Value);
                        //Freight = Freight.Replace("[%1]", "Freight");
                        //oRec.DoQuery(Freight);
                        //if (oRec.RecordCount > 0)
                        //{
                        //    //if (oCombo.Selected.Value == "FOR")
                        //    //{
                        //    //    if (Convert.ToDouble(objform.Items.Item("Et_FAmt").Specific.Value) > 0 && ExpCode == oRec.Fields.Item("ExpnsCode").Value)
                        //    //    {
                        //    //        oMatrix.Columns.Item("3").Cells.Item(i).Specific.Value = Convert.ToDouble(objform.Items.Item("Et_FAmt").Specific.Value);
                        //    //        oMatrix.Columns.Item("17").Cells.Item(i).Specific.Value = oRec.Fields.Item("TaxCode").Value;
                        //    //    }
                        //    //}
                        //    if (oCombo.Selected.Value == "EX-MILL")
                        //    {
                        //        if (Convert.ToDouble(objform.Items.Item("Et_FAdv").Specific.Value) > 0 && ExpCode == oRec.Fields.Item("ExpnsCode").Value)
                        //        {
                        //            oMatrix.Columns.Item("3").Cells.Item(i).Specific.Value = Convert.ToDouble(objform.Items.Item("Et_FAdv").Specific.Value);
                        //            oMatrix.Columns.Item("17").Cells.Item(i).Specific.Value = oRec.Fields.Item("TaxCode").Value;
                        //        }
                        //    }
                        //}
                        //else
                        //{
                        //    objSBOAPI.SBO_Appln.StatusBar.SetText("Expense Code is missing for Freight Charges", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //}

                        string Insurance = objSBOAPI.Query_Execute("SELECT \"QString\" FROM OUQR WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM OQCN WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='APInvoice'");
                        Insurance = Insurance.Replace("[%1]", "Insurance");
                        Insurance = Insurance.Replace("[%3]", objform.Items.Item("4").Specific.Value);
                        oRec1.DoQuery(Insurance);
                        if (oRec1.RecordCount > 0)
                        {
                            if (Convert.ToDouble(objform.Items.Item("Et_InsBal").Specific.Value) > 0 && ExpCode == oRec1.Fields.Item("ExpnsCode").Value)
                            {
                                oMatrix.Columns.Item("3").Cells.Item(i).Specific.Value = Convert.ToDouble(oRec1.Fields.Item("Amount").Value);
                                oMatrix.Columns.Item("17").Cells.Item(i).Specific.Value = oRec1.Fields.Item("TaxCode").Value;
                            }
                        }
                        else
                        {
                            objSBOAPI.SBO_Appln.StatusBar.SetText("Expense Code is missing for Insurance", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                    }
                }
                else
                {
                    objform.Freeze(false);
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Query is Missing, Please Update Query Manager.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }

                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                if (objSBOAPI.SBO_Appln.Forms.ActiveForm.Type == 3007)
                {
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region PO Creation
        //public void PO_Creation(string Docentry)
        //{
        //    try
        //    {
        //        SAPbobsCOM.Documents PO = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);
        //        SAPbobsCOM.Recordset oRec = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //        string Purchase = objSBOAPI.Query_Execute("SELECT \"QString\" FROM OUQR WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM OQCN WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='APInvoice'");
        //        if (Purchase != "")
        //        {
        //            Purchase = Purchase.Replace("[%1]", "Purchase");
        //            Purchase = Purchase.Replace("[%2]", Docentry);
        //            oRec.DoQuery(Purchase);
        //            if (oRec.RecordCount > 0)
        //            {
        //                PO.CardCode = oRec.Fields.Item("VENDOR").Value;
        //                PO.DocDate = objSBOAPI.oCompany.GetDBServerDate();
        //                PO.DocDueDate = objSBOAPI.oCompany.GetDBServerDate();
        //                PO.TaxDate = objSBOAPI.oCompany.GetDBServerDate();
        //                PO.NumAtCard = oRec.Fields.Item("DocNum").Value.ToString();
        //                PO.Confirmed = SAPbobsCOM.BoYesNoEnum.tYES;

        //                PO.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;
        //                PO.Lines.SACEntry = Convert.ToInt32(oRec.Fields.Item("SAC").Value);
        //                PO.Lines.AccountCode = oRec.Fields.Item("Account").Value.ToString();
        //                PO.Lines.ItemDescription = oRec.Fields.Item("Desc").Value.ToString();
        //                PO.Lines.LineTotal = Convert.ToDouble(oRec.Fields.Item("AMOUNT").Value);
        //                PO.Lines.LocationCode = Convert.ToInt32(oRec.Fields.Item("LOCATION").Value);
        //                PO.Lines.TaxCode = oRec.Fields.Item("TaxCode").Value.ToString();
        //                PO.UserFields.Fields.Item("U_AVA_SALESINVNO").Value = oRec.Fields.Item("DocNum").Value.ToString();
        //                PO.UserFields.Fields.Item("U_AVA_FRTCONDITION").Value = oRec.Fields.Item("U_AVA_FRTCONDITION").Value.ToString();
        //                PO.UserFields.Fields.Item("U_AVA_FREIGHTRATE").Value = Convert.ToDouble(oRec.Fields.Item("U_AVA_FREIGHTRATE").Value.ToString());
        //                PO.UserFields.Fields.Item("U_AVA_FREIGHTAMT").Value = Convert.ToDouble(oRec.Fields.Item("U_AVA_FREIGHTAMT").Value.ToString());
        //                PO.UserFields.Fields.Item("U_AVA_FREIGHTADVANCE").Value = Convert.ToDouble(oRec.Fields.Item("U_AVA_FREIGHTADVANCE").Value.ToString());
        //                PO.UserFields.Fields.Item("U_AVA_FREIGHTBALANCE").Value = Convert.ToDouble(oRec.Fields.Item("U_AVA_FREIGHTBALANCE").Value.ToString());
        //                PO.UserFields.Fields.Item("U_AVA_TRANSNAME").Value = oRec.Fields.Item("U_AVA_TRANSNAME").Value.ToString();
        //                PO.UserFields.Fields.Item("U_AVA_CUSCODE").Value = oRec.Fields.Item("CardCode").Value.ToString();
        //                PO.UserFields.Fields.Item("U_AVA_CUSNAME").Value = oRec.Fields.Item("CardName").Value.ToString();
        //                PO.UserFields.Fields.Item("U_AVA_LOADINGPERSON").Value = oRec.Fields.Item("U_AVA_LOADINGPERSON").Value.ToString();
        //                PO.UserFields.Fields.Item("U_AVA_DRIVERMOBILENO").Value = oRec.Fields.Item("U_AVA_DRIVERMOBILENO").Value.ToString();
        //                PO.UserFields.Fields.Item("U_AVA_DRIVERLNO").Value = oRec.Fields.Item("U_AVA_DRIVERLNO").Value.ToString();
        //                PO.UserFields.Fields.Item("U_AVA_OWNTRUCK").Value = oRec.Fields.Item("U_AVA_OWNTRUCK").Value.ToString();
        //                PO.UserFields.Fields.Item("U_AVA_SUPERVISOR").Value = oRec.Fields.Item("U_AVA_SUPERVISOR").Value.ToString();
        //                PO.UserFields.Fields.Item("U_AVA_VEHICLENO").Value = oRec.Fields.Item("U_AVA_VEHICLENO").Value.ToString();

        //                lRetCode = PO.Add();
        //                if (lRetCode == 0)
        //                {
        //                    objSBOAPI.SBO_Appln.StatusBar.SetText("Purchase Order Created successfully - " + objSBOAPI.oCompany.GetNewObjectKey(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        //                    DPR_Creation(objSBOAPI.oCompany.GetNewObjectKey());
        //                }
        //                else
        //                {
        //                    objSBOAPI.oCompany.GetLastError(out lErrCode, out sErrMsg);
        //                    objSBOAPI.SBO_Appln.StatusBar.SetText(lErrCode.ToString() + " - Error in While Creating Purchase Order - " + sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                }
        //            }
        //            else
        //            {
        //                objform.Freeze(false);
        //                objSBOAPI.SBO_Appln.StatusBar.SetText("No Records found.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //            }
        //        }
        //        else
        //        {
        //            objform.Freeze(false);
        //            objSBOAPI.SBO_Appln.StatusBar.SetText("APInvoice Query is Missing, Please Update Query Manager.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        objform.Freeze(false);
        //        objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //    }

        //}
        #endregion

        #region DPR Creation
        //public void DPR_Creation(string Docentry)
        //{
        //    try
        //    {
        //        SAPbobsCOM.Documents oDP = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDownPayments);
        //        SAPbobsCOM.Recordset oRec = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //        string Purchase = objSBOAPI.Query_Execute("SELECT \"QString\" FROM OUQR WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM OQCN WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='APInvoice'");
        //        if (Purchase != "")
        //        {
        //            Purchase = Purchase.Replace("[%1]", "DPR");
        //            Purchase = Purchase.Replace("[%2]", Docentry);
        //            oRec.DoQuery(Purchase);
        //            if (oRec.RecordCount > 0)
        //            {
        //                oDP.CardCode = oRec.Fields.Item("CardCode").Value;
        //                oDP.NumAtCard = oRec.Fields.Item("U_AVA_SALESINVNO").Value;
        //                oDP.DownPaymentType = SAPbobsCOM.DownPaymentTypeEnum.dptRequest;
        //                oDP.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;

        //                oDP.Lines.SACEntry = Convert.ToInt32(oRec.Fields.Item("SAC").Value);
        //                oDP.Lines.AccountCode = oRec.Fields.Item("Account").Value.ToString();
        //                oDP.Lines.ItemDescription = oRec.Fields.Item("Desc").Value.ToString();
        //                oDP.Lines.LineTotal = Convert.ToDouble(oRec.Fields.Item("LineTotal").Value);
        //                oDP.Lines.LocationCode = Convert.ToInt32(oRec.Fields.Item("LocCode").Value);
        //                oDP.Lines.TaxCode = oRec.Fields.Item("TaxCode").Value.ToString();
        //                oDP.Lines.BaseType = 22;
        //                oDP.Lines.BaseEntry = Convert.ToInt32(Docentry);
        //                oDP.Lines.BaseLine = oRec.Fields.Item("LineNum").Value;

        //                oDP.UserFields.Fields.Item("U_AVA_SALESINVNO").Value = oRec.Fields.Item("U_AVA_SALESINVNO").Value;
        //                oDP.UserFields.Fields.Item("U_AVA_FRTCONDITION").Value = oRec.Fields.Item("U_AVA_FRTCONDITION").Value.ToString();
        //                oDP.UserFields.Fields.Item("U_AVA_FREIGHTRATE").Value = Convert.ToDouble(oRec.Fields.Item("U_AVA_FREIGHTRATE").Value.ToString());
        //                oDP.UserFields.Fields.Item("U_AVA_FREIGHTAMT").Value = Convert.ToDouble(oRec.Fields.Item("U_AVA_FREIGHTAMT").Value.ToString());
        //                oDP.UserFields.Fields.Item("U_AVA_FREIGHTADVANCE").Value = Convert.ToDouble(oRec.Fields.Item("U_AVA_FREIGHTADVANCE").Value.ToString());
        //                oDP.UserFields.Fields.Item("U_AVA_FREIGHTBALANCE").Value = Convert.ToDouble(oRec.Fields.Item("U_AVA_FREIGHTBALANCE").Value.ToString());
        //                oDP.UserFields.Fields.Item("U_AVA_TRANSNAME").Value = oRec.Fields.Item("U_AVA_TRANSNAME").Value.ToString();
        //                oDP.UserFields.Fields.Item("U_AVA_CUSCODE").Value = oRec.Fields.Item("U_AVA_CUSCODE").Value.ToString();
        //                oDP.UserFields.Fields.Item("U_AVA_CUSNAME").Value = oRec.Fields.Item("U_AVA_CUSNAME").Value.ToString();
        //                oDP.UserFields.Fields.Item("U_AVA_LOADINGPERSON").Value = oRec.Fields.Item("U_AVA_LOADINGPERSON").Value.ToString();
        //                oDP.UserFields.Fields.Item("U_AVA_DRIVERMOBILENO").Value = oRec.Fields.Item("U_AVA_DRIVERMOBILENO").Value.ToString();
        //                oDP.UserFields.Fields.Item("U_AVA_DRIVERLNO").Value = oRec.Fields.Item("U_AVA_DRIVERLNO").Value.ToString();
        //                oDP.UserFields.Fields.Item("U_AVA_OWNTRUCK").Value = oRec.Fields.Item("U_AVA_OWNTRUCK").Value.ToString();
        //                oDP.UserFields.Fields.Item("U_AVA_SUPERVISOR").Value = oRec.Fields.Item("U_AVA_SUPERVISOR").Value.ToString();
        //                oDP.UserFields.Fields.Item("U_AVA_VEHICLENO").Value = oRec.Fields.Item("U_AVA_VEHICLENO").Value.ToString();

        //                oDP.Lines.Add();

        //                lRetCode = oDP.Add();
        //                if (lRetCode == 0)
        //                {
        //                    objSBOAPI.SBO_Appln.StatusBar.SetText("Down Payment Request  Created successfully - " + objSBOAPI.oCompany.GetNewObjectKey(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        //                }
        //                else
        //                {
        //                    objSBOAPI.oCompany.GetLastError(out lErrCode, out sErrMsg);
        //                    objSBOAPI.SBO_Appln.StatusBar.SetText(lErrCode.ToString() + " - Error in While Creating Down Payment Request - " + sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                }
        //            }
        //            else
        //            {
        //                objform.Freeze(false);
        //                objSBOAPI.SBO_Appln.StatusBar.SetText("No Records found.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //            }
        //        }
        //        else
        //        {
        //            objform.Freeze(false);
        //            objSBOAPI.SBO_Appln.StatusBar.SetText("APInvoice Query is Missing, Please Update Query Manager.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        objform.Freeze(false);
        //        objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //    }
        //}
        #endregion

        #region Insurance Update
        public void Insurance_Update(string Entry)
        {
            try
            {
                SAPbobsCOM.GeneralService oGeneralService = null;
                SAPbobsCOM.GeneralData oGeneralData = null;
                SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                SAPbobsCOM.GeneralData oChild = null;
                SAPbobsCOM.GeneralDataCollection oChildren = null;

                try
                {
                    SAPbobsCOM.CompanyService CmpServ = null;
                    CmpServ = objSBOAPI.oCompany.GetCompanyService();
                    SAPbobsCOM.Recordset oRec = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string Insurance = objSBOAPI.Query_Execute("SELECT \"QString\" FROM OUQR WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM OQCN WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='Insurance_Update'");
                    if (Insurance == "")
                    {
                        objform.Freeze(false);
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Insurance_Update Query is Missing, Please Update Query Manager.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                    else
                    {
                        Insurance = Insurance.Replace("[%1]", Entry);
                        oRec.DoQuery(Insurance);
                        if (oRec.RecordCount > 0)
                        {
                            oGeneralService = CmpServ.GetGeneralService("AVA_INSURANCE");

                            oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                            oGeneralParams.SetProperty("Code", oRec.Fields.Item("PolicyNo").Value);
                            oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                            oGeneralData.SetProperty("U_AVA_BalAmt", Convert.ToDouble(oRec.Fields.Item("Amount").Value));

                            oChildren = oGeneralData.Child("AVA_INSURANCEL");
                            oChild = oChildren.Add();
                            oChild.SetProperty("U_AVA_CardCode", oRec.Fields.Item("CardCode").Value);
                            oChild.SetProperty("U_AVA_CardName", oRec.Fields.Item("CardName").Value);
                            oChild.SetProperty("U_AVA_Type", oRec.Fields.Item("ObjType").Value);
                            oChild.SetProperty("U_AVA_DocEntry", oRec.Fields.Item("DocEntry").Value.ToString());
                            oChild.SetProperty("U_AVA_DocNum", oRec.Fields.Item("DocNum").Value.ToString());
                            oChild.SetProperty("U_AVA_DocStatus", oRec.Fields.Item("DocStatus").Value);
                            oChild.SetProperty("U_AVA_Date", oRec.Fields.Item("DocDate").Value);
                            if (oRec.Fields.Item("DocStatus").Value == "O")
                            {
                                oChild.SetProperty("U_AVA_Debit", Convert.ToDouble(oRec.Fields.Item("Total").Value));
                            }
                            else if (oRec.Fields.Item("DocStatus").Value == "C")
                            {
                                oChild.SetProperty("U_AVA_Credit", Convert.ToDouble(oRec.Fields.Item("Total").Value));
                            }
                            oGeneralService.Update(oGeneralData);

                            objSBOAPI.SBO_Appln.StatusBar.SetText("The Insurance Amount is updated for the Policy No : " + oRec.Fields.Item("PolicyNo").Value, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }
                    }
                }
                catch (Exception ex)
                {
                    objform.Freeze(false);
                    objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion
    }
}
