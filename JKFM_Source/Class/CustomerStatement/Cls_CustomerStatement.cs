using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JKFM_Source
{
    class Cls_CustomerStatement
    {
        #region Declaration
        SAPbouiCOM.Form objform;
        private ClsSBO objSBOAPI;

        SAPbouiCOM.Grid oGrid;
        SAPbouiCOM.DataTable Dt;

        string Type = "";
        string BtnType = "";
        string[] arr;

        SAPbouiCOM.CheckBox oChk;
        DialogResult Dialog;

        string FromMailId = "", SMTPServer = "", MailPwd = "";
        bool EnableSSL = false;
        int Port;
        string Subject = "", Body = "", ObjType = "";

        string ExcelQuery = "";
        #endregion

        #region Constructor
        public Cls_CustomerStatement(ClsSBO objSBO)
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
                    if ((pval.EventType == SAPbouiCOM.BoEventTypes.et_CLICK) | (pval.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) | (pval.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE) | (pval.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE) | (pval.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE))
                    {
                        try
                        {
                            SAPbouiCOM.StaticText oST;
                            oST = objform.Items.Item("ST1").Specific;

                            SAPbouiCOM.Form oform;
                            oform = objSBOAPI.SBO_Appln.Forms.Item(oST.Caption);
                            oform.Select();
                            bubbleevent = false;
                        }
                        catch (Exception)
                        {
                        }
                    }
                    switch (pval.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            switch (pval.ItemUID)
                            {
                                case "Item_19":
                                    DateTime Date1;
                                    DateTime.TryParseExact(objform.Items.Item("Item_8").Specific.value, "yyyyMMdd", new CultureInfo("en-US"), DateTimeStyles.None, out Date1);
                                    DateTime Date2;
                                    DateTime.TryParseExact(objform.Items.Item("Item_11").Specific.value, "yyyyMMdd", new CultureInfo("en-US"), DateTimeStyles.None, out Date2);
                                    if (objform.Items.Item("Item_8").Specific.value == "" | objform.Items.Item("Item_11").Specific.value == "")
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("From Date & To Date are Mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        bubbleevent = false;
                                    }
                                    else if (Date1 > Date2)
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("To Date Should be Greater than From Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        bubbleevent = false;
                                    }
                                    break;

                                case "Item_0":
                                case "Item_17":
                                case "Item_20":
                                    if (pval.ItemUID == "Item_0")
                                    {
                                        SAPbouiCOM.CheckBox oChk = objform.Items.Item("Item_28").Specific;
                                        if (oChk.Checked == false)
                                        {
                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Please Check the Email Checkbox to send mail..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                            bubbleevent = false;
                                            return;
                                        }
                                    }
                                    oGrid = objform.Items.Item("Item_2").Specific;
                                    objform.Freeze(true);
                                    int SltCnt = 0;
                                    if (oGrid.DataTable.IsEmpty == false)
                                    {
                                        for (int i = 0; i <= oGrid.DataTable.Rows.Count - 1; i++)
                                        {
                                            if (oGrid.DataTable.GetValue("CK_Select", i) == "Y")
                                            {
                                                SltCnt = 1;
                                                break;
                                            }
                                        }
                                        if (SltCnt == 0)
                                        {
                                            objSBOAPI.SBO_Appln.StatusBar.SetText("No Line is Selected...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                            bubbleevent = false;
                                        }
                                    }
                                    else
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("No records Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        bubbleevent = false;
                                    }
                                    objform.Freeze(false);
                                    break;
                            }
                            break;
                    }
                }
                else
                    switch (pval.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                            if (pval.CharPressed == Convert.ToInt32(Keys.Enter) & pval.ItemUID == "Item_22")
                            {
                                oGrid = objform.Items.Item("Item_2").Specific;
                                Dt = objform.DataSources.DataTables.Item("DT_0");
                                string DocNum = objform.Items.Item("Item_22").Specific.value.ToString();
                                objform.Freeze(true);
                                for (int i = 0; i <= Dt.Rows.Count - 1; i++)
                                {
                                    string Value = "";
                                    if (ObjType == "CustomerStatement")
                                    {
                                        Value = Convert.ToString(Dt.GetValue("NES_CardCode", oGrid.GetDataTableRowIndex(i)));
                                    }
                                    else
                                    {
                                        Value = Convert.ToString(Dt.GetValue("NES_DocNum", oGrid.GetDataTableRowIndex(i)));
                                    }
                                    if (Value.Equals(DocNum) == true)
                                    {
                                        if (Dt.GetValue("CK_Select", oGrid.GetDataTableRowIndex(i)) != "Y")
                                        {
                                            Dt.SetValue("CK_Select", oGrid.GetDataTableRowIndex(i), "Y");
                                            Selecting_Rows("Y", Convert.ToString(oGrid.GetDataTableRowIndex(i)));
                                        }
                                        oGrid.Columns.Item("CK_Select").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Descending);
                                        objform.Items.Item("Item_22").Specific.value = "";
                                        break;
                                    }
                                }
                                objform.Freeze(false);
                            }
                            break;



                        case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                            if (pval.ItemUID == "Item_16")
                            {
                                SAPbobsCOM.Recordset orec;
                                orec = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                string str = "";

                                SAPbouiCOM.ComboBox oCombo;
                                oCombo = objform.Items.Item("Item_16").Specific;
                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                {
                                    str = "SELECT \"Code\",\"Name\",IFNULL(\"U_AV_Deflt\",'N')\"U_AV_Deflt\" FROM \"@AV_LPST\" WHERE \"U_AV_Type\"='" + oCombo.Selected.Value + "'";
                                }
                                else
                                {
                                    str = "SELECT Code,Name,ISNULL(U_AV_Deflt,'N')'U_AV_Deflt' FROM [@AV_LPST] WHERE U_AV_Type='" + oCombo.Selected.Value + "'";
                                }
                                orec.DoQuery(str);

                                oCombo = objform.Items.Item("Item_18").Specific;

                                if (oCombo.ValidValues.Count > 0)
                                {
                                    while (oCombo.ValidValues.Count != 0)
                                    {
                                        oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }

                                while (!orec.EoF)
                                {
                                    oCombo.ValidValues.Add(orec.Fields.Item(0).Value, orec.Fields.Item(1).Value);

                                    if (orec.Fields.Item("U_AV_Deflt").Value == "Y")
                                    {
                                        oCombo.Select(orec.Fields.Item(0).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                    }
                                    orec.MoveNext();
                                }
                            }
                            break;

                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            switch (pval.ItemUID)
                            {
                                case "Item_0":
                                    oChk = objform.Items.Item("Item_30").Specific;
                                    if (oChk.Checked == false)
                                    {
                                        Dialog = (System.Windows.Forms.DialogResult)(objSBOAPI.SBO_Appln.MessageBox("Notification will be sent to the Customer. Do you wish to Proceed", 1, "Yes", "No"));
                                        if (Dialog == DialogResult.OK)
                                        {
                                            BtnType = "Mail";
                                            System_threading();
                                            BtnType = "";
                                        }
                                    }
                                    else
                                    {
                                        DialogResult Dialog;
                                        Dialog = (System.Windows.Forms.DialogResult)(objSBOAPI.SBO_Appln.MessageBox("Notification will be sent to the Customer by Grouping. Do you wish to Proceed", 1, "Yes", "No"));

                                        if (Dialog == DialogResult.OK)
                                        {
                                            BtnType = "Mail";
                                            System_threading();
                                            BtnType = "";
                                        }
                                    }
                                    break;

                                case "Item_19":
                                    arr = null;
                                    Loading_Query_Type();
                                    break;

                                case "Item_26":
                                    Dt = objform.DataSources.DataTables.Item("DT_0");
                                    objform.Freeze(true);
                                    arr = null;
                                    oGrid = objform.Items.Item("Item_2").Specific;
                                    for (int i = 0; i <= Dt.Rows.Count - 1; i++)
                                    {
                                        Dt.SetValue("CK_Select", i, "Y");
                                        Selecting_Rows("Y", Convert.ToString(oGrid.GetDataTableRowIndex(i)));
                                    }
                                    objform.Freeze(false);
                                    break;

                                case "Item_25":
                                    Dt = objform.DataSources.DataTables.Item("DT_0");
                                    objform.Freeze(true);
                                    arr = null;
                                    for (int i = 0; i <= Dt.Rows.Count - 1; i++)
                                        Dt.SetValue("CK_Select", i, "N");
                                    objform.Freeze(false);
                                    break;

                                case "Item_17":
                                    Dialog = (System.Windows.Forms.DialogResult)(objSBOAPI.SBO_Appln.MessageBox("Do you want to print the invoice?", 1, "Yes", "No"));
                                    if (Dialog == DialogResult.OK)
                                    {
                                        BtnType = "Print";
                                        System_threading();
                                        BtnType = "";
                                    }
                                    break;

                                case "Item_20":
                                    BtnType = "";
                                    System_threading();
                                    break;

                                case "Item_23":
                                    Dt = objform.DataSources.DataTables.Item("DT_0");
                                    objform.Freeze(true);
                                    string ccemail = objform.Items.Item("Item_15").Specific.value;
                                    for (var i = 0; i <= Dt.Rows.Count - 1; i++)
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait while copying the CCEmail ID...!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                        Dt.SetValue("E_CC Mail-ID", i, ccemail);
                                    }
                                    objSBOAPI.SBO_Appln.StatusBar.SetText("CCEmail ID Copied Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                    objform.Freeze(false);
                                    break;

                                case "Item_32":
                                    Dt = objform.DataSources.DataTables.Item("DT_0");
                                    objform.Freeze(true);
                                    string emailId = objform.Items.Item("Item_29").Specific.value;
                                    for (var i = 0; i <= Dt.Rows.Count - 1; i++)
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait while copying the Email ID...!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                        Dt.SetValue("E_Email-ID", i, emailId);
                                    }
                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Email ID Copied Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                    objform.Freeze(false);
                                    break;

                                case "Item_2":
                                    if (pval.ColUID == "CK_Select" && pval.Row != -1)
                                    {
                                        oGrid = objform.Items.Item("Item_2").Specific;
                                        SAPbouiCOM.CheckBoxColumn oChk = (SAPbouiCOM.CheckBoxColumn)oGrid.Columns.Item("CK_Select");
                                        if (oChk.IsChecked(pval.Row) == true)
                                            Selecting_Rows("Y", Convert.ToString(oGrid.GetDataTableRowIndex(pval.Row)));
                                        else if (oChk.IsChecked(pval.Row) == false)
                                            Selecting_Rows("N", Convert.ToString(oGrid.GetDataTableRowIndex(pval.Row)));
                                    }
                                    break;
                            }
                            break;

                        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                            SAPbouiCOM.DataTable dt;
                            SAPbouiCOM.ChooseFromListEvent cfl;
                            cfl = (SAPbouiCOM.ChooseFromListEvent)pval;
                            dt = cfl.SelectedObjects;

                            if (dt != null)
                            {
                                switch (pval.ItemUID)
                                {
                                    case "Item_3":
                                        try
                                        {
                                            objform.Items.Item("Item_3").Specific.value = dt.GetValue("CardCode", 0);
                                        }
                                        catch (Exception)
                                        {
                                        }
                                        break;

                                    case "Item_5":
                                        try
                                        {
                                            objform.Items.Item("Item_5").Specific.value = dt.GetValue("CardCode", 0);
                                        }
                                        catch (Exception)
                                        {
                                        }
                                        break;

                                    case "Item_13":
                                        try
                                        {
                                            objform.Items.Item("Item_13").Specific.@string = dt.GetValue("GroupName", 0);
                                        }
                                        catch (Exception)
                                        {
                                        }
                                        break;
                                }
                            }
                            break;
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
                switch (pval.MenuUID)
                {
                    case "AV_CUSTM":
                        objform = objSBOAPI.LoadForm("CustomerStatement.xml", "AV_CUSTF");
                        objform.Freeze(true);
                        Dt = objform.DataSources.DataTables.Item("DT_0");
                        objform.EnableMenu("1281", false);
                        objform.EnableMenu("1282", false);
                        objform.EnableMenu("784", true);
                        SAPbouiCOM.ComboBox oCombo;
                        oCombo = objform.Items.Item("Item_16").Specific;
                        oCombo.ValidValues.Add("AR", "A/R Invoice");
                        oCombo.ValidValues.Add("CN", "Credit Memo");
                        oCombo.ValidValues.Add("GP", "GRPO");
                        oCombo.ValidValues.Add("OP", "Outgoing Payments");
                        oCombo.ValidValues.Add("CS", "Customer Statement");
                        oCombo.Select("CS", SAPbouiCOM.BoSearchKey.psk_ByValue);
                        SAPbouiCOM.CheckBox oChk;
                        oChk = objform.Items.Item("Item_27").Specific;
                        oChk.Checked = true;
                        objform.Items.Item("Item_27").Enabled = false;
                        objform.Items.Item("Item_3").Click();
                        objform.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized;
                        objform.Visible = true;
                        objform.Freeze(false);
                        break;
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region LOADING QUERY BASED ON TYPE
        public void Loading_Query_Type()
        {
            try
            {
                SAPbobsCOM.Recordset orec;
                orec = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string str = "";

                SAPbouiCOM.ComboBox oCombo;
                oCombo = objform.Items.Item("Item_16").Specific;

                switch (oCombo.Value)
                {
                    case "AR":
                        ObjType = "13";
                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        {
                            str = objSBOAPI.Query_Execute("SELECT \"QString\" FROM \"OUQR\" WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM \"OQCN\" WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='AR_Invoice_HANA'");
                        }
                        else
                        {
                            str = objSBOAPI.Query_Execute("SELECT QString FROM OUQR WHERE QCategory=(SELECT CategoryId FROM OQCN WHERE CatName='AddonQuery_SQL') AND QName='AR_Invoice'");
                        }
                        break;

                    case "CN":
                        ObjType = "14";
                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        {
                            str = objSBOAPI.Query_Execute("SELECT \"QString\" FROM \"OUQR\" WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM \"OQCN\" WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='AR_CreditMemo_HANA'");
                        }
                        else
                        {
                            str = objSBOAPI.Query_Execute("SELECT QString FROM OUQR WHERE QCategory=(SELECT CategoryId FROM OQCN WHERE CatName='AddonQuery_SQL') AND QName='AR_CreditMemo'");
                        }
                        break;

                    case "GP":
                        ObjType = "20";
                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        {
                            str = objSBOAPI.Query_Execute("SELECT \"QString\" FROM \"OUQR\" WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM \"OQCN\" WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='GRPO_HANA'");
                        }
                        else
                        {
                            str = objSBOAPI.Query_Execute("SELECT QString FROM OUQR WHERE QCategory=(SELECT CategoryId FROM OQCN WHERE CatName='AddonQuery_SQL') AND QName='GRPO'");
                        }
                        break;

                    case "OP":
                        ObjType = "46";
                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        {
                            str = objSBOAPI.Query_Execute("SELECT \"QString\" FROM \"OUQR\" WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM \"OQCN\" WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='OutgoingPayment_HANA'");
                        }
                        else
                        {
                            str = objSBOAPI.Query_Execute("SELECT QString FROM OUQR WHERE QCategory=(SELECT CategoryId FROM OQCN WHERE CatName='AddonQuery_SQL') AND QName='OutgoingPayment_HANA'");
                        }
                        break;

                    case "CS":
                        ObjType = "CustomerStatement";
                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        {
                            str = objSBOAPI.Query_Execute("SELECT \"QString\" FROM \"OUQR\" WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM \"OQCN\" WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='CustomerStatement_HANA'");
                        }
                        else
                        {
                            str = objSBOAPI.Query_Execute("SELECT QString FROM OUQR WHERE QCategory=(SELECT CategoryId FROM OQCN WHERE CatName='AddonQuery_SQL') AND QName='CustomerStatement_HANA'");
                        }
                        break;

                    default:
                        ObjType = "0";
                        break;
                }

                if (str != "")
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait...Datas are loading...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    oGrid = objform.Items.Item("Item_2").Specific;

                    oGrid.DataTable.Clear();

                    objform.Freeze(true);

                    // 'FCC
                    str = str.Replace("%1", "" + objform.Items.Item("Item_3").Specific.value + "");
                    // 'TCC
                    str = str.Replace("%2", "" + objform.Items.Item("Item_5").Specific.value + "");
                    // 'FromDate
                    str = str.Replace("%3", "" + objform.Items.Item("Item_8").Specific.value + "");
                    // 'ToDate
                    str = str.Replace("%4", "" + objform.Items.Item("Item_11").Specific.value + "");
                    // 'BPGroup
                    str = str.Replace("%5", "" + objform.Items.Item("Item_13").Specific.value + "");
                    // 'CCMail
                    str = str.Replace("%6", "" + objform.Items.Item("Item_15").Specific.value + "");

                    oGrid.DataTable.ExecuteQuery(str);

                    if (oGrid.DataTable.IsEmpty == false)
                    {
                        oGrid.Columns.Item("RowsHeader").Visible = true;

                        for (int i = 0; i <= oGrid.Columns.Count - 1; i++)
                        {
                            string columncheck = "";

                            columncheck = oGrid.Columns.Item(i).TitleObject.Caption;

                            if (columncheck.ToString().StartsWith("NE_") == true)
                            {
                                oGrid.Columns.Item(columncheck).Editable = false;
                                oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("NE_", "");
                            }
                            else if (columncheck.ToString().StartsWith("NES_") == true)
                            {
                                oGrid.Columns.Item(columncheck).Editable = false;
                                oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("NES_", "");
                                oGrid.Columns.Item(columncheck).TitleObject.Sortable = true;
                            }
                            else if (columncheck.ToString().StartsWith("E_") == true)
                            {
                                oGrid.Columns.Item(columncheck).Editable = true;
                                oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("E_", "");
                            }
                            else if (columncheck.ToString().StartsWith("EC_") == true)
                            {
                                oGrid.Columns.Item(columncheck).Editable = true;
                                oGrid.Columns.Item(columncheck).BackColor = -1;
                                oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("EC_", "");
                            }
                            else if (columncheck.ToString().StartsWith("NEC_") == true)
                            {
                                oGrid.Columns.Item(columncheck).Editable = false;
                                oGrid.Columns.Item(columncheck).BackColor = -1;
                                oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("NEC_", "");
                            }
                            else if (columncheck.ToString().StartsWith("INV_") == true)
                            {
                                oGrid.Columns.Item(columncheck).Visible = false;
                                oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("INV_", "");
                            }
                            else if (columncheck.ToString().StartsWith("CK_") == true)
                            {
                                oGrid.Columns.Item(columncheck).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                                oGrid.Columns.Item(columncheck).Editable = true;
                                oGrid.Columns.Item(columncheck).TitleObject.Sortable = true;
                                oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("CK_", "");
                            }
                            else if (columncheck.ToString().StartsWith("ECMB_") == true)
                            {
                                oGrid.Columns.Item(columncheck).Editable = true;
                                oGrid.Columns.Item(columncheck).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                                oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace("ECMB_", "");
                            }
                        }

                        SAPbouiCOM.EditTextColumn oEdit;

                        oEdit = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("NES_CardCode");
                        oEdit.LinkedObjectType = "2";

                        if (ObjType != "CustomerStatement")
                        {
                            oEdit = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("NE_DocEntry");
                            oEdit.LinkedObjectType = ObjType;
                        }

                        oGrid.AutoResizeColumns();
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Datas Loaded Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                    else
                    {
                        objSBOAPI.SBO_Appln.StatusBar.SetText("No Datas are found", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }
                    objform.Freeze(false);
                }
                else
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("No Data are Found - Check the Addon Query in QM", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region SENDING MAIL BASED ON GRID
        public void Sending_Mail_Based_Grid()
        {
            try
            {
                oGrid = objform.Items.Item("Item_2").Specific;

                SAPbouiCOM.ComboBox objcombo;
                objcombo = objform.Items.Item("Item_18").Specific;

                string Layout_File = "";
                string Layout_Path = "";
                string Move_Path = "";
                string Str = "";

                SAPbobsCOM.Recordset Orec;
                Orec = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    Str = "Select \"U_AV_LPATH\",(Select \"U_AV_FPATH\" from \"OUSR\" where \"USER_CODE\"='" + objSBOAPI.oCompany.UserName + "') AS \"U_AV_FPATH\",\"U_AV_Domain\",\"U_AV_UserName\",\"U_AV_Password\",\"U_AV_WorkStation\",(SELECT \"U_AV_LayPt\" FROM \"@AV_LPST\" WHERE \"Code\"='" + objcombo.Selected.Value + "') AS \"U_AV_LayPt\" from \"OADM\"";
                }
                else
                {
                    Str = "Select U_AV_LPATH,(Select U_AV_FPATH from OUSR where User_Code='" + objSBOAPI.oCompany.UserName + "')'U_AV_FPATH',U_AV_Domain,U_AV_UserName,U_AV_Password,U_AV_WorkStation," + "(SELECT U_AV_LayPt FROM [@AV_LPST] WHERE Code='" + objcombo.Selected.Value + "')'U_AV_LayPt' from OADM";
                }
                Orec.DoQuery(Str);
                if (Orec.RecordCount > 0)
                {
                    if (Orec.Fields.Item("U_AV_LPATH").Value != "" & Orec.Fields.Item("U_AV_FPATH").Value != "" & Orec.Fields.Item("U_AV_LayPt").Value != "" & Orec.Fields.Item("U_AV_UserName").Value != "" & Orec.Fields.Item("U_AV_Password").Value != "")
                    {
                        Layout_File = Orec.Fields.Item("U_AV_LayPt").Value;
                        Layout_Path = Orec.Fields.Item("U_AV_LPATH").Value;
                        Move_Path = Orec.Fields.Item("U_AV_FPATH").Value;

                        if (System.IO.Directory.Exists(Layout_Path) == false)
                        {
                            objSBOAPI.SBO_Appln.MessageBox("Layout Path is Missing");
                            return;
                        }
                        else if (System.IO.File.Exists(Layout_Path + @"\" + Layout_File) == false)
                        {
                            objSBOAPI.SBO_Appln.MessageBox("Layout File Path is Missing - Check the File");
                            return;
                        }

                        if (System.IO.Directory.Exists(Move_Path) == false)
                        {
                            objSBOAPI.SBO_Appln.MessageBox("Move Path is Missing");
                            return;
                        }
                        if (Orec.Fields.Item("U_AV_WorkStation").Value == "Y")
                        {
                            try
                            {
                                NetworkCredential cred = new NetworkCredential(Orec.Fields.Item("U_AV_UserName").Value, Cls_Encrypt_Helper.Decrypt(Orec.Fields.Item("U_AV_Password").Value));

                                CredentialCache theNetCache = new CredentialCache();
                                theNetCache.Add(new Uri(@"\\computer"), "Basic", cred);

                                var dirs = Directory.GetFiles(Layout_Path);
                            }
                            catch (Exception ex)
                            {
                                objform.Freeze(false);
                                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                return;
                            }
                        }
                        else
                        {
                            if (Orec.Fields.Item("U_AV_Domain").Value != "")
                            {
                                UNCAccessWithCredentials unc = new UNCAccessWithCredentials();
                                if (unc.NetUseWithCredentials(Layout_Path, Orec.Fields.Item("U_AV_UserName").Value, Orec.Fields.Item("U_AV_Domain").Value, Cls_Encrypt_Helper.Decrypt(Orec.Fields.Item("U_AV_Password").Value)) == false)
                                {
                                    objSBOAPI.SBO_Appln.MessageBox("Layout Path - Server Credentials Failed");
                                    return;
                                }
                            }
                            else
                            {
                                objSBOAPI.SBO_Appln.MessageBox("Domain is Missing. Please provide it in Credentials Setup");
                                return;
                            }
                        }
                    }
                    else
                    {
                        objSBOAPI.SBO_Appln.MessageBox("Server and Folder Details are Missing. Please provide it in Credentials Setup");
                        return;
                    }
                }
                else
                {
                    objSBOAPI.SBO_Appln.MessageBox("Server and Folder Details are Missing. Please provide it in Credentials Setup");
                    return;
                }

                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    Str = "Select \"U_AV_SMTPS\",\"U_AV_EMID\",\"U_AV_Domain\",\"U_AV_UserName\",\"U_AV_Password\",\"U_AV_Port\",\"U_AV_EPswd\",\"U_AV_Essl\" from \"OADM\"";
                }
                else
                {
                    Str = "Select U_AV_SMTPS,U_AV_EMID,U_AV_Domain,U_AV_UserName,U_AV_Password,U_AV_Port,U_AV_EPswd,U_AV_Essl from OADM";
                }
                Orec.DoQuery(Str);
                if (Orec.RecordCount > 0)
                {
                    FromMailId = Orec.Fields.Item("U_AV_EMID").Value;
                    SMTPServer = Orec.Fields.Item("U_AV_SMTPS").Value;
                    Port = Orec.Fields.Item("U_AV_Port").Value != "" ? Convert.ToInt32(Orec.Fields.Item("U_AV_Port").Value) : 0;
                    MailPwd = Cls_Encrypt_Helper.Decrypt(Orec.Fields.Item("U_AV_EPswd").Value);
                    EnableSSL = Orec.Fields.Item("U_AV_Essl").Value == "Y" ? true : false;

                    SAPbouiCOM.ComboBox oCombo;
                    oCombo = objform.Items.Item("Item_16").Specific;

                    switch (oCombo.Value)
                    {
                        case "AR":
                            ObjType = "13";
                            break;

                        case "CN":
                            ObjType = "14";
                            break;

                        case "GP":
                            ObjType = "20";
                            break;

                        case "OP":
                            ObjType = "46";
                            break;

                        case "CS":
                            ObjType = "CustomerStatement";
                            break;

                        default:
                            ObjType = "0";
                            break;
                    }
                    if (ObjType == "CustomerStatement")
                    {
                        Subject = objSBOAPI.Query_Execute("Select \"U_AVA_EmailSbj\" from \"OADM\"");
                        Body = objSBOAPI.Query_Execute("Select \"U_AVA_EmailBody\" from \"OADM\"");
                    }
                    else if (ObjType != "0")
                    {
                        Subject = objSBOAPI.Query_Execute("Select \"EmailSbj\" from \"ADP1\" where \"ObjType\" ='" + ObjType.ToString() + "'");
                        Body = objSBOAPI.Query_Execute("Select \"EmailBody\" from \"ADP1\" where \"ObjType\" ='" + ObjType.ToString() + "'");
                    }
                }

                SAPbouiCOM.CheckBox oChk1, oChk2;
                oChk1 = objform.Items.Item("Item_27").Specific;
                oChk2 = objform.Items.Item("Item_28").Specific;

                if (oChk1.Checked == true & oChk2.Checked == true)
                    Type = "AL";
                else if (oChk1.Checked == true & oChk2.Checked == true)
                    Type = "PE";
                else if (oChk1.Checked == true)
                    Type = "PF";

                SAPbouiCOM.CheckBox oChk4;
                oChk4 = objform.Items.Item("Item_30").Specific;
                string Query = "";
                string[] Rows = arr;
                if (oChk4.Checked == true && (ObjType != "CustomerStatement"))
                {
                    if (oGrid.DataTable.IsEmpty == false)
                    {
                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        {
                            Query = objSBOAPI.Query_Execute("SELECT \"QString\" FROM \"OUQR\" WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM \"OQCN\" WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='CSGrouping_HANA'");
                        }
                        else
                        {
                            Query = objSBOAPI.Query_Execute("SELECT QString FROM OUQR WHERE QCategory=(SELECT CategoryId FROM OQCN WHERE CatName='AddonQuery_SQL') AND QName='CSGrouping'");
                        }
                        if (Query == "")
                        {
                            objSBOAPI.SBO_Appln.StatusBar.SetText("No Data are Found - Check the CSGrouping Addon Query in QM", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return;
                        }
                        else
                        {
                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                            {
                                objSBOAPI.Query_Execute("Delete from \"@AV_GRPTMP\" where \"Name\"='" + objSBOAPI.oCompany.UserName + "'");
                            }
                            else
                            {
                                objSBOAPI.Query_Execute("Delete from [@AV_GrpTmp] where Name='" + objSBOAPI.oCompany.UserName + "'");
                            }
                            foreach (var i in Rows)
                            {
                                if (i != "")
                                {
                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait Datas are inserting into temp table for grouping...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                    string oChk = Dt.GetValue("CK_Select", Convert.ToInt32(i));
                                    string EmailID = Dt.GetValue("E_Email-ID", Convert.ToInt32(i));
                                    if (EmailID == "")
                                    {
                                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                        {
                                            EmailID = objSBOAPI.Query_Execute("Select \"E_MailL\" from \"OCPR\" where \"CardCode\"='" + Dt.GetValue("NES_CardCode", Convert.ToInt32(i)) + "' and \"Name\"='" + Dt.GetValue("NE_ContactPerson", Convert.ToInt32(i)) + "'");
                                        }
                                        else
                                        {
                                            EmailID = objSBOAPI.Query_Execute("Select E_MailL from OCPR where CardCode='" + Dt.GetValue("NES_CardCode", Convert.ToInt32(i)) + "' and Name='" + Dt.GetValue("NE_ContactPerson", Convert.ToInt32(i)) + "'");
                                        }
                                    }

                                    SAPbouiCOM.ComboBox oCombo, oCombo1;
                                    oCombo = objform.Items.Item("Item_16").Specific;
                                    oCombo1 = objform.Items.Item("Item_18").Specific;
                                    string Insert_Query = "";
                                    string Code = objSBOAPI.Query_Execute("Select IFNULL(MAX(\"Code\"),0)+1 from \"@AV_GRPTMP\"");
                                    if (oChk == "Y")
                                    {
                                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                        {
                                            Insert_Query = @"INSERT INTO ""@AV_GRPTMP""(""Code"",""Name"",""U_AV_FrmDt"",""U_AV_ToDate"",""U_AV_RpTyp"",""U_AV_LtTyp"",""U_AV_Type"",""U_AV_Date"",""U_AV_DocEnt"",""U_AV_DocNm"",""U_AV_CrdCd"",""U_AV_CrdNm"",""U_AV_CntPn"",""U_AV_GrpNm"",""U_AV_DocTot"",""U_AV_PosDt"",""U_AV_DueDt"",""U_AV_EmlId"",""U_AV_GrdRw"",""U_AV_CCmlId"")VALUES('" + Code + "','" + objSBOAPI.oCompany.UserName + "','" + objSBOAPI.GetDateTimeValue(objform.Items.Item("Item_8").Specific.@string).ToString("yyyyMMdd") + "','" + objSBOAPI.GetDateTimeValue(objform.Items.Item("Item_11").Specific.@string).ToString("yyyyMMdd") + "','" + oCombo.Value.Trim() + "','" + oCombo1.Value.Trim() + "','" + Type + "','" + objSBOAPI.oCompany.GetDBServerDate().ToString("yyyyMMdd") + "','" + Dt.GetValue("NE_DocEntry", Convert.ToInt32(i)) + "','" + Dt.GetValue("NES_DocNum", Convert.ToInt32(i)) + "','" + Dt.GetValue("NES_CardCode", Convert.ToInt32(i)) + "','" + Dt.GetValue("NES_CardName", Convert.ToInt32(i)) + @"','" + Dt.GetValue("NE_ContactPerson", Convert.ToInt32(i)) + "', '" + Dt.GetValue("NE_GroupName", Convert.ToInt32(i)) + "','" + Dt.GetValue("NE_DocTotal", Convert.ToInt32(i)) + "', '" + Dt.GetValue("NES_Posting Date", Convert.ToInt32(i)).ToString("yyyyMMdd") + "', '" + Dt.GetValue("NE_Due Date", Convert.ToInt32(i)).ToString("yyyyMMdd") + "','" + EmailID + "','" + i + "','" + Dt.GetValue("E_CC Mail-ID", Convert.ToInt32(i)) + "')";
                                        }
                                        else
                                        {
                                            Insert_Query = @"Insert into [@AV_GrpTmp](Code,Name,U_AV_FrmDt,U_AV_ToDate,U_AV_RpTyp,U_AV_LtTyp,U_AV_Type,U_AV_Date,U_AV_DocEnt,U_AV_DocNm,U_AV_CrdCd,U_AV_CrdNm,U_AV_CntPn,U_AV_GrpNm,U_AV_DocTot,U_AV_PosDt,U_AV_DueDt,U_AV_EmlId,U_AV_GrdRw,U_AV_CCmlId)values('" + Code + "','" + objSBOAPI.oCompany.UserName + "','" + objSBOAPI.GetDateTimeValue(objform.Items.Item("Item_8").Specific.@string).ToString("yyyyMMdd") + "','" + objSBOAPI.GetDateTimeValue(objform.Items.Item("Item_11").Specific.@string).ToString("yyyyMMdd") + "','" + oCombo.Value.Trim() + "','" + oCombo1.Value.Trim() + "','" + Type + "','" + objSBOAPI.oCompany.GetDBServerDate() + "','" + Dt.GetValue("NE_DocEntry", Convert.ToInt32(i)) + "','" + Dt.GetValue("NES_DocNum", Convert.ToInt32(i)) + "','" + Dt.GetValue("NES_CardCode", Convert.ToInt32(i)) + "','" + Dt.GetValue("NES_CardName", Convert.ToInt32(i)) + @"','" + Dt.GetValue("NE_ContactPerson", Convert.ToInt32(i)) + "', '" + Dt.GetValue("NE_GroupName", Convert.ToInt32(i)) + "','" + Dt.GetValue("NE_DocTotal", Convert.ToInt32(i)) + "', '" + Dt.GetValue("NES_Posting Date", Convert.ToInt32(i)) + "', '" + Dt.GetValue("NE_Due Date", Convert.ToInt32(i)) + "','" + EmailID + "','" + i + "','" + Dt.GetValue("E_CC Mail-ID", Convert.ToInt32(i)) + "')";
                                        }
                                        objSBOAPI.Query_Execute(Insert_Query);
                                    }
                                }
                            }

                            SAPbobsCOM.Recordset oRecord = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            Int64 count = 0;

                            Query = Query.Replace("[%1]", objSBOAPI.oCompany.UserName);

                            oRecord.DoQuery(Query);
                            if (oRecord.RecordCount > 0)
                            {
                                SAPbouiCOM.ComboBox oCombo;
                                oCombo = objform.Items.Item("Item_16").Specific;
                                while (!oRecord.EoF)
                                {
                                    string EmailID = oRecord.Fields.Item("EmailID").Value;
                                    count = count + 1;

                                    if (BtnType == "Print")
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait.... Data is checking to print... - " + count + "", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                        Print_Invoice(Layout_Path + "\\" + Layout_File, Move_Path, oRecord.Fields.Item("DocEnt").Value.ToString(), 1, EmailID, oRecord.Fields.Item("CardCode").Value.ToString(), oRecord.Fields.Item("ContactPerson").Value.ToString());
                                    }
                                    else
                                    {
                                        if (BtnType == "Mail")
                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait.... Data is checking to send the Email... - " + count + "", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                        else
                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait.... Data is saving as PDF... - " + count + "", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                        Save_As_PDF(Layout_Path + "\\" + Layout_File, Move_Path, oRecord.Fields.Item("DocEnt").Value.ToString(), EmailID, 1, oRecord.Fields.Item("CardCode").Value.ToString(), oRecord.Fields.Item("ContactPerson").Value.ToString(), oRecord.Fields.Item("CCMailID").Value.ToString());
                                    }
                                    oRecord.MoveNext();
                                }
                                objSBOAPI.SBO_Appln.StatusBar.SetText("Notification Process will be Finished successfully - Check the Status Tab", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                                Grouping_Status();
                            }
                        }
                    }
                    else
                        objSBOAPI.SBO_Appln.StatusBar.SetText("No Datas are found", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
                else
                {
                    if (oGrid.DataTable.IsEmpty == false)
                    {
                        foreach (var i in Rows)
                        {
                            if (i != "")
                            {
                                string oChk = oGrid.DataTable.GetValue("CK_Select", Convert.ToInt32(i));
                                string EmailID = oGrid.DataTable.GetValue("E_Email-ID", Convert.ToInt32(i));
                                if (oChk == "Y")
                                {
                                    if (EmailID == "" && (Type == "AL" || Type == "PE") && BtnType == "Mail")
                                    {
                                        objform.Freeze(true);
                                        Dt.SetValue("E_Status", Convert.ToInt32(i), "Email ID is Missing");
                                        objform.Freeze(false);
                                        GeneralService_For_UDO_CustomerStatement_AddMethod(Dt.GetValue("E_Status", Convert.ToInt32(i)), Convert.ToInt32(i), "", "F");
                                    }
                                    else if (BtnType == "Print")
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait.... Data is checking to print... - " + (Convert.ToInt32(i) + 1) + "", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                                        if (ObjType == "CustomerStatement")
                                        {
                                            Print_Invoice(Layout_Path + "\\" + Layout_File, Move_Path, Dt.GetValue("NES_CardCode", Convert.ToInt32(i)).ToString(), Convert.ToInt32(i), EmailID, "", "");
                                        }
                                        else
                                        {
                                            Print_Invoice(Layout_Path + "\\" + Layout_File, Move_Path, Dt.GetValue("NE_DocEntry", Convert.ToInt32(i)).ToString(), Convert.ToInt32(i), EmailID, "", "");
                                        }
                                    }
                                    else
                                    {
                                        if (BtnType == "Mail")
                                        {
                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait.... Data is checking to send the Email... - " + (Convert.ToInt32(i) + 1) + "", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                        }
                                        else
                                        {
                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait.... Data is saving as PDF... - " + (Convert.ToInt32(i) + 1) + "", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                        }

                                        if (ObjType == "CustomerStatement")
                                        {
                                            Save_As_PDF(Layout_Path + "\\" + Layout_File, Move_Path, Dt.GetValue("NES_CardCode", Convert.ToInt32(i)).ToString(), EmailID, Convert.ToInt32(i), "", "", Dt.GetValue("E_CC Mail-ID", Convert.ToInt32(i)).ToString());
                                            //Save_As_PDF(Layout_Path + "\\" + Layout_File, Move_Path, "1", EmailID, Convert.ToInt32(i), "", "", Dt.GetValue("E_CC Mail-ID", Convert.ToInt32(i)).ToString());
                                        }
                                        else
                                        {
                                            Save_As_PDF(Layout_Path + "\\" + Layout_File, Move_Path, Dt.GetValue("NE_DocEntry", Convert.ToInt32(i)).ToString(), EmailID, Convert.ToInt32(i), "", "", Dt.GetValue("E_CC Mail-ID", Convert.ToInt32(i)).ToString());
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        objSBOAPI.SBO_Appln.StatusBar.SetText("No Datas are found", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }

                    objSBOAPI.SBO_Appln.StatusBar.SetText("Notification Process will be Finished successfully - Check the Status Tab", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region SAVE AS PDF
        public void Save_As_PDF(string File_Path, string Move_Path, string DocNum, string Email_ID, int i, string Group_Code, string Cntct_Prsn, string CCMailID)
        {
            oGrid = objform.Items.Item("Item_2").Specific;
            string oRPTConn = ConfigurationManager.AppSettings["RPTConn"];
            try
            {
                //objSBOAPI.objMain.HANA_UserID = "SYSTEM";
                //objSBOAPI.objMain.HANA_Pwd = "India@1947";
                //objSBOAPI.oCompany.CompanyDB = objSBOAPI.oCompany.CompanyDB;

                if (objSBOAPI.objMain.HANA_UserID != "" && objSBOAPI.objMain.HANA_Pwd != "")
                {
                    oRPTConn = oRPTConn.Replace("{0}", objSBOAPI.objMain.HANA_UserID);
                    oRPTConn = oRPTConn.Replace("{1}", objSBOAPI.objMain.HANA_Pwd);
                    oRPTConn = oRPTConn.Replace("{2}", objSBOAPI.oCompany.CompanyDB);

                    if (System.IO.File.Exists(File_Path) == true)
                    {
                        SAPbouiCOM.CheckBox oChk = objform.Items.Item("Item_30").Specific;
                       
                        DateTime DT1 = DateTime.Now;
                        string dateString;
                        dateString = DT1.ToString("yyyyMMddhhmmss");

                        string PDF_FileName = System.IO.Path.GetFileNameWithoutExtension(File_Path);
                        string PDF_Saving_Path = "";
                        if (oChk.Checked == true)
                        {
                            PDF_Saving_Path = Move_Path + @"\" + PDF_FileName + "_" + Group_Code + "_" + Cntct_Prsn + "_" + dateString + ".pdf";
                        }
                        else
                        {
                            PDF_Saving_Path = Move_Path + @"\" + PDF_FileName + "_" + DocNum + "_" + dateString + ".pdf";
                        }

                        ReportDocument rpt = new ReportDocument();
                        rpt.Load(File_Path);

                        TableLogOnInfo crtableLogoninfo = new TableLogOnInfo();
                        ConnectionInfo crConnectionInfo = new ConnectionInfo();
                        NameValuePairs2 logonProps2;

                        rpt.SetDatabaseLogon(objSBOAPI.objMain.HANA_UserID, objSBOAPI.objMain.HANA_Pwd, objSBOAPI.oCompany.Server, oRPTConn);

                        CrystalDecisions.Shared.ParameterValues pval1 = new ParameterValues();
                        ParameterDiscreteValue pdisval1 = new ParameterDiscreteValue();
                        pdisval1.Value = DocNum;
                        pval1.Add(pdisval1);

                        CrystalDecisions.Shared.ParameterValues pval2 = new ParameterValues();
                        ParameterDiscreteValue pdisval2 = new ParameterDiscreteValue();
                        pdisval2.Value = objSBOAPI.GetDateTimeValue(objSBOAPI.GetDateFromField(objform.Items.Item("Item_8").Specific.value));
                        pval2.Add(pdisval2);

                        CrystalDecisions.Shared.ParameterValues pval3 = new ParameterValues();
                        ParameterDiscreteValue pdisval3 = new ParameterDiscreteValue();
                        pdisval3.Value = objSBOAPI.GetDateTimeValue(objSBOAPI.GetDateFromField(objform.Items.Item("Item_11").Specific.value));
                        pval3.Add(pdisval3);

                        if (ObjType == "CustomerStatement")
                        {
                            rpt.DataDefinition.ParameterFields["@CC"].ApplyCurrentValues(pval1);
                            rpt.DataDefinition.ParameterFields["@FromDate"].ApplyCurrentValues(pval2);
                            rpt.DataDefinition.ParameterFields["@ToDate"].ApplyCurrentValues(pval3);
                        }
                        else
                        {
                            rpt.DataDefinition.ParameterFields["@DocNum"].ApplyCurrentValues(pval1);
                        }

                        NameValuePair2 r1;

                        logonProps2 = rpt.DataSourceConnections[0].LogonProperties;
                        r1 = (NameValuePair2)logonProps2[0];
                        r1.Value = oRPTConn;
                        rpt.DataSourceConnections[0].SetLogonProperties(logonProps2);
                        rpt.DataSourceConnections[0].SetConnection(objSBOAPI.oCompany.Server, objSBOAPI.oCompany.CompanyDB, objSBOAPI.objMain.HANA_UserID, objSBOAPI.objMain.HANA_Pwd);

                        CrystalDecisions.CrystalReports.Engine.ReportObjects crReportObjects;
                        CrystalDecisions.CrystalReports.Engine.SubreportObject crSubreportObject;

                        CrystalDecisions.CrystalReports.Engine.ReportDocument crSubreportDocument;
                        CrystalDecisions.CrystalReports.Engine.Database crDatabase;
                        CrystalDecisions.CrystalReports.Engine.Tables crTables;

                        foreach (CrystalDecisions.CrystalReports.Engine.Section crSection in rpt.ReportDefinition.Sections)
                        {
                            crReportObjects = crSection.ReportObjects;

                            foreach (CrystalDecisions.CrystalReports.Engine.ReportObject crReportObject in crReportObjects)
                            {
                                if (crReportObject.Kind == ReportObjectKind.SubreportObject)
                                {
                                    crSubreportObject = (CrystalDecisions.CrystalReports.Engine.SubreportObject)crReportObject;
                                    crSubreportDocument = crSubreportObject.OpenSubreport(crSubreportObject.SubreportName);
                                    crDatabase = crSubreportDocument.Database;
                                    crTables = crDatabase.Tables;

                                    if ((crSubreportDocument.DataSourceConnections.Count > 0))
                                    {
                                        logonProps2 = crSubreportDocument.DataSourceConnections[0].LogonProperties;
                                        NameValuePair2 r;
                                        r = (NameValuePair2)logonProps2[0];
                                        r.Value = oRPTConn;
                                        crSubreportDocument.DataSourceConnections[0].SetLogonProperties(logonProps2);
                                        crSubreportDocument.DataSourceConnections[0].SetConnection(objSBOAPI.oCompany.Server, objSBOAPI.oCompany.CompanyDB, objSBOAPI.objMain.HANA_UserID, objSBOAPI.objMain.HANA_Pwd);

                                        Tables rptTables = crDatabase.Tables;
                                        for (int j = 0; j < rptTables.Count; j++)
                                        {
                                            CrystalDecisions.CrystalReports.Engine.Table rptTable = rptTables[j];
                                            TableLogOnInfo tblInfo = rptTable.LogOnInfo;
                                            System.Diagnostics.Debug.Print(tblInfo.ConnectionInfo.LogonProperties[0].ToString());
                                            // next table
                                        }
                                    }
                                }
                            }
                        }

                        ExportOptions rptExportOption;
                        DiskFileDestinationOptions rptFileDestOption = new DiskFileDestinationOptions();
                        PdfRtfWordFormatOptions rptFormatOption = new PdfRtfWordFormatOptions();
                        rptFileDestOption.DiskFileName = PDF_Saving_Path;
                        rptExportOption = rpt.ExportOptions;
                        rptExportOption.ExportDestinationType = ExportDestinationType.DiskFile;
                        rptExportOption.ExportFormatType = ExportFormatType.PortableDocFormat;
                        rptExportOption.ExportDestinationOptions = rptFileDestOption;
                        rptExportOption.ExportFormatOptions = rptFormatOption;

                        rpt.Export();
                        rpt.Dispose();
                        rpt.Close();

                        if ((Type == "AL" || Type == "PE") && BtnType == "Mail")
                        {
                            Mail_Service(Email_ID, PDF_Saving_Path, i, Group_Code, CCMailID);
                        }
                        else if (oChk.Checked == false)
                        {
                            objform.Freeze(true);
                            Dt.SetValue("E_Status", i, "PDF Saved Successfully");
                            objform.Freeze(false);
                            GeneralService_For_UDO_CustomerStatement_AddMethod(Dt.GetValue("E_Status", i), i, PDF_Saving_Path, "S");
                        }
                        else
                            Group_Log(Group_Code, PDF_Saving_Path, "PDF Saved Successfully", Email_ID, "S");
                    }
                    else
                    {
                        objform.Freeze(true);
                        SAPbouiCOM.CheckBox oChk = objform.Items.Item("Item_30").Specific;
                        if (oChk.Checked == false)
                        {
                            objform.Freeze(true);
                            Dt.SetValue("E_Status", i, "File Path is Missing - Check the File");
                            objform.Freeze(false);
                            GeneralService_For_UDO_CustomerStatement_AddMethod(Dt.GetValue("E_Status", i), i, "", "F");
                        }
                        else
                            Group_Log(Group_Code, "", "File Path is Missing - Check the File", Email_ID, "F");
                        objform.Freeze(false);
                    }
                }
                else
                {
                    objform.Freeze(true);
                    SAPbouiCOM.CheckBox oChk = objform.Items.Item("Item_30").Specific;
                    if (oChk.Checked == false)
                    {
                        objform.Freeze(true);
                        Dt.SetValue("E_Status", i, "DataBase Details are Missing. Please provide it in Credentials Setup");
                        objform.Freeze(false);
                        GeneralService_For_UDO_CustomerStatement_AddMethod(Dt.GetValue("E_Status", i), i, "", "F");
                    }
                    else
                        Group_Log(Group_Code, "", "DataBase Details are Missing. Please provide it in Credentials Setup", Email_ID, "F");
                    objform.Freeze(false);
                }

            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                SAPbouiCOM.CheckBox oChk = objform.Items.Item("Item_30").Specific;
                if (oChk.Checked == false)
                {
                    objform.Freeze(true);
                    Dt.SetValue("E_Status", i, ex.Message.ToString());
                    objform.Freeze(false);
                    GeneralService_For_UDO_CustomerStatement_AddMethod(Dt.GetValue("E_Status", i), i, "", "F");
                }
                else
                {
                    Group_Log(Group_Code, "", ex.Message.ToString(), Email_ID, "F");
                }
            }
        }
        #endregion

        #region MAIL SERVICE

        public void Mail_Service(string EmailID, string File_Name, int i, string Group_Code, string CCMailID/*, string Excel_Path*/)
        {
            oGrid = objform.Items.Item("Item_2").Specific;
            try
            {
                if (SMTPServer != "" && FromMailId != "" && MailPwd != "")
                {
                    string ToMailId = EmailID;
                    try
                    {
                        SAPbouiCOM.ComboBox oCombo;
                        oCombo = objform.Items.Item("Item_16").Specific;

                        if (ObjType != "CustomerStatement")
                        {
                            if (Group_Code != "")
                            {
                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                {
                                    Body = Body.Replace("[%1]", objSBOAPI.Query_Execute("Select \"CardName\" from \"OCRD\" where \"CardCode\"='" + Group_Code + "'"));
                                }
                                else
                                {
                                    Body = Body.Replace("[%1]", objSBOAPI.Query_Execute("Select CardName from OCRD where CardCode='" + Group_Code + "'"));
                                }
                            }
                            else
                            {
                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                {
                                    Body = Body.Replace("[%1]", objSBOAPI.Query_Execute("Select \"CardName\" from \"OCRD\" where \"CardCode\"='" + oGrid.DataTable.Columns.Item("NES_CardCode").Cells.Item(oGrid.GetDataTableRowIndex(i)).Value + "'"));
                                }
                                else
                                {
                                    Body = Body.Replace("[%1]", objSBOAPI.Query_Execute("Select CardName from OCRD where CardCode='" + oGrid.DataTable.Columns.Item("NES_CardCode").Cells.Item(oGrid.GetDataTableRowIndex(i)).Value + "'"));
                                }
                            }
                        }

                        MailMessage message = new MailMessage(FromMailId.Trim(), ToMailId.Trim(), Subject, Body);
                        if (CCMailID != "")
                        {
                            message.CC.Add(CCMailID);
                        }
                        if (File_Name != "")
                        {
                            message.Attachments.Add(new Attachment(File_Name));
                        }

                        SmtpClient emailClient = new SmtpClient(SMTPServer, Convert.ToInt32(Port));
                        emailClient.EnableSsl = EnableSSL;

                        message.Priority = MailPriority.High;
                        NetworkCredential SMTPUserInfo = new NetworkCredential(FromMailId, MailPwd);
                        emailClient.UseDefaultCredentials = false;
                        emailClient.Credentials = SMTPUserInfo;
                        ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(customCertValidation);
                        emailClient.Send(message);
                        message.Attachments.Dispose();
                        // System.IO.File.Delete(File_Name)
                        SAPbouiCOM.CheckBox oChk = objform.Items.Item("Item_30").Specific;
                        if (oChk.Checked == false)
                        {
                            objform.Freeze(true);
                            Dt.SetValue("E_Status", i, "Email Sent Successfully");
                            objform.Freeze(false);
                            GeneralService_For_UDO_CustomerStatement_AddMethod(Dt.GetValue("E_Status", i), i, File_Name, "S");
                        }
                        else
                        {
                            Group_Log(Group_Code, File_Name, "Email Sent Successfully", EmailID, "S");
                        }
                    }
                    catch (Exception ex)
                    {
                        SAPbouiCOM.CheckBox oChk = objform.Items.Item("Item_30").Specific;
                        if (oChk.Checked == false)
                        {
                            objform.Freeze(true);
                            Dt.SetValue("E_Status", i, ex.Message.ToString());
                            objform.Freeze(false);
                            GeneralService_For_UDO_CustomerStatement_AddMethod(Dt.GetValue("E_Status", i), i, File_Name, "F");
                        }
                        else
                        {
                            Group_Log(Group_Code, File_Name, ex.Message.ToString(), EmailID, "F");
                        }
                    }
                    finally
                    {
                    }
                }
                else
                {
                    objform.Freeze(true);
                    SAPbouiCOM.CheckBox oChk = objform.Items.Item("Item_30").Specific;
                    if (oChk.Checked == false)
                    {
                        objform.Freeze(true);
                        Dt.SetValue("E_Status", i, "Email is not sent, Please Check the Details in Credentials Setup");
                        objform.Freeze(false);
                        GeneralService_For_UDO_CustomerStatement_AddMethod(Dt.GetValue("E_Status", i), i, File_Name, "F");
                    }
                    else
                        Group_Log(Group_Code, File_Name, "Email is not sent, Please Check the Details in Credentials Setup", EmailID, "F");
                    objform.Freeze(false);
                }
            }
            catch (Exception ex)
            {
                SAPbouiCOM.CheckBox oChk = objform.Items.Item("Item_30").Specific;
                if (oChk.Checked == false)
                {
                    objform.Freeze(true);
                    Dt.SetValue("E_Status", i, ex.Message.ToString());
                    objform.Freeze(false);
                    GeneralService_For_UDO_CustomerStatement_AddMethod(Dt.GetValue("E_Status", i), i, File_Name, "F");
                }
                else
                    Group_Log(Group_Code, "", ex.Message.ToString(), EmailID, "F");
            }

        }
        public static bool customCertValidation(object sender, X509Certificate cert, X509Chain chain, SslPolicyErrors errors)
        {
            return true;
        }
        #endregion

        #region System Threading
        public bool System_threading()
        {
            try
            {
                System.Threading.Timer tm = new System.Threading.Timer(new System.Threading.TimerCallback(KeepUIAlive));
                tm.Change(1000 * 60, 0);
                System.Threading.Thread t = new System.Threading.Thread(new System.Threading.ThreadStart(Sending_Mail_Based_Grid)); // this calls the NET form but it could be just a heavy load method.  
                t.TrySetApartmentState(System.Threading.ApartmentState.STA);
                t.Start();
                t.Join();
                tm.Dispose();
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        public void KeepUIAlive(object info)
        {
            System.Threading.Timer t = (System.Threading.Timer)info;
            objSBOAPI.SBO_Appln.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, true);
            t.Change(1000 * 60, 0);
        }
        #endregion

        #region General Service For UDO Customer Statement Add Method
        public void GeneralService_For_UDO_CustomerStatement_AddMethod(string Status, int k, string PDFPath, string MainStatus)
        {
            SAPbobsCOM.GeneralService oGeneralService;
            SAPbobsCOM.GeneralData oGeneralData;

            try
            {
                SAPbobsCOM.CompanyService sCmp;
                sCmp = objSBOAPI.oCompany.GetCompanyService();

                Status = Status.Replace("'", "\"");

                SAPbouiCOM.ComboBox oCombo, oCombo1;
                oCombo = objform.Items.Item("Item_16").Specific;
                oCombo1 = objform.Items.Item("Item_18").Specific;

                oGrid = objform.Items.Item("Item_2").Specific;

                oGeneralService = sCmp.GetGeneralService("AV_CSLGT");

                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    oGeneralData.SetProperty("Code", "" + objSBOAPI.Query_Execute("Select IFNULL(Max(Cast(\"Code\" as Int)),0)+1 from \"@AV_CSLGT\"") + "");
                }
                else
                {
                    oGeneralData.SetProperty("Code", "" + objSBOAPI.Query_Execute("Select Isnull(Max(Cast(Code as Int)),0)+1 from [@AV_CSLGT]") + "");
                }
                oGeneralData.SetProperty("Name", "" + objSBOAPI.oCompany.UserName + "");
                oGeneralData.SetProperty("U_AV_FrmDt", "" + objform.Items.Item("Item_8").Specific.value + "");
                oGeneralData.SetProperty("U_AV_ToDate", "" + objform.Items.Item("Item_11").Specific.value + "");
                oGeneralData.SetProperty("U_AV_RpTyp", "" + oCombo.Value.Trim() + "");
                oGeneralData.SetProperty("U_AV_LtTyp", "" + oCombo1.Value.Trim() + "");
                oGeneralData.SetProperty("U_AV_Type", "" + Type + "");
                oGeneralData.SetProperty("U_AV_Date", "" + objSBOAPI.oCompany.GetDBServerDate() + "");

                if (ObjType == "CustomerStatement")
                {
                    oGeneralData.SetProperty("U_AV_CrdCd", "" + Dt.GetValue("NES_CardCode", k) + "");
                    oGeneralData.SetProperty("U_AV_CrdNm", "" + Dt.GetValue("NES_CardName", k) + "");
                    oGeneralData.SetProperty("U_AV_CntPn", "" + Dt.GetValue("NE_ContactPerson", k) + "");

                    oGeneralData.SetProperty("U_AV_OpnBl", "" + Dt.GetValue("NE_Opening", k) + "");
                    oGeneralData.SetProperty("U_AV_ClsBl", "" + Dt.GetValue("NE_Closing", k) + "");
                    oGeneralData.SetProperty("U_AV_DebBl", "" + Dt.GetValue("NE_Debit", k) + "");
                    oGeneralData.SetProperty("U_AV_CrdBl", "" + Dt.GetValue("NE_Credit", k) + "");
                }
                else
                {
                    oGeneralData.SetProperty("U_AV_CrdCd", "" + Dt.GetValue("NES_CardCode", k) + "");
                    oGeneralData.SetProperty("U_AV_CrdNm", "" + Dt.GetValue("NES_CardName", k) + "");
                    oGeneralData.SetProperty("U_AV_CntPn", "" + Dt.GetValue("NE_ContactPerson", k) + "");
                    oGeneralData.SetProperty("U_AV_GrpNm", "" + Dt.GetValue("NE_GroupName", k) + "");
                    oGeneralData.SetProperty("U_AV_DocEnt", "" + Dt.GetValue("NE_DocEntry", k) + "");
                    oGeneralData.SetProperty("U_AV_DocNm", "" + Dt.GetValue("NES_DocNum", k) + "");
                    oGeneralData.SetProperty("U_AV_DocTot", "" + Dt.GetValue("NE_DocTotal", k) + "");

                    oGeneralData.SetProperty("U_AV_PosDt", "" + Dt.GetValue("NES_Posting Date", k) + "");
                    oGeneralData.SetProperty("U_AV_DueDt", "" + Dt.GetValue("NE_Due Date", k) + "");
                }

                oGeneralData.SetProperty("U_AV_EmlId", "" + Dt.GetValue("E_Email-ID", k) + "");
                oGeneralData.SetProperty("U_AV_CCmlId", "" + Dt.GetValue("E_CC Mail-ID", k) + "");
                oGeneralData.SetProperty("U_AV_Stats", "" + Status + "");
                oGeneralData.SetProperty("U_AV_PDFPath", "" + PDFPath + "");
                oGeneralData.SetProperty("U_AV_Status", "" + MainStatus + "");
                oGeneralData.SetProperty("U_AV_MailType", "Man");
                oGeneralService.Add(oGeneralData);
            }
            catch (Exception ex)
            {
                objSBOAPI.SBO_Appln.MessageBox(ex.Message);
            }

        }

        #endregion

        #region For Grouping Insert into Log
        public void Group_Log(string CardCode, string PDF_Saving_Path, string Status, string EmailID, string MainStatus)
        {
            try
            {
                string Str = "";
                SAPbobsCOM.Recordset oRec = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    Str = "Select \"U_AV_GrdRw\" AS \"Row\" from \"@AV_GRPTMP\" where \"U_AV_CrdCd\"='" + CardCode + "' ";
                }
                else
                {
                    Str = "Select U_AV_GrdRw[Row] from [@AV_GRPTMP] where U_AV_CrdCd='" + CardCode + "' ";
                }
                oRec.DoQuery(Str);
                if (oRec.RecordCount > 0)
                {
                    while (!oRec.EoF)
                    {
                        GeneralService_For_UDO_CustomerStatement_AddMethod(Status, oRec.Fields.Item("Row").Value, PDF_Saving_Path, MainStatus);

                        oRec.MoveNext();
                    }
                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        objSBOAPI.Query_Execute("Update \"@AV_GRPTMP\" Set \"U_AV_Stats\"='" + Status + "',\"U_AV_PDFPath\"='" + PDF_Saving_Path + "',\"U_AV_Status\"='" + MainStatus + "' where \"U_AV_CrdCd\"='" + CardCode + "' and \"U_AV_EmlId\"='" + EmailID + "'");
                    }
                    else
                    {
                        objSBOAPI.Query_Execute("Update [@AV_GrpTmp] Set U_AV_Stats='" + Status + "',U_AV_PDFPath='" + PDF_Saving_Path + "',U_AV_Status='" + MainStatus + "' where U_AV_CrdCd='" + CardCode + "' and U_AV_EmlId='" + EmailID + "'");
                    }
                }
            }
            catch (Exception)
            {
            }
        }
        #endregion

        #region Grouping Status
        public void Grouping_Status()
        {
            SAPbouiCOM.Form oForm = null;
            string str = "";
            try
            {
                objform.Freeze(true);
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    str = "Select CAST(ROW_NUMBER() OVER ( ORDER BY \"U_AV_GrdRw\") AS varchar) AS \"#\",\"U_AV_CrdCd\" AS \"CardCode\",\"U_AV_DocEnt\" AS \"DocEntry\",\"U_AV_CntPn\" AS \"ContactPerson\",\"U_AV_EmlId\" AS \"Mail\",\"U_AV_Stats\" AS \"Status\",\"U_AV_PDFPath\" AS \"PDF Path\" from \"@AV_GRPTMP\" where \"Name\"='" + objSBOAPI.oCompany.UserName + "' ";
                }
                else
                {
                    str = "Select CONVERT(VARCHAR(10),ROW_NUMBER() OVER(Order By U_AV_GrdRw))'#',U_AV_CrdCd[CardCode],U_AV_DocEnt[DocEntry],U_AV_CntPn[ContactPerson],U_AV_EmlId[Mail],U_AV_Stats[Status],U_AV_PDFPath[PDF Path] from [@AV_GRPTMP] where Name='" + objSBOAPI.oCompany.UserName + "' ";
                }
                oForm = objSBOAPI.LoadForm("Group_Status.xml", "AV_STATF");
                oGrid = oForm.Items.Item("Item_0").Specific;
                oGrid.DataTable.ExecuteQuery(str);

                oForm.Freeze(true);
                if (oGrid.DataTable.IsEmpty == false)
                {
                    oGrid.Columns.Item("RowsHeader").Visible = false;

                    SAPbouiCOM.EditTextColumn oEdit;
                    oEdit = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("CardCode");
                    oEdit.LinkedObjectType = "2";

                    oGrid.Columns.Item("#").Editable = false;

                    oEdit = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("DocEntry");
                    oEdit.LinkedObjectType = "13";

                    oGrid.AutoResizeColumns();
                    oForm.Select();
                }
                else
                    objSBOAPI.SBO_Appln.StatusBar.SetText("No Datas are found", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                SAPbouiCOM.StaticText oST;
                oST = objform.Items.Item("ST1").Specific;
                oST.Caption = oForm.UniqueID;
                oForm.Freeze(false);
                objform.Freeze(false);
            }

            catch (Exception)
            {
                oForm.Freeze(false);
                objform.Freeze(false);
            }
        }
        #endregion

        #region Printing
        public void Print_Invoice(string File_Path, string Move_Path, string DocNum, int i, string Email_ID, string Group_Code, string Cntct_Prsn)
        {
            string oRPTConn = ConfigurationManager.AppSettings["RPTConn"];
            try
            {

                if (objSBOAPI.objMain.HANA_UserID != "" && objSBOAPI.objMain.HANA_Pwd != "")
                {
                    oRPTConn = oRPTConn.Replace("{0}", objSBOAPI.objMain.HANA_UserID);
                    oRPTConn = oRPTConn.Replace("{1}", objSBOAPI.objMain.HANA_Pwd);
                    oRPTConn = oRPTConn.Replace("{2}", objSBOAPI.oCompany.CompanyDB);

                    if (System.IO.File.Exists(File_Path) == true)
                    {
                        SAPbouiCOM.CheckBox oChk = objform.Items.Item("Item_30").Specific;
                        // '
                        DateTime DT1 = DateTime.Now;
                        string dateString;
                        dateString = DT1.ToString("yyyyMMddhhmmss");

                        string PDF_FileName = System.IO.Path.GetFileNameWithoutExtension(File_Path);
                        string PDF_Saving_Path = "";
                        if (oChk.Checked == true)
                            PDF_Saving_Path = Move_Path + @"\" + PDF_FileName + "_" + Group_Code + "_" + Cntct_Prsn + "_" + dateString + ".pdf";
                        else
                            PDF_Saving_Path = Move_Path + @"\" + PDF_FileName + "_" + DocNum + "_" + dateString + ".pdf";

                        ReportDocument rpt = new ReportDocument();
                        rpt.Load(File_Path);

                        TableLogOnInfo crtableLogoninfo = new TableLogOnInfo();
                        ConnectionInfo crConnectionInfo = new ConnectionInfo();
                        NameValuePairs2 logonProps2;

                        rpt.SetDatabaseLogon(objSBOAPI.objMain.HANA_UserID, objSBOAPI.objMain.HANA_Pwd, objSBOAPI.oCompany.Server, oRPTConn);

                        CrystalDecisions.Shared.ParameterValues pval1 = new ParameterValues();
                        ParameterDiscreteValue pdisval1 = new ParameterDiscreteValue();
                        pdisval1.Value = DocNum;
                        pval1.Add(pdisval1);

                        rpt.DataDefinition.ParameterFields["DocKey@"].ApplyCurrentValues(pval1);

                        NameValuePair2 r1;

                        logonProps2 = rpt.DataSourceConnections[0].LogonProperties;
                        r1 = (NameValuePair2)logonProps2[0];
                        r1.Value = oRPTConn;
                        rpt.DataSourceConnections[0].SetLogonProperties(logonProps2);
                        rpt.DataSourceConnections[0].SetConnection(objSBOAPI.oCompany.Server, objSBOAPI.oCompany.CompanyDB, objSBOAPI.objMain.HANA_UserID, objSBOAPI.objMain.HANA_Pwd);

                        CrystalDecisions.CrystalReports.Engine.ReportObjects crReportObjects;
                        CrystalDecisions.CrystalReports.Engine.SubreportObject crSubreportObject;

                        CrystalDecisions.CrystalReports.Engine.ReportDocument crSubreportDocument;
                        CrystalDecisions.CrystalReports.Engine.Database crDatabase;
                        CrystalDecisions.CrystalReports.Engine.Tables crTables;

                        foreach (CrystalDecisions.CrystalReports.Engine.Section crSection in rpt.ReportDefinition.Sections)
                        {
                            crReportObjects = crSection.ReportObjects;

                            foreach (CrystalDecisions.CrystalReports.Engine.ReportObject crReportObject in crReportObjects)
                            {
                                if (crReportObject.Kind == ReportObjectKind.SubreportObject)
                                {
                                    crSubreportObject = (CrystalDecisions.CrystalReports.Engine.SubreportObject)crReportObject;
                                    crSubreportDocument = crSubreportObject.OpenSubreport(crSubreportObject.SubreportName);
                                    crDatabase = crSubreportDocument.Database;
                                    crTables = crDatabase.Tables;

                                    if (crSubreportDocument.DataSourceConnections.Count > 0)
                                    {
                                        logonProps2 = crSubreportDocument.DataSourceConnections[0].LogonProperties;
                                        NameValuePair2 r;
                                        r = (NameValuePair2)logonProps2[0];
                                        r.Value = oRPTConn;
                                        crSubreportDocument.DataSourceConnections[0].SetLogonProperties(logonProps2);
                                        crSubreportDocument.DataSourceConnections[0].SetConnection(objSBOAPI.oCompany.Server, objSBOAPI.oCompany.CompanyDB, objSBOAPI.objMain.HANA_UserID, objSBOAPI.objMain.HANA_Pwd);

                                        Tables rptTables = crDatabase.Tables;
                                        for (int j = 0; j < rptTables.Count; j++)
                                        {
                                            CrystalDecisions.CrystalReports.Engine.Table rptTable = rptTables[j];
                                            TableLogOnInfo tblInfo = rptTable.LogOnInfo;
                                            System.Diagnostics.Debug.Print(tblInfo.ConnectionInfo.LogonProperties[0].ToString());
                                        }
                                    }
                                }
                            }
                        }

                        System.Drawing.Printing.PrintDocument localPrinter = new System.Drawing.Printing.PrintDocument();
                        rpt.PrintOptions.PrinterName = localPrinter.PrinterSettings.PrinterName;
                        rpt.PrintToPrinter(1, false, 0, 0);

                        rpt.Dispose();

                        if (oChk.Checked == false)
                        {
                            objform.Freeze(true);
                            Dt.SetValue("E_Status", i, "Printed Successfully");
                            objform.Freeze(false);
                            GeneralService_For_UDO_CustomerStatement_AddMethod(Dt.GetValue("E_Status", i), i, "", "S");
                        }
                        else
                            Group_Log(Group_Code, "", "Printed Successfully", Email_ID, "S");
                    }
                    else
                    {
                        objform.Freeze(true);
                        SAPbouiCOM.CheckBox oChk = objform.Items.Item("Item_30").Specific;
                        if (oChk.Checked == false)
                        {
                            objform.Freeze(true);
                            Dt.SetValue("E_Status", i, "File Path is Missing - Check the File");
                            objform.Freeze(false);
                            GeneralService_For_UDO_CustomerStatement_AddMethod(Dt.GetValue("E_Status", i), i, "", "F");
                        }
                        else
                            Group_Log(Group_Code, "", "File Path is Missing - Check the File", Email_ID, "F");
                        objform.Freeze(false);
                    }
                }
                else
                {
                    objform.Freeze(true);
                    SAPbouiCOM.CheckBox oChk = objform.Items.Item("Item_30").Specific;
                    if (oChk.Checked == false)
                    {
                        objform.Freeze(true);
                        Dt.SetValue("E_Status", i, "DataBase Details are Missing. Please provide it in Credentials Setup");
                        objform.Freeze(false);
                        GeneralService_For_UDO_CustomerStatement_AddMethod(Dt.GetValue("E_Status", i), i, "", "F");
                    }
                    else
                        Group_Log(Group_Code, "", "DataBase Details are Missing. Please provide it in Credentials Setup", Email_ID, "F");
                    objform.Freeze(false);
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                SAPbouiCOM.CheckBox oChk = objform.Items.Item("Item_30").Specific;
                if (oChk.Checked == false)
                {
                    objform.Freeze(true);
                    Dt.SetValue("E_Status", i, ex.Message.ToString());
                    objform.Freeze(false);
                    GeneralService_For_UDO_CustomerStatement_AddMethod(Dt.GetValue("E_Status", i), i, "", "F");
                }
                else
                    Group_Log(Group_Code, "", ex.Message.ToString(), Email_ID, "F");
            }
        }
        #endregion

        #region Selected Rows
        public void Selecting_Rows(string Chk, string Row)
        {
            if (Chk == "Y")
            {
                if (arr == null)
                    Array.Resize(ref arr, 1);
                else
                    Array.Resize(ref arr, arr.Length + 1);
                arr[arr.Length - 1] = Row;
            }
            else
            {
                List<string> strList = arr.ToList();
                strList.Remove(Row);
                arr = strList.ToArray();
            }
        }
        #endregion
    }
}
