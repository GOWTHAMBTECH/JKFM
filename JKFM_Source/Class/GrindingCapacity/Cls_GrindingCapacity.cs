using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.VisualBasic;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Collections;
using System.Media;
using System.Text.RegularExpressions;

namespace JKFM_Source
{
    class Cls_GrindingCapacity
    {

        #region Declaration
        SAPbouiCOM.Form objform;
        ClsSBO objSBOAPI;

        SAPbouiCOM.DBDataSource oDBDSHeader;
        SAPbouiCOM.DBDataSource oDBDSDetail;
        SAPbouiCOM.Matrix oMatrix;
        string sFormUID;
        string UDOID;
        bool sFormLoade;
        public static bool sPressed;
        #endregion        

        #region Constructor
        public Cls_GrindingCapacity(ClsSBO objSBO)
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
                        case BoEventTypes.et_CLICK:
                            switch(pval.ItemUID)
                            {
                                case "1":
                                    if ((objform.Mode == BoFormMode.fm_ADD_MODE || objform.Mode == BoFormMode.fm_UPDATE_MODE) && !Validation())
                                    {
                                        SystemSounds.Asterisk.Play();
                                        bubbleevent = false;
                                    }
                                    break;
                            }
                            break;

                        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
                            SAPbouiCOM.Conditions oCons = null;
                            SAPbouiCOM.Condition oCon = null;
                            oCFLs = objform.ChooseFromLists;
                            SAPbouiCOM.ChooseFromList oCFL = null;
                            switch (pval.ItemUID)
                            {
                                case "t_Customer":
                                    oCFL = oCFLs.Item("CFL_RES");
                                    oCFL.SetConditions(null);
                                    oCons = oCFL.GetConditions();
                                    if (oCons.Count == 0)
                                    {
                                        oCon = oCons.Add();
                                    }
                                    else
                                    {
                                        oCon = oCons.Item(0);
                                    }
                                    oCon.BracketOpenNum = 1;
                                    oCon.Alias = "SellItem";
                                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                    oCon.CondVal = "Y";
                                    oCon.BracketCloseNum = 1;
                                    oCFL.SetConditions(oCons);
                                    GC.Collect();
                                    break;
                            }
                            break;
                    }
                }
                else
                {
                    switch (pval.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            switch(pval.ItemUID)
                            {
                                case "1":
                                    if (pval.ActionSuccess == true && objform.Mode == BoFormMode.fm_ADD_MODE)
                                        InitForm();
                                    break;
                            }
                            break;

                        case BoEventTypes.et_LOST_FOCUS:
                            switch(pval.ItemUID)
                            {
                                case "T_CType":
                                    if (sFormLoade)
                                    {
                                        CheckBox checkBox = (CheckBox)objform.Items.Item("c_MRP").Specific;
                                        string QueryStr = null;
                                        if (checkBox.Checked)
                                        {
                                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                            {
                                                QueryStr = "SELECT  \"StartDate\" ,\"EndDate\"  FROM \"DBO\".\"OMSN\" T0  Where \"MsnCode\" = '" + objform.Items.Item("T_CType").Specific.Value.ToString() + "'";
                                            }
                                            else
                                            {
                                                QueryStr = "SELECT  StartDate ,EndDate  FROM DBO.OMSN T0  Where MsnCode = '" + objform.Items.Item("T_CType").Specific.Value.ToString() + "'";
                                            }
                                                
                                            Recordset Orec = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            Orec.DoQuery(QueryStr);
                                            if (Orec.RecordCount > 0)
                                            {
                                                Orec.MoveFirst();
                                                DBDataSource oDbdsHeader1 = oDBDSHeader;
                                                DateTime date = Convert.ToDateTime(Orec.Fields.Item("StartDate").Value);
                                                string newVal1 = date.ToString("yyyyMMdd");
                                                oDbdsHeader1.SetValue("U_CPeriodFrom", 0, newVal1);
                                                DBDataSource oDbdsHeader2 = oDBDSHeader;
                                                date = Convert.ToDateTime(Orec.Fields.Item("EndDate").Value);
                                                string newVal2 = date.ToString("yyyyMMdd");
                                                oDbdsHeader2.SetValue("U_CPeriodTo", 0, newVal2);
                                            }
                                        }
                                        if (objform.Mode == BoFormMode.fm_ADD_MODE)
                                            ColumnName();
                                    }
                                    break;

                                case "c_series":
                                    if (objform.Mode == BoFormMode.fm_ADD_MODE && sFormLoade)
                                    {
                                        ComboBox comboBox = (ComboBox)objform.Items.Item("c_series").Specific;
                                        oDBDSHeader.SetValue("DocNum", 0, Convert.ToString((long)objform.BusinessObject.GetNextSerialNumber(comboBox.Selected.Value, "OCAC")));
                                    }
                                    break;
                            }
                            break;

                        case BoEventTypes.et_VALIDATE:
                            switch(pval.ItemUID)
                            {
                                case "T_CPTto":
                                    if (!(pval.ItemChanged && sFormLoade))
                                        break;
                                    break;
                                case "Matrix":
                                    switch(pval.ColUID)
                                    {
                                        case "PDate":
                                            objSBOAPI.AddNewLine(oMatrix, oDBDSDetail, oMatrix.VisualRowCount, pval.ColUID);
                                            oMatrix.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(BoCellClickType.ct_Regular, 0);
                                            break;
                                        case "DPlan":
                                            if (pval.ItemChanged)
                                            {
                                                double num1 = Convert.ToDouble(oMatrix.Columns.Item("DPlan").Cells.Item(pval.Row).Specific.Value);
                                                double num2 = Convert.ToDouble(oMatrix.Columns.Item("NPlan").Cells.Item(pval.Row).Specific.Value) + num1;
                                                oMatrix.FlushToDataSource();
                                                oDBDSDetail.SetValue("U_Available", (pval.Row - 1), num2.ToString());
                                                oDBDSDetail.SetValue("U_Balance", (pval.Row - 1), num2.ToString());
                                                oMatrix.LoadFromDataSource();
                                                oMatrix.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(BoCellClickType.ct_Regular, 0);
                                            }
                                            break;
                                        case "NPlan":
                                            if (pval.ItemChanged)
                                            {
                                                double num1 = Convert.ToDouble(oMatrix.Columns.Item("DPlan").Cells.Item(pval.Row).Specific.Value);
                                                double num2 = Convert.ToDouble(oMatrix.Columns.Item("NPlan").Cells.Item(pval.Row).Specific.Value) + num1;
                                                oMatrix.FlushToDataSource();
                                                oDBDSDetail.SetValue("U_Available", (pval.Row - 1), num2.ToString());
                                                oDBDSDetail.SetValue("U_Balance", (pval.Row - 1), num2.ToString());
                                                oMatrix.LoadFromDataSource();
                                                oMatrix.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(BoCellClickType.ct_Regular, 0);
                                            }
                                            break;
                                    }
                                    break;
                            }
                            break;

                        case BoEventTypes.et_CHOOSE_FROM_LIST:
                            ChooseFromListEvent chooseFromListEvent = (ChooseFromListEvent)pval;
                            DataTable selectedObjects = chooseFromListEvent.SelectedObjects;
                            switch(pval.ItemUID)
                            {
                                case "t_Customer":
                                    if (selectedObjects != null)
                                    {
                                        oDBDSHeader.SetValue("U_Customer", 0, Convert.ToString(selectedObjects.GetValue("CardCode", 0)).Trim());
                                        oDBDSHeader.SetValue("U_Name", 0, Convert.ToString(selectedObjects.GetValue("CardName", 0)).Trim());
                                    }
                                    break;
                            }
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

        #region Menu Event
        public void MenuEvent(ref SAPbouiCOM.MenuEvent pval, ref bool BubbleEvent)
        {
            try
            {
                if (pval.BeforeAction == true)
                {
                }
                else if (pval.BeforeAction == false)
                {
                    switch(pval.MenuUID)
                    {
                        case "OCAC":
                            objform = objSBOAPI.LoadForm("GrindingCapacity.xml", "OCAC");
                            objform = objSBOAPI.SBO_Appln.Forms.Item("OCAC");
                            oDBDSHeader = objform.DataSources.DBDataSources.Item("@AIS_OCAC");
                            oDBDSDetail = objform.DataSources.DBDataSources.Item("@AIS_CAC1");
                            oMatrix = (Matrix)objform.Items.Item("Matrix").Specific;
                            sFormLoade = true;
                            objform.Mode = BoFormMode.fm_ADD_MODE;
                            InitForm();
                            DefineModeForFields();
                            break;

                        case "1282":
                            InitForm();
                            if (objform.Mode != BoFormMode.fm_ADD_MODE)
                                return;
                            sPressed = true;
                            break;

                        case "1281":
                            FindInitForm();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #region InitForm
        public void InitForm()
        {
            try
            {
                objform.Freeze(true);
                objSBOAPI.LoadComboBoxSeries((ComboBox)objform.Items.Item("c_series").Specific, "OCAC");
                objSBOAPI.LoadDocumentDate((EditText)objform.Items.Item("t_DocDate").Specific);
                oMatrix.Clear();
                objSBOAPI.AddNewLine(oMatrix, oDBDSDetail, 1, "");
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    objSBOAPI.setComboBoxValue((ComboBox)objform.Items.Item("cmb_Unit").Specific, "SELECT  \"U_BranchCode\" , \"U_BranchName\" FROM \"@AIS_BRN1\"  Where \"U_Active\"='Y'");
                }
                else
                {
                    objSBOAPI.setComboBoxValue((ComboBox)objform.Items.Item("cmb_Unit").Specific, "SELECT  U_BranchCode , U_BranchName FROM [@AIS_BRN1]  Where U_Active='Y'");
                }
                    
                objform.ActiveItem = "T_CPFrom";
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("InitForm Method Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                objform.Freeze(false);
            }
        }
        #endregion

        #region DefineModeForFields
        public void DefineModeForFields()
        {
            try
            {
                objform.Items.Item("c_series").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);
                objform.Items.Item("t_DocNum").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);
                objform.Items.Item("c_series").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
                objform.Items.Item("t_DocDate").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);
                objform.Items.Item("t_DocDate").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);
                objform.Items.Item("T_CPTto").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
                objform.Items.Item("T_CPFrom").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
                objform.Items.Item("cmb_Unit").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Define Mode For Fields Failed: " + ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region FindInitForm
        public void FindInitForm()
        {
            try
            {
                objform.Freeze(true);
                objform.Mode = BoFormMode.fm_FIND_MODE;
                objform.ActiveItem = "t_DocNum";
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Find InitForm Method Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                objform.Freeze(false);
            }
        }
        #endregion

        #region ColumnName
        public void ColumnName()
        {
            try
            {
                objform.Freeze(true);
                int num1 = 1;
                do
                {
                    Column column1 = oMatrix.Columns.Item(("d" + Convert.ToString(num1)));
                    column1.Visible = false;
                    Column column2 = oMatrix.Columns.Item(("a" + Convert.ToString(num1)));
                    column2.Visible = false;
                    Column column3 = oMatrix.Columns.Item(("b" + Convert.ToString(num1)));
                    column3.Visible = false;
                    ++num1;
                }
                while (num1 <= 31);
                string Left1 = oDBDSHeader.GetValue("U_CPeriodFrom", 0);
                string Left2 = oDBDSHeader.GetValue("U_CPeriodTo", 0);
                Recordset recordset = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    recordset.DoQuery("SELECT CAST(thedate AS varchar(10)) AS 'Date', DAYOFMONTH(thedate) AS 'Day' FROM \"dbo\".\"ExplodeDates\"('" + Left1 + "','" + Left2 + "')  ");
                }
                else
                {
                    recordset.DoQuery("SELECT Convert(varchar(10),thedate,104)  as 'Date',day(thedate) as 'Day' FROM dbo.ExplodeDates('" + Left1 + "','" + Left2 + "')  ");
                }
                    
                if (Left1 != string.Empty && Left2 != string.Empty)
                {
                    if (recordset.RecordCount > 0)
                    {
                        recordset.MoveFirst();
                        int num2 = (recordset.RecordCount - 1);
                        int num3 = 0;
                        while (num3 <= num2)
                        {
                            Column column = oMatrix.Columns.Item(("a" + Convert.ToString(recordset.Fields.Item("Day").Value).Trim()));
                            column.TitleObject.Caption = "Available For " + recordset.Fields.Item("Date").Value.ToString().Trim();
                            column.Visible = true;
                            column.Editable = true;
                            recordset.MoveNext();
                            { ++num3; }
                        }
                    }
                    else
                    {
                        int num2 = 1;
                        do
                        {
                            oMatrix.Columns.Item(("a" + num2.ToString())).TitleObject.Caption = "" + num2.ToString();
                            recordset.MoveNext();
                            { ++num2; }
                        }
                        while (num2 <= 31);
                    }
                    if (recordset.RecordCount > 0)
                    {
                        recordset.MoveFirst();
                        int num2 = (recordset.RecordCount - 1);
                        int num3 = 0;
                        while (num3 <= num2)
                        {
                            Column column = oMatrix.Columns.Item(("d" + Convert.ToString(recordset.Fields.Item("Day").Value).Trim()));
                            column.TitleObject.Caption = "Allocated For " + recordset.Fields.Item("Date").Value.ToString().Trim();
                            column.Visible = true;
                            column.Editable = false;
                            recordset.MoveNext();
                            { ++num3; }
                        }
                    }
                    else
                    {
                        int num2 = 1;
                        do
                        {
                            oMatrix.Columns.Item(("d" + num2.ToString())).TitleObject.Caption = "" + num2.ToString();
                            recordset.MoveNext();
                            { ++num2; }
                        }
                        while (num2 <= 31);
                    }
                }
                if (recordset.RecordCount > 0)
                {
                    recordset.MoveFirst();
                    int num2 = (recordset.RecordCount - 1);
                    int num3 = 0;
                    while (num3 <= num2)
                    {
                        Column column = oMatrix.Columns.Item(("b" + Convert.ToString(recordset.Fields.Item("Day").Value).Trim()));
                        column.TitleObject.Caption = "Balance For " + recordset.Fields.Item("Date").Value.ToString().Trim();
                        column.Visible = true;
                        column.Editable = false;
                        recordset.MoveNext();
                        { ++num3; }
                    }
                }
                else
                {
                    int num2 = 1;
                    do
                    {
                        oMatrix.Columns.Item(("b" + num2.ToString())).TitleObject.Caption = "" + num2.ToString();
                        recordset.MoveNext();
                        { ++num2; }
                    }
                    while (num2 <= 31);
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Column Name Changing Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                objform.Freeze(false);
            }
        }
        #endregion

        #region Validation

        public bool Validation()
        {
            bool flag = false;
            try
            {
                string Capacity_To_Date = objform.Items.Item("T_CPTto").Specific.Value.ToString().Trim();
                string Capacity_From_Date = objform.Items.Item("T_CPFrom").Specific.Value.ToString().Trim();
                string Unit = oDBDSHeader.GetValue("U_Unit", 0).ToString().Trim();
                if (Capacity_From_Date == string.Empty)
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Capacity From Date Should Not Be Empty", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
                    flag = false;
                }
                else if (Capacity_To_Date == string.Empty)
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Capacity To Date Should Not Be Empty", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
                    flag = false;
                }
                else if (Unit == string.Empty)
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Unit Should Not Be Empty", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
                    flag = false;
                }
                else
                    flag = true;
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Validation Function Failed: " + ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
            }
            return flag;
        }
        #endregion

    }
}
