//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//using Microsoft.VisualBasic;
//using SAPbobsCOM;
//using SAPbouiCOM;
//using System.Collections;
//using System.Drawing;
//using System.Media;
//using System.Runtime.CompilerServices;
//using System.Runtime.InteropServices;

//namespace JKFM_Source
//{
//    class Cls_SalesPlan
//    {
//        #region Declaration
//        SAPbouiCOM.Form objform;
//        ClsSBO objSBOAPI;

//        SAPbouiCOM.Form frmSalesPlan;
//        SAPbouiCOM.Form frmSubPlanningDetailsSubMatrix;
//        SAPbouiCOM.Form frmSubPlanningDetailsSubMatrix2;
//        SAPbouiCOM.DBDataSource oDBDSHeader;
//        SAPbouiCOM.DBDataSource oDBDSDetail1;
//        SAPbouiCOM.DBDataSource oDBDSDetail2;
//        SAPbouiCOM.DBDataSource oDBDSDetail3;
//        SAPbouiCOM.DBDataSource oDBDSDetail4;
//        SAPbouiCOM.DBDataSource oDBDSDetail5;
//        SAPbouiCOM.DBDataSource oDBDSDetail6;
//        SAPbouiCOM.DBDataSource oDBDSDetail7;
//        SAPbouiCOM.DBDataSource oDBDSDetail8;
//        SAPbouiCOM.DBDataSource oDBDSDetail9;
//        SAPbouiCOM.DBDataSource oDBDSMainPlanDetails;
//        SAPbouiCOM.DBDataSource oDBDSMainPlanDetails2;
//        SAPbouiCOM.DBDataSource oDBDSSubPlanDetails;
//        SAPbouiCOM.DBDataSource oDBDSSubPlanDetails2;
//        SAPbouiCOM.Matrix oMatrix1;
//        SAPbouiCOM.Matrix oMatrix2;
//        SAPbouiCOM.Matrix oMatrix3;
//        SAPbouiCOM.Matrix oMatrix4;
//        SAPbouiCOM.Matrix oMatrix5;
//        SAPbouiCOM.Matrix oMatrix6;
//        SAPbouiCOM.Matrix oMatrix7;
//        SAPbouiCOM.Matrix oMatrix8;
//        SAPbouiCOM.Matrix oMatrix9;
//        SAPbouiCOM.Matrix oMatMainplanDetails;
//        SAPbouiCOM.Matrix oMatMainplanDetails2;
//        SAPbouiCOM.Matrix oMatSubPlanDetails;
//        SAPbouiCOM.Matrix oMatSubPlanDetails2;
//        SAPbouiCOM.Grid oGrid;
//        SAPbouiCOM.Grid oGrid1;
//        SAPbouiCOM.Grid oGrid2;
//        SAPbouiCOM.Grid oGrid3;
//        SAPbouiCOM.Grid oGrid4;
//        SAPbouiCOM.Grid oGrid5;
//        SAPbouiCOM.Grid oGrid6;
//        SAPbouiCOM.Grid oGrid7;
//        SAPbouiCOM.DataTable dtloadgrid;
//        SAPbouiCOM.DataTable dtloadgrid1;
//        SAPbouiCOM.DataTable dtloadgrid2;
//        SAPbouiCOM.DataTable dtloadgrid3;
//        SAPbouiCOM.DataTable dtloadgrid4;
//        SAPbouiCOM.DataTable dtloadgrid5;
//        SAPbouiCOM.DataTable dtloadgrid6;
//        SAPbouiCOM.DataTable dtloadgrid7;
//        string sFormUID;
//        string UDOID;
//        string SubRowID_PlanningDetails;
//        string SubRowID_PlanningDetails_PO;
//        string sQuery;
//        bool sFormLoade;
//        string SubRowID_PlanDate;
//        string SubRowID_ItemCode;
//        string SubRowID_PlanDate_PO;
//        private bool BubbleEvent;

//        #endregion

//        #region Constructor
//        public Cls_SalesPlan(ClsSBO objSBO)
//        {
//            objSBOAPI = objSBO;
//        }
//        #endregion

//        #region Item Event
//        public void itemevent(string formuid, ref SAPbouiCOM.ItemEvent pval, ref bool bubbleevent)
//        {
//            try
//            {
//                if (pval.BeforeAction == true)
//                {
//                    objform = objSBOAPI.SBO_Appln.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount);
//                    switch (pval.EventType)
//                    {
//                        case BoEventTypes.et_COMBO_SELECT:
//                            switch (pval.ItemUID)
//                            {
//                                case "cmb_Unit":
//                                    if (sFormLoade)
//                                    {
//                                        DateTime Planned_From_Date = frmSalesPlan.Items.Item("T_CPFrom").Specific.Value.ToString().Trim();

//                                        DateTime Planned_To_Date = frmSalesPlan.Items.Item("T_CPTto").Specific.Value.ToString().Trim();
//                                        DateTime Current_Date = Convert.ToDateTime(objSBOAPI.Query_Execute("Select  Convert(Varchar(10),getDate(),112)"));
//                                        string str1 = string.Empty;


//                                        ComboBox comboBox1 = (ComboBox)frmSalesPlan.Items.Item("c_Category").Specific;
//                                        if (comboBox1.Selected != null)
//                                            str1 = comboBox1.Selected.Value.ToString().Trim();
//                                        string str2 = string.Empty;


//                                        ComboBox comboBox2 = (ComboBox)frmSalesPlan.Items.Item("c_QtyType").Specific;
//                                        if (comboBox2.Selected != null)
//                                            str2 = comboBox2.Selected.Value.ToString().Trim();
//                                        if (Planned_From_Date < Current_Date && Planned_From_Date != null)
//                                        {

//                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Planned Date Shouldn't be less than Current Date...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                            SystemSounds.Asterisk.Play();
//                                            bubbleevent = false;
//                                        }
//                                        if (Planned_To_Date < Current_Date && Planned_To_Date != null)
//                                        {

//                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Planned Date Shouldn't be less than Current Date...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                            SystemSounds.Asterisk.Play();
//                                            bubbleevent = false;
//                                        }
//                                    }


//                                    break;

//                                case "c_Category":
//                                    if (sFormLoade)
//                                    {
//                                        DateTime Planned_From_Date = frmSalesPlan.Items.Item("T_CPFrom").Specific.Value.ToString().Trim();

//                                        DateTime Planned_To_Date = frmSalesPlan.Items.Item("T_CPTto").Specific.Value.ToString().Trim();
//                                        DateTime Current_Date = Convert.ToDateTime(objSBOAPI.Query_Execute("Select  Convert(Varchar(10),getDate(),112)"));
//                                        string str1 = string.Empty;


//                                        ComboBox comboBox1 = (ComboBox)frmSalesPlan.Items.Item("c_Category").Specific;
//                                        if (comboBox1.Selected != null)
//                                            str1 = comboBox1.Selected.Value.ToString().Trim();
//                                        string str2 = string.Empty;


//                                        ComboBox comboBox2 = (ComboBox)frmSalesPlan.Items.Item("c_QtyType").Specific;
//                                        if (comboBox2.Selected != null)
//                                            str2 = comboBox2.Selected.Value.ToString().Trim();
//                                        if (Planned_From_Date < Current_Date && Planned_From_Date != null)
//                                        {

//                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Planned Date Shouldn't be less than Current Date...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                            SystemSounds.Asterisk.Play();
//                                            bubbleevent = false;
//                                        }
//                                        if (Planned_To_Date < Current_Date && Planned_To_Date != null)
//                                        {

//                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Planned Date Shouldn't be less than Current Date...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                            SystemSounds.Asterisk.Play();
//                                            bubbleevent = false;
//                                        }
//                                    }
//                                    break;

//                                case "c_QtyType":
//                                    if (sFormLoade)
//                                    {
//                                        DateTime Planned_From_Date = frmSalesPlan.Items.Item("T_CPFrom").Specific.Value.ToString().Trim();

//                                        DateTime Planned_To_Date = frmSalesPlan.Items.Item("T_CPTto").Specific.Value.ToString().Trim();
//                                        DateTime Current_Date = Convert.ToDateTime(objSBOAPI.Query_Execute("Select  Convert(Varchar(10),getDate(),112)"));
//                                        string str1 = string.Empty;


//                                        ComboBox comboBox1 = (ComboBox)frmSalesPlan.Items.Item("c_Category").Specific;
//                                        if (comboBox1.Selected != null)
//                                            str1 = comboBox1.Selected.Value.ToString().Trim();
//                                        string str2 = string.Empty;


//                                        ComboBox comboBox2 = (ComboBox)frmSalesPlan.Items.Item("c_QtyType").Specific;
//                                        if (comboBox2.Selected != null)
//                                            str2 = comboBox2.Selected.Value.ToString().Trim();
//                                        if (Planned_From_Date < Current_Date && Planned_From_Date != null)
//                                        {

//                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Planned Date Shouldn't be less than Current Date...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                            SystemSounds.Asterisk.Play();
//                                            bubbleevent = false;
//                                        }
//                                        if (Planned_To_Date < Current_Date && Planned_To_Date != null)
//                                        {

//                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Planned Date Shouldn't be less than Current Date...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                            SystemSounds.Asterisk.Play();
//                                            bubbleevent = false;
//                                        }
//                                    }
//                                    break;

//                            }
//                            break;
//                        case BoEventTypes.et_ITEM_PRESSED:
//                            switch (pval.ItemUID)
//                            {
//                                case "1":
//                                    if (frmSalesPlan.Mode == BoFormMode.fm_ADD_MODE || frmSalesPlan.Mode == BoFormMode.fm_UPDATE_MODE)
//                                    {
//                                        if (frmSalesPlan.Mode == BoFormMode.fm_ADD_MODE)
//                                        {
//                                            if (!Validation())
//                                            {
//                                                SystemSounds.Asterisk.Play();
//                                                bubbleevent = false;
//                                            }
//                                            else if (oMatrix1.VisualRowCount > 1)
//                                                objSBOAPI.DeleteEmptyRowInFormDataEvent(oMatrix1, "ICode", oDBDSDetail1);
//                                        }
//                                        //else if (frmSalesPlan.Mode == BoFormMode.fm_UPDATE_MODE)
//                                            //UpdatePresalesDetails();
//                                    }
//                                    break;
//                            }
//                            break;

//                        case BoEventTypes.et_VALIDATE:
//                            switch (pval.ItemUID)
//                            {
//                                case "T_CPTto":
//                                    if (sFormLoade)
//                                    {
//                                        DateTime Planned_From_Date = Convert.ToString(frmSalesPlan.Items.Item("T_CPFrom").Specific.Value);

//                                        DateTime Planned_To_Date = Convert.ToString(frmSalesPlan.Items.Item("T_CPTto").Specific.Value);
//                                        DateTime Current_Date = Convert.ToDateTime(objSBOAPI.Query_Execute("Select  Convert(Varchar(10),getDate(),112)"));
//                                        string str1 = string.Empty;
//                                        ComboBox comboBox1 = (ComboBox)frmSalesPlan.Items.Item("c_Category").Specific;
//                                        if (comboBox1.Selected != null)
//                                            str1 = comboBox1.Selected.Value.ToString().Trim();
//                                        string str2 = string.Empty;

//                                        ComboBox comboBox2 = (ComboBox)frmSalesPlan.Items.Item("c_QtyType").Specific;
//                                        if (comboBox2.Selected != null)
//                                            str2 = comboBox2.Selected.Value.ToString().Trim();
//                                        if (Planned_From_Date < Current_Date && Planned_From_Date != null)
//                                        {
//                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Planned Date Shouldn't be less than Current Date...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                            SystemSounds.Asterisk.Play();
//                                            bubbleevent = false;
//                                        }
//                                        if (Planned_To_Date < Current_Date && Planned_To_Date != null)
//                                        {
//                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Planned Date Shouldn't be less than Current Date...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                            SystemSounds.Asterisk.Play();
//                                            bubbleevent = false;
//                                        }
//                                    }

//                                    break;
//                                case "T_CPFrom":
//                                    if (sFormLoade)
//                                    {
//                                        DateTime Planned_From_Date = Convert.ToString(frmSalesPlan.Items.Item("T_CPFrom").Specific.Value);

//                                        DateTime Planned_To_Date = Convert.ToString(frmSalesPlan.Items.Item("T_CPTto").Specific.Value);
//                                        DateTime Current_Date = Convert.ToDateTime(objSBOAPI.Query_Execute("Select  Convert(Varchar(10),getDate(),112)"));
//                                        string str1 = string.Empty;
//                                        ComboBox comboBox1 = (ComboBox)frmSalesPlan.Items.Item("c_Category").Specific;
//                                        if (comboBox1.Selected != null)
//                                            str1 = comboBox1.Selected.Value.ToString().Trim();
//                                        string str2 = string.Empty;

//                                        ComboBox comboBox2 = (ComboBox)frmSalesPlan.Items.Item("c_QtyType").Specific;
//                                        if (comboBox2.Selected != null)
//                                            str2 = comboBox2.Selected.Value.ToString().Trim();
//                                        if (Planned_From_Date < Current_Date && Planned_From_Date != null)
//                                        {
//                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Planned Date Shouldn't be less than Current Date...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                            SystemSounds.Asterisk.Play();
//                                            bubbleevent = false;
//                                        }
//                                        if (Planned_To_Date < Current_Date && Planned_To_Date != null)
//                                        {
//                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Planned Date Shouldn't be less than Current Date...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                            SystemSounds.Asterisk.Play();
//                                            bubbleevent = false;
//                                        }
//                                    }
//                                    break;
//                            }
//                            break;
//                    }
//                }
//                else
//                {
//                    switch (pval.EventType)
//                    {
//                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

//                            switch (pval.ItemUID)
//                            {
//                                case "1":
//                                    if (pval.Action_Success == true && frmSalesPlan.Mode == BoFormMode.fm_ADD_MODE)
//                                        InitForm();
//                                    break;
//                                case "23":
//                                    frmSalesPlan.PaneLevel = 1;
//                                    frmSalesPlan.Items.Item(pval.ItemUID).AffectsFormMode = false;
//                                    break;
//                                case "24":
//                                    frmSalesPlan.PaneLevel = 2;
//                                    frmSalesPlan.Items.Item(pval.ItemUID).AffectsFormMode = false;
//                                    break;
//                                case "fld_DayCap":
//                                    frmSalesPlan.PaneLevel = 4;
//                                    frmSalesPlan.Items.Item(pval.ItemUID).AffectsFormMode = false;
//                                    break;
//                                case "fld_SOQty":
//                                    frmSalesPlan.PaneLevel = 5;
//                                    frmSalesPlan.Items.Item(pval.ItemUID).AffectsFormMode = false;
//                                    break;
//                                case "fld_Feasi":
//                                    frmSalesPlan.PaneLevel = 6;
//                                    frmSalesPlan.Items.Item(pval.ItemUID).AffectsFormMode = false;
//                                    break;
//                                case "fld_Pack":
//                                    frmSalesPlan.PaneLevel = 7;
//                                    frmSalesPlan.Items.Item(pval.ItemUID).AffectsFormMode = false;
//                                    break;
//                                case "fld_ShrtEx":
//                                    frmSalesPlan.PaneLevel = 8;
//                                    frmSalesPlan.Items.Item(pval.ItemUID).AffectsFormMode = false;
//                                    break;
//                                case "fld_Slob":
//                                    frmSalesPlan.PaneLevel = 9;
//                                    frmSalesPlan.Items.Item(pval.ItemUID).AffectsFormMode = false;
//                                    break;
//                                case "IPre":
//                                    frmSalesPlan.PaneLevel = 3;
//                                    frmSalesPlan.Settings.MatrixUID = "Matrix2";
//                                    frmSalesPlan.Items.Item(pval.ItemUID).AffectsFormMode = false;
//                                    break;
//                                case "b_Color":
//                                    //ColourUpdation();
//                                    break;
//                                case "b_SO":
//                                    if (frmSalesPlan.Items.Item(pval.ItemUID).Enabled)
//                                    {
//                                        if (frmSalesPlan.Mode == BoFormMode.fm_OK_MODE)
//                                        {
//                                            if (!objSBOAPI.oCompany.InTransaction)
//                                            {
//                                                objSBOAPI.oCompany.StartTransaction();
//                                            }
//                                            //if (!TransactionManagement(true))
//                                            //{
//                                            //    if (objSBOAPI.oCompany.InTransaction)
//                                            //    {
//                                            //        objSBOAPI.oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
//                                            //    }
//                                            //    //bubbleevent = false;
//                                            //}
//                                            //else if (objSBOAPI.oCompany.InTransaction)
//                                            //{
//                                            //    objSBOAPI.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
//                                            //}
//                                        }
//                                        else
//                                        {
//                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Add/Update the Document...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                            //bubbleevent = false;
//                                        }
//                                    }
//                                    break;
//                                case "b_WO":
//                                    if (frmSalesPlan.Mode == BoFormMode.fm_OK_MODE & frmSalesPlan.Items.Item(pval.ItemUID).Enabled)
//                                    {
//                                        int num1 = (oGrid5.Rows.Count - 1);
//                                        int num2 = 0;
//                                        while (num2 <= num1)
//                                        {

//                                            if (oGrid5.Rows.IsSelected(num2))
//                                            {

//                                                string str1 = oDBDSHeader.GetValue("DocEntry", 0).Trim();

//                                                string str2 = oDBDSHeader.GetValue("U_CPeriodFrom", 0);
//                                                string singleValue = objSBOAPI.Query_Execute("Select Convert(Varchar(10), dateadd(day," + Convert.ToString(num2) + " ,'" + str2 + "') ,112) ");
//                                                if (Convert.ToInt32(objSBOAPI.Query_Execute("Select Count(DocEntry) from OWOR Where U_BaseEntry='" + str1 + "' and U_BaseObject ='OSAP' and Convert(Varchar(10),U_BaseDate,112) ='" + singleValue + "'")) > 0)
//                                                {
//                                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Production Order Already Created for this Date...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                                    //bubbleevent = false;
//                                                }
//                                                else
//                                                {
//                                                    if (!objSBOAPI.oCompany.InTransaction)
//                                                    {
//                                                        objSBOAPI.oCompany.StartTransaction();
//                                                    }
//                                                    //if (CreateProductionOrder(num2))
//                                                    //{
//                                                    //    if (objSBOAPI.oCompany.InTransaction)
//                                                    //    {
//                                                    //        objSBOAPI.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
//                                                    //    }
//                                                    //}
//                                                    //else
//                                                    //{
//                                                    //    if (objSBOAPI.oCompany.InTransaction)
//                                                    //    {
//                                                    //        objSBOAPI.oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
//                                                    //    }
//                                                    //    //bubbleevent = false;
//                                                    //}
//                                                }
//                                            }
//                                            { ++num2; }
//                                        }
//                                    }
//                                    break;
//                                case "b_Fes":
//                                    if (frmSalesPlan.Mode == BoFormMode.fm_ADD_MODE)
//                                    {
//                                        string QtyType = string.Empty;
//                                        ComboBox comboBox = (ComboBox)frmSalesPlan.Items.Item("c_QtyType").Specific;
//                                        if (comboBox.Selected != null)
//                                            QtyType = comboBox.Selected.Value.ToString().Trim();
//                                        //if (QtyType == "B")
//                                        //    //LoadFeasibilty_Bag();
//                                        //else if (QtyType == "T")
//                                        //    //LoadFeasibilty_Tonnage();
//                                    }
//                                    break;
//                            }
//                            break;

//                        case BoEventTypes.et_COMBO_SELECT:
//                            switch (pval.ItemUID)
//                            {
//                                case "cmb_Unit":
//                                    if (frmSalesPlan.Mode == BoFormMode.fm_ADD_MODE && sFormLoade)
//                                    {
//                                        string str = string.Empty;
//                                        ComboBox comboBox = (ComboBox)frmSalesPlan.Items.Item("cmb_Unit").Specific;
//                                        if (comboBox.Selected != null)
//                                            str = comboBox.Selected.Value.ToString().Trim();
//                                        //LoadMatrixDetails1();
//                                    }
//                                    break;
//                                case "c_Category":
//                                    if (frmSalesPlan.Mode == BoFormMode.fm_ADD_MODE && sFormLoade)
//                                    {
//                                        string str = string.Empty;
//                                        ComboBox comboBox = (ComboBox)frmSalesPlan.Items.Item("cmb_Unit").Specific;
//                                        if (comboBox.Selected != null)
//                                            str = comboBox.Selected.Value.ToString().Trim();
//                                        //LoadMatrixDetails1();
//                                    }
//                                    break;
//                                case "c_QtyType":
//                                    if (frmSalesPlan.Mode == BoFormMode.fm_ADD_MODE && sFormLoade)
//                                    {
//                                        string str = string.Empty;
//                                        ComboBox comboBox = (ComboBox)frmSalesPlan.Items.Item("cmb_Unit").Specific;
//                                        if (comboBox.Selected != null)
//                                            str = comboBox.Selected.Value.ToString().Trim();
//                                        //LoadMatrixDetails1();
//                                    }
//                                    break;

//                                case "c_series":
//                                    if (frmSalesPlan.Mode == BoFormMode.fm_ADD_MODE && sFormLoade)
//                                    {
//                                        ComboBox comboBox = (ComboBox)frmSalesPlan.Items.Item("c_series").Specific;
//                                        oDBDSHeader.SetValue("DocNum", 0, Convert.ToString((long)frmSalesPlan.BusinessObject.GetNextSerialNumber(comboBox.Selected.Value, UDOID)));
//                                    }
//                                    break;
//                            }
//                            break;

//                        case BoEventTypes.et_DOUBLE_CLICK:
//                            switch (pval.ItemUID)
//                            {
//                                case "Grid5":
//                                    switch (pval.ColUID)
//                                    {
//                                        case "PlanDate":
//                                            if (pval.ColUID != string.Empty)
//                                            {
//                                                SubRowID_PlanningDetails = pval.ColUID;

//                                                EditTextColumn editTextColumn = (EditTextColumn)oGrid5.Columns.Item(pval.ColUID);
//                                                string strItemCode = ((IEditTextColumn)editTextColumn).TitleObject.Caption.ToString().Trim();

//                                                string str = oDBDSHeader.GetValue("U_CPeriodFrom", 0);

//                                                oGrid5.DataTable.GetValue("PlanDate", pval.Row).ToString().Trim();
//                                                SubRowID_ItemCode = strItemCode;
//                                                SubRowID_PlanDate = objSBOAPI.Query_Execute(" Select Convert(Varchar(10), dateadd(day," + Convert.ToString(pval.Row) + " ,'" + str + "'  ) ,112)  ");
//                                                //LoadSubMatrixDetailsForFeasibilityPopup(SubRowID_PlanDate, SubRowID_PlanningDetails, strItemCode);
//                                            }
//                                            break;
//                                    }
//                                    break;
//                            }
//                            break;

//                        case BoEventTypes.et_VALIDATE:
//                            switch (pval.ItemUID)
//                            {
//                                case "T_CPTto":
//                                    if (frmSalesPlan.Mode == BoFormMode.fm_ADD_MODE && pval.ItemChanged && sFormLoade)
//                                    {
//                                        DateTime Planned_From_Date = Convert.ToString(frmSalesPlan.Items.Item("T_CPFrom").Specific.Value);

//                                        DateTime Planned_To_Date = Convert.ToString(frmSalesPlan.Items.Item("T_CPTto").Specific.Value);
//                                        DateTime Current_Date = Convert.ToDateTime(objSBOAPI.Query_Execute("Select  Convert(Varchar(10),getDate(),112)"));
//                                        string Category = string.Empty;

//                                        ComboBox comboBox1 = (ComboBox)frmSalesPlan.Items.Item("c_Category").Specific;
//                                        if (comboBox1.Selected != null)
//                                            Category = comboBox1.Selected.Value.ToString().Trim();
//                                        string str2 = string.Empty;

//                                        ComboBox comboBox2 = (ComboBox)frmSalesPlan.Items.Item("c_QtyType").Specific;
//                                        if (comboBox2.Selected != null)
//                                            str2 = comboBox2.Selected.Value.ToString().Trim();
//                                        if (Planned_From_Date < Current_Date && Planned_From_Date != null)
//                                        {

//                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Planned Date Shouldn't be less than Current Date...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                            SystemSounds.Asterisk.Play();
//                                            //bubbleevent = false;
//                                        }
//                                        if (Planned_To_Date < Current_Date && Planned_To_Date != null)
//                                        {

//                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Planned Date Shouldn't be less than Current Date...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                            SystemSounds.Asterisk.Play();
//                                            //bubbleevent = false;
//                                        }
//                                        //LoadMatrixDetails1();
//                                    }
//                                    break;
//                                case "T_CPFrom":
//                                    if (frmSalesPlan.Mode == BoFormMode.fm_ADD_MODE && pval.ItemChanged && sFormLoade)
//                                    {
//                                        DateTime Planned_From_Date = Convert.ToString(frmSalesPlan.Items.Item("T_CPFrom").Specific.Value);

//                                        DateTime Planned_To_Date = Convert.ToString(frmSalesPlan.Items.Item("T_CPTto").Specific.Value);
//                                        DateTime Current_Date = Convert.ToDateTime(objSBOAPI.Query_Execute("Select  Convert(Varchar(10),getDate(),112)"));
//                                        string Category = string.Empty;

//                                        ComboBox comboBox1 = (ComboBox)frmSalesPlan.Items.Item("c_Category").Specific;
//                                        if (comboBox1.Selected != null)
//                                            Category = comboBox1.Selected.Value.ToString().Trim();
//                                        string str2 = string.Empty;

//                                        ComboBox comboBox2 = (ComboBox)frmSalesPlan.Items.Item("c_QtyType").Specific;
//                                        if (comboBox2.Selected != null)
//                                            str2 = comboBox2.Selected.Value.ToString().Trim();
//                                        if (Planned_From_Date < Current_Date && Planned_From_Date != null)
//                                        {

//                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Planned Date Shouldn't be less than Current Date...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                            SystemSounds.Asterisk.Play();
//                                            //bubbleevent = false;
//                                        }
//                                        if (Planned_To_Date < Current_Date && Planned_To_Date != null)
//                                        {

//                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Planned Date Shouldn't be less than Current Date...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                            SystemSounds.Asterisk.Play();
//                                            //bubbleevent = false;
//                                        }
//                                        //LoadMatrixDetails1();
//                                    }
//                                    break;

//                                case "matPreSale":
//                                    switch (pval.ColUID)
//                                    {
//                                        case "V_1":
//                                            if (pval.ItemChanged)
//                                            {
//                                                oMatrix1.FlushToDataSource();
//                                                double num = Convert.ToDouble(oDBDSDetail1.GetValue("U_Quantity", (pval.Row - 1))) * Convert.ToDouble(oDBDSDetail1.GetValue("U_UnitPrice", (pval.Row - 1)));
//                                                oDBDSDetail1.SetValue("U_BaseAmount", (pval.Row - 1), Convert.ToString(num));
//                                                oMatrix1.LoadFromDataSource();
//                                                oMatrix1.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(BoCellClickType.ct_Regular, 0);
//                                            }
//                                            break;
//                                        case "V_8":
//                                            if (pval.ItemChanged)
//                                            {
//                                                oMatrix1.FlushToDataSource();
//                                                double num = Convert.ToDouble(oDBDSDetail1.GetValue("U_Quantity", (pval.Row - 1))) * Convert.ToDouble(oDBDSDetail1.GetValue("U_UnitPrice", (pval.Row - 1)));
//                                                oDBDSDetail1.SetValue("U_BaseAmount", (pval.Row - 1), Convert.ToString(num));
//                                                oMatrix1.LoadFromDataSource();
//                                                oMatrix1.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(BoCellClickType.ct_Regular, 0);
//                                            }
//                                            break;
//                                        case "V_6":
//                                            if (pval.ItemChanged)
//                                            {
//                                                oMatrix1.FlushToDataSource();
//                                                double num = Convert.ToDouble(oDBDSDetail1.GetValue("U_BaseAmount", (pval.Row - 1))) + Convert.ToDouble(oDBDSDetail1.GetValue("U_TaxAmount", (pval.Row - 1)));
//                                                oDBDSDetail1.SetValue("U_TotalPrice", (pval.Row - 1), Convert.ToString(num));
//                                                oMatrix1.LoadFromDataSource();
//                                                oMatrix1.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(BoCellClickType.ct_Regular, 0);
//                                            }
//                                            break;
//                                    }
//                                    break;
//                            }
//                            break;

//                        case BoEventTypes.et_CHOOSE_FROM_LIST:
//                            ChooseFromListEvent chooseFromListEvent = (ChooseFromListEvent)pval;
//                            DataTable selectedObjects = chooseFromListEvent.SelectedObjects;
//                            if (selectedObjects != null)
//                            {
//                                switch (pval.ItemUID)
//                                {
//                                    case "t_Customer":
//                                        oDBDSHeader.SetValue("U_Customer", 0, Convert.ToString(selectedObjects.GetValue("CardCode", 0)).Trim());
//                                        oDBDSHeader.SetValue("U_Name", 0, Convert.ToString(selectedObjects.GetValue("CardName", 0)).Trim());
//                                        break;
//                                    case "Matrix":
//                                        switch(pval.ColUID)
//                                        {
//                                            case "CCode":
//                                                oMatrix1.FlushToDataSource();
//                                                oDBDSDetail1.SetValue("U_CardCode", (pval.Row - 1), Convert.ToString(selectedObjects.GetValue("CardCode", 0)).Trim());
//                                                oDBDSDetail1.SetValue("U_CardName", (pval.Row - 1), Convert.ToString(selectedObjects.GetValue("CardName", 0)).Trim());
//                                                oMatrix1.LoadFromDataSource();
//                                                oMatrix1.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(BoCellClickType.ct_Regular, 0);
//                                                break;
//                                            case "ICode":
//                                                oMatrix1.FlushToDataSource();
//                                                oDBDSDetail1.SetValue("U_ItemCode", (pval.Row - 1), Convert.ToString(selectedObjects.GetValue("ItemCode", 0)).Trim());
//                                                oDBDSDetail1.SetValue("U_ItemName", (pval.Row - 1), Convert.ToString(selectedObjects.GetValue("ItemName", 0)).Trim());
//                                                oMatrix1.LoadFromDataSource();
//                                                objSBOAPI.AddNewLine(oMatrix1, oDBDSDetail1, oMatrix1.VisualRowCount, pval.ColUID);
//                                                break;
//                                        }
//                                        break;
//                                }
//                            }
//                            break;
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                objform.Freeze(false);
//                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//            }
//        }
//        #endregion

//        #region Menu Event
//        public void MenuEvent(ref SAPbouiCOM.MenuEvent pval, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pval.BeforeAction == true)
//                {
//                }
//                else if (pval.BeforeAction == false)
//                {
//                    switch (pval.MenuUID)
//                    {
//                        case "OSAP":
//                            objform = objSBOAPI.LoadForm("SalesPlanning.xml", "OSAP");
//                            objform = objSBOAPI.SBO_Appln.Forms.Item("OSAP");

//                            oDBDSHeader = frmSalesPlan.DataSources.DBDataSources.Item("@AIS_OSAP");

//                            oDBDSDetail1 = frmSalesPlan.DataSources.DBDataSources.Item("@AIS_SAP1");

//                            oDBDSDetail2 = frmSalesPlan.DataSources.DBDataSources.Item("@AIS_SAP2");

//                            oDBDSMainPlanDetails = frmSalesPlan.DataSources.DBDataSources.Item("@AIS_SAP3");

//                            oDBDSMainPlanDetails2 = frmSalesPlan.DataSources.DBDataSources.Item("@AIS_SAP4");

//                            oDBDSDetail3 = frmSalesPlan.DataSources.DBDataSources.Item("@AIS_SAP5");

//                            oDBDSDetail4 = frmSalesPlan.DataSources.DBDataSources.Item("@AIS_SAP6");

//                            oDBDSDetail5 = frmSalesPlan.DataSources.DBDataSources.Item("@AIS_SAP7");

//                            oDBDSDetail6 = frmSalesPlan.DataSources.DBDataSources.Item("@AIS_SAP8");

//                            oDBDSDetail7 = frmSalesPlan.DataSources.DBDataSources.Item("@AIS_SAP9");

//                            oDBDSDetail8 = frmSalesPlan.DataSources.DBDataSources.Item("@AIS_SAP10");

//                            oDBDSDetail9 = frmSalesPlan.DataSources.DBDataSources.Item("@AIS_SAP11");

//                            oMatrix1 = (SAPbouiCOM.Matrix)frmSalesPlan.Items.Item("Matrix").Specific;

//                            oMatrix2 = (Matrix)frmSalesPlan.Items.Item("Matrix2").Specific;

//                            oMatrix3 = (Matrix)frmSalesPlan.Items.Item("Matrix3").Specific;

//                            oMatrix4 = (Matrix)frmSalesPlan.Items.Item("Matrix4").Specific;

//                            oMatrix5 = (Matrix)frmSalesPlan.Items.Item("Matrix5").Specific;

//                            oMatrix6 = (Matrix)frmSalesPlan.Items.Item("Matrix6").Specific;

//                            oMatrix7 = (Matrix)frmSalesPlan.Items.Item("Matrix7").Specific;

//                            oMatrix8 = (Matrix)frmSalesPlan.Items.Item("Matrix8").Specific;

//                            oMatrix9 = (Matrix)frmSalesPlan.Items.Item("Matrix9").Specific;

//                            oMatMainplanDetails = (Matrix)frmSalesPlan.Items.Item("subMatrix").Specific;

//                            oMatMainplanDetails2 = (Matrix)frmSalesPlan.Items.Item("subMatrix2").Specific;

//                            frmSalesPlan.Items.Item("subMatrix").Visible = false;

//                            oGrid = (Grid)frmSalesPlan.Items.Item("Grid").Specific;

//                            dtloadgrid = frmSalesPlan.DataSources.DataTables.Add("DataTable");

//                            oGrid1 = (Grid)frmSalesPlan.Items.Item("Grid1").Specific;

//                            dtloadgrid1 = frmSalesPlan.DataSources.DataTables.Add("DataTable1");

//                            oGrid2 = (Grid)frmSalesPlan.Items.Item("Grid2").Specific;

//                            dtloadgrid2 = frmSalesPlan.DataSources.DataTables.Add("DataTable2");

//                            oGrid3 = (Grid)frmSalesPlan.Items.Item("Grid3").Specific;

//                            dtloadgrid3 = frmSalesPlan.DataSources.DataTables.Add("DataTable3");

//                            oGrid4 = (Grid)frmSalesPlan.Items.Item("Grid4").Specific;

//                            dtloadgrid4 = frmSalesPlan.DataSources.DataTables.Add("DataTable4");

//                            oGrid5 = (Grid)frmSalesPlan.Items.Item("Grid5").Specific;

//                            dtloadgrid5 = frmSalesPlan.DataSources.DataTables.Add("DataTable5");

//                            oGrid6 = (Grid)frmSalesPlan.Items.Item("Grid6").Specific;

//                            dtloadgrid6 = frmSalesPlan.DataSources.DataTables.Add("DataTable6");

//                            oGrid7 = (Grid)frmSalesPlan.Items.Item("Grid7").Specific;

//                            dtloadgrid7 = frmSalesPlan.DataSources.DataTables.Add("DataTable7");
//                            sFormLoade = true;
//                            frmSalesPlan.Mode = BoFormMode.fm_ADD_MODE;
//                            InitForm();
//                            DefineModeForFields();
//                            break;
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                objform.Freeze(false);
//                objSBOAPI.SBO_Appln.MessageBox(ex.Message);
//            }
//        }
//        #endregion

//        public void InitForm()
//        {
//            try
//            {

//                frmSalesPlan.Freeze(true);

//                objSBOAPI.LoadComboBoxSeries((ComboBox)frmSalesPlan.Items.Item("c_series").Specific, "OSAP");

//                objSBOAPI.LoadDocumentDate((EditText)frmSalesPlan.Items.Item("t_DocDate").Specific);

//                oMatrix1.Clear();
//                oGrid.DataTable = dtloadgrid;

//                oGrid.DataTable.Clear();

//                dtloadgrid.Clear();
//                oGrid1.DataTable = dtloadgrid1;

//                oGrid1.DataTable.Clear();

//                dtloadgrid1.Clear();
//                oGrid2.DataTable = dtloadgrid2;

//                oGrid2.DataTable.Clear();

//                dtloadgrid2.Clear();
//                oGrid3.DataTable = dtloadgrid3;

//                oGrid3.DataTable.Clear();

//                dtloadgrid3.Clear();
//                oGrid4.DataTable = dtloadgrid4;

//                oGrid4.DataTable.Clear();

//                dtloadgrid4.Clear();
//                oGrid5.DataTable = dtloadgrid5;

//                oGrid5.DataTable.Clear();

//                dtloadgrid5.Clear();

//                objSBOAPI.setComboBoxValue((ComboBox)frmSalesPlan.Items.Item("cmb_Shift").Specific, "SELECT  Code , Name FROM [@AIS_OSFT] ");

//                objSBOAPI.setComboBoxValue((ComboBox)frmSalesPlan.Items.Item("cmb_Unit").Specific, "SELECT  U_BranchCode , U_BranchName FROM [@AIS_BRN1]  Where U_Active='Y'");


//                frmSalesPlan.Items.Item("23").Click(BoCellClickType.ct_Regular);
//                frmSalesPlan.ActiveItem = "T_CPFrom";
//            }
//            catch (Exception ex)
//            {
//                frmSalesPlan.Freeze(false);

//                objSBOAPI.SBO_Appln.StatusBar.SetText("InitForm Method Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                frmSalesPlan.Freeze(false);
//            }
//        }

//        public void DefineModeForFields()
//        {
//            try
//            {


//                frmSalesPlan.Items.Item("c_series").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);


//                frmSalesPlan.Items.Item("t_DocNum").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);


//                frmSalesPlan.Items.Item("c_series").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);


//                frmSalesPlan.Items.Item("t_DocDate").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);


//                frmSalesPlan.Items.Item("T_CPTto").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);


//                frmSalesPlan.Items.Item("T_CPFrom").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);


//                frmSalesPlan.Items.Item("cmb_Unit").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);


//                frmSalesPlan.Items.Item("b_Fes").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);


//                frmSalesPlan.Items.Item("b_SO").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 2, BoModeVisualBehavior.mvb_False);


//                frmSalesPlan.Items.Item("c_QtyType").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);


//                frmSalesPlan.Items.Item("c_Category").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);


//                frmSalesPlan.Items.Item("cmb_Type").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
//            }
//            catch (Exception ex)
//            {
//                frmSalesPlan.Freeze(false);
//                objSBOAPI.SBO_Appln.StatusBar.SetText("Define Mode For Fields Failed: " + ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        public bool Validation()
//        {
//            bool flag = false;
//            try
//            {
                
//                string Left1 = frmSalesPlan.Items.Item("T_CPTto").Specific.Value.ToString().Trim();
                
//                string str1 = frmSalesPlan.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                
                
//                ComboBox comboBox = (ComboBox)frmSalesPlan.Items.Item("21").Specific;
                
//                string Left2 = oDBDSHeader.GetValue("U_Unit", 0).ToString().Trim();
//                if (Convert.ToInt32(objSBOAPI.Query_Execute(" Select Count(U_Status ) as Status From [@AIS_OSAP]  Where U_Status ='A' And DocEntry != '" + str1 + "'")) > 0 && comboBox.Selected.Value.ToString() == "A")
//                {
//                    objSBOAPI.SBO_Appln.StatusBar.SetText("At a Time Only One Planning Should be Active", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
//                    flag = false;
//                }
//                else if (Left1 == string.Empty)
//                {
//                    objSBOAPI.SBO_Appln.StatusBar.SetText("Capacity To Date Should Not Be Empty", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
//                    flag = false;
//                }
//                else if (Left1 == string.Empty)
//                {
//                    objSBOAPI.SBO_Appln.StatusBar.SetText("Capacity To Date Should Not Be Empty", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
//                    flag = false;
//                }
//                else if (Left2 == string.Empty)
//                {
//                    objSBOAPI.SBO_Appln.StatusBar.SetText("Unit Should Not Be Empty", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
//                    flag = false;
//                }
//                else
//                {
//                    if (Left2 == "01")
//                    {
//                        Recordset recordset = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
//                        recordset.DoQuery("Select U_Category ,U_Perntage ,U_Tolarence  from [@AIS_REM1]  Where U_Unit   ='" + Left2 + "'");
//                        if (recordset.RecordCount > 0)
//                        {
//                            recordset.MoveFirst();
//                        }
//                        double num1 = 0.0;
//                        double num2 = 0.0;
//                        double num3 = 0.0;
//                        double num4 = 0.0;
//                        double num5 = 0.0;
//                        double num6 = 0.0;
//                        double num7 = 0.0;
//                        double num8 = 0.0;
//                        double num9 = 0.0;
//                        double num10 = 0.0;
//                        double num11 = 0.0;
//                        double num12 = 0.0;
//                        int recordCount = recordset.RecordCount;
//                        int num13 = 1;
//                        while (num13 <= recordCount)
//                        {
                            
//                            string str2 = Convert.ToString(recordset.Fields.Item("U_Category").Value);
                            
//                            string str3 = Convert.ToString(recordset.Fields.Item("U_Perntage").Value);
                            
//                            string str4 = Convert.ToString(recordset.Fields.Item("U_Tolarence").Value);
//                            double num14 = Convert.ToDouble(str4) + Convert.ToDouble(str3);
//                            if (str2.ToUpper().Trim() == "M1")
//                            {
//                                num1 = num14;
//                                num2 = Convert.ToDouble(str3) - Convert.ToDouble(str4);
//                            }
//                            else if (str2.ToUpper().Trim() == "M2")
//                            {
//                                num3 = num14;
//                                num4 = Convert.ToDouble(str3) - Convert.ToDouble(str4);
//                            }
//                            else if (str2.ToUpper().Trim() == "M3")
//                            {
//                                num5 = num14;
//                                num6 = Convert.ToDouble(str3) - Convert.ToDouble(str4);
//                            }
//                            else if (str2.ToUpper().Trim() == "ATTA")
//                            {
//                                num7 = num14;
//                                num8 = Convert.ToDouble(str3) - Convert.ToDouble(str4);
//                            }
//                            else if (str2.ToUpper().Trim() == "SUJI")
//                            {
//                                num9 = num14;
//                                num10 = Convert.ToDouble(str3) - Convert.ToDouble(str4);
//                            }
//                            else if (str2.ToUpper().Trim() == "BRAN")
//                            {
//                                num11 = num14;
//                                num12 = Convert.ToDouble(str3) - Convert.ToDouble(str4);
//                            }
                            
//                            recordset.MoveNext();
//                            ++num13;
//                        }
//                        int visualRowCount = oMatrix2.VisualRowCount;
//                        int num15 = 1;
//                        while (num15 <= visualRowCount)
//                        {
                            
//                            string Left3 = oDBDSDetail2.GetValue("U_Date", checked(num15 - 1)).ToString().Trim();
                            
//                            string str2 = oDBDSDetail2.GetValue("U_M1", checked(num15 - 1)).ToString().Trim();
                            
//                            string str3 = oDBDSDetail2.GetValue("U_M2", checked(num15 - 1)).ToString().Trim();
                            
//                            string str4 = oDBDSDetail2.GetValue("U_M3", checked(num15 - 1)).ToString().Trim();
                            
//                            string str5 = oDBDSDetail2.GetValue("U_Atta", checked(num15 - 1)).ToString().Trim();
                            
//                            string str6 = oDBDSDetail2.GetValue("U_Suji", checked(num15 - 1)).ToString().Trim();
                            
//                            string str7 = oDBDSDetail2.GetValue("U_Bran", checked(num15 - 1)).ToString().Trim();
                            
//                            oDBDSDetail2.GetValue("U_Total", checked(num15 - 1)).ToString().Trim();
//                            double num14 = Convert.ToDouble(str2) + Convert.ToDouble(str3) + Convert.ToDouble(str4);
//                            int integer = Convert.ToInt32(objSBOAPI.Query_Execute("select Count(T0.DocEntry) as Available  from [@AIS_OCAC] T0 Inner Join [@AIS_CAC1] T1 On T0.DocEntry =T1.DocEntry   Where  T0.U_Unit ='" + Left2 + "' And  CONVERT (Varchar(10),U_PlanDate,112) = '" + Left3 + "'"));
//                            if (Left3 != string.Empty & integer > 0)
//                            {
//                                double num16 = num2 + num4 + num6;
//                                double num17 = num1 + num3 + num5;
//                                if (!(num14 >= num16 & num14 <= num17))
//                                {
                                    
//                                    objSBOAPI.SBO_Appln.StatusBar.SetText("M1,M2 & M3 Output Percentage is not matched in Yield Percentage for " + Left3, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                    flag = false;
//                                    return flag;
//                                }
//                                else if (num8 > 0.0 && !(Math.Round(Convert.ToDouble(str5), 1) >= num8 & Math.Round(Convert.ToDouble(str5), 1) <= num7))
//                                {
                                    
//                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Atta Output Percentage is not matched in Yield Percentage for " + Left3, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                    flag = false;
//                                    return flag;
//                                }
//                                else if (num10 > 0.0 && !(Math.Round(Convert.ToDouble(str6), 1) >= Math.Round(num10, 1) & Math.Round(Convert.ToDouble(str6), 1) <= Math.Round(num9, 1)))
//                                {
                                    
//                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Suji Output Percentage is not matched in Yield Percentage for " + Left3, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                    flag = false;
//                                    return flag;
//                                }
//                                else if (num12 > 0.0 && !(Math.Round(Convert.ToDouble(str7), 0) >= num12 & Math.Round(Convert.ToDouble(str7), 0) <= num11))
//                                {
                                    
//                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Bran Output Percentage is not matched in Yield Percentage for " + Left3, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                                    flag = false;
//                                    return flag;
//                                }
//                            }
//                            ++num15;
//                        }
//                    }
//                    //if (!LoadMatrixFromGrid_1())
//                    //    flag = false;
//                    //else if (!LoadMatrixFromGrid_2())
//                    //    flag = false;
//                    //else if (!LoadMatrixFromGrid_3())
//                    //    flag = false;
//                    //else if (!LoadMatrixFromGrid_4())
//                    //    flag = false;
//                    //else if (!LoadMatrixFromGrid_5())
//                    //    flag = false;
//                    //else if (!LoadMatrixFromGrid_6())
//                    //    flag = false;
//                    //else if (!LoadMatrixFromGrid_7())
//                    //{
//                    //    flag = false;
//                    //}
//                    //else
//                    //{
//                    //    //UpdatePresalesDetails();
//                    //    flag = true;
//                    //}
//                }
//            }
//            catch (Exception ex)
//            {
//                objform.Freeze(false);
//                objSBOAPI.SBO_Appln.StatusBar.SetText("Validation Function Failed: " + ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
//            }
//            return flag;
//        }
//    }
//}
