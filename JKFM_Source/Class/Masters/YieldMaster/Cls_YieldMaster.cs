using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Media;

namespace JKFM_Source
{
    class Cls_YieldMaster
    {
        #region Declaration
        SAPbouiCOM.Form objform;
        ClsSBO objSBOAPI;

        SAPbouiCOM.DBDataSource oDBDSDetail;
        SAPbouiCOM.DBDataSource oDBDSHeader;
        string UDOID;
        SAPbouiCOM.Matrix oMatSeason;
        bool sFormID;
        #endregion        

        #region Constructor
        public Cls_YieldMaster(ClsSBO objSBO)
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
                        case BoEventTypes.et_ITEM_PRESSED:
                            switch(pval.ItemUID)
                            {
                                case "1":
                                    if (objform.Mode == BoFormMode.fm_ADD_MODE || objform.Mode == BoFormMode.fm_UPDATE_MODE)
                                    {

                                        objform.Freeze(true);
                                        if (!ValidateAll())
                                        {
                                            SystemSounds.Asterisk.Play();
                                            bubbleevent = false;
                                        }
                                        else if (oMatSeason.VisualRowCount > 1)
                                            objSBOAPI.DeleteEmptyRowInFormDataEvent(oMatSeason, "Category", oDBDSDetail);

                                        objform.Freeze(false);
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
                                case "MatSeason":
                                    switch (pval.ColUID)
                                    {
                                        case "Category":
                                            oCFL = oCFLs.Item("CFL_Item");
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
                                            oCon.Alias = "U_IsProd";
                                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                            oCon.CondVal = "Y";
                                            oCon.BracketCloseNum = 1;
                                            oCFL.SetConditions(oCons);
                                            GC.Collect();
                                            break;
                                    }
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

                            switch (pval.ItemUID)
                            {
                                case "1":
                                    if ((pval.Action_Success == true && objform.Mode == BoFormMode.fm_ADD_MODE) || objform.Mode == BoFormMode.fm_UPDATE_MODE)
                                    {
                                        if (objform.Mode == BoFormMode.fm_ADD_MODE)
                                            InitForm();
                                        else if (objform.Mode == BoFormMode.fm_OK_MODE)
                                        {
                                            objform.Freeze(true);
                                            objSBOAPI.AddNewLine(oMatSeason, oDBDSDetail, oMatSeason.VisualRowCount, "Category");
                                            objform.Freeze(false);
                                        }
                                    }
                                    break;
                            }
                            break;

                        //case BoEventTypes.et_GOT_FOCUS:

                        //    switch (pval.ItemUID)
                        //    {
                        //        case "MatSeason":
                        //            switch (pval.ColUID)
                        //            {
                        //                case "Category":
                        //                    GFun.ChooseFromListFilteration(objform, "CFL_Item", "U_IsProd", "Select 'Y'");
                        //                    break;
                        //            }
                        //            break;
                        //    }
                        //    break;

                        case BoEventTypes.et_LOST_FOCUS:

                            switch (pval.ItemUID)
                            {
                                case "MatSeason":
                                    switch (pval.ColUID)
                                    {
                                        case "Category":
                                            objSBOAPI.AddNewLine(oMatSeason, oDBDSDetail, oMatSeason.VisualRowCount, pval.ColUID);
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
                                case "MatSeason":
                                    switch (pval.ColUID)
                                    {
                                        case "Category":
                                            if (selectedObjects != null)
                                            {
                                                oMatSeason.FlushToDataSource();
                                                oDBDSDetail.SetValue("U_Category", pval.Row - 1, Convert.ToString(selectedObjects.GetValue("ItemCode", 0)).Trim());
                                                oMatSeason.LoadFromDataSource();
                                                oMatSeason.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(BoCellClickType.ct_Regular, 0);
                                                objSBOAPI.AddNewLine(oMatSeason, oDBDSDetail, pval.Row, pval.ColUID);
                                            }
                                            break;
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
                    switch (pval.MenuUID)
                    {
                        case "OREM":
                            objform = objSBOAPI.LoadForm("YieldMaster.xml", "OREM");
                            objform = objSBOAPI.SBO_Appln.Forms.Item("OREM");
                            
                            //objform.Items.Item("t_Code").Visible = false;
                            //objform.Items.Item("t_Code").Enabled = true;
                            oDBDSHeader = objform.DataSources.DBDataSources.Item("@AIS_OREM");
                            oDBDSDetail = objform.DataSources.DBDataSources.Item("@AIS_REM1");
                            oMatSeason = (Matrix)objform.Items.Item("MatSeason").Specific;
                            DefineModesForFields();
                            sFormID = true;
                            InitForm();
                            break;

                        case "1282":
                            InitForm();
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
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            if (BusinessObjectInfo.ActionSuccess == true)                            {
                                int visualRowCount = oMatSeason.VisualRowCount;
                                int RowNum = 1;
                                while (RowNum <= visualRowCount)
                                {
                                    if (oMatSeason.Columns.Item("Category").Cells.Item(RowNum).Specific.Value != string.Empty)
                                    {
                                        oMatSeason.CommonSetting.SetCellEditable(RowNum, 1, false);
                                    }
                                    else
                                    {
                                        oMatSeason.CommonSetting.SetCellEditable(RowNum, 1, true);
                                    }
                                    ++RowNum;
                                }
                            }                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Form Data Event Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
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
                oMatSeason.Clear();
                string QueryStr = null;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    QueryStr = "Select top 1  \"Code\" from \"@AIS_OREM\"";
                }
                else
                {
                    QueryStr = "Select top 1  Code from [@AIS_OREM]";
                }
                    
                Recordset Orec = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                Orec.DoQuery(QueryStr);
                if (Orec.RecordCount > 0)
                {
                    Orec.MoveFirst();
                    oDBDSDetail.Clear();
                    objSBOAPI.AddNewLine(oMatSeason, oDBDSDetail, 1, "");
                    if (objform.Mode != BoFormMode.fm_FIND_MODE)
                        objform.Mode = BoFormMode.fm_FIND_MODE;
                    objform.Items.Item("t_Code").Specific.Value = Orec.Fields.Item("Code").Value.ToString().Trim();
                   
                    objform.Items.Item("1").Click(BoCellClickType.ct_Regular);
                    objform.Mode = BoFormMode.fm_OK_MODE;
                }
                else if (Orec.RecordCount == 0)
                {
                    if (objform.Mode == BoFormMode.fm_FIND_MODE)
                        objform.Mode = BoFormMode.fm_ADD_MODE;
                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        oDBDSHeader.SetValue("Code", 0, Convert.ToString(objSBOAPI.GetCodeGeneration("@AIS_OREM")));
                        //oDBDSHeader.SetValue("DocEntry", 0, Convert.ToString(objSBOAPI.GetCodeGeneration("@AIS_OREM")));
                    }
                    else
                    {
                        oDBDSHeader.SetValue("Code", 0, Convert.ToString(objSBOAPI.GetCodeGeneration("[@AIS_OREM]")));
                    }
                        
                }
                objSBOAPI.AddNewLine(oMatSeason, oDBDSDetail, oMatSeason.VisualRowCount, "Category");
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    objSBOAPI.setComboBoxValue((ComboBox)oMatSeason.Columns.Item("Unit").Cells.Item(1).Specific, "SELECT  \"U_BranchCode\" , \"U_BranchName\" FROM \"@AIS_BRN1\"  Where \"U_Active\"='Y'");
                }
                else
                {
                    objSBOAPI.setComboBoxValue((ComboBox)oMatSeason.Columns.Item("Unit").Cells.Item(1).Specific, "SELECT  U_BranchCode , U_BranchName FROM [@AIS_BRN1]  Where U_Active='Y'");
                }
                    
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Init Form Method Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
            }
            finally
            {
                objform.Freeze(false);
            }
        }
        #endregion

        #region DefineModesForFields
        public void DefineModesForFields()
        {
            try
            {
                objform.Items.Item("t_Code").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Define Modes For Fields Method Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
            }
        }
        #endregion

        #region ValidateAll
        public bool ValidateAll()
        {
            bool flag;
            try
            {
                if (oMatSeason.VisualRowCount <= 1)
                {
                    objSBOAPI.SBO_Appln.SetStatusBarMessage("Empty Document should not be Added...", BoMessageTime.bmt_Short, true);
                    flag = false;
                    return flag;
                }
                else
                {
                    int visualRowCount = oMatSeason.VisualRowCount;
                    int num = 1;
                    while (num <= visualRowCount)
                    {
                        string Category = oMatSeason.Columns.Item("Category").Cells.Item(num).Specific.Value.ToString().ToUpper().Trim();
                        string Unit = oMatSeason.Columns.Item("Unit").Cells.Item(num).Specific.Value.ToString().Trim();
                        if (!(Category == string.Empty & Unit == string.Empty) && Category == string.Empty | Unit == string.Empty)
                        {
                            objSBOAPI.SBO_Appln.SetStatusBarMessage("For LineId " + Convert.ToString(num) + ", Category and Unit Should Not be Empty...", BoMessageTime.bmt_Short, true);
                            flag = false;
                            return flag;
                        }
                        else
                            ++num;
                    }
                    flag = true;
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Validate Function Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                flag = false;
            }
            return flag;
        }
        #endregion
    }
}
