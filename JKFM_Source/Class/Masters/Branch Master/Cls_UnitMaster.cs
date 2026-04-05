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
    class Cls_UnitMaster
    {
        #region Declaration
        SAPbouiCOM.Form objform;
        ClsSBO objSBOAPI;
        SAPbouiCOM.DBDataSource oDBDSDetail;
        SAPbouiCOM.DBDataSource oDBDSHeader;
        string UDOID;
        SAPbouiCOM.Matrix oMatrix;
        #endregion        

        #region Constructor
        public Cls_UnitMaster(ClsSBO objSBO)
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
                                    if ((objform.Mode == BoFormMode.fm_ADD_MODE || objform.Mode == BoFormMode.fm_UPDATE_MODE) && !ValidateAll())
                                    {
                                        SystemSounds.Asterisk.Play();
                                        bubbleevent = false;
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
                            switch(pval.ItemUID)
                            {
                                case "1":
                                    if (pval.ActionSuccess == true & objform.Mode == BoFormMode.fm_ADD_MODE)
                                        InitForm();
                                    if (pval.ActionSuccess == true & objform.Mode == BoFormMode.fm_OK_MODE)
                                        FormEnableDisableFields();
                                    break;
                            }
                            break;

                        case BoEventTypes.et_LOST_FOCUS:
                            switch(pval.ItemUID)
                            {
                                case "Matrix":
                                    switch (pval.ColUID)
                                    {
                                        case "BranchCode":
                                            objSBOAPI.AddNewLine(oMatrix, oDBDSDetail, pval.Row, pval.ColUID);
                                            break;
                                    }
                                    break;
                            }
                            break;

                        case BoEventTypes.et_CHOOSE_FROM_LIST:
                            ChooseFromListEvent chooseFromListEvent = (ChooseFromListEvent)pval;
                            DataTable selectedObjects = chooseFromListEvent.SelectedObjects;
                            if (selectedObjects != null)
                            {
                                switch(pval.ItemUID)
                                {
                                    case "Matrix":
                                        switch (pval.ColUID)
                                        {
                                            case "WhsCode":
                                                oMatrix.FlushToDataSource();
                                                oDBDSDetail.SetValue("U_WhsCode", pval.Row - 1, Convert.ToString(selectedObjects.GetValue("WhsCode", 0)).Trim());
                                                oMatrix.LoadFromDataSource();
                                                break;
                                        }
                                        break;
                                }
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
                        case "OBRN":
                            objform = objSBOAPI.LoadForm("BranchMaster.xml", "OBRN");
                            objform = objSBOAPI.SBO_Appln.Forms.Item("OBRN");
                            oDBDSHeader = objform.DataSources.DBDataSources.Item("@AIS_OBRN");
                            oDBDSDetail = objform.DataSources.DBDataSources.Item("@AIS_BRN1");
                            oMatrix = (Matrix)objform.Items.Item("Matrix").Specific;
                            objform.Mode = BoFormMode.fm_ADD_MODE;
                            DefineModesForFields();
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
                                FormEnableDisableFields();
                            }                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                objSBOAPI.SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #region DefineModesForFields
        public void DefineModesForFields()
        {
            try
            {
                objform.Items.Item("t_Code").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);
                objform.Items.Item("t_Code").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Visible, 4, BoModeVisualBehavior.mvb_True);
                objform.Items.Item("t_Code").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Visible, 1, BoModeVisualBehavior.mvb_False);
                objform.Items.Item("t_Code").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Visible, 2, BoModeVisualBehavior.mvb_False);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Define Modes For Fields Method Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region InitForm
        public void InitForm()
        {
            try
            {
                objform.Freeze(true);
                oMatrix.Clear();
                string QueryStr = null;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    QueryStr = "select top 1 \"Code\" from \"@AIS_OBRN\" ";
                }
                else
                {
                    QueryStr = "select top 1 Code from [@AIS_OBRN] ";
                }   
                Recordset Orec = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                Orec.DoQuery(QueryStr);
                if (Orec.RecordCount > 0)
                {
                    Orec.MoveFirst();
                    objform.Mode = BoFormMode.fm_FIND_MODE;
                    objform.Items.Item("t_Code").Specific.Value = Orec.Fields.Item("Code").Value.ToString().Trim();
                    objform.Items.Item("1").Click(BoCellClickType.ct_Regular);
                }
                else
                {
                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        oDBDSHeader.SetValue("Code", 0, Convert.ToString(objSBOAPI.GetCodeGeneration("@AIS_OBRN")));
                    }
                    else
                    {
                        oDBDSHeader.SetValue("Code", 0, Convert.ToString(objSBOAPI.GetCodeGeneration("[@AIS_OBRN]")));
                    }
                        
                    objSBOAPI.AddNewLine(oMatrix, oDBDSDetail, 1, "");
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Init Form Method Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                objform.Freeze(false);
            }
        }
        #endregion

        #region FormEnableDisableFields
        public void FormEnableDisableFields()
        {
            try
            {
                int visualRowCount = oMatrix.VisualRowCount;
                int RowNum = 1;
                while (RowNum <= visualRowCount)
                {
                    if (oMatrix.Columns.Item("BranchCode").Cells.Item(RowNum).Specific.Value!=string.Empty)
                    {
                        oMatrix.CommonSetting.SetCellEditable(RowNum, 1, false);
                    }
                    else
                    {
                        oMatrix.CommonSetting.SetCellEditable(RowNum, 1, true);
                    }
                    ++RowNum; 
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Form Enable Disable Event Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region ValidateAll
        public bool ValidateAll()
        {
            bool flag;
            try
            {
                if (oMatrix.VisualRowCount <= 1)
                {
                    objSBOAPI.SBO_Appln.SetStatusBarMessage("Empty Document should not be Added...", BoMessageTime.bmt_Short, true);
                    flag = false;
                    return flag;
                }
                else
                {
                    int visualRowCount = oMatrix.VisualRowCount;
                    int num1 = 1;
                    while (num1 <= visualRowCount)
                    {
                        string str = oMatrix.Columns.Item("BranchCode").Cells.Item(num1).Specific.Value.ToString().ToUpper().Trim();
                        string Left = oMatrix.Columns.Item("BranchName").Cells.Item(num1).Specific.Value.ToString().Trim();
                        if (!(str == string.Empty && Left == string.Empty))
                        {
                            if (str == string.Empty || Left == string.Empty)
                            {
                                objSBOAPI.SBO_Appln.StatusBar.SetText("For LineId " + Convert.ToString(num1) + ", Unit Code and Unit Name Should Not be Empty...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                flag = false;
                                return flag;
                            }
                            else
                            {
                                int num2 = num1 + 1;
                                int num3 = oMatrix.VisualRowCount - 1;
                                int num4 = num2;
                                while (num4 <= num3)
                                {
                                    if (!(oMatrix.Columns.Item("BranchCode").Cells.Item(num4).Specific.Value.ToString().Trim() == string.Empty && oMatrix.Columns.Item("BranchName").Cells.Item(num4).Specific.Value.ToString().Trim() == string.Empty) && oMatrix.Columns.Item("BranchCode").Cells.Item(num4).Specific.Value.ToString().ToUpper().Trim() == str)
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("For LineId " + Convert.ToString(num4) + ", Dublicate Entry shouldn't be allowed...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        flag = false;
                                        return flag;
                                    }
                                    else
                                        ++num4; 
                                }
                            }
                        }
                        ++num1; 
                    }
                    flag = true;
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Validate Function Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                flag = false;
            }
            return flag;
        }
        #endregion
    }
}
