using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using SAPbouiCOM;
using System.Media;

namespace JKFM_Source
{
    class Cls_CustomerWiseItemAllocation
    {
        #region Declaration
        SAPbouiCOM.Form objform;
        ClsSBO objSBOAPI;
        SAPbouiCOM.Form frmBranchMaster;
        SAPbouiCOM.DBDataSource oDBDSDetail;
        SAPbouiCOM.DBDataSource oDBDSHeader;
        string UDOID;
        SAPbouiCOM.Matrix oMatrix;
        #endregion        

        #region Constructor
        public Cls_CustomerWiseItemAllocation(ClsSBO objSBO)
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
                    frmBranchMaster = objSBOAPI.SBO_Appln.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount);
                    switch (pval.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_CLICK:
                            switch(pval.ItemUID)
                            {
                                case "1":
                                    if((frmBranchMaster.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || frmBranchMaster.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) && !ValidateAll())
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
                                    if (pval.ActionSuccess == true && frmBranchMaster.Mode == BoFormMode.fm_ADD_MODE)
                                    {
                                        InitForm();
                                    }
                                    break;
                            }
                            break;
                        case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                            switch(pval.ItemUID)
                            {
                                case "t_CardCode":
                                    string Broker_Name = "";
                                    if(oDBDSHeader.GetValue("U_CardCode",0) != "")
                                    {
                                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                        {
                                            Broker_Name = objSBOAPI.Query_Execute("Select B.\"SlpName\" from \"OCRD\" A INNER JOIN \"OSLP\" B ON A.\"SlpCode\"=B.\"SlpCode\" Where A.\"CardCode\" = '" + oDBDSHeader.GetValue("U_CardCode", 0) + "'");
                                        }
                                        else
                                        {
                                            Broker_Name = objSBOAPI.Query_Execute("Select B.SlpName from OCRD A INNER JOIN OSLP B ON A.SlpCode=B.SlpCode Where A.CardCode='" + oDBDSHeader.GetValue("U_CardCode", 0) + "'");
                                        }
                                        oDBDSHeader.SetValue("U_AVA_BrName", 0, Convert.ToString(Broker_Name).Trim());
                                        //oMatrix.Columns.Item("ItemCode").Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                                    }
                                    
                                    break;
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

                        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:

                            SAPbouiCOM.ChooseFromListEvent chooseFromListEvent = (ChooseFromListEvent)pval;
                            SAPbouiCOM.DataTable selectedObjects = chooseFromListEvent.SelectedObjects;
                            if (selectedObjects != null)
                            {
                                switch (pval.ItemUID)
                                {
                                    case "t_CardCode":
                                        try
                                        {
                                            oDBDSHeader.SetValue("U_CardCode", 0, Convert.ToString(selectedObjects.GetValue("CardCode", 0)).Trim());
                                        }
                                        catch(Exception ex)
                                        {

                                        }
                                        
                                        try
                                        {
                                            oDBDSHeader.SetValue("U_CardName", 0, Convert.ToString(selectedObjects.GetValue("CardName", 0)).Trim());
                                        }
                                        catch(Exception ex)
                                        {

                                        }
                                        
                                        if (frmBranchMaster.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                        {
                                            frmBranchMaster.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                        }
                                        break;
                                    case "Matrix":
                                        switch (pval.ColUID)
                                        {
                                            case "ItemCode":
                                                oMatrix.FlushToDataSource();
                                                oDBDSDetail.SetValue("U_ItemCode", pval.Row - 1, Convert.ToString(selectedObjects.GetValue("ItemCode", 0)).Trim());
                                                oDBDSDetail.SetValue("U_ItemName", pval.Row - 1, Convert.ToString(selectedObjects.GetValue("ItemName", 0)).Trim());
                                                oMatrix.LoadFromDataSource();
                                                oMatrix.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                                                objSBOAPI.AddNewLine(oMatrix, oDBDSDetail, pval.Row, pval.ColUID);
                                                if (frmBranchMaster.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                                {
                                                    frmBranchMaster.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                                }
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
                frmBranchMaster.Freeze(false);
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
                        case "OCIA":
                            frmBranchMaster = objSBOAPI.LoadForm("CustomerWiseItemAllocation.xml", "OCIA");
                            //frmBranchMaster = EventHandler.oApplication.Forms.Item(LVariables.CustomerWiseItemAllocationFormID);
                            oDBDSHeader = frmBranchMaster.DataSources.DBDataSources.Item("@AIS_OCIA");
                            oDBDSDetail = frmBranchMaster.DataSources.DBDataSources.Item("@AIS_CIA1");
                            oMatrix = (SAPbouiCOM.Matrix)frmBranchMaster.Items.Item("Matrix").Specific;
                            frmBranchMaster.Mode = BoFormMode.fm_ADD_MODE;
                            DefineModesForFields();
                            InitForm();
                            break;

                        case "1282":
                            InitForm();
                            break;

                        case "1292":
                            objSBOAPI.AddNewLine(oMatrix, oDBDSDetail, 1, "");
                            oMatrix.ClearRowData(oMatrix.VisualRowCount);
                            break;

                        case "":
                            if (pval.MenuUID  != "1293")
                                return;
                            DeleteRow(oMatrix, oDBDSDetail);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                frmBranchMaster.Freeze(false);
                objSBOAPI.SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #region InitForm
        public void InitForm()
        {
            try
            {
                frmBranchMaster.Freeze(true);
                oMatrix.Clear();
                oDBDSDetail.Clear();
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    oDBDSHeader.SetValue("Code", 0, Convert.ToString(objSBOAPI.GetCodeGeneration("@AIS_OCIA")));
                }
                else
                {
                    oDBDSHeader.SetValue("Code", 0, Convert.ToString(objSBOAPI.GetCodeGeneration("[@AIS_OCIA]")));
                }
                    
                objSBOAPI.AddNewLine(oMatrix, oDBDSDetail, 1, "");
                frmBranchMaster.ActiveItem = "t_CardCode";
            }
            catch (Exception ex)
            {
                frmBranchMaster.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Init Form Method Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                frmBranchMaster.Freeze(false);
            }
        }
        #endregion

        #region DefineModesForFields
        public void DefineModesForFields()
        {
            try
            {
                frmBranchMaster.Items.Item("t_Code").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);
                frmBranchMaster.Items.Item("t_CardName").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);
            }
            catch (Exception ex)
            {
                frmBranchMaster.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Define Modes For Fields Method Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region DeleteRow
        public void DeleteRow(SAPbouiCOM.Matrix oMatrix, SAPbouiCOM.DBDataSource oDBDSDetail)
        {
            try
            {
                oMatrix.FlushToDataSource();
                int visualRowCount = oMatrix.VisualRowCount;
                int RowNum = 1;
                while (RowNum <= visualRowCount)
                {
                    oMatrix.GetLineData(RowNum);
                    oDBDSDetail.Offset = RowNum - 1;
                    oDBDSDetail.SetValue("LineID", oDBDSDetail.Offset, RowNum.ToString());
                    oMatrix.SetLineData(RowNum);
                    oMatrix.FlushToDataSource();
                    ++RowNum;
                }
                oDBDSDetail.RemoveRecord(oDBDSDetail.Size - 1);
                oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                frmBranchMaster.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Delete Row Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region ValidateAll
        public bool ValidateAll()
        {
            bool flag;
            try
            {
                flag = true;
            }
            catch (Exception ex)
            {
                flag = false;
                frmBranchMaster.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Validate Function Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            return flag;
        }
        #endregion
    }
}
