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
    class Cls_PackingCapacity
    {
        #region Declaration
        SAPbouiCOM.Form objform;
        ClsSBO objSBOAPI;

        SAPbouiCOM.DBDataSource oDBDSDetail1;
        SAPbouiCOM.DBDataSource oDBDSDetail2;
        SAPbouiCOM.DBDataSource oDBDSHeader;
        string UDOID;
        SAPbouiCOM.Matrix oMatrix1;
        SAPbouiCOM.Matrix oMatrix2;
        string strMatrixID;
        #endregion        

        #region Constructor
        public Cls_PackingCapacity(ClsSBO objSBO)
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
                                    if (pval.ActionSuccess == true && objform.Mode == BoFormMode.fm_ADD_MODE)
                                    {
                                        InitForm();
                                    }
                                    break;
                                case "fld_Unit1":
                                    objform.PaneLevel = 1;
                                    objform.Settings.MatrixUID = "Matrix1";
                                    objform.Items.Item(pval.ItemUID).AffectsFormMode = false;
                                    break;
                                case "fld_Unit2":
                                    objform.PaneLevel = 2;
                                    objform.Settings.MatrixUID = "Matrix2";
                                    objform.Items.Item(pval.ItemUID).AffectsFormMode = false;
                                    break;
                           }
                           break;

                        case BoEventTypes.et_LOST_FOCUS:
                            switch(pval.ItemUID)
                            {
                                case "Matrix1":
                                    switch (pval.ColUID)
                                    {
                                        case "ItemCode":
                                            objSBOAPI.AddNewLine(oMatrix1, oDBDSDetail1, pval.Row, pval.ColUID);
                                            break;
                                    }
                                    break;
                            }
                            break;

                        case BoEventTypes.et_VALIDATE:
                            switch(pval.ItemUID)
                            {
                                case "Matrix2":
                                    switch(pval.ColUID)
                                    {
                                        case "Unt2QtyBag":
                                            if (pval.ItemChanged)
                                            {
                                                string singleValue = null;
                                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                                {
                                                    singleValue = objSBOAPI.Query_Execute(" Select \"SalPackUn\" from \"OITM\" Where  \"ItemCode\" = '" + oMatrix2.Columns.Item("ItemCode").Cells.Item(pval.Row).Specific.Value.ToString().Trim() + "'  ");
                                                }
                                                else
                                                {
                                                    singleValue = objSBOAPI.Query_Execute(" Select SalPackUn from OITM Where  ItemCode = '" + oMatrix2.Columns.Item("ItemCode").Cells.Item(pval.Row).Specific.Value.ToString().Trim() + "'  ");
                                                }
                                                    
                                                double num = Convert.ToDouble(oMatrix2.Columns.Item("Unt2QtyBag").Cells.Item(pval.Row).Specific.Value) * Convert.ToDouble(singleValue) / 1000.0;
                                                oMatrix2.FlushToDataSource();
                                                oDBDSDetail2.SetValue("U_Unt2QtyTon", pval.Row - 1, num.ToString());
                                                oMatrix2.LoadFromDataSource();
                                            }
                                            break;
                                    }
                                    break;
                                case "Matrix1":
                                    switch(pval.ColUID)
                                    {
                                        case "Unt1QtyBag":
                                            if (pval.ItemChanged)
                                            {
                                                string singleValue = null;
                                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                                {
                                                    singleValue = objSBOAPI.Query_Execute(" Select \"SalPackUn\" from \"OITM\" Where  \"ItemCode\" = '" + oMatrix1.Columns.Item("ItemCode").Cells.Item(pval.Row).Specific.Value.ToString().Trim() + "'  ");
                                                }
                                                else
                                                {
                                                    singleValue = objSBOAPI.Query_Execute(" Select SalPackUn from OITM Where  ItemCode = '" + oMatrix1.Columns.Item("ItemCode").Cells.Item(pval.Row).Specific.Value.ToString().Trim() + "'  ");
                                                }
                                                    
                                                double num = Convert.ToDouble(oMatrix1.Columns.Item("Unt1QtyBag").Cells.Item(pval.Row).Specific.Value) * Convert.ToDouble(singleValue) / 1000.0;
                                                oMatrix1.FlushToDataSource();
                                                oDBDSDetail1.SetValue("U_Unt1QtyTon", pval.Row - 1, num.ToString());
                                                oMatrix1.LoadFromDataSource();
                                            }
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
                                switch (pval.ItemUID)
                                {
                                    case "Matrix1":
                                        switch (pval.ColUID)
                                        {
                                            case "ItemCode":
                                                oMatrix1.FlushToDataSource();
                                                oDBDSDetail1.SetValue("U_ItemCode", (pval.Row - 1), Convert.ToString(selectedObjects.GetValue("ItemCode", 0)).Trim());
                                                oDBDSDetail1.SetValue("U_ItemName", (pval.Row - 1), Convert.ToString(selectedObjects.GetValue("ItemName", 0)).Trim());
                                                oMatrix1.LoadFromDataSource();
                                                oMatrix1.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(BoCellClickType.ct_Regular, 0);
                                                objSBOAPI.AddNewLine(oMatrix1, oDBDSDetail1, pval.Row, pval.ColUID);
                                                if (objform.Mode == BoFormMode.fm_OK_MODE)
                                                    objform.Mode = BoFormMode.fm_UPDATE_MODE;
                                                break;
                                        }
                                        break;
                                    case "Matrix2":
                                        switch (pval.ColUID)
                                        {
                                            case "ItemCode":
                                                oMatrix2.FlushToDataSource();
                                                oDBDSDetail2.SetValue("U_ItemCode", (pval.Row - 1), Convert.ToString(selectedObjects.GetValue("ItemCode", 0)).Trim());
                                                oDBDSDetail2.SetValue("U_ItemName", (pval.Row - 1), Convert.ToString(selectedObjects.GetValue("ItemName", 0)).Trim());
                                                oMatrix2.LoadFromDataSource();
                                                oMatrix2.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(BoCellClickType.ct_Regular, 0);
                                                objSBOAPI.AddNewLine(oMatrix2, oDBDSDetail2, pval.Row, pval.ColUID);
                                                if (objform.Mode == BoFormMode.fm_OK_MODE)
                                                    objform.Mode = BoFormMode.fm_UPDATE_MODE;
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
                        case "ORMC":
                            objform = objSBOAPI.LoadForm("PackingCapacity.xml", "ORMC");
                            objform = objSBOAPI.SBO_Appln.Forms.Item("ORMC");
                            oDBDSHeader = objform.DataSources.DBDataSources.Item("@AIS_ORMC");
                            oDBDSDetail1 = objform.DataSources.DBDataSources.Item("@AIS_RMC1");
                            oDBDSDetail2 = objform.DataSources.DBDataSources.Item("@AIS_RMC2");
                            oMatrix1 = (Matrix)objform.Items.Item("Matrix1").Specific;
                            oMatrix2 = (Matrix)objform.Items.Item("Matrix2").Specific;
                            objform.Mode = BoFormMode.fm_ADD_MODE;
                            DefineModesForFields();
                            InitForm();
                            objform.EnableMenu("1292", true);
                            objform.EnableMenu("1293", true);
                            break;

                        case "1282":
                            InitForm();
                            break;

                        case "1281":
                            InitForm();
                            break;

                        case "1292":
                            if (strMatrixID == "Matrix1")
                            {
                                objSBOAPI.AddNewLine(oMatrix1, oDBDSDetail1, 1, "");
                                oMatrix1.ClearRowData(oMatrix1.VisualRowCount);
                            }
                            else
                            {
                                if (strMatrixID  != "Matrix2")
                                    return;
                                objSBOAPI.AddNewLine(oMatrix2, oDBDSDetail2, 1, "");
                                oMatrix2.ClearRowData(oMatrix2.VisualRowCount);
                            }
                            break;

                        case "1293":
                            if (strMatrixID == "Matrix1")
                                DeleteRow(oMatrix1, oDBDSDetail1);
                            else if (strMatrixID == "Matrix2")
                                DeleteRow(oMatrix2, oDBDSDetail2);
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

        #region Right Click
        /// <summary>
        /// Enabling Right click button for Add and Delete row in Routing Form
        /// </summary>
        /// <param name="eventInfo">The object that holds the event information</param>
        /// <param name="BubbleEvent">Indicates how the application handles the event. Relevant only when ContextMenuInfo.BeforeAction is true.</param>
        public void Right_Click(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (eventInfo.BeforeAction == true)
                {
                    switch (eventInfo.EventType)
                    {
                        case BoEventTypes.et_RIGHT_CLICK:
                            strMatrixID = eventInfo.ItemUID;
                            break;
                    }
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region InitForm
        public void InitForm()
        {
            try
            {
                objform.Freeze(true);
                oMatrix1.Clear();
                string QueryStr = null;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    QueryStr = "Select top 1 \"Code\" from \"@AIS_ORMC\" ";
                }
                else
                {
                    QueryStr = "Select top 1 Code from [@AIS_ORMC] ";
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
                        oDBDSHeader.SetValue("Code", 0, Convert.ToString(objSBOAPI.GetCodeGeneration("@AIS_ORMC")));
                    }
                    else
                    {
                        oDBDSHeader.SetValue("Code", 0, Convert.ToString(objSBOAPI.GetCodeGeneration("[@AIS_ORMC]")));
                    }
                        
                    objSBOAPI.AddNewLine(oMatrix1, oDBDSDetail1, 1, "");
                    objSBOAPI.AddNewLine(oMatrix2, oDBDSDetail2, 1, "");
                }
                objform.Items.Item("fld_Unit1").Click(BoCellClickType.ct_Regular);
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

        #region DeleteRow
        public void DeleteRow(Matrix oMatrix, DBDataSource oDBDSDetail)
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
                objform.Freeze(false);
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
                if (oMatrix1.VisualRowCount <= 1)
                {
                    objSBOAPI.SBO_Appln.SetStatusBarMessage("Empty Document should not be Added...", BoMessageTime.bmt_Short, true);
                    flag = false;
                    return flag;
                }
                else
                    flag = true;
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
