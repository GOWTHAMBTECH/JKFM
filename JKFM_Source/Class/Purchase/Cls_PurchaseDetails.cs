using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JKFM_Source
{
    class Cls_PurchaseDetails
    {
        #region Declaration
        SAPbouiCOM.Form objform, oForm;
        ClsSBO objSBOAPI;

        SAPbouiCOM.DBDataSource oDBDSHeader;
        SAPbouiCOM.DBDataSource oDBDSDetail;
        SAPbouiCOM.DBDataSource oDBDSSubPlanDetails;

        SAPbouiCOM.DBDataSource oDBDSMainPlanDetails;
        SAPbouiCOM.Matrix oMatrix;
        SAPbouiCOM.Matrix oMatSubPlanDetails;
        SAPbouiCOM.Matrix oMatMainplanDetails;
        bool sFormUID;
        string UDOID;
        int SubRowID_PlanningDetails;
        string SubRowID_DayID;
        string sQuery;
        string lretcode;

        #endregion        

        #region Constructor
        public Cls_PurchaseDetails(ClsSBO objSBO)
        {
            objSBOAPI = objSBO;
        }
        #endregion

        #region Item Event
        public void itemevent(string formuid, ref SAPbouiCOM.ItemEvent pval, ref bool bubbleevent)
        {
            try
            {
                objform = objSBOAPI.SBO_Appln.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount);
                if (pval.BeforeAction == true)
                {
                    
                    try
                    {
                        switch (pval.EventType)
                        {
                            case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                                SAPbouiCOM.ChooseFromListEvent chooseFromListEvent = (SAPbouiCOM.ChooseFromListEvent)pval;
                                SAPbouiCOM.DataTable selectedObjects = chooseFromListEvent.SelectedObjects;
                                if (selectedObjects == null || pval.ItemUID != "subMatrix" || pval.ColUID != "WhsCode")
                                {
                                    break;
                                }
                                break;

                            case SAPbouiCOM.BoEventTypes.et_CLICK:
                                switch(pval.ItemUID)
                                {
                                    case "btn_Ok":
                                        if (!Validation_SubGrid())
                                        {
                                            bubbleevent = false;
                                        }
                                        break;
                                }
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        bubbleevent = false;
                    }
                    
                }
                else
                {
                    switch (pval.EventType)
                    {


                        //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        //    switch (pval.FormTypeEx)
                        //    {
                        //        case "PRE2":
                        //            oDBDSSubPlanDetails = objform.DataSources.DBDataSources.Item("@AIS_PRE2");
                        //            oMatSubPlanDetails = (SAPbouiCOM.Matrix)objform.Items.Item("subMatrix").Specific;
                        //            break;
                        //    }
                        //    break;



                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                            break;

                        

                        case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                            switch(pval.ItemUID)
                            {
                                case "subMatrix":
                                    switch (pval.ColUID)
                                    {
                                        case "a1":
                                            objform.Freeze(true);
                                            oDBDSSubPlanDetails = (SAPbouiCOM.DBDataSource)objform.DataSources.DBDataSources.Item("@AIS_PRE2");
                                            oMatSubPlanDetails = (SAPbouiCOM.Matrix)objform.Items.Item("subMatrix").Specific;
                                            oForm = objSBOAPI.SBO_Appln.Forms.Item(objform.DataSources.UserDataSources.Item("UD_0").Value);
                                            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Matrix").Specific;
                                            SubRowID_PlanningDetails = Convert.ToInt32(objform.DataSources.UserDataSources.Item("UD_1").Value);
                                            SubRowID_DayID = objform.DataSources.UserDataSources.Item("UD_2").Value;
                                            objSBOAPI.SetNewLineSubGridForSubgrid(SubRowID_PlanningDetails, oMatSubPlanDetails, oDBDSSubPlanDetails, pval.Row, SubRowID_DayID, pval.ColUID, (string[,])null);
                                            oMatSubPlanDetails.FlushToDataSource();
                                            string singleValue1, singleValue2, singleValue3 = null;
                                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                            {
                                                singleValue1 = objSBOAPI.Query_Execute(" Select IFNULL(\"U_UNIT\",'01') from \"OITM\" Where \"ItemCode\" ='" + oMatrix.Columns.Item("ICode").Cells.Item(SubRowID_PlanningDetails).Specific.Value.ToString().Trim() + "' ");
                                            }
                                            else
                                            {
                                                singleValue1 = objSBOAPI.Query_Execute(" Select isnull(U_UNIT,'01') from OITM Where ItemCode ='" + oMatrix.Columns.Item("ICode").Cells.Item(SubRowID_PlanningDetails).Specific.Value.ToString().Trim() + "' ");
                                            }
                                                
                                            oDBDSSubPlanDetails.SetValue("U_DefUnit", pval.Row, singleValue1);
                                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                            {
                                                singleValue2 = objSBOAPI.Query_Execute(" Select \"U_WhsCode\"  from \"@AIS_BRN1\" Where \"U_BranchCode\" ='" + singleValue1 + "' ");
                                            }
                                            else
                                            {
                                                singleValue2 = objSBOAPI.Query_Execute(" Select U_WhsCode  from [@AIS_BRN1] Where U_BranchCode ='" + singleValue1 + "' ");
                                            }
                                                
                                            oDBDSSubPlanDetails.SetValue("U_WhsCode", pval.Row, singleValue2);
                                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                            {
                                                singleValue3 = objSBOAPI.Query_Execute(" Select \"WhsName\"  from \"OWHS\" Where \"WhsCode\" ='" + singleValue2 + "' ");
                                            }
                                            else
                                            {
                                                singleValue3 = objSBOAPI.Query_Execute(" Select WhsName  from OWHS Where WhsCode ='" + singleValue2 + "' ");
                                            }
                                                
                                            oDBDSSubPlanDetails.SetValue("U_WhsName", pval.Row, singleValue3);
                                            oMatSubPlanDetails.LoadFromDataSource();
                                            objform.Freeze(false);
                                            break;
                                    }
                                    break;
                            }
                            break;

                        case SAPbouiCOM.BoEventTypes.et_CLICK:
                            switch(pval.ItemUID)
                            {
                                case "btn_Ok":
                                    objform.Freeze(true);
                                    oForm = objSBOAPI.SBO_Appln.Forms.Item(objform.DataSources.UserDataSources.Item("UD_0").Value);
                                    oDBDSSubPlanDetails = (SAPbouiCOM.DBDataSource)objform.DataSources.DBDataSources.Item("@AIS_PRE2");
                                    oMatSubPlanDetails = (SAPbouiCOM.Matrix)objform.Items.Item("subMatrix").Specific;
                                    oDBDSHeader = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@AIS_OPRE");
                                    oDBDSDetail = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@AIS_PRE1");
                                    oDBDSMainPlanDetails = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@AIS_PRE2");//objform.DataSources.DBDataSources.Item("@AIS_PRE2");
                                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("Matrix").Specific;
                                    oMatMainplanDetails = (SAPbouiCOM.Matrix)oForm.Items.Item("subMatrix").Specific;//(SAPbouiCOM.Matrix)objform.Items.Item("subMatrix").Specific;
                                    //Thiru - 18.03.2021
                                    SubRowID_PlanningDetails = Convert.ToInt32(objform.DataSources.UserDataSources.Item("UD_1").Value);
                                    //Thiru - 18.03.2021
                                    if (SaveSubGrid_WithoutEmptyLine(oMatMainplanDetails, oDBDSMainPlanDetails, oMatSubPlanDetails, oDBDSSubPlanDetails, SubRowID_PlanningDetails))
                                    {
                                        double num1 = 0.0;
                                        double num3 = 0.0;
                                        int rowCount = oMatSubPlanDetails.RowCount;
                                        int num2 = 1;
                                        while (num2 <= rowCount)
                                        {
                                            string Left = Convert.ToString(oMatSubPlanDetails.Columns.Item("CompStatus").Cells.Item(num2).Specific.Value);
                                            if (!(oMatSubPlanDetails.Columns.Item("MobStatus").Cells.Item(num2).Specific.Value.ToString() == "C" || Left == "C"))
                                            {
                                                num1 = num1 + Convert.ToDouble(oMatSubPlanDetails.Columns.Item("a2").Cells.Item(num2).Specific.Value);
                                            }
                                            ++num2;
                                        }
                                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                        {
                                            num3 = Convert.ToDouble(objSBOAPI.Query_Execute("Select IFNULL(\"SalPackUn\",'0') from \"OITM\" Where  \"ItemCode\" = '" + oMatrix.Columns.Item("ICode").Cells.Item(SubRowID_PlanningDetails).Specific.Value.ToString() + "'  "));
                                        }
                                        else
                                        {
                                            num3 = Convert.ToDouble(objSBOAPI.Query_Execute("Select ISNULL(SalPackUn,'0') from OITM Where  ItemCode = '" + oMatrix.Columns.Item("ICode").Cells.Item(SubRowID_PlanningDetails).Specific.Value.ToString() + "'  "));
                                        }
                                            
                                        oMatrix.FlushToDataSource();
                                        oDBDSDetail.SetValue("U_ReqQty", SubRowID_PlanningDetails - 1, (Convert.ToDouble(num1 * num3) / 1000.0).ToString());
                                        oDBDSDetail.SetValue("U_ReqDate", SubRowID_PlanningDetails - 1, num1.ToString());
                                        double num4 = num1;
                                        double num5 = Convert.ToDouble(oMatrix.Columns.Item("UPrice").Cells.Item(SubRowID_PlanningDetails).Specific.Value.ToString());
                                        double num6 = Convert.ToDouble(oMatrix.Columns.Item("DisCount").Cells.Item(SubRowID_PlanningDetails).Specific.Value.ToString());
                                        double num7 = num5 * num6 / 100.0;
                                        oDBDSDetail.SetValue("U_LineTotal", SubRowID_PlanningDetails - 1, (num4 * num5 - num7).ToString());
                                        oMatrix.LoadFromDataSource();
                                        int visualRowCount1 = oMatMainplanDetails.VisualRowCount;
                                        int num8 = 1;
                                        while (num8 <= visualRowCount1)
                                        {
                                            string Left1 = oDBDSMainPlanDetails.GetValue("U_DefUnit", num8 - 1).ToString().Trim();
                                            Convert.ToDouble(oDBDSMainPlanDetails.GetValue("U_Qty", num8 - 1).ToString().Trim());
                                            string Left2 = oDBDSMainPlanDetails.GetValue("U_PlanDate", num8 - 1).ToString().Trim();
                                            if (!(oDBDSMainPlanDetails.GetValue("U_CompStatus", num8 - 1).ToString().Trim() == "C" || oDBDSMainPlanDetails.GetValue("U_MobStatus", num8 - 1).ToString().Trim() == "C") && Left1 == "03")
                                            {
                                                double num9 = 0.0;
                                                double num10 = 0.0;
                                                double num11 = 0.0;
                                                string str1 = string.Empty;
                                                string str2 = string.Empty;
                                                string empty = string.Empty;
                                                int visualRowCount2 = oMatMainplanDetails.VisualRowCount;
                                                int num12 = 1;
                                                while (num12 <= visualRowCount2)
                                                {
                                                    string Right = oDBDSMainPlanDetails.GetValue("U_PlanDate", num12 - 1).ToString().Trim();
                                                    string Left3 = oDBDSMainPlanDetails.GetValue("U_DefUnit", num12 - 1).ToString().Trim();
                                                    double num13 = Convert.ToDouble(oDBDSMainPlanDetails.GetValue("U_Qty", num12 - 1).ToString().Trim());
                                                    if (!(oDBDSMainPlanDetails.GetValue("U_CompStatus", num12 - 1).ToString().Trim() == "C" || oDBDSMainPlanDetails.GetValue("U_MobStatus", num12 - 1).ToString().Trim() == "C") && Left2 == Right)
                                                    {
                                                        if (Left3 == "01")
                                                        {
                                                            num9 += num13;
                                                            if (oDBDSMainPlanDetails.GetValue("U_WhsCode", num12 - 1).ToString().Trim() != string.Empty)
                                                            {
                                                                str1 = oDBDSMainPlanDetails.GetValue("U_WhsCode", num12 - 1).ToString().Trim();
                                                            }
                                                        }
                                                        else if (Left3 == "02")
                                                        {
                                                            num10 += num13;
                                                            if (oDBDSMainPlanDetails.GetValue("U_WhsCode", num12 - 1).ToString().Trim() != string.Empty)
                                                            {
                                                                str2 = oDBDSMainPlanDetails.GetValue("U_WhsCode", num12 - 1).ToString().Trim();
                                                            }
                                                        }
                                                        else if (Left3 == "03")
                                                            num11 += num13;
                                                    }
                                                    ++num12;
                                                }
                                            }
                                            ++num8;
                                        }
                                        GetDocTotal();
                                        objform.Close();
                                        
                                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                        {
                                            //Thiru Changes
                                            //oForm.Items.Item("Matrix").Visible = true;
                                            //oForm.Items.Item("subMatrix").Visible = false;
                                            //Thiru Changes
                                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                        }
                                        objform.Freeze(false);
                                    }
                                    break;
                            }
                            break;
                            
                        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                            SAPbouiCOM.ChooseFromListEvent chooseFromListEvent = (SAPbouiCOM.ChooseFromListEvent)pval;
                            SAPbouiCOM.DataTable selectedObjects = chooseFromListEvent.SelectedObjects;
                            switch (pval.ItemUID)
                            {
                                case "subMatrix":
                                    switch (pval.ColUID)
                                    {
                                        case "WhsCode":
                                            oMatSubPlanDetails.FlushToDataSource();
                                            oDBDSSubPlanDetails.SetValue("U_WhsCode", pval.Row - 1, Convert.ToString(selectedObjects.GetValue("WhsCode", 0)));
                                            oDBDSSubPlanDetails.SetValue("U_WhsName", pval.Row - 1, Convert.ToString(selectedObjects.GetValue("WhsName", 0)));
                                            oMatSubPlanDetails.LoadFromDataSource();
                                            oMatSubPlanDetails.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
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
                    if (pval.MenuUID == "1281")
                    {

                    }
                    else if (pval.MenuUID == "1282")
                    {

                    }
                    else if ((pval.MenuUID == "1288") || (pval.MenuUID == "1289") || (pval.MenuUID == "1290") || (pval.MenuUID == "1291"))
                    {

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

        #region Validation_SubGrid
        public bool Validation_SubGrid()
        {
            bool flag = false;
            try
            {
                flag = true;
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            return flag;
        }
        #endregion

        #region GetDocTotal
        public void GetDocTotal()
        {
            double num1 = 0.0;
            double num2 = 0.0;
            double num3 = 0.0;
            int num4 = 1;
            while (num4 <= oMatrix.VisualRowCount)
            {
                double num5 = Convert.ToDouble(oMatrix.Columns.Item("LineTotal").Cells.Item(num4).Specific.Value.ToString());

                double num6 = Convert.ToDouble(oMatrix.Columns.Item("ReqDate").Cells.Item(num4).Specific.Value.ToString());

                double num7 = Convert.ToDouble(oMatrix.Columns.Item("RQty").Cells.Item(num4).Specific.Value.ToString());
                num1 += num5;
                num2 += num6;
                num3 += num7;
                ++num4; 
            }

            oDBDSHeader.SetValue("U_DocTotal", 0, Convert.ToString(num1));
            oDBDSHeader.SetValue("U_BDiscount", 0, Convert.ToString(num1));
            oDBDSHeader.SetValue("U_TotalBag", 0, Convert.ToString(num2));
            oDBDSHeader.SetValue("U_TotalTon", 0, Convert.ToString(num3));
        }
        #endregion

        #region SaveSubGrid_WithoutEmptyLine
        public bool SaveSubGrid_WithoutEmptyLine(
      SAPbouiCOM.Matrix oMatMainGrid,
      SAPbouiCOM.DBDataSource oDBDSMainSubGrid,
      SAPbouiCOM.Matrix oMatSubGrid,
      SAPbouiCOM.DBDataSource oDBDsSubGrid,
      int RowID)
        {
            bool flag1;
            try
            {
                SAPbobsCOM.Recordset recordset = null;
                recordset = (SAPbobsCOM.Recordset)objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string strQry = null;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    strQry = "Select \"AliasID\" From \"CUFD\" Where  \"AliasID\" <> 'UniqID' AND \"TableID\" ='" + oDBDsSubGrid.TableName + "' ";
                }
                else
                {
                    strQry = "Select AliasID From CUFD Where  AliasID <> 'UniqID' AND TableID ='" + oDBDsSubGrid.TableName + "' ";
                }
                    
                recordset.DoQuery(strQry);
                int size1 = oDBDSMainSubGrid.Size;
                int size2 = oDBDSMainSubGrid.Size;
                int num1 = size1 - 1;
                int num2 = 0;
                while (num2 <= num1)// && num1!=0)
                {
                    oDBDSMainSubGrid.Offset = num2 - (size1 - size2);
                    oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset).Trim();
                    if (oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset).Trim() != "")
                    {
                        if (Convert.ToInt32(oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset)) == RowID)
                            oDBDSMainSubGrid.RemoveRecord(oDBDSMainSubGrid.Offset);
                    }
                    else if (oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset).Trim().Equals(""))
                        oDBDSMainSubGrid.RemoveRecord(oDBDSMainSubGrid.Offset);
                    size2 = oDBDSMainSubGrid.Size;
                    ++num2; 
                }
                if (oDBDSMainSubGrid.Size == 0)
                    oDBDSMainSubGrid.InsertRecord(oDBDSMainSubGrid.Size);
                oDBDSMainSubGrid.Offset = 0;
                bool flag2 = oDBDSMainSubGrid.Size == 1 && oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset).Equals("");
                oMatSubGrid.FlushToDataSource();
                oMatSubGrid.LoadFromDataSource();
                int num3 = oMatSubGrid.VisualRowCount - 1;
                int RecordNumber = 0;
                while (RecordNumber < num3)//Anbu
                {
                    oDBDsSubGrid.Offset = RecordNumber;
                    if (flag2)
                    {
                        oDBDSMainSubGrid.Offset = oDBDSMainSubGrid.Size - 1;
                        oDBDSMainSubGrid.SetValue("U_UniqID", oDBDSMainSubGrid.Offset, Convert.ToString(RowID));
                        recordset.MoveFirst();
                        int num4 = recordset.RecordCount - 1;
                        int num5 = 0;
                        while (num5 <= num4)
                        {
                            string str = Convert.ToString(string.Concat("U_", recordset.Fields.Item(0).Value));
                            oDBDSMainSubGrid.SetValue(str, oDBDSMainSubGrid.Offset, oDBDsSubGrid.GetValue(str, RecordNumber).Trim());
                            recordset.MoveNext();
                            ++num5; 
                        }
                        if (RecordNumber+1 < num3)//Anbu
                            oDBDSMainSubGrid.InsertRecord(oDBDSMainSubGrid.Size);
                    }
                    else if (!flag2)
                    {
                        oDBDSMainSubGrid.InsertRecord(oDBDSMainSubGrid.Size);
                        oDBDSMainSubGrid.Offset = oDBDSMainSubGrid.Size - 1;
                        oDBDSMainSubGrid.SetValue("U_UniqID", oDBDSMainSubGrid.Offset, Convert.ToString(RowID));
                        recordset.MoveFirst();
                        int num4 = recordset.RecordCount - 1;
                        int num5 = 0;
                        while (num5 <= num4)
                        {
                            string str = Convert.ToString(string.Concat("U_", recordset.Fields.Item(0).Value));
                            if (oDBDsSubGrid.GetValue("U_PlanDate", RecordNumber) != string.Empty)
                                oDBDSMainSubGrid.SetValue(str, oDBDSMainSubGrid.Offset, oDBDsSubGrid.GetValue(str, RecordNumber).Trim());
                            recordset.MoveNext();
                            ++num5; 
                        }
                    }
                    ++RecordNumber; 
                }
                oMatMainGrid.LoadFromDataSource();
                flag1 = true;
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                flag1 = false;
            }
            return flag1;
        }
        #endregion
    }
}
