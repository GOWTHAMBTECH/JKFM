using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Media;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JKFM_Source
{
    class Cls_PreSalesOrder
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

        #endregion        

        #region Constructor
        public Cls_PreSalesOrder(ClsSBO objSBO)
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
                    try
                    {
                        switch (pval.EventType)
                        {
                            case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                                SAPbouiCOM.ChooseFromListCollection oCFLs = null;
                                SAPbouiCOM.Conditions oCons = null;
                                SAPbouiCOM.Condition oCon = null;
                                oCFLs = objform.ChooseFromLists;
                                SAPbouiCOM.ChooseFromList oCFL = null;
                                switch(pval.ItemUID)
                                {
                                    case "t_ccode":
                                        oCFL = oCFLs.Item("CFL_OCRD");
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
                                        oCon.Alias = "CardType";
                                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                        oCon.CondVal = "C";
                                        oCon.BracketCloseNum = 1;
                                        oCFL.SetConditions(oCons);
                                        GC.Collect();
                                        break;
                                    case "Matrix":
                                        switch (pval.ColUID)
                                        {
                                            case "ICode":
                                                string CardCode = objform.Items.Item("t_ccode").Specific.Value;
                                                if(CardCode != "")
                                                {
                                                    SAPbobsCOM.Recordset oRecordSet;
                                                    oRecordSet = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                    string StrSql;
                                                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                                    {
                                                        StrSql = objSBOAPI.Query_Execute("SELECT \"QString\" FROM \"OUQR\" WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM \"OQCN\" WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='ItemCode_CFL_Loading_HANA'");
                                                    }
                                                    else
                                                    {
                                                        StrSql = objSBOAPI.Query_Execute("SELECT QString FROM OUQR WHERE QCategory=(SELECT CategoryId FROM OQCN WHERE CatName='AddonQuery_SQL') AND QName='ItemCode_CFL_Loading_SQL'");
                                                    }
                                                        
                                                    StrSql = StrSql.Replace("[%1]", CardCode);
                                                    oRecordSet.DoQuery(StrSql);
                                                    oCFL = oCFLs.Item("CFL_Item");
                                                    oCFL.SetConditions(null);
                                                    oCons = oCFL.GetConditions();

                                                    if (oRecordSet.RecordCount > 0)
                                                    {
                                                        int Count = oRecordSet.RecordCount;
                                                        for (int i = 1; i <= oRecordSet.RecordCount; i++)
                                                        {
                                                            if (i != 1)
                                                            {
                                                                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                                                            }
                                                            oCon = oCons.Add();
                                                            oCon.BracketOpenNum = 1;
                                                            oCon.Alias = "ItemCode";
                                                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                                            oCon.CondVal = oRecordSet.Fields.Item("ItemCode").Value;
                                                            oCon.BracketCloseNum = 1;
                                                            oCFL.SetConditions(oCons);
                                                            GC.Collect();
                                                            oRecordSet.MoveNext();
                                                        }
                                                    }
                                                    else
                                                    {
                                                        // Thiru Addon Changes - Start - 19.04.2021
                                                        if (oCons.Count == 0)
                                                        {
                                                            oCon = oCons.Add();
                                                        }
                                                        else
                                                        {
                                                            oCon = oCons.Item(0);
                                                        }
                                                        oCon.BracketOpenNum = 1;
                                                        oCon.Alias = "ItemCode";
                                                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                                        oCon.CondVal = oRecordSet.Fields.Item("ItemCode").Value;
                                                        oCon.BracketCloseNum = 1;
                                                        oCFL.SetConditions(oCons);
                                                        GC.Collect();

                                                        //    if (oCons.Count == 0)
                                                        //    {
                                                        //        oCon = oCons.Add();
                                                        //    }
                                                        //    else
                                                        //    {
                                                        //        oCon = oCons.Item(0);
                                                        //    }
                                                        //    oCon.BracketOpenNum = 1;
                                                        //    oCon.Alias = "SellItem";
                                                        //    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                                        //    oCon.CondVal = "Y";
                                                        //    oCon.BracketCloseNum = 1;
                                                        //    oCFL.SetConditions(oCons);
                                                        //    GC.Collect();
                                                        // Thiru Addon Changes - End - 19.04.2021
                                                    }
                                                }
                                                else
                                                {
                                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Card Code shouldn't be Empty...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                                    objform.Items.Item("t_ccode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                                    bubbleevent = false;
                                                }
                                                break;
                                        }
                                        break;
                                }
                                break;

                            case SAPbouiCOM.BoEventTypes.et_CLICK:
                                switch(pval.ItemUID)
                                {
                                    case "1":
                                        if(pval.FormMode == Convert.ToInt32(SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                                        {
                                            if (Validation() == false)
                                            {
                                                bubbleevent = false;
                                                return;
                                            }
                                        }
                                        break;
                                }
                                break;
                        }
                    }
                    catch(Exception ex)
                    {
                        bubbleevent = false;
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
                                    if (objform.Mode == BoFormMode.fm_ADD_MODE && pval.ActionSuccess == true)
                                    {
                                        InitForm();
                                    }
                                    break;
                                //case "Btn_Temp":
                                //    if (objform.Items.Item("Matrix").Visible == true)
                                //    {
                                //        objform.Items.Item("Matrix").Visible = false;
                                //        objform.Items.Item("subMatrix").Visible = true;
                                //    }
                                //    else
                                //    {
                                //        objform.Items.Item("Matrix").Visible = true;
                                //        objform.Items.Item("subMatrix").Visible = false;
                                //    }
                                //    break;
                            }
                            break;

                        //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
                        //    SAPbouiCOM.ChooseFromListCollection oCFLs;
                        //    SAPbouiCOM.Conditions oCons;
                        //    SAPbouiCOM.Condition oCon;
                        //    oCFLs = objform.ChooseFromLists;
                        //    SAPbouiCOM.ChooseFromList oCFL = null;
                        //    if (pval.ItemUID == "t_ccode")
                        //    {
                        //        oCFL = oCFLs.Item("CFL_OCRD");
                        //        oCFL.SetConditions(null);
                        //        oCons = oCFL.GetConditions();
                        //        if (oCons.Count == 0)
                        //        {
                        //            oCon = oCons.Add();
                        //        }
                        //        else
                        //        {
                        //            oCon = oCons.Item(0);
                        //        }
                        //        oCon.BracketOpenNum = 1;
                        //        oCon.Alias = "CardType";
                        //        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        //        oCon.CondVal = "C";
                        //        oCon.BracketCloseNum = 1;
                        //        oCFL.SetConditions(oCons);
                        //        GC.Collect();
                        //    }
                        //    break;

                        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                            SAPbouiCOM.DataTable dt;
                            SAPbouiCOM.ChooseFromListEvent cfl;
                            cfl = (SAPbouiCOM.ChooseFromListEvent)pval;
                            dt = cfl.SelectedObjects;
                            if (dt != null)
                            {
                                switch (pval.ItemUID)
                                {
                                    case "t_ccode":
                                        
                                        oDBDSHeader.SetValue("U_CardCode", 0, dt.GetValue("CardCode", 0));
                                        
                                        oDBDSHeader.SetValue("U_CardName", 0, dt.GetValue("CardName", 0));
                                        
                                        //if (oMatrix.VisualRowCount == 0)
                                        //{
                                        //    oMatrix.AddRow();
                                        //    oMatrix.ClearRowData(oMatrix.VisualRowCount);
                                        //    oMatrix.Columns.Item("LineId").Cells.Item(oMatrix.VisualRowCount).Specific.Value = oMatrix.VisualRowCount;
                                        //}
                                        //oMatrix.Columns.Item("ICode").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                        break;
                                    case "Matrix":
                                        switch(pval.ColUID)
                                        {
                                            case "ICode":
                                                string str = oDBDSHeader.GetValue("U_CardCode", 0);
                                                oMatrix.FlushToDataSource();

                                                oDBDSDetail.SetValue("U_ItemCode", pval.Row - 1, dt.GetValue("ItemCode", 0));
                                                oDBDSDetail.SetValue("U_ItemName", pval.Row - 1, dt.GetValue("ItemName", 0));
                                                oDBDSDetail.SetValue("U_UomCode", pval.Row - 1, dt.GetValue("SalUnitMsr", 0));
                                                string singleValue1 = null;
                                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                                {
                                                    singleValue1 = objSBOAPI.Query_Execute("CALL \"@AIS_SalesPlanning_GetUnitPrice\"('" + dt.GetValue("ItemCode", 0) + "', '" + str + "', '" + dt.GetValue("SalUnitMsr", 0) + "') ");
                                                }
                                                else
                                                {
                                                    singleValue1 = objSBOAPI.Query_Execute("Exec dbo.[@AIS_SalesPlanning_GetUnitPrice]'" + dt.GetValue("ItemCode", 0) + "','" + str + "','" + dt.GetValue("SalUnitMsr", 0) + "' ");
                                                }
                                                
                                                oDBDSDetail.SetValue("U_UnitPrice", pval.Row - 1, singleValue1);
                                                string singleValue2 = null;
                                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                                {
                                                    singleValue2 = objSBOAPI.Query_Execute("CALL \"@AIS_SalesPlanning_GetTaxCodeDetermination\"('" + dt.GetValue("ItemCode", 0) + "', '" + str + "') ");
                                                }
                                                else
                                                {
                                                    singleValue2 = objSBOAPI.Query_Execute("EXEC dbo.[@AIS_SalesPlanning_GetTaxCodeDetermination] '" + dt.GetValue("ItemCode", 0) + "','" + str + "' ");
                                                }
                                                
                                                oDBDSDetail.SetValue("U_TaxCode", pval.Row - 1, singleValue2);
                                                oMatrix.LoadFromDataSource();

                                                oMatrix.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                                                objSBOAPI.AddNewLine(oMatrix, oDBDSDetail, pval.Row, pval.ColUID);
                                                break;

                                            case "WhsCode":
                                                oMatrix.FlushToDataSource();
                                                oDBDSDetail.SetValue("U_WhsCode", pval.Row - 1, dt.GetValue("WhsCode", 0));
                                                oDBDSDetail.SetValue("U_WhsName", pval.Row - 1, dt.GetValue("WhsName", 0));
                                                oMatrix.LoadFromDataSource();
                                                oMatrix.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                                                break;

                                            case "TaxCode":
                                                oMatrix.FlushToDataSource();
                                                oDBDSDetail.SetValue("U_TaxCode", pval.Row - 1, dt.GetValue("Code", 0));
                                                oMatrix.LoadFromDataSource();
                                                oMatrix.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                                                break;
                                        }
                                        break;
                                }
                            }
                            break;

                        case BoEventTypes.et_VALIDATE:
                            double num1 = 0.0, num2 = 0.0, num3 = 0.0, num4 = 0.0;
                            switch (pval.ItemUID)
                            {
                                
                                case "t_Discount":
                                    objform.Freeze(true);
                                    num1 = Convert.ToDouble(objform.Items.Item("t_Discount").Specific.Value);
                                    num2 = Convert.ToDouble(objform.Items.Item("t_DocTotal").Specific.Value);
                                    num3 = num2 * num1 / 100.0;
                                    oDBDSHeader.SetValue("U_DocTotal", 0, Convert.ToString(num2 - num3));
                                    objform.Freeze(false);
                                    break;

                                case "Matrix":

                                    if ((pval.ColUID == "RQty" || pval.ColUID == "UPrice" || pval.ColUID == "DisCount") && pval.ItemChanged == true)
                                    {
                                        objform.Freeze(true);
                                        num1 = Convert.ToDouble(oMatrix.Columns.Item("ReqDate").Cells.Item(pval.Row).Specific.Value.ToString());
                                        num2 = Convert.ToDouble(oMatrix.Columns.Item("UPrice").Cells.Item(pval.Row).Specific.Value.ToString());
                                        num3 = Convert.ToDouble(oMatrix.Columns.Item("DisCount").Cells.Item(pval.Row).Specific.Value.ToString());
                                        oMatrix.FlushToDataSource();
                                        num4 = num2 * num3 / 100.0;
                                        oDBDSDetail.SetValue("U_LineTotal", pval.Row - 1, Convert.ToString(num1 * num2 - num4));
                                        oMatrix.LoadFromDataSource();
                                        oMatrix.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                        GetDocTotal();
                                        objform.Freeze(false);
                                    }
                                    break;
                            }
                            break;

                        case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                            switch (pval.ItemUID)
                            {
                                case "c_Series":
                                    if (objform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                    {
                                        SAPbouiCOM.ComboBox comboBox = (SAPbouiCOM.ComboBox)objform.Items.Item("c_Series").Specific;
                                        oDBDSHeader.SetValue("DocNum", 0, Convert.ToString((long)objform.BusinessObject.GetNextSerialNumber(comboBox.Selected.Value, UDOID)));
                                    }
                                    break;
                                case "Matrix":
                                    switch (pval.ColUID)
                                    {
                                        case "DefUnit":
                                            SAPbouiCOM.ComboBox comboBox = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("DefUnit").Cells.Item(pval.Row).Specific;
                                            string str = string.Empty;
                                            string singleValue1 = null;
                                            string singleValue2 = null;
                                            if (comboBox.Selected != null)
                                            {
                                                str = comboBox.Selected.Value.ToString().Trim();
                                            }
                                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                            {
                                                singleValue1 = objSBOAPI.Query_Execute("Select \"U_WhsCode\" from \"@AIS_BRN1\" Where \"U_BranchCode\" ='" + str + "' and \"U_Active\" ='Y'");
                                            }
                                            else
                                            {
                                                singleValue1 = objSBOAPI.Query_Execute("Select U_WhsCode from \"@AIS_BRN1\" Where U_BranchCode ='" + str + "' and U_Active ='Y'");
                                            }
                                                
                                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                            {
                                                singleValue2 = objSBOAPI.Query_Execute("Select \"WhsName\" from \"OWHS\" Where \"WhsCode\" ='" + singleValue1 + "'");
                                            }
                                            else
                                            {
                                                singleValue2 = objSBOAPI.Query_Execute("Select WhsName from OWHS Where WhsCode ='" + singleValue1 + "'");
                                            }
                                                
                                            oMatrix.FlushToDataSource();
                                            oDBDSDetail.SetValue("U_WhsCode", pval.Row - 1, singleValue1);
                                            oDBDSDetail.SetValue("U_WhsName", pval.Row - 1, singleValue2);
                                            oMatrix.LoadFromDataSource();
                                            break;
                                    }
                                    break;
                            }
                            break;

                        case BoEventTypes.et_DOUBLE_CLICK:
                            switch (pval.ItemUID)
                            {
                                case "Matrix":
                                    switch (pval.ColUID)
                                    {
                                        case "ReqDate":

                                            if(pval.Row != 0)
                                            {
                                                //Thiru Changes
                                                //objform.Items.Item("Btn_Temp").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                                //Thiru Changes
                                                SubRowID_PlanningDetails = pval.Row;
                                                SubRowID_DayID = pval.ColUID;
                                                LoadSubMatrixDetails(SubRowID_PlanningDetails, pval.ColUID);
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
                    switch(pval.MenuUID)
                    {
                        case "OPRE":
                            objform = objSBOAPI.LoadForm("PreSalesOrder.xml", "OPRE");
                            oDBDSHeader = objform.DataSources.DBDataSources.Item("@AIS_OPRE");
                            oDBDSDetail = objform.DataSources.DBDataSources.Item("@AIS_PRE1");
                            oDBDSMainPlanDetails = objform.DataSources.DBDataSources.Item("@AIS_PRE2");
                            oMatrix = objform.Items.Item("Matrix").Specific;
                            oMatMainplanDetails = objform.Items.Item("subMatrix").Specific;
                            SAPbouiCOM.ComboBox oCombo, oCombo1;
                            oCombo = objform.Items.Item("c_Type").Specific;
                            oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                            oCombo1 = objform.Items.Item("c_TrkReq").Specific;
                            oCombo1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                            InitForm();
                            DefineModeForFields();
                            sFormUID = true;
                            objform.EnableMenu("1292",true);
                            objform.EnableMenu("1293", true);
                            break;

                        case "4870":
                            objform.Items.Item("Matrix").Enabled = false;
                            break;

                        case "1282":
                            InitForm();
                            break;

                        case "1292":
                            objSBOAPI.AddNewLine(oMatrix, oDBDSDetail, 1, "");
                            oMatrix.ClearRowData(oMatrix.VisualRowCount);
                            break;

                        case "1293":
                            oMatrix.FlushToDataSource();
                            int visualRowCount = oMatrix.VisualRowCount;
                            int RowNum = 1;
                            while (RowNum <= visualRowCount)
                            {
                                oMatrix.GetLineData(RowNum);
                                oDBDSDetail.Offset = RowNum - 1;
                                oDBDSDetail.SetValue("LineID", oDBDSDetail.Offset, Convert.ToString(RowNum));
                                oMatrix.SetLineData(RowNum);
                                oMatrix.FlushToDataSource();
                                ++RowNum;
                            }
                            oDBDSDetail.RemoveRecord(oDBDSDetail.Size - 1);
                            oMatrix.LoadFromDataSource();
                            break;

                        case "":
                            objform.ActiveItem = "t_DocNum";
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
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            
                            break;
                    }
                          
                }
                if(BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            if (BusinessObjectInfo.ActionSuccess == true)                            {
                                string DocEntry = null;
                                //string MyXml = BusinessObjectInfo.ObjectKey;
                                //System.Xml.XmlDocument MyDoc = new System.Xml.XmlDocument();
                                //MyDoc.LoadXml(MyXml);
                                //DocEntry = MyDoc.SelectSingleNode("//DocumentParams/DocEntry").InnerText;
                                string CardCode = objform.DataSources.DBDataSources.Item("@AIS_OPRE").GetValue("U_CardCode", 0).ToString();
                                DocEntry = objform.DataSources.DBDataSources.Item("@AIS_OPRE").GetValue("DocEntry", 0).ToString();
                                if (DocEntry != "")
                                {
                                    SAPbobsCOM.Recordset oRecordSet;
                                    oRecordSet = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    string Update_Query;
                                    //if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                    //{
                                    //    Update_Query = objSBOAPI.Query_Execute("SELECT \"QString\" FROM \"OUQR\" WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM \"OQCN\" WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='PreSalesOrder_Update_Fields_HANA'");
                                    //}
                                    //else
                                    //{
                                    //    Update_Query = objSBOAPI.Query_Execute("SELECT QString FROM OUQR WHERE QCategory=(SELECT CategoryId FROM OQCN WHERE CatName='AddonQuery_SQL') AND QName='PreSalesOrder_Update_Fields_SQL'");
                                    //}
                                    //Update_Query = Update_Query.Replace("[%1]", CardCode);
                                    //Update_Query = Update_Query.Replace("[%3]", DocEntry);
                                    int Row = 1;
                                    while (Row <= oMatrix.VisualRowCount)
                                    {
                                        string ItemCode = oMatrix.Columns.Item("ICode").Cells.Item(Row).Specific.Value;
                                        string LineId = oMatrix.Columns.Item("LineId").Cells.Item(Row).Specific.Value;
                                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                        {
                                            Update_Query = objSBOAPI.Query_Execute("SELECT \"QString\" FROM \"OUQR\" WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM \"OQCN\" WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='PreSalesOrder_Update_Fields_HANA'");
                                        }
                                        else
                                        {
                                            Update_Query = objSBOAPI.Query_Execute("SELECT QString FROM OUQR WHERE QCategory=(SELECT CategoryId FROM OQCN WHERE CatName='AddonQuery_SQL') AND QName='PreSalesOrder_Update_Fields_SQL'");
                                        }
                                        Update_Query = Update_Query.Replace("[%1]", CardCode);
                                        Update_Query = Update_Query.Replace("[%2]", ItemCode);
                                        Update_Query = Update_Query.Replace("[%3]", DocEntry);
                                        Update_Query = Update_Query.Replace("[%4]", LineId);
                                        oRecordSet.DoQuery(Update_Query);
                                        //oRecordSet.MoveFirst();
                                        ++Row;
                                    }
                                    //SAPbouiCOM.Menus oMenus = objSBOAPI.SBO_Appln.Menus;
                                    //if (oMenus.Item("1304").Enabled == true)
                                    //{
                                        //objSBOAPI.SBO_Appln.ActivateMenuItem("1304");
                                    //}
                                }
                                
                                
                                //string CardCode = objform.Items.Item("t_ccode").Specific.Value;
                                //string CardName = null;
                                //if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                //{
                                //    CardName = objSBOAPI.Query_Execute("Select \"CardName\" from \"OCRD\" where \"CardCode\" = '" + CardCode + "'");
                                //}
                                //else
                                //{
                                //    CardName = objSBOAPI.Query_Execute("Select CardName from OCRD where CardCode = '" + CardCode + "'");
                                //}
                                //objform.Items.Item("t_cname").Specific.Value = CardName;
                                ////oDBDSHeader.SetValue("U_CardName", 0, CardName);                                //int Row = 1;
                                //while (Row <= oMatrix.VisualRowCount)
                                //{
                                //    string ItemCode = oMatrix.Columns.Item("ICode").Cells.Item(Row).Specific.Value;
                                //    string ItemName = null;
                                //    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                //    {
                                //        ItemName = objSBOAPI.Query_Execute("Select \"ItemName\" from \"OITM\" where \"ItemCode\" = '" + ItemCode + "'");
                                //    }
                                //    else
                                //    {
                                //        ItemName = objSBOAPI.Query_Execute("Select ItemName from OITM where ItemCode = '" + ItemCode + "'");
                                //    }
                                //    oMatrix.Columns.Item("IName").Cells.Item(Row).Specific.Value = ItemName;
                                //    //oDBDSDetail.SetValue("U_ItemName", Row, ItemName);

                                //    double Required_Date = 0.0;
                                //    string DocEntry = null;
                                //    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                //    {
                                //        DocEntry = objSBOAPI.Query_Execute("Select \"DocEntry\" from \"@AIS_OPRE\" WHERE \"DocNum\"='" + objform.Items.Item("t_DocNum").Specific.Value + "'");
                                //        Required_Date = Convert.ToDouble(objSBOAPI.Query_Execute("Select IFNULL(SUM(T2.\"U_Qty\"),'0') AS \"QUANTITY\" from \"@AIS_PRE1\" T1 INNER JOIN \"@AIS_PRE2\" T2 ON T1.\"DocEntry\" = T2.\"DocEntry\" AND T1.\"LineId\" = T2.\"U_UniqID\"   Where T1.\"U_ItemCode\" = '" + oMatrix.Columns.Item("ICode").Cells.Item(Row).Specific.Value + "' AND T1.\"DocEntry\" = '" + DocEntry + "' AND T1.\"LineId\" = '" + oMatrix.Columns.Item("LineId").Cells.Item(Row).Specific.Value + "' AND T2.\"U_CompStatus\" != 'C'"));
                                //    }
                                //    else
                                //    {
                                //        DocEntry = objSBOAPI.Query_Execute("Select DocEntry from [@AIS_OPRE] WHERE DocNum='" + objform.Items.Item("t_DocNum").Specific.Value + "'");
                                //        Required_Date = Convert.ToDouble(objSBOAPI.Query_Execute("select ISNULL(SUM(T2.U_Qty),0) AS 'QUANTITY' from[@AIS_PRE1] T1 INNER JOIN[@AIS_PRE2] T2 ON T1.DocEntry = T2.DocEntry AND T1.LineId = T2.U_UniqID WHERE T1.U_ItemCode = '" + oMatrix.Columns.Item("ICode").Cells.Item(Row).Specific.Value + "' AND T1.DocEntry = '" + DocEntry + "' AND T1.LineId = '" + oMatrix.Columns.Item("LineId").Cells.Item(Row).Specific.Value + "' AND T2.U_CompStatus != 'C'"));
                                //    }
                                //    oMatrix.Columns.Item("ReqDate").Cells.Item(Row).Specific.Value = Required_Date;
                                //    ++Row;
                                //}
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

        #region InitForm
        public void InitForm()
        {
            try
            {
                objform.Freeze(true);
                if (objform.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                }
                SAPbouiCOM.ComboBox oComboBox, oComboBox1, oComboBox2, oComboBox3;
                //GFun.LoadComboBoxSeries((ComboBox)objform.Items.Item((object)"c_Series")[], this.UDOID);
                oComboBox = objform.Items.Item("c_Series").Specific;
                oComboBox.ValidValues.LoadSeries("OPRE", BoSeriesMode.sf_Add);
                oComboBox.Select(0, BoSearchKey.psk_Index);
                oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                //GFun.LoadDocumentDate((EditText)objform.Items.Item((object)"t_DocDate")[]);
                SAPbouiCOM.EditText oEditText = objform.Items.Item("t_DocDate").Specific;
                oEditText.Active = true;
                oEditText.String = "A";

                objSBOAPI.AddNewLine(oMatrix, oDBDSDetail, 1, "");
                
                //GFun.setComboBoxValue((ComboBox)oMatrix[].Item("PlanUnit").Cells.Item(1)[], "SELECT  U_BranchCode , U_BranchName FROM [@AIS_BRN1]  Where U_Active='Y'");
                oComboBox = oMatrix.Columns.Item("PlanUnit").Cells.Item(1).Specific;
                if (oComboBox.ValidValues.Count == 0)
                {
                    SAPbobsCOM.Recordset recordset = null;
                    recordset = (SAPbobsCOM.Recordset)objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string strQry = null;
                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        strQry = "SELECT \"U_BranchCode\", \"U_BranchName\" FROM \"@AIS_BRN1\"  Where \"U_Active\" = 'Y'";
                    }
                    else
                    {
                        strQry = "SELECT U_BranchCode, U_BranchName FROM [@AIS_BRN1]  Where U_Active = 'Y'";
                    }
                        
                    recordset.DoQuery(strQry);
                    recordset.MoveFirst();
                    int num1 = recordset.RecordCount - 1;
                    int num2 = 0;
                    while (num2 <= num1)
                    {
                        oComboBox.ValidValues.Add(Convert.ToString(recordset.Fields.Item(0).Value), Convert.ToString(recordset.Fields.Item(1).Value));
                        recordset.MoveNext();
                        ++num2; 
                    }
                }
                oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                oComboBox1 = oMatrix.Columns.Item("Status").Cells.Item(1).Specific;
                oComboBox2 = oMatrix.Columns.Item("PStatus").Cells.Item(1).Specific;
                oComboBox3 = oMatrix.Columns.Item("MOBStatus").Cells.Item(1).Specific;
                oComboBox1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                oComboBox2.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                oComboBox3.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                //GFun.ChooseFromListFilteration(objform, "CFL_Item", "SellItem", "Select 'Y'");

                objform.ActiveItem = "t_ccode";
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
                objform.Items.Item("c_Series").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
                objform.Items.Item("c_Series").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);
                objform.Items.Item("t_DocNum").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);
                objform.Items.Item("t_DocDate").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Validation
        public bool Validation()
        {
            bool flag;
            try
            {
                if (objform.Items.Item("t_ccode").Specific.Value == "")
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Card Code shouldn't be Empty...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    objform.Items.Item("t_ccode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    flag = false;
                    return false;
                }
                else if (oMatrix.VisualRowCount == 0)
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Empty Document shouldn't be Empty...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oMatrix.Columns.Item("ICode").Cells.Item(oMatrix.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    flag = false;
                    return false;
                }
                else
                {
                    
                    if (oMatrix.VisualRowCount > 0 && oMatrix.Columns.Item("ICode").Cells.Item(1).Specific.Value == "")
                    {
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Empty Document shouldn't be Empty...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        oMatrix.Columns.Item("ICode").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        flag = false;
                        return false;
                    }
                    else
                        flag = true;
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
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

        #region LoadSubMatrixDetails
        public void LoadSubMatrixDetails(int Row, string ColumID)
        {
            int num1 = 0;
            try
            {
                
                oForm = objSBOAPI.LoadForm("PurchaseDetails.xml", "PRE2");
                oDBDSSubPlanDetails = oForm.DataSources.DBDataSources.Item("@AIS_PRE2");
                oMatSubPlanDetails = (SAPbouiCOM.Matrix)oForm.Items.Item("subMatrix").Specific;
                oForm.DataSources.UserDataSources.Item("UD_0").Value = objform.TypeEx;
                oForm.DataSources.UserDataSources.Item("UD_1").Value = Convert.ToString(Row);
                oForm.DataSources.UserDataSources.Item("UD_2").Value = ColumID;
                //Thiru - 18.03.2021
                SubRowID_PlanningDetails = Convert.ToInt32(oForm.DataSources.UserDataSources.Item("UD_1").Value);
                //Thiru - 18.03.2021
                //oForm.Items.Item("Et_Row").Specific.Value = Row;
                //oForm.Items.Item("Et_Col").Specific.Value = ColumID;
                int size = oDBDSMainPlanDetails.Size;
                oForm.Freeze(true);
                LoadSubGrid(oMatSubPlanDetails, oDBDSSubPlanDetails, oDBDSMainPlanDetails, Row, (string[,])null);
                SAPbouiCOM.ComboBox oComboBox, oComboBox1, oComboBox2;
                oComboBox = (SAPbouiCOM.ComboBox)oMatSubPlanDetails.Columns.Item("DefUnit").Cells.Item(1).Specific;
                if (oComboBox.ValidValues.Count == 0)
                {
                    SAPbobsCOM.Recordset recordset = null;
                    recordset = (SAPbobsCOM.Recordset)objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string strQry = null;
                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        strQry = "SELECT  \"Code\",\"Name\" FROM \"@AIS_OUNT\" ";
                    }
                    else
                    {
                        strQry = "SELECT  Code,Name FROM [@AIS_OUNT] ";
                    }
                        
                    recordset.DoQuery(strQry);
                    recordset.MoveFirst();
                    int num3 = recordset.RecordCount - 1;
                    int num4 = 0;
                    while (num4 <= num3)
                    {
                        oComboBox.ValidValues.Add(Convert.ToString(recordset.Fields.Item(0).Value), Convert.ToString(recordset.Fields.Item(1).Value));
                        recordset.MoveNext();
                        ++num4; 
                    }
                }
                oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                oComboBox1 = (SAPbouiCOM.ComboBox)oMatSubPlanDetails.Columns.Item("MobStatus").Cells.Item(1).Specific;
                oComboBox2 = (SAPbouiCOM.ComboBox)oMatSubPlanDetails.Columns.Item("CompStatus").Cells.Item(1).Specific;
                oComboBox1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                oComboBox2.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                bool flag = false;
                oMatSubPlanDetails.FlushToDataSource();
                int num2 = oMatSubPlanDetails.VisualRowCount - 1;
                int RecordNumber = 0;
                while (RecordNumber <= num2)
                {
                    oDBDSSubPlanDetails.GetValue("U_PlanDate", RecordNumber);
                    if (!oDBDSSubPlanDetails.GetValue("U_PlanDate", RecordNumber).Equals(""))
                    {
                        flag = true;
                        num1 = 47;
                        break;
                    }
                    ++RecordNumber; 
                }
                oMatSubPlanDetails.FlushToDataSource();
                string singleValue1 = null;
                string singleValue2 = null;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    singleValue1 = objSBOAPI.Query_Execute(" Select IFNULL(\"U_UNIT\",'01') from \"OITM\" Where \"ItemCode\" ='" + oMatrix.Columns.Item("ICode").Cells.Item(Row).Specific.Value.ToString().Trim() + "' ");
                }
                else
                {
                    singleValue1 = objSBOAPI.Query_Execute(" Select isnull(U_UNIT,'01') from OITM Where ItemCode ='" + oMatrix.Columns.Item("ICode").Cells.Item(Row).Specific.Value.ToString().Trim() + "' ");
                }
                    
                oDBDSSubPlanDetails.SetValue("U_DefUnit", oMatSubPlanDetails.VisualRowCount - 1, singleValue1);
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    singleValue2 = objSBOAPI.Query_Execute(" Select \"U_WhsCode\"  from \"@AIS_BRN1\" Where \"U_BranchCode\" ='" + singleValue1 + "' ");
                }
                else
                {
                    singleValue2 = objSBOAPI.Query_Execute(" Select U_WhsCode  from [@AIS_BRN1] Where U_BranchCode ='" + singleValue1 + "' ");
                }
                    
                oDBDSSubPlanDetails.SetValue("U_WhsCode", oMatSubPlanDetails.VisualRowCount - 1, singleValue2);
                oDBDSSubPlanDetails.SetValue("U_WhsName", oMatSubPlanDetails.VisualRowCount - 1, objSBOAPI.Query_Execute(" Select WhsName  from OWHS Where WhsCode ='" + singleValue2 + "' "));
                oMatSubPlanDetails.LoadFromDataSource();
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    objSBOAPI.Query_Execute("select 'True' from \"@AIS_OUSR\" where \"U_UNAME\"='" + objSBOAPI.oCompany.UserName + "'");
                }
                else
                {
                    objSBOAPI.Query_Execute("select 'True' from [@AIS_OUSR] where U_UNAME='" + objSBOAPI.oCompany.UserName + "'");
                }
                    
                //GVariables.boolModelForm = true;
                //GVariables.123 = Convert.ToString(objSBOAPI.PlanningDetailsFormID);
            }
            catch (Exception ex)
            {
                int lErl = num1;
                oForm.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        #endregion

        #region LoadSubGrid
        public bool LoadSubGrid(
      Matrix oMatSubGrid,
      DBDataSource oDBDSSubGrid,
      DBDataSource oDBDSMainSubGrid,
      int UniqID,
      string[,] DefaulFields = null)
        {
            bool flag1;
            try
            {
                oMatSubGrid.Clear();
                oDBDSSubGrid.Clear();
                SAPbobsCOM.Recordset recordset = null;
                recordset = (SAPbobsCOM.Recordset)objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string strQry = null;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    strQry = "Select \"AliasID\" From \"CUFD\" Where  \"AliasID\" <> 'UniqID' AND \"TableID\" ='" + oDBDSSubGrid.TableName + "' ";
                }
                else
                {
                    strQry = "Select AliasID From CUFD Where  AliasID <> 'UniqID' AND TableID ='" + oDBDSSubGrid.TableName + "' ";
                }
                   
                recordset.DoQuery(strQry);
                bool flag2 = false;
                int num1 = oDBDSMainSubGrid.Size - 1;
                int RecordNumber = 0;
                while (RecordNumber <= num1)
                {
                    if (oDBDSMainSubGrid.GetValue("U_UniqID", RecordNumber).Equals(UniqID.ToString()))
                    {
                        flag2 = true;
                        break;
                    }
                    ++RecordNumber; 
                }
                if (!flag2)
                {
                    objSBOAPI.SetNewLineSubGrid(UniqID, oMatSubGrid, oDBDSSubGrid, 1, "", DefaulFields);
                }
                    

                else if (oDBDSMainSubGrid.Size >= 1 & flag2)
                {
                    int num2 = oDBDSMainSubGrid.Size - 1;
                    int num3 = 0;
                    while (num3 <= num2)
                    {
                        oDBDSMainSubGrid.Offset = num3;
                        if (!oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset).Equals("") & !oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset).Trim().Equals("") && Convert.ToInt32(oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset)) == UniqID)
                        {
                            oDBDSSubGrid.InsertRecord(oDBDSSubGrid.Size);
                            oDBDSSubGrid.Offset = oDBDSSubGrid.Size - 1;
                            oDBDSSubGrid.SetValue("LineID", oDBDSSubGrid.Offset, Convert.ToString(oDBDSSubGrid.Offset + 1));
                            oDBDSSubGrid.SetValue("U_UniqID", oDBDSSubGrid.Offset, Convert.ToString(UniqID));
                            if (DefaulFields != null)
                            {
                                short num4 = (short)(DefaulFields.GetLength(0) - 1);
                                short num5 = 0;
                                while ((int)num5 <= (int)num4)
                                {
                                    oDBDSSubGrid.SetValue(DefaulFields[(int)num5, 0], oDBDSSubGrid.Offset, DefaulFields[(int)num5, 1]);
                                    ++num5; 
                                }
                            }
                            recordset.MoveFirst();
                            int num6 = recordset.RecordCount - 1;
                            int num7 = 0;
                            while (num7 <= num6)
                            {
                                string str = Convert.ToString(string.Concat("U_", recordset.Fields.Item(0).Value));
                                oDBDSSubGrid.SetValue(str, oDBDSSubGrid.Offset, oDBDSMainSubGrid.GetValue(str, oDBDSMainSubGrid.Offset).Trim());
                                recordset.MoveNext();
                                ++num7; 
                            }
                        }
                        ++num3; 
                    }
                    oMatSubGrid.LoadFromDataSource();
                    oMatSubGrid.FlushToDataSource();
                    objSBOAPI.SetNewLineSubGrid(UniqID, oMatSubGrid, oDBDSSubGrid, 1, "", DefaulFields);
                }
                flag1 = true;
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                flag1 = false;
            }
            return flag1;
        }
        #endregion
    }
}
