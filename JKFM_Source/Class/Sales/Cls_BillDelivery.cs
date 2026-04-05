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
using System.Runtime.CompilerServices;

namespace JKFM_Source
{
    class Cls_BillDelivery
    {

        #region Declaration
        SAPbouiCOM.Form objform;
        ClsSBO objSBOAPI;
        Cls_PickListManager obj_PickListManager;

        SAPbouiCOM.Form frmPurchaseDetails;
        SAPbouiCOM.Form frmInstockDetails;
        SAPbouiCOM.Form frmMultipleSelection;
        SAPbobsCOM.Recordset oRS;
        SAPbobsCOM.Recordset oRS1;
        SAPbouiCOM.DBDataSource oDBDSDetail;
        SAPbouiCOM.DBDataSource oDBDSHeader;
        SAPbouiCOM.DBDataSource oDBDSSubPlanDetails;
        SAPbouiCOM.DBDataSource oDBDSSub_Instock;
        //public UIXML objUIXml;
        SAPbouiCOM.ComboBox oCombo;
        SAPbouiCOM.Matrix oMatrix;
        SAPbouiCOM.Matrix oMainSubMatrix;
        SAPbouiCOM.Matrix oSubPlanMatrix;
        SAPbouiCOM.CheckBox oCheck;
        string QryStr;
        SAPbobsCOM.Documents oDoc;
        string SQL;
        string Form;
        int chckFlag;
        int SubRowID_PlanningDetails;
        SAPbouiCOM.Grid oGrid;
        string oPickListNo;
        #endregion        

        #region Constructor
        public Cls_BillDelivery(ClsSBO objSBO)
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
                    if (objform == null)
                        objform = objSBOAPI.SBO_Appln.Forms.Item(formuid);
                    switch (pval.EventType)
                    {
                        case BoEventTypes.et_CLICK:
                            switch (pval.ItemUID)
                            {
                                case "1":
                                    if ((objform.Mode == BoFormMode.fm_ADD_MODE || objform.Mode == BoFormMode.fm_UPDATE_MODE) && !ValidateAll())
                                    {
                                        SystemSounds.Asterisk.Play();
                                        bubbleevent = false;
                                    }
                                    break;
                                case "Matrix":
                                    switch (pval.ColUID)
                                    {
                                        case "V_6":
                                            if (((ICheckBox)oMatrix.Columns.Item("V_6").Cells.Item(pval.Row).Specific).Checked && Convert.ToString(objSBOAPI.Query_Execute(" Select ISNULL(\"U_Return\",'N')  from \"@AIS_OPDE\" m left outer join \"@AIS_OPDE1\" d on m.\"DocEntry\" =d.\"DocEntry\" where \"DocNum\"='" + objform.Items.Item("9").Specific.Value.ToString().Trim() + "' and \"LineId\" ='" + Convert.ToString(pval.Row) + "'")).ToString().Trim() == "Y")
                                            {
                                                objSBOAPI.SBO_Appln.SetStatusBarMessage("You Cannot change....", BoMessageTime.bmt_Short, true);
                                                bubbleevent = false;
                                                return;
                                            }
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
                                    if (pval.Action_Success == true && objform.Mode == BoFormMode.fm_ADD_MODE)
                                    {
                                        InitForm();
                                    }
                                    break;
                            }
                            break;

                        case BoEventTypes.et_COMBO_SELECT:
                            switch(pval.ItemUID)
                            {
                                case "c_series":
                                    if (objform.Mode == BoFormMode.fm_ADD_MODE)
                                    {
                                        oDBDSHeader.SetValue("DocNum", 0, Convert.ToString((long)objform.BusinessObject.GetNextSerialNumber(((IComboBox)objform.Items.Item("c_series").Specific).Value.ToString().Trim(), "OPDE")));
                                    }
                                    break;
                                case "t_Location":
                                    string Location = ((IComboBox)objform.Items.Item("t_Location").Specific).Value.ToString().Trim();
                                    double Rate = 0.00;
                                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                    {
                                        Rate = Convert.ToDouble(objSBOAPI.Query_Execute("Select \"U_Rate\" from \"@AIS_OLOC\" Where \"Code\" ='" + Location + "'"));
                                    }
                                    else
                                    {
                                        Rate = Convert.ToDouble(objSBOAPI.Query_Execute("Select \"U_Rate\" from \"@AIS_OLOC\" Where \"Code\" ='" + Location + "'"));
                                    }
                                        
                                    double Weight = 0.0;
                                    int visualRowCount = oMatrix.VisualRowCount;
                                    int Row = 1;
                                    while (Row <= visualRowCount)
                                    {
                                        string Weights = Convert.ToString(oMatrix.Columns.Item("Weight").Cells.Item(Row).Specific.Value);
                                        Weight += Convert.ToDouble(Weights);
                                        { ++Row; }
                                    }
                                    oDBDSHeader.SetValue("U_LFrieght", 0, Convert.ToString(Rate * Weight));
                                    break;
                                case "c_FrighTyp":
                                    if (pval.ItemChanged == true)
                                    {
                                        objform.Freeze(true);
                                        int MatrixVisualRowCount = oMatrix.VisualRowCount;
                                        int Rows = 1;
                                        while (Rows < MatrixVisualRowCount)
                                        {
                                            double Price = Convert.ToDouble(oMatrix.Columns.Item("Price").Cells.Item(Rows).Specific.Value);
                                            double LineTotal = Convert.ToDouble(oMatrix.Columns.Item("Qty").Cells.Item(Rows).Specific.Value) * Price;
                                            oMatrix.FlushToDataSource();
                                            oDBDSDetail.SetValue("U_LineTotal", (Rows - 1), Convert.ToString(LineTotal));
                                            oMatrix.LoadFromDataSource();

                                            //Thiru Addon Changes - Start - 06.08.2021
                                            //oDBDSDetail.SetValue("U_AVA_SalVal", (Rows - 1), Convert.ToString(oRS.Fields.Item("Sales value for this FY").Value));
                                            string TaxCode = oDBDSDetail.GetValue("U_TaxCode", (Rows - 1)).ToString().Trim();//Convert.ToString(oRS.Fields.Item("TaxCode").Value);
                                            double Rates = 0.00;
                                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                            {
                                                Rates = Convert.ToDouble(objSBOAPI.Query_Execute("Select \"Rate\" from \"OSTC\" Where \"Code\"='" + TaxCode + "' "));
                                            }
                                            else
                                            {
                                                Rates = Convert.ToDouble(objSBOAPI.Query_Execute("Select Rate from OSTC Where Code='" + TaxCode + "' "));
                                            }
                                            double Tax_Amount = (LineTotal * Rates) / 100;
                                            oMatrix.Columns.Item("Col_TaxAmt").Cells.Item(Rows).Specific.Value = Tax_Amount;
                                            //oDBDSDetail.SetValue("U_AVA_TaxAmt", (Rows - 1), Tax_Amount.ToString());
                                            string Freight_Type = oDBDSHeader.GetValue("U_FrightType", 0).ToString().Trim();
                                            if (Freight_Type == "O")
                                            {
                                                double Tonnage = Convert.ToDouble(oDBDSDetail.GetValue("U_SOWeight", (Rows - 1)).ToString().Trim());
                                                string Location_Code = ((IComboBox)objform.Items.Item("t_Location").Specific).Value.ToString().Trim();
                                                double Tax_Rate = 0.00;
                                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                                {
                                                    Tax_Rate = Convert.ToDouble(objSBOAPI.Query_Execute("Select \"U_Rate\" from \"@AIS_OLOC\" Where \"Code\" ='" + Location_Code + "'"));
                                                }
                                                else
                                                {
                                                    Tax_Rate = Convert.ToDouble(objSBOAPI.Query_Execute("Select U_Rate from [@AIS_OLOC] Where Code ='" + Location_Code + "'"));
                                                }
                                                double Main_Loading_Charges = Tonnage * Tax_Rate;
                                                double Sub_Loading_Charges = (Main_Loading_Charges * Rates) / 100;
                                                double Loading_Charges = Main_Loading_Charges + Sub_Loading_Charges;
                                                double Doc_Total = LineTotal + Tax_Amount + Loading_Charges;
                                                oMatrix.Columns.Item("Col_DocTot").Cells.Item(Rows).Specific.Value = Doc_Total;
                                                //oDBDSDetail.SetValue("U_AVA_DocTot", (Rows - 1), Convert.ToString(Doc_Total));
                                            }
                                            else if (Freight_Type == "H" || Freight_Type == "C")
                                            {
                                                double Doc_Total = LineTotal + Tax_Amount;
                                                oMatrix.Columns.Item("Col_DocTot").Cells.Item(Rows).Specific.Value = Doc_Total;
                                                //oDBDSDetail.SetValue("U_AVA_DocTot", (Rows - 1), Convert.ToString(Doc_Total));
                                            }

                                            //Thiru Addon Changes - End - 06.08.2021
                                            { ++Rows; }
                                        }
                                        objform.Freeze(false);
                                    }
                                    break;
                            }
                            break;

                        case BoEventTypes.et_CLICK:
                            
                            switch(pval.ItemUID)
                            {
                                case "b_Ref":
                                    string Pick_List_No = string.Empty;
                                    if (oMatrix.VisualRowCount > 0)
                                        Pick_List_No = oMatrix.Columns.Item("PLN").Cells.Item(1).Specific.Value.ToString().Trim();
                                    string Vehicle_No = objform.Items.Item("t_VNo").Specific.Value.ToString().Trim();
                                    string Unit = oDBDSHeader.GetValue("U_Unit", 0).ToString().Trim();
                                    string WBType = oDBDSHeader.GetValue("U_WBType", 0).ToString().Trim();
                                    if (Pick_List_No == string.Empty)
                                    {
                                        objSBOAPI.SBO_Appln.SetStatusBarMessage("Pick List No. isn't available....", BoMessageTime.bmt_Short, true);
                                    }
                                    else if (WBType == string.Empty)
                                    {
                                        objSBOAPI.SBO_Appln.SetStatusBarMessage("Weigh Bridge Type Shouldn't be Empty....", BoMessageTime.bmt_Short, true);
                                    }
                                    else
                                    {
                                        Recordset recordset = null;
                                        recordset = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                        {
                                            recordset.DoQuery("CALL \"@AIS_BillDelivery_LoadWeighBridgeDetails\"('" + Pick_List_No + "','" + Vehicle_No + "','" + Unit + "','" + WBType + "')");
                                        }
                                        else
                                        {
                                            recordset.DoQuery("Exec [dbo].[@AIS_BillDelivery_LoadWeighBridgeDetails]'" + Pick_List_No + "','" + Vehicle_No + "','" + Unit + "','" + WBType + "'");
                                        }
                                            
                                        if (recordset.RecordCount > 0)
                                        {
                                            recordset.MoveFirst();
                                            //oDBDSHeader.SetValue("U_LorryIn", 0, Strings.Replace(Convert.ToString(recordset.Fields.Item("Weight1Date").Value), ":", "", 1, -1, CompareMethod.Binary));
                                            //oDBDSHeader.SetValue("U_LorryOut", 0, Strings.Replace(Convert.ToString(recordset.Fields.Item("Weight2Date").Value), ":", "", 1, -1, CompareMethod.Binary));
                                            oDBDSHeader.SetValue("U_LorryIn", 0, Convert.ToString(recordset.Fields.Item("Weight1Date").Value));
                                            oDBDSHeader.SetValue("U_LorryOut", 0, Convert.ToString(recordset.Fields.Item("Weight2Date").Value));
                                            oDBDSHeader.SetValue("U_Unit", 0, Convert.ToString(recordset.Fields.Item("Weight1Indicator").Value));
                                            oDBDSHeader.SetValue("U_FWeight", 0, Convert.ToString(recordset.Fields.Item("Weight1").Value));
                                            oDBDSHeader.SetValue("U_SWeight", 0, Convert.ToString(recordset.Fields.Item("Weight2").Value));
                                            oDBDSHeader.SetValue("U_WBNetWt", 0, Convert.ToString(recordset.Fields.Item("ItemWeight").Value));
                                            if (recordset.Fields.Item("DriverName").Value != "")
                                                oDBDSHeader.SetValue("U_DriverName", 0, Convert.ToString(recordset.Fields.Item("DriverName").Value));
                                            if (recordset.Fields.Item("MobileNo").Value != "")
                                                oDBDSHeader.SetValue("U_Mobile", 0, Convert.ToString(recordset.Fields.Item("MobileNo").Value));
                                        }
                                        else
                                        {
                                            objSBOAPI.SBO_Appln.SetStatusBarMessage("Weigh Bridge Detail's aren't available....", BoMessageTime.bmt_Short, true);
                                        }
                                    }
                                    break;

                                case "b_Del":
                                    if (objform.Items.Item(pval.ItemUID).Enabled)
                                    {

                                        if (((IComboBox)objform.Items.Item("c_Appr").Specific).Selected.Value != "Y")
                                        {
                                            double ActWeight = Convert.ToDouble(oDBDSHeader.GetValue("U_ActWeight", 0));
                                            double TNetWeight = Convert.ToDouble(oDBDSHeader.GetValue("U_TNetWeight", 0));
                                            double WBNetWt = Convert.ToDouble(oDBDSHeader.GetValue("U_WBNetWt", 0));

                                            string Value = null;
                                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                            {
                                                Value = "CALL \"Bill_Delivery_Weight\"('" + oDBDSHeader.GetValue("DocEntry", 0).ToString().Trim() + "')";
                                            }
                                            else
                                            {
                                                Value = "Exec dbo.[Bill_Delivery_Weight]'" + oDBDSHeader.GetValue("DocEntry", 0).ToString().Trim() + "'";
                                            }

                                            Recordset oRecordset = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            oRecordset.DoQuery(Value);
                                            if(oRecordset.RecordCount > 0)
                                            {
                                                if (ActWeight - TNetWeight > oRecordset.Fields.Item("VAL").Value)
                                                {
                                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Actual Weight field not equal to Tem Net Weight field (tolerance plus or minus '" + oRecordset.Fields.Item("VAL").Value + "' kg)", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                                    break;
                                                }
                                                if (ActWeight - TNetWeight < -(oRecordset.Fields.Item("VAL").Value))
                                                {
                                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Actual Weight field not equal to Tem Net Weight field (tolerance plus or minus '" + oRecordset.Fields.Item("VAL").Value + "' kg)", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                                    break;
                                                }
                                                if (WBNetWt != TNetWeight)
                                                {
                                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Tem Weight field not equal to WB Net Weight field ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                                    break;
                                                }
                                            }
                                        }
                                        else if (Convert.ToDouble(oDBDSHeader.GetValue("U_WBNetWt", 0)) <= 0.0)
                                        {
                                            objSBOAPI.SBO_Appln.StatusBar.SetText("WB Net Weight should be greater than zero...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                            break;
                                        }
                                        if (!objSBOAPI.oCompany.InTransaction)
                                            objSBOAPI.oCompany.StartTransaction();
                                        if (PostingARInvoice())
                                        {
                                            if (objSBOAPI.oCompany.InTransaction)
                                                objSBOAPI.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                                            oDBDSHeader.SetValue("U_Status", 0, "C");
                                            if (objform.Mode == BoFormMode.fm_OK_MODE)
                                                objform.Mode = BoFormMode.fm_UPDATE_MODE;
                                            if (objform.Mode == BoFormMode.fm_UPDATE_MODE)
                                                objform.Items.Item("1").Click(BoCellClickType.ct_Regular);
                                            ArrayList arrayList = new ArrayList();
                                            int visualRowCount = oMatrix.VisualRowCount;
                                            int Row = 1;
                                            while (Row <= visualRowCount)
                                            {
                                                string TargetEntry = oDBDSDetail.GetValue("U_TargetEntry", (Row - 1));
                                                if (TargetEntry != string.Empty && !arrayList.Contains(TargetEntry))
                                                {
                                                    Recordset Orec = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                                    string str = null;
                                                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                                    {
                                                        str = "CALL \"@AIS_BillDelivery_XMLGeneration\"('" + TargetEntry + "')";
                                                    }
                                                    else
                                                    {
                                                        str = "Exec dbo.[@AIS_BillDelivery_XMLGeneration]'" + TargetEntry + "'";
                                                    }
                                                        
                                                    objSBOAPI.SBO_Appln.StatusBar.SetText(str, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                                    Orec.DoQuery(str);
                                                    arrayList.Add(TargetEntry);
                                                }
                                                { ++Row; }
                                            }
                                        }
                                        else
                                        {
                                            int visualRowCount = oMatrix.VisualRowCount;
                                            int Row = 1;
                                            while (Row <= visualRowCount)
                                            {
                                                oMatrix.FlushToDataSource();
                                                oDBDSDetail.SetValue("U_TargetEntry", checked(Row - 1), string.Empty);
                                                oDBDSDetail.SetValue("U_TargetNum", checked(Row - 1), string.Empty);
                                                oDBDSDetail.SetValue("U_TargetObject", checked(Row - 1), string.Empty);
                                                oMatrix.LoadFromDataSource();
                                                { ++Row; }
                                            }
                                            if (objSBOAPI.oCompany.InTransaction)
                                                objSBOAPI.oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                                            SystemSounds.Asterisk.Play();
                                        }

                                    }
                                    break;

                            }
                            break;

                        case BoEventTypes.et_VALIDATE:
                            switch(pval.ItemUID)
                            {
                                case "t_TNetWt":
                                    if (pval.ItemChanged == true)
                                    {
                                        oDBDSHeader.SetValue("U_DiffWeight", 0, Convert.ToString(Convert.ToDouble(oDBDSHeader.GetValue("U_ActWeight", 0)) - Convert.ToDouble(oDBDSHeader.GetValue("U_TNetWeight", 0))));
                                    }
                                    break;
                                case "Matrix":
                                    switch(pval.ColUID)
                                    {
                                        case "PQty":
                                            if (pval.ItemChanged == true)
                                            {
                                                LoadingCharges();
                                            }
                                            break;
                                        case "Price":
                                            if (pval.ItemChanged == true)
                                            {
                                                double Price = Convert.ToDouble(oMatrix.Columns.Item("Price").Cells.Item(pval.Row).Specific.Value);
                                                double LineTotal = Convert.ToDouble(oMatrix.Columns.Item("Qty").Cells.Item(pval.Row).Specific.Value) * Price;
                                                oMatrix.FlushToDataSource();
                                                oDBDSDetail.SetValue("U_LineTotal", (pval.Row - 1), Convert.ToString(LineTotal));
                                                oMatrix.LoadFromDataSource();

                                                //Thiru Addon Changes - Start - 06.08.2021
                                                //oDBDSDetail.SetValue("U_AVA_SalVal", (pval.Row - 1), Convert.ToString(oRS.Fields.Item("Sales value for this FY").Value));
                                                string TaxCode = oDBDSDetail.GetValue("U_TaxCode", (pval.Row - 1)).ToString().Trim();//Convert.ToString(oRS.Fields.Item("TaxCode").Value);
                                                double Rates = 0.00;
                                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                                {
                                                    Rates = Convert.ToDouble(objSBOAPI.Query_Execute("Select \"Rate\" from \"OSTC\" Where \"Code\"='" + TaxCode + "' "));
                                                }
                                                else
                                                {
                                                    Rates = Convert.ToDouble(objSBOAPI.Query_Execute("Select Rate from OSTC Where Code='" + TaxCode + "' "));
                                                }
                                                double Tax_Amount = (LineTotal * Rates) / 100;
                                                oMatrix.Columns.Item("Col_TaxAmt").Cells.Item(pval.Row).Specific.Value= Tax_Amount;
                                                //oDBDSDetail.SetValue("U_AVA_TaxAmt", (pval.Row - 1), Tax_Amount.ToString());
                                                string Freight_Type = oDBDSHeader.GetValue("U_FrightType", 0).ToString().Trim();
                                                if (Freight_Type == "O")
                                                {
                                                    double Tonnage = Convert.ToDouble(oDBDSDetail.GetValue("U_SOWeight", (pval.Row - 1)).ToString().Trim());
                                                    string Location_Code = ((IComboBox)objform.Items.Item("t_Location").Specific).Value.ToString().Trim();
                                                    double Rate = 0.00;
                                                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                                    {
                                                        Rate = Convert.ToDouble(objSBOAPI.Query_Execute("Select \"U_Rate\" from \"@AIS_OLOC\" Where \"Code\" ='" + Location_Code + "'"));
                                                    }
                                                    else
                                                    {
                                                        Rate = Convert.ToDouble(objSBOAPI.Query_Execute("Select U_Rate from [@AIS_OLOC] Where Code ='" + Location_Code + "'"));
                                                    }
                                                    double Main_Loading_Charges = Tonnage * Rate;
                                                    double Sub_Loading_Charges = (Main_Loading_Charges * Rates) / 100;
                                                    double Loading_Charges = Main_Loading_Charges + Sub_Loading_Charges;
                                                    double Doc_Total = LineTotal + Tax_Amount + Loading_Charges;
                                                    oMatrix.Columns.Item("Col_DocTot").Cells.Item(pval.Row).Specific.Value = Doc_Total;
                                                    //oDBDSDetail.SetValue("U_AVA_DocTot", (pval.Row - 1), Convert.ToString(Doc_Total));
                                                }
                                                else if (Freight_Type == "H" || Freight_Type == "C")
                                                {
                                                    double Doc_Total = LineTotal + Tax_Amount;
                                                    oMatrix.Columns.Item("Col_DocTot").Cells.Item(pval.Row).Specific.Value = Doc_Total;
                                                    //oDBDSDetail.SetValue("U_AVA_DocTot", (pval.Row - 1), Convert.ToString(Doc_Total));
                                                }

                                                //Thiru Addon Changes - End - 06.08.2021

                                            }
                                            break;
                                    }
                                    break;
                                case "25":
                                    if (pval.ItemChanged == true)
                                    {
                                        if (((IComboBox)objform.Items.Item("28").Specific).Selected.Value == "C")
                                            break;
                                        else
                                            LoadBills();
                                    }
                                    break;
                            }
                            break;

                        case BoEventTypes.et_CHOOSE_FROM_LIST:
                            DataTable selectedObjects = ((IChooseFromListEvent)pval).SelectedObjects;
                            if (selectedObjects != null)
                            {
                                switch(pval.ItemUID)
                                {
                                    case "Matrix":
                                        switch(pval.ColUID)
                                        {
                                            case "TaxCode":
                                                oMatrix.FlushToDataSource();
                                                oDBDSDetail.SetValue("U_TaxCode", (pval.Row - 1), Convert.ToString(selectedObjects.GetValue("Code", 0)).Trim());
                                                oMatrix.LoadFromDataSource();
                                                if (objform.Mode == BoFormMode.fm_OK_MODE)
                                                    objform.Mode = BoFormMode.fm_UPDATE_MODE;
                                                break;
                                            case "WhsCode":
                                                oMatrix.FlushToDataSource();
                                                oDBDSDetail.SetValue("U_WhsCode", (pval.Row - 1), Convert.ToString(selectedObjects.GetValue("WhsCode", 0)).Trim());
                                                oMatrix.LoadFromDataSource();
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
                        case "OPDE":
                            objform = objSBOAPI.LoadForm("Delivery.xml", "OPDE");
                            objform.Freeze(true);
                            oDBDSHeader = objform.DataSources.DBDataSources.Item("@AIS_OPDE");
                            oDBDSDetail = objform.DataSources.DBDataSources.Item("@AIS_PDE1");
                            oMatrix = (SAPbouiCOM.Matrix)objform.Items.Item("Matrix").Specific;
                            SAPbouiCOM.ComboBox oCombo, oCombo1, oCombo2;
                            oCombo = objform.Items.Item("c_Unit").Specific;
                            oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                            oCombo1 = objform.Items.Item("c_WBType").Specific;
                            oCombo1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                            oCombo2 = objform.Items.Item("t_Location").Specific;
                            oCombo2.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                            oCombo = objform.Items.Item("c_TVehicle").Specific;
                            oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                            oCombo1 = objform.Items.Item("c_series").Specific;
                            oCombo1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                            oCombo2 = objform.Items.Item("28").Specific;
                            oCombo2.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                            oCombo = objform.Items.Item("c_Appr").Specific;
                            oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                            oCombo1 = objform.Items.Item("c_FrighTyp").Specific;
                            oCombo1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                            if (objform.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            }
                            
                            addmode();
                            DefineModeForFields();
                            LoadBills();
                            objform.ReportType = "RCRI";
                            DefineModeForFields();
                            objform.Freeze(false);
                            break;
                            
                        case "1281":
                            if (objform.Title == "Bill Delivery")
                            {
                                objform.Items.Item("9").Enabled = true;
                                //oMatrix.Columns.Item("V_6").Editable = true;
                                //oMatrix.Columns.Item("Credit").Editable = true;

                                ComboBox comboBox = (ComboBox)objform.Items.Item("28").Specific;

                                if (((IComboBox)objform.Items.Item("28").Specific).Selected.Value == "C")
                                    objform.Items.Item("28").Enabled = false;
                                else
                                    objform.Items.Item("28").Enabled = true;
                            }
                            else
                                objform.Items.Item("28").Enabled = false;
                            break;

                        case "1282":
                            if (objform.Title == "Bill Delivery")
                            {
                                addmode();
                            }
                            else
                            {
                                BubbleEvent = true;
                                objform.Mode = BoFormMode.fm_FIND_MODE;
                                objform.Items.Item("9").Enabled = true;
                                oMatrix.Columns.Item("V_6").Editable = true;
                                oMatrix.Columns.Item("Credit").Editable = true;
                            }
                            break;
                        case "":
                            if (objform.Title == "Bill Delivery")
                            {
                                ComboBox comboBox1 = (ComboBox)objform.Items.Item("28").Specific;
                            }
                            else
                                objform.Items.Item("28").Enabled = false;
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
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            
                            break;

                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            string str = oDBDSHeader.GetValue("U_Unit", 0).ToString().Trim();
                            oCombo = (ComboBox)objform.Items.Item("c_WBType").Specific;
                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                            {
                                QryStr = "SELECT  \"Code\",\"Name\" FROM \"@AIS_OWBT\" Where \"U_Unit\"='" + str + "'";
                            }
                            else
                            {
                                QryStr = "SELECT  Code,Name FROM [@AIS_OWBT] Where U_Unit='" + str + "'";
                            }
                                
                            objSBOAPI.SetComboBoxValueRefresh(oCombo, QryStr);
                            if (objform.Mode != BoFormMode.fm_ADD_MODE)
                                oMatrix.Columns.Item("V_5").Editable = false;
                            else
                                oMatrix.Columns.Item("V_5").Editable = true;
                            if (oDBDSHeader.GetValue("U_Status", 0).ToString().Trim() == "C")
                                objform.Mode = BoFormMode.fm_VIEW_MODE;
                            else
                                objform.Mode = BoFormMode.fm_OK_MODE;

                            //Thiru - Price Field Editable and Non - Editable Conditions Check Based on Super User - Start - 20210805
                            string Bill_Delivery_Price_Edit = objSBOAPI.Query_Execute("Select \"U_AVA_BDPEUR\" From \"OUSR\" Where \"USER_CODE\" = '" + objSBOAPI.oCompany.UserName + "'");
                            if (Bill_Delivery_Price_Edit == "Y")
                            {
                                oMatrix.Columns.Item("Price").Editable = true;
                            }
                            else
                            {
                                oMatrix.Columns.Item("Price").Editable = false;
                            }

                            //Thiru - Price Field Editable and Non - Editable Conditions Check Based on Super User - End - 20210805
                            
                            SAPbobsCOM.Recordset oRec;
                            oRec = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            if (((IComboBox)objform.Items.Item("28").Specific).Selected.Value != "C")
                            {
                                string SelectQuery = "Select \"Name\" from \"@AVA_ENEST\"";
                                oRec.DoQuery(SelectQuery);
                                if (oRec.RecordCount > 0)
                                {
                                    while (!oRec.EoF)
                                    {
                                        if (Bill_Delivery_Price_Edit == "Y")
                                        {
                                            objform.Items.Item(oRec.Fields.Item("Name").Value.ToString().Trim()).Enabled = true;
                                        }
                                        else
                                        {
                                            objform.Items.Item(oRec.Fields.Item("Name").Value.ToString().Trim()).Enabled = false;
                                        }
                                        oRec.MoveNext();
                                    }
                                }
                            }
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

        #region addmode
        public void addmode()
        {
            try
            {
                objform.Items.Item("9").Enabled = false;
                objSBOAPI.LoadDocumentDate((EditText)objform.Items.Item("10").Specific);
                objSBOAPI.LoadComboBoxSeries((ComboBox)objform.Items.Item("c_series").Specific, "OPDE");
                oCombo = (ComboBox)objform.Items.Item("c_TVehicle").Specific;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    QryStr = "Select \"Code\",\"U_TVehicle\" from \"@AIS_OVEH\"";
                }
                else
                {
                    QryStr = "Select \"Code\",\"U_TVehicle\" from \"@AIS_OVEH\"";
                }
                    
                objSBOAPI.setComboBoxValue(oCombo, QryStr);
                oCombo = (ComboBox)objform.Items.Item("t_Location").Specific;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    QryStr = " Select \"Code\",\"U_Location\" From \"@AIS_OLOC\"";
                }
                else
                {
                    QryStr = " Select Code,U_Location From [@AIS_OLOC]";
                }
                    
                objSBOAPI.setComboBoxValue(oCombo, QryStr);
                oCombo = (ComboBox)objform.Items.Item("c_Unit").Specific;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    QryStr = " SELECT  \"Code\",\"Name\" FROM \"@AIS_OUNT\" ";
                }
                else
                {
                    QryStr = " SELECT  Code,Name FROM [@AIS_OUNT] ";
                }
                    
                objSBOAPI.setComboBoxValue(oCombo, QryStr);
                oCombo = (ComboBox)objform.Items.Item("c_WBType").Specific;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    QryStr = "SELECT  \"Code\",\"Name\" FROM \"@AIS_OWBT\"";
                }
                else
                {
                    QryStr = "SELECT  Code,Name FROM [@AIS_OWBT]";
                }
                    
                objSBOAPI.setComboBoxValue(oCombo, QryStr);
                objform.ActiveItem = "c_Unit";
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Add Mode Function Failure - " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region DefineModeForFields
        public void DefineModeForFields()
        {
            try
            {
                objform.Items.Item("b_Del").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Visible, 2, BoModeVisualBehavior.mvb_False);
                objform.Items.Item("b_Del").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Visible, 1, BoModeVisualBehavior.mvb_True);
                objform.Items.Item("c_series").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);
                objform.Items.Item("9").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);
                objform.Items.Item("10").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);
                objform.Items.Item("28").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);
                objform.Items.Item("28").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_True);
                objform.Items.Item("c_series").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Define Mode For Fields Failed: " + ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region LoadBills
        private void LoadBills()
        {
            object obj;
            
            try
            {
                objform.Freeze(true);
                oRS = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRS1 = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                if (Cls_PickListManager.oPickListNo == null)
                {
                    Cls_PickListManager.oPickListNo = "0";
                }
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    oRS.DoQuery("CALL \"@AIS_BillDelivery_LoadDelivery\"('" + Cls_PickListManager.oPickListNo + "')");
                }
                else
                {
                    oRS.DoQuery("Exec dbo.[@AIS_BillDelivery_LoadDelivery]'" + Cls_PickListManager.oPickListNo + "'");
                }
                    
                oMatrix.Clear();
                oDBDSDetail.Clear();
                chckFlag = 1;
                oMatrix.FlushToDataSource();
                if (objform.Mode != BoFormMode.fm_ADD_MODE)
                    oMatrix.Columns.Item("V_5").Editable = false;
                else
                    oMatrix.Columns.Item("V_5").Editable = true;
                double DocTotal = 0.0;
                double TotalBag = 0.0;
                double TotalTon = 0.0;
                int size = oDBDSDetail.Size;
                if (oRS.RecordCount > 0)
                {
                    oRS.MoveFirst();
                    //Thiru - 20210728
                    oDBDSHeader.SetValue("U_Unit", 0, Convert.ToString(oRS.Fields.Item("U_UNIT").Value));

                    //string Unit = objSBOAPI.Query_Execute("Select \"U_AVA_UNIT\" from \"OPKL\" Where \"AbsEntry\"='" + Cls_PickListManager.oPickListNo + "'");
                    //oDBDSHeader.SetValue("U_Unit", 0, Convert.ToString(Unit));
                    //Thiru - 20210728
                    oCombo = (ComboBox)objform.Items.Item("c_WBType").Specific;
                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        QryStr = Convert.ToString("SELECT  \"Code\",\"Name\" FROM \"@AIS_OWBT\" Where \"U_Unit\"='" + oRS.Fields.Item("U_UNIT").Value + "'");
                    }
                    else
                    {
                        QryStr = Convert.ToString("SELECT  Code,Name FROM [@AIS_OWBT] Where U_Unit='" + oRS.Fields.Item("U_UNIT").Value + "'");
                    }
                        
                    objSBOAPI.SetComboBoxValueRefresh(oCombo, QryStr);

                    //Thiru - 20210728
                    oDBDSHeader.SetValue("U_Mobile", 0, Convert.ToString(oRS.Fields.Item("MobileNo").Value));
                    oDBDSHeader.SetValue("U_VehicleNo", 0, Convert.ToString(oRS.Fields.Item("TruckNo").Value));
                    oDBDSHeader.SetValue("U_DriverName", 0, Convert.ToString(oRS.Fields.Item("DriverName").Value));

                    //string MobileNumber = objSBOAPI.Query_Execute("Select \"U_MobileNum\" from \"OPKL\" Where \"AbsEntry\"='" + Cls_PickListManager.oPickListNo + "'");
                    //oDBDSHeader.SetValue("U_Mobile", 0, Convert.ToString(MobileNumber));
                    //string VehicleNumber = objSBOAPI.Query_Execute("Select \"U_TruckNo\" from \"OPKL\" Where \"AbsEntry\"='" + Cls_PickListManager.oPickListNo + "'");
                    //oDBDSHeader.SetValue("U_VehicleNo", 0, Convert.ToString(VehicleNumber));
                    //string DriverName = objSBOAPI.Query_Execute("Select \"U_Drivname\" from \"OPKL\" Where \"AbsEntry\"='" + Cls_PickListManager.oPickListNo + "'");
                    //oDBDSHeader.SetValue("U_DriverName", 0, Convert.ToString(DriverName));
                    //Thiru - 20210728
                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        //Thiru - 20210728
                        //oDBDSHeader.SetValue("U_TFVehicle", 0, Convert.ToString(objSBOAPI.Query_Execute(Convert.ToString("select \"Code\"  from \"@AIS_OVEH\"  where \"U_TVehicle\" ='" + oRS.Fields.Item("U_Vehicle").Value + "'"))));
                        oDBDSHeader.SetValue("U_Location", 0, Convert.ToString(objSBOAPI.Query_Execute(Convert.ToString("select \"Code\"  from \"@AIS_OLOC\" T0 inner join \"CRD1\" T1 on T1.\"City\" =T0.\"U_Location\" where \"CardCode\" ='" + oRS.Fields.Item("CardCode").Value) + "' and \"AdresType\" ='S'")));
                        string VehicleType = objSBOAPI.Query_Execute("Select \"U_AVA_VEHICLETYPE\" from \"OPKL\" Where \"AbsEntry\"='" + Cls_PickListManager.oPickListNo + "'");
                        //string VehicleType = objSBOAPI.Query_Execute("Select T0.\"Code\" from \"@AIS_OVEH\" T0 INNER JOIN \"OPKL\" T1 ON T1.\"U_Vehicle\"=T0.\"U_TVehicle\" Where T0.\"U_TVehicle\" ='" + Vehicle + "' and T1.\"AbsEntry\"='" + Cls_PickListManager.oPickListNo + "'");
                        oDBDSHeader.SetValue("U_TFVehicle", 0, Convert.ToString(VehicleType));
                        //string Location = objSBOAPI.Query_Execute("Select T0.\"Code\" from \"@AIS_OLOC\" T0 inner join \"OPKL\" T1 on T1.\"U_AVA_LOCN\" =T0.\"U_Location\"  Where T1.\"AbsEntry\"='" + Cls_PickListManager.oPickListNo + "'");
                        //oDBDSHeader.SetValue("U_Location", 0, Convert.ToString(Location));
                        //Thiru - 20210728
                    }
                    else
                    {
                        oDBDSHeader.SetValue("U_TFVehicle", 0, Convert.ToString(objSBOAPI.Query_Execute(Convert.ToString("select Code  from [@AIS_OVEH]  where U_TVehicle ='" + oRS.Fields.Item("U_Vehicle").Value + "'"))));
                        oDBDSHeader.SetValue("U_Location", 0, Convert.ToString(objSBOAPI.Query_Execute(Convert.ToString("select Code  from [@AIS_OLOC] T0 inner join CRD1 T1 on T1.City =T0.U_Location where CardCode ='" + oRS.Fields.Item("CardCode").Value) + "' and AdresType ='S'")));
                    }

                    //Thiru - 18.03.2021
                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        //Thiru - 20210728
                        //oDBDSHeader.SetValue("U_FrightType", 0, Convert.ToString(objSBOAPI.Query_Execute(Convert.ToString("Select \"U_FRETY\" from \"@AIS_TRAN\" Where \"U_TRNO\"='" + Convert.ToString(oRS.Fields.Item("TruckNo").Value) + "'"))));

                        string FreightType = objSBOAPI.Query_Execute("Select \"U_FRTY\" from \"OPKL\" Where \"AbsEntry\"='" + Cls_PickListManager.oPickListNo + "' and \"U_TruckNo\"='" + Convert.ToString(oRS.Fields.Item("TruckNo").Value) + "'");
                        oDBDSHeader.SetValue("U_FrightType", 0, Convert.ToString(FreightType));
                        //Thiru - 20210728

                    }
                    else
                    {
                        oDBDSHeader.SetValue("U_FrightType", 0, Convert.ToString(objSBOAPI.Query_Execute(Convert.ToString("Select U_FRETY from [@AIS_TRAN] Where U_TRNO='" + Convert.ToString(oRS.Fields.Item("TruckNo").Value) + "'"))));
                    }
                    //Thiru - 18.03.2021

                    int recordCount = oRS.RecordCount;
                    int Row = 1;
                    while (Row <= recordCount)
                    {
                        oDBDSDetail.InsertRecord(oDBDSDetail.Size);
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait Loading - " + Convert.ToString(size), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                        oDBDSDetail.SetValue("LineId", (size - 1), Convert.ToString(size));
                        oDBDSDetail.SetValue("U_BillNo", (size - 1), Convert.ToString(oRS.Fields.Item("DocNum").Value));
                        oDBDSDetail.SetValue("U_BDocEntry", (size - 1), Convert.ToString(oRS.Fields.Item("DocEntry").Value));
                        oDBDSDetail.SetValue("U_PListNo", (size - 1), Convert.ToString(oRS.Fields.Item("AbsEntry").Value));
                        oDBDSDetail.SetValue("U_BDate", (size - 1), Convert.ToDateTime(oRS.Fields.Item("DocDate").Value.ToString().Trim()).ToString("yyyyMMdd"));
                        oDBDSDetail.SetValue("U_CusCode", (size - 1), Convert.ToString(oRS.Fields.Item("CardCode").Value));
                        oDBDSDetail.SetValue("U_CusName", (size - 1), Convert.ToString(oRS.Fields.Item("CardName").Value));
                        oDBDSDetail.SetValue("U_ItemCode", (size - 1), Convert.ToString(oRS.Fields.Item("ItemCode").Value));
                        oDBDSDetail.SetValue("U_ItemName", (size - 1), Convert.ToString(oRS.Fields.Item("ItemName").Value));
                        oDBDSDetail.SetValue("U_DefUnit", (size - 1), Convert.ToString(oRS.Fields.Item("U_UNIT").Value));
                        oDBDSDetail.SetValue("U_BaseLine", (size - 1), Convert.ToString(oRS.Fields.Item("LineNum").Value));
                        oDBDSDetail.SetValue("U_Qty", (size - 1), Convert.ToString(oRS.Fields.Item("PendingQty").Value));
                        oDBDSDetail.SetValue("U_TreeType", (size - 1), Convert.ToString(oRS.Fields.Item("TreeType").Value));
                        oDBDSDetail.SetValue("U_UomCode", (size - 1), Convert.ToString(oRS.Fields.Item("SalUnitMsr").Value));



                        double PendingQty = Convert.ToDouble(oRS.Fields.Item("PendingQty").Value);

                        //Thiru - Price Field Editable and Non - Editable Conditions Check Based on Super User - Start - 20210805
                        string Bill_Delivery_Price_Edit = objSBOAPI.Query_Execute("Select \"U_AVA_BDPEUR\" From \"OUSR\" Where \"USER_CODE\" = '" + objSBOAPI.oCompany.UserName + "'");
                        if (Bill_Delivery_Price_Edit == "Y")
                        {
                            oMatrix.Columns.Item("Price").Editable = true;
                        }
                        else
                        {
                            oMatrix.Columns.Item("Price").Editable = false;
                        }

                        SAPbobsCOM.Recordset oRec;
                        oRec = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        if (((IComboBox)objform.Items.Item("28").Specific).Selected.Value != "C")
                        {
                            string SelectQuery = "Select \"Name\" from \"@AVA_ENEST\"";
                            oRec.DoQuery(SelectQuery);
                            if (oRec.RecordCount > 0)
                            {
                                while (!oRec.EoF)
                                {
                                    if (Bill_Delivery_Price_Edit == "Y")
                                    {
                                        objform.Items.Item(oRec.Fields.Item("Name").Value.ToString().Trim()).Enabled = true;
                                    }
                                    else
                                    {
                                        objform.Items.Item(oRec.Fields.Item("Name").Value.ToString().Trim()).Enabled = false;
                                    }
                                    oRec.MoveNext();
                                }
                            }
                        }
                           
                        //Thiru - Price Field Editable and Non - Editable Conditions Check Based on Super User - End - 20210805

                        Convert.ToDouble(oRS.Fields.Item("Price").Value);
                        string GetUnitPrice_Query = null;
                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        {
                            GetUnitPrice_Query = Convert.ToString("CALL \"@AIS_SalesPlanning_GetUnitPrice\"('" + oRS.Fields.Item("ItemCode").Value + "','" + oRS.Fields.Item("CardCode").Value + "','" + Convert.ToString(oRS.Fields.Item("SalUnitMsr").Value) + "') ");
                        }
                        else
                        {
                            GetUnitPrice_Query = Convert.ToString("Exec dbo.[@AIS_SalesPlanning_GetUnitPrice]'" + oRS.Fields.Item("ItemCode").Value + "','" + oRS.Fields.Item("CardCode").Value + "','" + Convert.ToString(oRS.Fields.Item("SalUnitMsr").Value) + "' ");
                        }
                            
                        string GetUnitPrice = Convert.ToString(objSBOAPI.Query_Execute(GetUnitPrice_Query));
                        double UnitPrice = GetUnitPrice != string.Empty ? Convert.ToDouble(GetUnitPrice) : 0.0;
                        oDBDSDetail.SetValue("U_Price", (size - 1), Convert.ToString(UnitPrice));
                        double LineTotal = PendingQty * UnitPrice;
                        oDBDSDetail.SetValue("U_LineTotal", (size - 1), Convert.ToString(LineTotal));
                        if (oRS.Fields.Item("TreeType").Value != "I")
                        {
                            TotalBag += PendingQty;
                            DocTotal += LineTotal;
                            TotalTon = Convert.ToDouble(TotalTon + oRS.Fields.Item("TonInWeight").Value);
                        }
                        oDBDSDetail.SetValue("U_TaxCode", (size - 1), Convert.ToString(oRS.Fields.Item("TaxCode").Value));
                        oDBDSDetail.SetValue("U_WhsCode", (size - 1), Convert.ToString(oRS.Fields.Item("WhsCode").Value));
                        oDBDSDetail.SetValue("U_SOWeight", (size - 1), Convert.ToString(oRS.Fields.Item("TonInWeight").Value));
                        oDBDSDetail.SetValue("U_LoadingChg", (size - 1), Convert.ToString(oRS.Fields.Item("LoadingCharges").Value));
                        oDBDSDetail.SetValue("U_PlanQty", (size - 1), Convert.ToString(oRS.Fields.Item("PendingQty").Value));

                        //Thiru Addon Changes - Start - 02.07.2021
                        oDBDSDetail.SetValue("U_AVA_SalVal", (size - 1), Convert.ToString(oRS.Fields.Item("Sales value for this FY").Value));
                        string TaxCode = Convert.ToString(oRS.Fields.Item("TaxCode").Value);
                        double Rates = 0.00;
                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        {
                            Rates = Convert.ToDouble(objSBOAPI.Query_Execute("Select \"Rate\" from \"OSTC\" Where \"Code\"='" + TaxCode + "' "));
                        }
                        else
                        {
                            Rates = Convert.ToDouble(objSBOAPI.Query_Execute("Select Rate from OSTC Where Code='" + TaxCode + "' "));
                        }
                        double Tax_Amount = (LineTotal * Rates) / 100;
                        //oMatrix.Columns.Item("Col_TaxAmt").Cells.Item(size).Specific.Value = Tax_Amount;
                        oDBDSDetail.SetValue("U_AVA_TaxAmt", (size - 1), Convert.ToString(Tax_Amount));
                        string Freight_Type = oDBDSHeader.GetValue("U_FrightType", 0).ToString().Trim();
                        if(Freight_Type == "O")
                        {
                            double Tonnage = Convert.ToDouble(oDBDSDetail.GetValue("U_SOWeight", (size - 1)).ToString().Trim());
                            string Location_Code = ((IComboBox)objform.Items.Item("t_Location").Specific).Value.ToString().Trim();
                            double Rate = 0.00;
                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                            {
                                Rate = Convert.ToDouble(objSBOAPI.Query_Execute("Select \"U_Rate\" from \"@AIS_OLOC\" Where \"Code\" ='" + Location_Code + "'"));
                            }
                            else
                            {
                                Rate = Convert.ToDouble(objSBOAPI.Query_Execute("Select U_Rate from [@AIS_OLOC] Where Code ='" + Location_Code + "'"));
                            }
                            double Main_Loading_Charges = Tonnage * Rate;
                            double Sub_Loading_Charges = (Main_Loading_Charges * Rates) / 100;
                            double Loading_Charges = Main_Loading_Charges + Sub_Loading_Charges;
                            double Doc_Total = LineTotal + Tax_Amount + Loading_Charges;
                            //oMatrix.Columns.Item("Col_DocTot").Cells.Item(size).Specific.Value = Doc_Total;
                            oDBDSDetail.SetValue("U_AVA_DocTot", (size - 1), Convert.ToString(Doc_Total));
                        }
                        else if (Freight_Type == "H" || Freight_Type == "C")
                        {
                            double Doc_Total = LineTotal + Tax_Amount;
                            //oMatrix.Columns.Item("Col_DocTot").Cells.Item(size).Specific.Value = Doc_Total;
                            oDBDSDetail.SetValue("U_AVA_DocTot", (size - 1), Convert.ToString(Doc_Total));
                        }
                        
                        //Thiru Addon Changes - End - 02.07.2021

                        { ++size; }
                        oRS.MoveNext();
                        { ++Row; }
                    }
                    oDBDSHeader.SetValue("U_DocTotal", 0, Convert.ToString(DocTotal));
                    oDBDSHeader.SetValue("U_TotalBag", 0, Convert.ToString(TotalBag));
                    oDBDSHeader.SetValue("U_TotalTon", 0, Convert.ToString(TotalTon));
                    oDBDSHeader.SetValue("U_DiffWeight", 0, Convert.ToString(Convert.ToDouble(oDBDSHeader.GetValue("U_ActWeight", 0)) - Convert.ToDouble(oDBDSHeader.GetValue("U_TNetWeight", 0))));

                    


                    oMatrix.LoadFromDataSource();
                    LoadFreight();
                    LoadingCharges();
                    chckFlag = 0;
                    obj = true;
                }
                else
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Billing Details Not available for this Selection...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    obj = false;
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Load Bills Method Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                obj = false;
            }
            finally
            {
                oPickListNo = string.Empty;
                objform.Freeze(false);
            }
            return;
        }
        #endregion

        #region LoadFreight
        public void LoadFreight()
        {
            string Location_Code = ((IComboBox)objform.Items.Item("t_Location").Specific).Value.ToString().Trim();
            double Rate = 0.00;
            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
            {
                Rate = Convert.ToDouble(objSBOAPI.Query_Execute("Select \"U_Rate\" from \"@AIS_OLOC\" Where \"Code\" ='" + Location_Code + "'"));
            }
            else
            {
                Rate = Convert.ToDouble(objSBOAPI.Query_Execute("Select U_Rate from [@AIS_OLOC] Where Code ='" + Location_Code + "'"));
            }
                
            double Weight = 0.0;
            int visualRowCount = oMatrix.VisualRowCount;
            int Row = 1;
            while (Row <= visualRowCount)
            {
                string Weights = Convert.ToString(oMatrix.Columns.Item("Weight").Cells.Item(Row).Specific.Value);
                Weight += Convert.ToDouble(Weights);
                { ++Row; }
            }
            oDBDSHeader.SetValue("U_LFrieght", 0, Convert.ToString(Rate * Weight));
        }
        #endregion

        #region LoadingCharges
        public void LoadingCharges()
        {
            oMatrix.FlushToDataSource();
            double Coolie = 0.0;
            double Actual_Weight = 0.0;
            int Row = (oMatrix.VisualRowCount - 1);
            int RecordNumber = 0;
            while (RecordNumber <= Row)
            {
                if (oDBDSDetail.GetValue("U_PListNo", RecordNumber)  != string.Empty && oDBDSDetail.GetValue("U_TreeType", RecordNumber)  != "I")
                {
                    double LoadingChg = 0.0;
                    if (oDBDSDetail.GetValue("U_LoadingChg", RecordNumber) != string.Empty)
                        LoadingChg = Convert.ToDouble(oDBDSDetail.GetValue("U_LoadingChg", RecordNumber));
                    double PlanQty = 0.0;
                    if (oDBDSDetail.GetValue("U_PlanQty", RecordNumber) != string.Empty)
                        PlanQty = Convert.ToDouble(oDBDSDetail.GetValue("U_PlanQty", RecordNumber));
                    double SOWeight = 0.0;
                    if (oDBDSDetail.GetValue("U_SOWeight", RecordNumber) != string.Empty)
                        SOWeight = Convert.ToDouble(oDBDSDetail.GetValue("U_SOWeight", RecordNumber));
                    Actual_Weight += SOWeight;
                    Coolie += PlanQty * LoadingChg;
                }
                { ++RecordNumber; }
            }
            oDBDSHeader.SetValue("U_ActWeight", 0, Convert.ToString(Actual_Weight * 1000.0));
            oDBDSHeader.SetValue("U_DiffWeight", 0, Convert.ToString(Convert.ToDouble(oDBDSHeader.GetValue("U_ActWeight", 0)) - Convert.ToDouble(oDBDSHeader.GetValue("U_TNetWeight", 0))));
            oDBDSHeader.SetValue("U_Coolie", 0, Convert.ToString(Coolie));
            oMatrix.LoadFromDataSource();
        }
        #endregion

        #region ValidateAll
        public bool ValidateAll()
        {
            bool flag;
            try
            {
                string Temp1 = string.Empty;
                if (objform.Items.Item("c_Unit").Specific.Value.ToString().Trim() == string.Empty)
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Unit Shouldn't be Empty...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    flag = false;
                    return flag;
                }
                else
                {
                    string Temp2 = string.Empty;
                    string Vehicle_No = objform.Items.Item("t_VNo").Specific.Value.ToString().Trim();
                    if (Vehicle_No == string.Empty)
                    {
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Vehicle No  Shouldn't be Empty...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        flag = false;
                        return flag;
                    }
                    else
                    {
                        string Location = string.Empty;
                        ComboBox Location_comboBox1 = (ComboBox)objform.Items.Item("t_Location").Specific;
                        if (Location_comboBox1.Selected != null)
                            Location = Location_comboBox1.Selected.Value.ToString().Trim();
                        if (Location == string.Empty)
                        {
                            objSBOAPI.SBO_Appln.StatusBar.SetText("Location Shouldn't be Empty...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            flag = false;
                            return flag;
                        }
                        else
                        {
                            string Vehicle_Type = string.Empty;
                            if (((IComboBox)objform.Items.Item("c_TVehicle").Specific).Selected != null)
                                Vehicle_Type = ((IComboBox)objform.Items.Item("c_TVehicle").Specific).Selected.Value.ToString().Trim();
                            if (Vehicle_Type == string.Empty)
                            {
                                objSBOAPI.SBO_Appln.StatusBar.SetText("Vehicle Type Shouldn't be Empty...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                flag = false;
                                return flag;
                            }
                            else if (objform.Items.Item("t_DName").Specific.Value.ToString().Trim() == string.Empty)
                            {
                                objSBOAPI.SBO_Appln.StatusBar.SetText("Driver Name Shouldn't be Empty...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                flag = false;
                                return flag;
                            }
                            else
                            {
                                //Thiru - Lorry In Time , Lorry Out Time Validations Remove - Start - 20210804

                                //string Lorry_In_Time = Convert.ToString(objform.Items.Item("t_LorryIn").Specific.Value);
                                //if (Lorry_In_Time == string.Empty)
                                //{
                                //    objSBOAPI.SBO_Appln.StatusBar.SetText("Lorry In Time Shouldn't be Empty...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                //    flag = false;
                                //    return flag;
                                //}
                                //else
                                //{
                                //    string Lorry_Out_Time = Convert.ToString(objform.Items.Item("t_LorryOut").Specific.Value);
                                //    if (Lorry_Out_Time == string.Empty)
                                //    {
                                //        objSBOAPI.SBO_Appln.StatusBar.SetText("Lorry Out Time Shouldn't be Empty...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                //        flag = false;
                                //        return flag;
                                //    }
                                //    else
                                //    {
                                        ComboBox Approval_comboBox2 = (ComboBox)objform.Items.Item("c_Appr").Specific;
                                        string Approval = string.Empty;
                                        if (Approval_comboBox2.Selected != null)
                                            Approval = Approval_comboBox2.Selected.Value.ToString().Trim();
                                        if (Approval == string.Empty)
                                        {
                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Approval Shouldn't be Empty...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                            flag = false;
                                            return flag;
                                        }
                                        else if (Convert.ToDouble(objform.Items.Item("t_LFrght").Specific.Value) == 0.0)
                                        {
                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Lorry Frieght Shouldn't be Empty...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                            flag = false;
                                            return flag;
                                        }
                                        else
                                        {
                                            double Diff_Weight = 0.0;
                                            Diff_Weight = Convert.ToDouble(objform.Items.Item("t_DWeight").Specific.Value);
                                            if (oMatrix.VisualRowCount == 0)
                                            {
                                                objSBOAPI.SBO_Appln.StatusBar.SetText("Empty Document Shouldn't be Added...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                                flag = false;
                                                return flag;
                                            }
                                            else
                                                flag = true;
                                        }
                                //}
                                //}
                                //Thiru - Lorry In Time , Lorry Out Time Validations Remove - End - 20210804

                            }
                        }
                    }
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

        #region InitForm
        public void InitForm()
        {
            try
            {
                objform.Freeze(true);
                object Doc_Num = null;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    Doc_Num = RuntimeHelpers.GetObjectValue(objSBOAPI.Query_Execute("select Top 1 \"DocNum\"  from \"@AIS_OPDE\" order by \"DocNum\"  desc  "));
                }
                else
                {
                    Doc_Num = RuntimeHelpers.GetObjectValue(objSBOAPI.Query_Execute("select Top 1 DocNum  from \"@AIS_oPDE\" order by DocNum  desc  "));
                }
                    
                object Doc_Number = RuntimeHelpers.GetObjectValue(Doc_Num);
                oMatrix.Clear();
                oDBDSDetail.Clear();
                if (Convert.ToDouble(Doc_Number) == 0.0)
                    return;
                objform.Mode = BoFormMode.fm_FIND_MODE;
                objform.Items.Item("9").Specific.Value = RuntimeHelpers.GetObjectValue(Doc_Number);
                
                objform.Items.Item("1").Click(BoCellClickType.ct_Regular);
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

        #region PostingARInvoice
        public bool PostingARInvoice()
        {
            bool flag1;
            try
            {
                objform.Freeze(true);
                string Temp1 = string.Empty;
                string Temp2 = string.Empty;
                string series = string.Empty;
                string DocEntry = oDBDSHeader.GetValue("DocEntry", 0).ToString().Trim();
                string DocNum = oDBDSHeader.GetValue("DocNum", 0).ToString().Trim();
                ComboBox series_comboBox1 = (ComboBox)objform.Items.Item("c_series").Specific;
                if (series_comboBox1.Selected != null)
                    series = series_comboBox1.Selected.Value.ToString().Trim();
                ComboBox Location_comboBox2 = (ComboBox)objform.Items.Item("t_Location").Specific;
                string Location_Code = string.Empty;
                if (Location_comboBox2.Selected != null)
                    Location_Code = Location_comboBox2.Selected.Value.ToString().Trim();
                double Rate = 0.00;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    Rate = Convert.ToDouble(objSBOAPI.Query_Execute("Select \"U_Rate\" from \"@AIS_OLOC\" Where \"Code\" ='" + Location_Code + "'"));
                }
                else
                {
                    Rate = Convert.ToDouble(objSBOAPI.Query_Execute("Select \"U_Rate\" from \"@AIS_OLOC\" Where \"Code\" ='" + Location_Code + "'"));
                }
                    
                string FrightType = oDBDSHeader.GetValue("U_FrightType", 0).ToString().Trim();
                ArrayList arrayList = new ArrayList();
                Recordset Orec1 = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string QueryStr1 = null;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    QueryStr1 = "SELECT   T1.\"U_BDocEntry\", T1.\"U_CusCode\" ,T1.\"U_TaxCode\",T1.\"U_WhsCode\"   FROM \"@AIS_PDE1\" T1  Where T1.\"DocEntry\" ='" + DocEntry + "' and IFNULL(T1.\"U_CusCode\",'')!='' group by T1.\"U_BDocEntry\", T1.\"U_CusCode\" ,T1.\"U_TaxCode\",T1.\"U_WhsCode\"";
                }
                else
                {
                    QueryStr1 = "SELECT   T1.U_BDocEntry, T1.U_CusCode ,T1.U_TaxCode,T1.U_WhsCode   FROM [@AIS_PDE1] T1  Where T1.DocEntry ='" + DocEntry + "' and isnull(T1.U_CusCode,'')!='' group by T1.U_BDocEntry, T1.U_CusCode ,T1.U_TaxCode,T1.U_WhsCode";
                }
                    
                Orec1.DoQuery(QueryStr1);
                if (Orec1.RecordCount > 0)
                    Orec1.MoveFirst();
                int recordCount1 = Orec1.RecordCount;
                int Rows = 1;
                while (Rows <= recordCount1)
                {
                    string Pick_List_No = string.Empty;
                    string BDocEntry = Orec1.Fields.Item("U_BDocEntry").Value.ToString().Trim();
                    string TaxCode = Orec1.Fields.Item("U_TaxCode").Value.ToString().Trim();
                    string CusCode = Orec1.Fields.Item("U_CusCode").Value.ToString().Trim();
                    string QueryStr2 = null;
                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        QueryStr2 = "CALL \"@AIS_BillDelivery_TCSDetails\"('" + CusCode + "','" + TaxCode + "')";
                    }
                    else
                    {
                        QueryStr2 = "Exec dbo.[@AIS_BillDelivery_TCSDetails]'" + CusCode + "','" + TaxCode + "'";
                    }
                        
                    Recordset businessObject2 = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    businessObject2.DoQuery(QueryStr2);
                    double TransValue = 0.0;
                    if (businessObject2.RecordCount > 0)
                        TransValue = Convert.ToDouble(businessObject2.Fields.Item("TransValue").Value);
                    double num4 = 0.0;
                    string WhsCode = Orec1.Fields.Item("U_WhsCode").Value.ToString().Trim();
                    Documents oInvoices = (Documents)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.oInvoices);
                    oInvoices.CardCode = Orec1.Fields.Item("U_CusCode").Value.ToString().Trim();
                    oInvoices.TaxDate = DateTime.Now;
                    oInvoices.DocDueDate = DateTime.Now;
                    oInvoices.DocDate = DateTime.Now;
                    oInvoices.Comments = "Invoice Document Posted from Bill Delivery Doc No. " + DocNum + "/" + series;
                    //Thiru - 26.03.2021
                    oInvoices.UserFields.Fields.Item("U_AV_TransName").Value = objform.Items.Item("t_DName").Specific.Value;
                    oInvoices.UserFields.Fields.Item("U_AV_VehNo").Value = objform.Items.Item("t_VNo").Specific.Value;
                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        oInvoices.UserFields.Fields.Item("U_AV_VehType").Value = objSBOAPI.Query_Execute("Select \"U_TVehicle\" from \"@AIS_OVEH\" WHERE \"Code\" = '" + objform.Items.Item("c_TVehicle").Specific.Value + "'");
                    }
                    else
                    {
                        oInvoices.UserFields.Fields.Item("U_AV_VehType").Value = objSBOAPI.Query_Execute("Select U_TVehicle from [@AIS_OVEH] WHERE Code = '" + objform.Items.Item("c_TVehicle").Specific.Value + "'");
                    }
                    //Thiru - 26.03.2021

                    oInvoices.UserFields.Fields.Item("U_BaseEntry").Value = DocEntry;
                    oInvoices.UserFields.Fields.Item("U_BaseNum").Value = DocNum;
                    oInvoices.UserFields.Fields.Item("U_BaseObject").Value = "OCDL";
                    int visualRowCount1 = oMatrix.VisualRowCount;
                    int RowCount = 1;
                    while (RowCount <= visualRowCount1)
                    {
                        if (oMatrix.Columns.Item("V_2").Cells.Item(RowCount).Specific.Value.ToString().Trim() != string.Empty)
                        {
                            string BEn = oMatrix.Columns.Item("BEn").Cells.Item(RowCount).Specific.Value.ToString().Trim();
                            string TreeType = oMatrix.Columns.Item("TreeType").Cells.Item(RowCount).Specific.Value.ToString().Trim();
                            Pick_List_No = oMatrix.Columns.Item("PLN").Cells.Item(RowCount).Specific.Value.ToString().Trim();
                            string Percentage = null;
                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                            {
                                Percentage = Convert.ToString(objSBOAPI.Query_Execute("Select IFNULL(\"U_Percentage\",0.0)  from \"OPKL\" Where \"AbsEntry\" ='" + Pick_List_No + "'"));
                            }
                            else
                            {
                                Percentage = Convert.ToString(objSBOAPI.Query_Execute("Select isnull(U_Percentage,0.0)  from OPKL Where AbsEntry ='" + Pick_List_No + "'"));
                            }
                                
                            string TargetEnt = oMatrix.Columns.Item("TargetEnt").Cells.Item(RowCount).Specific.Value.ToString().Trim();
                            string TaxCodes = oMatrix.Columns.Item("TaxCode").Cells.Item(RowCount).Specific.Value.ToString().Trim();
                            string WhsCodes = oMatrix.Columns.Item("WhsCode").Cells.Item(RowCount).Specific.Value.ToString().Trim();
                            if (TargetEnt == string.Empty && BDocEntry == BEn && TaxCode == TaxCodes && WhsCode == WhsCodes)
                            {
                                string pItemCode = oMatrix.Columns.Item("ItemCode").Cells.Item(RowCount).Specific.Value.ToString().Trim();
                                string TaxCategory = null;
                                string BranchCode = null;
                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                {
                                    TaxCategory = Convert.ToString(objSBOAPI.Query_Execute("select \"U_TCategory\" from \"OITM\" where \"ItemCode\" ='" + pItemCode + "'"));
                                    BranchCode = Convert.ToString(objSBOAPI.Query_Execute("select \"U_BranchCode\"  from \"@AIS_BRN1\"  where \"U_WhsCode\" ='" + WhsCode + "'"));
                                }
                                else
                                {
                                    TaxCategory = Convert.ToString(objSBOAPI.Query_Execute("select U_TCategory from OITM where ItemCode ='" + pItemCode + "'"));
                                    BranchCode = Convert.ToString(objSBOAPI.Query_Execute("select U_BranchCode  from [@AIS_BRN1]  where U_WhsCode ='" + WhsCode + "'"));
                                }
                                    
                                string Series = string.Empty;
                                if (TaxCategory == "Taxable")
                                {
                                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                    {
                                        Series = Convert.ToString(objSBOAPI.Query_Execute("select \"Series\" from \"NNM1\" T0 inner join \"@AIS_OUNT\" T1 on T1.\"U_ARTax\" =T0.\"SeriesName\" where \"ObjectCode\" ='13' and T1.\"Code\"='" + BranchCode + "' and T0.\"DocSubType\" ='GA' and T0.\"Locked\" ='N'"));
                                    }
                                    else
                                    {
                                        Series = Convert.ToString(objSBOAPI.Query_Execute("select Series from NNM1 T0 inner join [@AIS_OUNT] T1 on T1.U_ARTax =T0.SeriesName where ObjectCode ='13' and T1.Code='" + BranchCode + "' and T0.DocSubType ='GA' and T0.Locked ='N'"));
                                    }
                                }
                                    
                                        
                                else if (TaxCategory == "Zero tax")
                                {
                                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                    {
                                        Series = Convert.ToString(objSBOAPI.Query_Execute("select \"Series\" from \"NNM1\" T0 inner join \"@AIS_OUNT\" T1 on T1.\"U_ARNTax\" =T0.\"SeriesName\" where \"ObjectCode\" ='13' and T1.\"Code\"='" + BranchCode + "' and T0.\"DocSubType\" ='GA' and T0.\"Locked\" ='N'"));
                                    }
                                    else
                                    {
                                        Series = Convert.ToString(objSBOAPI.Query_Execute("select Series from NNM1 T0 inner join [@AIS_OUNT] T1 on T1.U_ARNTax =T0.SeriesName where ObjectCode ='13' and T1.Code='" + BranchCode + "' and T0.DocSubType ='GA' and T0.Locked ='N'"));
                                    }
                                }
                                    
                                else if (TaxCategory == "Non Tax")
                                {
                                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                    {
                                        Series = Convert.ToString(objSBOAPI.Query_Execute("select \"Series\" from \"NNM1\" T0 inner join \"@AIS_OUNT\" T1 on T1.\"U_ARExmpt\" =T0.\"SeriesName\" where \"ObjectCode\" ='13' and T1.\"Code\"='" + BranchCode + "' and T0.\"DocSubType\" ='--' and T0.\"Locked\" ='N'"));
                                    }
                                    else
                                    {
                                        Series = Convert.ToString(objSBOAPI.Query_Execute("select Series from NNM1 T0 inner join [@AIS_OUNT] T1 on T1.U_ARExmpt =T0.SeriesName where ObjectCode ='13' and T1.Code='" + BranchCode + "' and T0.DocSubType ='--' and T0.Locked ='N'"));
                                    }
                                }
                                    
                                if (Series == string.Empty)
                                {
                                    objSBOAPI.SBO_Appln.StatusBar.SetText("For " + BranchCode + "-Unit, Numbering Series should be assigned...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    flag1 = false;
                                    return flag1;
                                }
                                else
                                {
                                    oInvoices.Series = Convert.ToInt32(Series);
                                    string BaseLine = oMatrix.Columns.Item("BaseLine").Cells.Item(RowCount).Specific.Value.ToString().Trim();
                                    string WarehouseCode = oMatrix.Columns.Item("WhsCode").Cells.Item(RowCount).Specific.Value.ToString().Trim();
                                    string ChapterID = null;
                                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                    {
                                        ChapterID = Convert.ToString(objSBOAPI.Query_Execute("SELECT IFNULL(\"ChapterID\",'0')  FROM \"OITM\" Where \"ItemCode\" ='" + pItemCode + "'"));
                                    }
                                    else
                                    {
                                        ChapterID = Convert.ToString(objSBOAPI.Query_Execute("SELECT isnull(ChapterID,'')  FROM OITM Where ItemCode ='" + pItemCode + "'"));
                                    }
                                        
                                    oInvoices.Lines.BaseType = 17;
                                    oInvoices.Lines.BaseEntry = Convert.ToInt32(BDocEntry);
                                    oInvoices.Lines.BaseLine = Convert.ToInt32(BaseLine);
                                    oInvoices.Lines.ItemCode = pItemCode;
                                    oInvoices.Lines.HSNEntry = Convert.ToInt32(ChapterID);
                                    oInvoices.Lines.TaxCode = oMatrix.Columns.Item("TaxCode").Cells.Item(RowCount).Specific.Value.ToString().Trim();
                                    oInvoices.Lines.WarehouseCode = WarehouseCode;
                                    double Qty = Convert.ToDouble(oMatrix.Columns.Item("Qty").Cells.Item(RowCount).Specific.Value.ToString().Trim());
                                    double Price = Convert.ToDouble(oMatrix.Columns.Item("Price").Cells.Item(RowCount).Specific.Value.ToString().Trim());
                                    oInvoices.Lines.Quantity = Qty;
                                    oInvoices.Lines.UnitPrice = Price;
                                    double num7 = Qty * Price;
                                    string Rates = null;
                                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                    {
                                        Rates = Convert.ToString(objSBOAPI.Query_Execute("Select \"Rate\" from \"OSTC\" Where \"Code\"='" + TaxCode + "' "));
                                    }
                                    else
                                    {
                                        Rates = Convert.ToString(objSBOAPI.Query_Execute("Select Rate from OSTC Where Code='" + TaxCode + "' "));
                                    }
                                        
                                    if (Rates == string.Empty)
                                        Rates = Convert.ToString(0);
                                    double num8 = num7 * Convert.ToDouble(Rates) / 100.0;
                                    double num9 = num7 + num8;
                                    TransValue += num9;
                                    num4 += num9;
                                    double num10 = Qty * Convert.ToDouble(Percentage) / 100.0;
                                    if (TreeType != "S")
                                    {
                                        if (FindIsBatchItem(pItemCode) == "Y")
                                        {
                                            bool flag2 = false;
                                            double PendingQuantity = 0.0;
                                            Recordset recordset1 = null;
                                            recordset1=(Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                            {
                                                recordset1.DoQuery("Select \"BATCHNO\",\"PENDINGQTY\" From \"TEMP_PKL1\" Where \"ABSENTRY\"='" + Pick_List_No + "' And \"[LINENO]\" ='" + BaseLine + "'");
                                            }
                                            else
                                            {
                                                recordset1.DoQuery("Select [BatchNo],[PendingQty] From [Temp_PKL1] Where [AbsEntry]='" + Pick_List_No + "' And [LineNo] ='" + BaseLine + "'");
                                            }
                                                
                                            if (recordset1.RecordCount > 0)
                                                recordset1.MoveFirst();
                                            int RecordCount = (recordset1.RecordCount - 1);
                                            int Row = 0;
                                            while (Row <= RecordCount)
                                            {
                                                double PendingQty = Convert.ToDouble(Convert.ToString(recordset1.Fields.Item("PendingQty").Value).Trim());
                                                string BatchNo = Convert.ToString(recordset1.Fields.Item("BatchNo").Value).Trim();
                                                flag2 = true;
                                                oInvoices.Lines.BatchNumbers.BatchNumber = BatchNo;
                                                PendingQuantity += PendingQty;
                                                oInvoices.Lines.BatchNumbers.Quantity = PendingQty;
                                                oInvoices.Lines.BatchNumbers.Add();
                                                if (Qty > PendingQuantity)
                                                {
                                                    recordset1.MoveNext();
                                                    { ++Row; }
                                                }
                                                else
                                                    break;
                                            }
                                            double num15 = Qty - PendingQuantity;
                                            if (num15 > 0.0)
                                            {
                                                Recordset recordset2 = null;
                                                recordset2= (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                                {
                                                    recordset2.DoQuery("CALL \"AIS_BillDelivery_GetBatchDetails\"('" + pItemCode + "','" + Convert.ToString(num15) + "','" + WarehouseCode + "') ");
                                                }
                                                else
                                                {
                                                    recordset2.DoQuery("Exec dbo.[@AIS_BillDelivery_GetBatchDetails]'" + pItemCode + "','" + Convert.ToString(num15) + "','" + WarehouseCode + "' ");
                                                }
                                                    
                                                if (recordset2.RecordCount > 0)
                                                {
                                                    recordset2.MoveFirst();
                                                    int num14 = (recordset2.RecordCount - 1);
                                                    int num16 = 0;
                                                    while (num16 <= num14)
                                                    {
                                                        double BatchQty = Convert.ToDouble(Convert.ToString(recordset2.Fields.Item("BatchQty").Value).Trim());
                                                        string BatchNum = Convert.ToString(recordset2.Fields.Item("BatchNum").Value).Trim();
                                                        flag2 = true;
                                                        oInvoices.Lines.BatchNumbers.BatchNumber = BatchNum;
                                                        oInvoices.Lines.BatchNumbers.Quantity = BatchQty;
                                                        oInvoices.Lines.BatchNumbers.Add();
                                                        PendingQuantity += BatchQty;
                                                        if (Qty > PendingQuantity)
                                                        {
                                                            recordset2.MoveNext();
                                                            checked { ++num16; }
                                                        }
                                                        else
                                                            break;
                                                    }
                                                }
                                            }
                                            if (!flag2)
                                            {
                                                objSBOAPI.SBO_Appln.StatusBar.SetText("Batch No Not Allocated for Item " + pItemCode, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                                flag1 = false;
                                                return flag1;
                                            }
                                            else if (Qty > PendingQuantity)
                                            {
                                                objSBOAPI.SBO_Appln.StatusBar.SetText("Stock isn't available for Item " + pItemCode, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                                flag1 = false;
                                                return flag1;
                                            }
                                        }
                                        if (FindIsSerialItem(pItemCode) == "Y")
                                        {
                                            if (Convert.ToDouble(Percentage) > 0.0)
                                            {
                                                Recordset recordset = null;
                                                recordset = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                                {
                                                    recordset.DoQuery("Select \"ABSENTRY\" From \"TEMP_PKL1\" Where \"ABSENTRY\"='" + Pick_List_No + "'");
                                                }
                                                else
                                                {
                                                    recordset.DoQuery("Select AbsEntry From [Temp_PKL1] Where [AbsEntry]='" + Pick_List_No + "'");
                                                }
                                                    
                                                if (Convert.ToString(Math.Round(Convert.ToDouble(Convert.ToString(objSBOAPI.Query_Execute("Select Sum(T1.PickQtty) * isnull(T0.U_Percentage,0.0)/100  from OPKL T0 inner join PKL1  T1 on t1.AbsEntry =T0.AbsEntry Where T0.AbsEntry='" + Pick_List_No + "' Group by T0.U_Percentage"))), 0)) != recordset.RecordCount.ToString())
                                                {
                                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Pick List Percentage isn't equal to WeighBridge Scanning Percentage...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                                    flag1 = false;
                                                    return flag1;
                                                }
                                            }
                                            bool flag2 = false;
                                            if (Convert.ToDouble(Percentage) > 0.0)
                                            {
                                                Recordset recordset = null;
                                                recordset= (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                                {
                                                    recordset.DoQuery("Select \"BATCHNO\",\"PENDINGQTY\" From \"TEMP_PKL1\" Where   \"ABSENTRY\"='" + Pick_List_No + "' And \"[LINENO]\" ='" + BaseLine + "' Order by \"BATCHNO\"");
                                                }
                                                else
                                                {
                                                    recordset.DoQuery("Select [BatchNo],[PendingQty] From [Temp_PKL1] Where   [AbsEntry]='" + Pick_List_No + "' And [LineNo] ='" + BaseLine + "' Order by BatchNo");
                                                }
                                                    
                                                if (recordset.RecordCount > 0)
                                                    recordset.MoveFirst();
                                                int num11 = 0;
                                                double num12 = 0.0;
                                                int num13 = (recordset.RecordCount - 1);
                                                int num14 = 0;
                                                while (num14 <= num13)
                                                {
                                                    string PendingQty = Convert.ToDouble(Convert.ToString(recordset.Fields.Item("PendingQty").Value).Trim());
                                                    string BatchNo = Convert.ToString(recordset.Fields.Item("BatchNo").Value).Trim();
                                                    flag2 = true;
                                                    oInvoices.Lines.SerialNumbers.InternalSerialNumber = BatchNo;
                                                    oInvoices.Lines.SerialNumbers.Quantity = 1.0;
                                                    ++num12;
                                                    num11 = ((int)Math.Round((Convert.ToDouble(BatchNo) + 1.0)));
                                                    oInvoices.Lines.SerialNumbers.Add();
                                                    recordset.MoveNext();
                                                    { ++num14; }
                                                }
                                                int num15 = ((int)Math.Round((num12 + 1.0)));
                                                int num16 = ((int)Math.Round(Qty));
                                                int num17 = num15;
                                                while (num17 <= num16)
                                                {
                                                    oInvoices.Lines.SerialNumbers.InternalSerialNumber = Convert.ToString(num11);
                                                    oInvoices.Lines.SerialNumbers.Quantity = 1.0;
                                                    ++num12;
                                                    { ++num11; }
                                                    if (num12 < Qty)
                                                    {
                                                        oInvoices.Lines.SerialNumbers.Add();
                                                        { ++num17; }
                                                    }
                                                    else
                                                        break;
                                                }
                                                if (!flag2)
                                                {
                                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Batch No Not Allocated for Item " + pItemCode, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                                    flag1 = false;
                                                    return flag1;
                                                }
                                            }
                                            else
                                            {
                                                int num11 = 0;
                                                Recordset Orec4 = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                                string QueryStr3 = null;
                                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                                {
                                                    QueryStr3 = "CALL \"@AIS_BillDelivery_GetSerialNumbers\"('" + pItemCode + "','" + WarehouseCode + "')";
                                                }
                                                else
                                                {
                                                    QueryStr3 = "Exec dbo.[@AIS_BillDelivery_GetSerialNumbers]'" + pItemCode + "','" + WarehouseCode + "'";
                                                }
                                                    
                                                Orec4.DoQuery(QueryStr3);
                                                if (Orec4.RecordCount > 0)
                                                {
                                                    Orec4.MoveFirst();
                                                    int recordCount2 = Orec4.RecordCount;
                                                    int Row = 1;
                                                    while (Row <= recordCount2)
                                                    {
                                                        oInvoices.Lines.SerialNumbers.InternalSerialNumber = Orec4.Fields.Item("IntrSerial").Value.ToString().Trim();
                                                        oInvoices.Lines.SerialNumbers.Quantity = 1.0;
                                                        oInvoices.Lines.SerialNumbers.Add();
                                                        { ++num11; }
                                                        if ((double)num11 < Qty)
                                                        {
                                                            Orec4.MoveNext();
                                                            { ++Row; }
                                                        }
                                                        else
                                                            break;
                                                    }
                                                    if ((double)num11 < Qty)
                                                    {
                                                        objSBOAPI.SBO_Appln.StatusBar.SetText("Stock isn't available for serial Item : " + pItemCode, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                                        flag1 = false;
                                                        return flag1;
                                                    }
                                                }
                                                else
                                                {
                                                   objSBOAPI.SBO_Appln.StatusBar.SetText("Stock isn't available for serial Item : " + pItemCode, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                                    flag1 = false;
                                                    return flag1;
                                                }
                                            }
                                        }
                                    }
                                    oInvoices.Lines.Add();
                                }
                            }
                        }
                        { ++RowCount; }
                    }
                    if (FrightType == "O" && !arrayList.Contains(Orec1.Fields.Item("U_CusCode").Value.ToString().Trim()))
                    {
                        string QueryStr3 = null;
                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        {
                            QueryStr3 = "SELECT   T1.\"U_CusCode\",Sum(T1.\"U_SOWeight\") AS \"Tonnage\" ,Min(T2.\"Rate\" ) AS \"TaxRate\"   FROM \"@AIS_PDE1\" T1 inner join \"OSTC\" T2 on T2.\"Code\"=T1.\"U_TaxCode\" Where T1.\"DocEntry\" ='" + DocEntry + "' and T1.\"U_CusCode\" ='" + Orec1.Fields.Item("U_CusCode").Value.ToString().Trim() + "' group by T1.\"U_CusCode\"-- ,T1.\"U_TaxCode\" ";
                        }
                        else
                        {
                            QueryStr3 = "SELECT   T1.U_CusCode,Sum(T1.U_SOWeight) AS Tonnage ,Min(T2.Rate ) as TaxRate   FROM [@AIS_PDE1] T1 inner join OSTC T2 on T2.Code=T1.U_TaxCode Where T1.DocEntry ='" + DocEntry + "' and T1.U_CusCode ='" + Orec1.Fields.Item("U_CusCode").Value.ToString().Trim() + "' group by T1.U_CusCode-- ,T1.U_TaxCode ";
                        }

                        Recordset businessObject4 = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        businessObject4.DoQuery(QueryStr3);
                        double LineTotal = 0.0;
                        if (businessObject4.RecordCount > 0)
                        {
                            businessObject4.MoveFirst();
                            double Tonnage = Convert.ToDouble(businessObject4.Fields.Item("Tonnage").Value);
                            double TaxRate = Convert.ToDouble(businessObject4.Fields.Item("TaxRate").Value);
                            string Rates = null;
                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                            {
                                Rates = Convert.ToString(objSBOAPI.Query_Execute("Select \"Rate\" from \"OSTC\" Where \"Code\"='" + TaxCode + "' "));
                            }
                            else
                            {
                                Rates = Convert.ToString(objSBOAPI.Query_Execute("Select Rate from OSTC Where Code='" + TaxCode + "' "));
                            }

                            if (Rates == string.Empty)
                                Rates = Convert.ToString(0);
                            if (TaxRate == Convert.ToDouble(Rates))
                                LineTotal = Tonnage * Rate;
                        }
                        if (LineTotal > 0.0)
                        {
                            string ExpnsCode = null;
                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                            {
                                //ExpnsCode = Convert.ToString(objSBOAPI.Query_Execute("Select \"ExpnsCode\" from \"OEXD\" Where IFNULL(\"U_IsTCS\",'N')='N'"));
                                ExpnsCode = Convert.ToString(objSBOAPI.Query_Execute("Select \"ExpnsCode\" from \"OEXD\" Where IFNULL(\"U_AVA_FREIGHT\",'N')='Y'"));
                            }
                            else
                            {
                                //ExpnsCode = Convert.ToString(objSBOAPI.Query_Execute("Select ExpnsCode from OEXD Where isnull(U_IsTCS,'N')='N'"));
                                ExpnsCode = Convert.ToString(objSBOAPI.Query_Execute("Select ExpnsCode from OEXD Where isnull(U_AVA_FREIGHT,'N')='Y'"));
                            }

                            if (ExpnsCode == string.Empty)
                            {
                                objSBOAPI.SBO_Appln.StatusBar.SetText("Expnse Code shouldn't be available for " + TaxCode + " Tax", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                flag1 = false;
                                return flag1;
                            }
                            else if (FrightType == "O")
                            {
                                oInvoices.Expenses.ExpenseCode = Convert.ToInt32(ExpnsCode);
                                oInvoices.Expenses.TaxCode = TaxCode;
                                oInvoices.Expenses.LineTotal = Math.Round(LineTotal, 2);
                                string Rate1 = null;
                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                {
                                    Rate1 = Convert.ToString(objSBOAPI.Query_Execute("Select \"Rate\" from \"OSTC\" Where \"Code\"='" + TaxCode + "' "));
                                }
                                else
                                {
                                    Rate1 = Convert.ToString(objSBOAPI.Query_Execute("Select Rate from OSTC Where Code='" + TaxCode + "' "));
                                }

                                if (Rate1 == string.Empty)
                                    Rate1 = Convert.ToString(0);
                                string str7 = Convert.ToString(Math.Round(LineTotal, 2) * Convert.ToDouble(Rate1) / 100.0);
                                TransValue = TransValue + Convert.ToDouble(str7) + Math.Round(LineTotal, 2);
                                num4 = num4 + Math.Round(LineTotal, 2) + Convert.ToDouble(str7);
                                oInvoices.Expenses.Add();
                                arrayList.Add(Orec1.Fields.Item("U_CusCode").Value.ToString().Trim());
                            }
                            else if (FrightType == "H")
                            {
                                oInvoices.Expenses.ExpenseCode = Convert.ToInt32(ExpnsCode);
                                string FreightTax = null;
                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                {
                                    FreightTax = Convert.ToString(objSBOAPI.Query_Execute("Select \"U_FreightTax\" from \"OSTC\" Where \"Code\"='" + TaxCode + "' "));
                                }
                                else
                                {
                                    FreightTax = Convert.ToString(objSBOAPI.Query_Execute("Select U_FreightTax from OSTC Where Code='" + TaxCode + "' "));
                                }

                                if (FreightTax == string.Empty)
                                {
                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Hired Freight Tax should be available for " + TaxCode + " Tax", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    flag1 = false;
                                    return flag1;
                                }
                                else
                                {
                                    oInvoices.Expenses.TaxCode = FreightTax;
                                    oInvoices.Expenses.LineTotal = Math.Round(LineTotal, 2);
                                    TransValue += Math.Round(LineTotal, 2);
                                    num4 += Math.Round(LineTotal, 2);
                                    oInvoices.Expenses.Add();
                                    arrayList.Add(Orec1.Fields.Item("U_CusCode").Value.ToString().Trim());
                                }
                            }
                        }
                    }
                    if (TransValue > 5000000.0)
                    {
                        string ExpnsCode = null;
                        string TCSDetails_QueryStr3 = null;
                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        {
                            ExpnsCode = Convert.ToString(objSBOAPI.Query_Execute("Select \"ExpnsCode\" from \"OEXD\" Where \"U_IsTCS\"='Y' "));
                            TCSDetails_QueryStr3 = "CALL \"@AIS_BillDelivery_TCSDetails\"('" + CusCode + "','" + TaxCode + "')";
                        }
                        else
                        {
                            ExpnsCode = Convert.ToString(objSBOAPI.Query_Execute("Select ExpnsCode from OEXD Where \"U_IsTCS\"='Y' "));
                            TCSDetails_QueryStr3 = "Exec dbo.[@AIS_BillDelivery_TCSDetails]'" + CusCode + "','" + TaxCode + "'";
                        }
                            
                        Recordset Orec4 = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        Orec4.DoQuery(TCSDetails_QueryStr3);
                        if (Orec4.RecordCount > 0)
                        {
                            //Thiru Addon Changes - Start - 02.07.2021 ( AR Invoice Freight Calculation Changes )
                            double Sales_Value_for_this_FY = Convert.ToDouble(objSBOAPI.Query_Execute("Select Distinct IFNULL(B.\"U_AVA_SalVal\",'0.00') AS \"Sales Value for this FY\" From \"@AIS_OPDE\" A INNER JOIN \"@AIS_PDE1\" B ON A.\"DocEntry\"=B.\"DocEntry\" WHERE A.\"DocEntry\"='" + DocEntry + "' and B.\"U_BDocEntry\" = '" + BDocEntry + "' and B.\"U_CusCode\"='"+CusCode+"' and B.\"U_TaxCode\"='"+TaxCode+"'and B.\"U_WhsCode\"='"+WhsCode+"'"));
                            double DocTotal = Convert.ToDouble(objSBOAPI.Query_Execute("Select IFNULL(SUM(B.\"U_AVA_DocTot\"),'0.00') AS \"DocTotal\" From \"@AIS_OPDE\" A INNER JOIN \"@AIS_PDE1\" B ON A.\"DocEntry\"=B.\"DocEntry\" WHERE A.\"DocEntry\"='" + DocEntry + "' and B.\"U_BDocEntry\" = '" + BDocEntry + "'and B.\"U_CusCode\"='" + CusCode + "' and B.\"U_TaxCode\"='" + TaxCode + "'and B.\"U_WhsCode\"='" + WhsCode + "'"));
                            double Total = Sales_Value_for_this_FY + DocTotal;
                            double Customer_TCS_Value = Convert.ToDouble(objSBOAPI.Query_Execute("Select \"U_AVA_CusTCSVal\" From \"OADM\""));
                            string Last_2_FY_IT_Return_Filed = objSBOAPI.Query_Execute("Select \"U_AVA_ITReturnFiled\" From \"OCRD\" Where \"CardCode\" = '" + CusCode + "'");
                            string Sum_Of_TDS_TCS_Greater_than_50000_for_Last_2_FY = objSBOAPI.Query_Execute("Select \"U_AVA_SumTDSTCS\" From \"OCRD\" Where \"CardCode\" = '" + CusCode + "'");
                            string Customer_Deducting_TDS = objSBOAPI.Query_Execute("Select \"U_AVA_CusTDSDed\" From \"OCRD\" Where \"CardCode\" = '" + CusCode + "'");
                            string PAN_No = objSBOAPI.Query_Execute("Select \"TaxId0\" From \"CRD7\" Where \"CardCode\"='" + CusCode + "'");

                            //Thiru - Start - 20211203
                            int Count = 0;
                            string Scrap_Item = "N";
                            string Sales_Order_No = string.Empty;
                            Recordset Orec_ItemCode = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            Recordset Orec_Scrap = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            string ItemCode = "Select B.\"U_ItemCode\" AS \"ItemCode\",B.\"U_BillNo\" AS \"Sales Order No\" From \"@AIS_OPDE\" A INNER JOIN \"@AIS_PDE1\" B ON A.\"DocEntry\"=B.\"DocEntry\" WHERE A.\"DocEntry\"='" + DocEntry + "' and B.\"U_BDocEntry\" = '" + BDocEntry + "'and B.\"U_CusCode\"='" + CusCode + "' and B.\"U_TaxCode\"='" + TaxCode + "'and B.\"U_WhsCode\"='" + WhsCode + "'";
                            Orec_ItemCode.DoQuery(ItemCode);
                            if(Orec_ItemCode.RecordCount > 0)
                            {
                                while(!Orec_ItemCode.EoF)
                                {
                                    Sales_Order_No = Orec_ItemCode.Fields.Item("Sales Order No").Value.ToString();
                                    string ItmsGrpCod = objSBOAPI.Query_Execute("Select \"ItmsGrpCod\" from \"OITM\" Where \"ItemCode\" = '" + Orec_ItemCode.Fields.Item("ItemCode").Value.ToString() + "'");
                                    string Scrap_Item_Group = "Select \"ItmsGrpCod\" from \"OITB\" Where \"ItmsGrpNam\" like '%Scrap%'";
                                    Orec_Scrap.DoQuery(Scrap_Item_Group);
                                    if(Orec_Scrap.RecordCount > 0)
                                    {
                                        while(!Orec_Scrap.EoF)
                                        {
                                            if(ItmsGrpCod == Orec_Scrap.Fields.Item("ItmsGrpCod").Value.ToString())
                                            {
                                                Count = Count + 1;
                                                break;
                                            }
                                            Orec_Scrap.MoveNext();
                                        }
                                    }
                                    Orec_ItemCode.MoveNext();
                                }
                                if(Orec_ItemCode.RecordCount == Count)
                                {
                                    Scrap_Item = "Yes";
                                }
                                else if(Count == 0)
                                {
                                    Scrap_Item = "No";
                                }
                                else if(Orec_ItemCode.RecordCount != Count)
                                {
                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Sales Order No : " + Sales_Order_No + " - Contains both Scrap Items and Normal Items...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    flag1 = false;
                                    return flag1;
                                }
                            }

                            if(Scrap_Item == "Yes")
                            {
                                double TCS_Percentage = 0.00;
                                string TCS_Rate_for_Scrap_Items = objSBOAPI.Query_Execute("Select \"U_AVA_TCSRSI\" from \"OADM\"");
                                if(TCS_Rate_for_Scrap_Items != "" && TCS_Rate_for_Scrap_Items != "0")
                                {
                                    TCS_Percentage = Convert.ToDouble(TCS_Rate_for_Scrap_Items);
                                    oInvoices.Expenses.ExpenseCode = Convert.ToInt32(ExpnsCode);
                                    oInvoices.Expenses.TaxCode = Convert.ToString(Orec4.Fields.Item("TaxCode").Value);
                                    double TCSAmt = DocTotal * TCS_Percentage / 100;
                                    oInvoices.Expenses.LineTotal = Math.Round(TCSAmt, 2);
                                    oInvoices.Expenses.Add();
                                }
                                else
                                {
                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Please Give TCS Rate Greater than Zero for Scrap Items in Company Details...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    flag1 = false;
                                    return flag1;
                                }
                            }
                            else if (Scrap_Item == "No")
                            {
                                if (Total > Customer_TCS_Value)
                                {
                                    if (PAN_No == "")
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("Please Give the PAN No. in Business Partner Master... ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        flag1 = false;
                                        return flag1;
                                    }
                                    if (Last_2_FY_IT_Return_Filed == "Unknown" && Sum_Of_TDS_TCS_Greater_than_50000_for_Last_2_FY == "Unknown" && Customer_Deducting_TDS == "Unknown")
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("Please Give the Value As Yes or No in Business Partner Master UDF Fields Last_2_FY_IT_Return_Filed,Sum_Of_TDS_TCS_Greater_than_50000_for_Last_2_FY and Customer_Deducting_TDS... ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        flag1 = false;
                                        return flag1;
                                    }
                                    else if (Last_2_FY_IT_Return_Filed == "Unknown")
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("Please Give the Value As Yes or No in Business Partner Master UDF Field Last_2_FY_IT_Return_Filed... ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        flag1 = false;
                                        return flag1;
                                    }
                                    else if (Sum_Of_TDS_TCS_Greater_than_50000_for_Last_2_FY == "Unknown")
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("Please Give the Value As Yes or No in Business Partner Master UDF Field Sum_Of_TDS_TCS_Greater_than_50000_for_Last_2_FY... ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        flag1 = false;
                                        return flag1;
                                    }
                                    else if (Customer_Deducting_TDS == "Unknown")
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("Please Give the Value As Yes or No in Business Partner Master UDF Field Customer_Deducting_TDS... ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        flag1 = false;
                                        return flag1;
                                    }
                                    else
                                    {
                                        if (Customer_Deducting_TDS == "N")
                                        {
                                            oInvoices.Expenses.ExpenseCode = Convert.ToInt32(ExpnsCode);
                                            oInvoices.Expenses.TaxCode = Convert.ToString(Orec4.Fields.Item("TaxCode").Value);

                                            if (Last_2_FY_IT_Return_Filed == "Y" && Sum_Of_TDS_TCS_Greater_than_50000_for_Last_2_FY == "N")
                                            {
                                                double TCSAmt = DocTotal * 0.1 / 100;//Convert.ToDouble(Orec4.Fields.Item("TCSAmt").Value) * num4 / 100.0;
                                                oInvoices.Expenses.LineTotal = Math.Round(TCSAmt, 2);
                                            }
                                            else if (Last_2_FY_IT_Return_Filed == "N" && Sum_Of_TDS_TCS_Greater_than_50000_for_Last_2_FY == "Y")
                                            {
                                                double TCSAmt = DocTotal * 5 / 100;//Convert.ToDouble(Orec4.Fields.Item("TCSAmt").Value) * num4 / 100.0;
                                                oInvoices.Expenses.LineTotal = Math.Round(TCSAmt, 2);
                                            }
                                            else if (Last_2_FY_IT_Return_Filed == "Y" && Sum_Of_TDS_TCS_Greater_than_50000_for_Last_2_FY == "Y")
                                            {
                                                double TCSAmt = DocTotal * 0.1 / 100;//Convert.ToDouble(Orec4.Fields.Item("TCSAmt").Value) * num4 / 100.0;
                                                oInvoices.Expenses.LineTotal = Math.Round(TCSAmt, 2);
                                            }
                                            else if (Last_2_FY_IT_Return_Filed == "N" && Sum_Of_TDS_TCS_Greater_than_50000_for_Last_2_FY == "N")
                                            {
                                                double TCSAmt = DocTotal * 0.1 / 100;//Convert.ToDouble(Orec4.Fields.Item("TCSAmt").Value) * num4 / 100.0;
                                                oInvoices.Expenses.LineTotal = Math.Round(TCSAmt, 2);
                                            }
                                            oInvoices.Expenses.Add();
                                        }
                                        //else
                                        //{
                                        //    oInvoices.Expenses.ExpenseCode = 0;//Convert.ToInt32(ExpnsCode);
                                        //    oInvoices.Expenses.TaxCode = ""; //Convert.ToString(Orec4.Fields.Item("TaxCode").Value);
                                        //    //double TCSAmt = DocTotal * 0.1 / 100;//Convert.ToDouble(Orec4.Fields.Item("TCSAmt").Value) * num4 / 100.0;
                                        //    oInvoices.Expenses.LineTotal = 0.00;
                                        //    oInvoices.Expenses.Add();
                                        //}
                                    }
                                }
                            }
                            //Thiru - End - 20211203
                            
                            //oInvoices.Expenses.ExpenseCode = Convert.ToInt32(ExpnsCode);
                            //oInvoices.Expenses.TaxCode = Convert.ToString(Orec4.Fields.Item("TaxCode").Value);
                            //double TCSAmt = Convert.ToDouble(Orec4.Fields.Item("TCSAmt").Value) * num4 / 100.0;
                            //oInvoices.Expenses.LineTotal = Math.Round(TCSAmt, 2);
                            //oInvoices.Expenses.Add();
                            //Thiru Addon Changes - End - 02.07.2021 ( AR Invoice Freight Calculation Changes )
                        }
                    }
                    string errMsg = string.Empty;
                    int errCode = oInvoices.Add();
                    if (errCode != 0)
                    {
                        objSBOAPI.oCompany.GetLastError(out errCode, out errMsg);
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Error : (" + errMsg + ") , Pick List No: " + Pick_List_No, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        flag1 = false;
                        return flag1;
                    }
                    else
                    {
                        int TargetEntry = 0;
                        string Entry = Convert.ToString(TargetEntry);
                        ref string local = ref Entry;
                        objSBOAPI.oCompany.GetNewObjectCode(out local);
                        TargetEntry = Convert.ToInt32(Entry);
                        Recordset Orec4 = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        {
                            Orec4.DoQuery("Select \"DocNum\" from \"OINV\" Where \"DocEntry\"='" + Convert.ToString(TargetEntry) + "'");
                        }
                        else
                        {
                            Orec4.DoQuery("Select \"DocNum\" from OINV Where \"DocEntry\"='" + Convert.ToString(TargetEntry) + "'");
                        }
                            
                        int TargetNum = 0;
                        if (Orec4.RecordCount > 0)
                        {
                            Orec4.MoveFirst();
                            TargetNum = Convert.ToInt32(Orec4.Fields.Item("DocNum").Value.ToString().Trim());
                        }
                        int visualRowCount2 = oMatrix.VisualRowCount;
                        int Row = 1;
                        while (Row <= visualRowCount2)
                        {
                            CheckBox checkBox = (CheckBox)oMatrix.Columns.Item("V_5").Cells.Item(Row).Specific;
                            string BEn = oMatrix.Columns.Item("BEn").Cells.Item(Row).Specific.Value.ToString().Trim();
                            string TargetEnt = oMatrix.Columns.Item("TargetEnt").Cells.Item(Row).Specific.Value.ToString().Trim();
                            string TaxCodes = oMatrix.Columns.Item("TaxCode").Cells.Item(Row).Specific.Value.ToString().Trim();
                            if (TargetEnt == string.Empty && BDocEntry == BEn && TaxCode == TaxCodes)
                            {
                                oMatrix.FlushToDataSource();
                                oDBDSDetail.SetValue("U_TargetEntry", (Row - 1), Convert.ToString(TargetEntry));
                                oDBDSDetail.SetValue("U_TargetNum", (Row - 1), Convert.ToString(TargetNum));
                                oDBDSDetail.SetValue("U_TargetObject", (Row - 1), "13");
                                oMatrix.LoadFromDataSource();
                            }
                            { ++Row; }
                        }
                        objSBOAPI.SBO_Appln.StatusBar.SetText("A/R Invoice Posting Successfully , Pick List No: " + Pick_List_No, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                        Orec1.MoveNext();
                        { ++Rows; }
                    }
                }
                flag1 = true;
            }
            catch (Exception ex)
            {
                if (objSBOAPI.oCompany.InTransaction)
                    objSBOAPI.oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("A/R Invoice Posting Method Failed: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                flag1 = false;
            }
            finally
            {
                objform.Freeze(false);
                string QueryStr = null;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    QueryStr = "Delete from \"BATCHTEMPTAB\"";
                }
                else
                {
                    QueryStr = "Delete from BatchTempTab";
                }
                    
                ((IRecordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)).DoQuery(QueryStr);
            }
            return flag1;
        }
        #endregion

        #region FindIsBatchItem
        private string FindIsBatchItem(string pItemCode)
        {
            string str=string.Empty;
            try
            {
                Recordset Orec = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    Orec.DoQuery(" Select \"ManBtchNum\" from \"OITM\" where \"ItemCode\" = '" + pItemCode + "'");
                }
                else
                {
                    Orec.DoQuery(" select \"ManBtchNum\" from oitm where \"ItemCode\" = '" + pItemCode + "'");
                }
                    
                str = Convert.ToString(Orec.Fields.Item("ManBtchNum").Value);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("FindIsBatchItem Error " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            return str;
        }
        #endregion

        #region FindIsSerialItem
        private string FindIsSerialItem(string pItemCode)
        {
            string str=string.Empty;
            try
            {
                Recordset Orec = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    Orec.DoQuery(" select \"ManSerNum\" from \"OITM\" where \"ItemCode\" = '" + pItemCode + "'");
                }
                else
                {
                    Orec.DoQuery(" select \"ManSerNum\" from oitm where \"ItemCode\" = '" + pItemCode + "'");
                }
                    
                str = Convert.ToString(Orec.Fields.Item("ManSerNum").Value);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("FindIsSerialItem Error " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            return str;
        }
        #endregion
    }
}
