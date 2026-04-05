using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Collections;
using System.Diagnostics;
using System.Media;

namespace JKFM_Source
{
    class Cls_PickListManager
    {
        #region Declaration
        SAPbouiCOM.Form objform;
        ClsSBO objSBOAPI;

        SAPbouiCOM.Matrix oMatrix;
        SAPbouiCOM.Matrix oMatrix1;
        SAPbouiCOM.DBDataSource oDBDSHeader;
        SAPbouiCOM.DBDataSource oDBDSHeaderPick;
        
        SAPbouiCOM.DBDataSource oDBDSDetail;
        SAPbouiCOM.ComboBox oCombo;
        string QryStr;
        public static string oPickListNo;
        #endregion        

        #region Constructor
        public Cls_PickListManager(ClsSBO objSBO)
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
                                case "11":
                                    if(ReleasePickList(formuid, pval.Row) == false)
                                    {
                                        SystemSounds.Asterisk.Play();
                                        bubbleevent = false;
                                    }
                                    break;
                            }
                            break;
                        case BoEventTypes.et_FORM_LOAD:
                            LoadForm();
                            break;
                    }
                }
                else
                {
                    switch (pval.EventType)
                    {

                        case BoEventTypes.et_CLICK:
                            switch(pval.ItemUID)
                            {
                                case "b_Del":
                                    
                                    string Left = string.Empty;
                                    ArrayList arrayList1 = new ArrayList();
                                    ArrayList arrayList2 = new ArrayList();
                                    oMatrix = (Matrix)objform.Items.Item("19").Specific;
                                    int visualRowCount = oMatrix.VisualRowCount;
                                    int Row = 1;
                                    while (Row <= visualRowCount)
                                    {
                                        CheckBox checkBox = (CheckBox)oMatrix.Columns.Item("1").Cells.Item(Row).Specific;
                                        oMatrix.Columns.Item("11").Cells.Item(Row).Specific.Value.ToString();
                                        if (checkBox.Checked)
                                        {
                                            string str = oMatrix.Columns.Item("14").Cells.Item(Row).Specific.Value.ToString();
                                            if (!arrayList1.Contains(str))
                                            {
                                                if (arrayList1.Count == 0)
                                                {
                                                    arrayList1.Add(str);
                                                    Left = Left == string.Empty ? str : Left + "," + str;
                                                }
                                                else
                                                {
                                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Select Only One Pick List  ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                                    return;
                                                }
                                            }
                                            else
                                            {
                                                objSBOAPI.SBO_Appln.StatusBar.SetText("Duplicate Pick List isn't allowed", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                                return;
                                            }
                                        }
                                        { ++Row; }
                                    }
                                    if (Left == string.Empty)
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("Select Atleast Only One Pick List  ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    }
                                    else
                                    {
                                        oPickListNo = Left;
                                        objSBOAPI.SBO_Appln.ActivateMenuItem("OPDE");
                                    }
                                    break;

                                case "11":
                                    if (ReleasePickList(formuid, pval.Row) == false)
                                    {
                                        SystemSounds.Asterisk.Play();
                                    }
                                    break;

                                case "10":
                                    switch(pval.ColUID)
                                    {
                                        case "1":
                                            oMatrix1 = (Matrix)objform.Items.Item("10").Specific;
                                            double Total_Weight = 0.0;
                                            int Rows = 1;
                                            while (Rows <= oMatrix1.VisualRowCount)
                                            {
                                                if (((ICheckBox)oMatrix1.Columns.Item("1").Cells.Item(Rows).Specific).Checked)
                                                {
                                                    string Weight = Convert.ToString(oMatrix1.Columns.Item("1320000122").Cells.Item(Rows).Specific.Value);
                                                    Weight = Weight.Replace("T", "");
                                                    if (Weight == string.Empty)
                                                    {
                                                        string Weights = Convert.ToString(0.0);
                                                        Total_Weight += Convert.ToDouble(Weights);
                                                    }
                                                    else
                                                        Total_Weight += Convert.ToDouble(Weight);
                                                }
                                                { ++Rows; }
                                            }
                                            objform.Items.Item("t_weight").Specific.Value = Total_Weight;
                                            
                                            break;
                                    }
                                    break;

                                case "7":
                                    objform.Items.Item("b_Del").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
                                    objform.Items.Item("13").Visible = false;
                                    break;

                                case "6":
                                    objform.Items.Item("b_Del").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
                                    objform.Items.Item("13").Visible = false;
                                    break;

                                case "8":
                                    objform.Items.Item("b_Del").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_True);
                                    objform.Items.Item("13").Visible = false;
                                    break;
                            }
                            break;

                        case BoEventTypes.et_DOUBLE_CLICK:
                            switch(pval.ItemUID)
                            {
                                case "10":
                                    switch(pval.ColUID)
                                    {
                                        case "1":
                                            oMatrix1 = (Matrix)objform.Items.Item("10").Specific;
                                            double Total_Weight = 0.0;
                                            int visualRowCount = oMatrix1.VisualRowCount;
                                            int Row = 1;
                                            while (Row <= visualRowCount)
                                            {
                                                if (((ICheckBox)oMatrix1.Columns.Item("1").Cells.Item(Row).Specific).Checked)
                                                {
                                                    string Left = Convert.ToString(oMatrix1.Columns.Item("1320000122").Cells.Item(Row).Specific.Value);
                                                    Left = Left.Replace("T", "");
                                                    if (Left == string.Empty)
                                                    {
                                                        string str = Convert.ToString(0.0);
                                                        Total_Weight += Convert.ToDouble(str);
                                                    }
                                                    else
                                                        Total_Weight += Convert.ToDouble(Left);
                                                }
                                                { ++Row; }
                                            }
                                            objform.Items.Item("t_weight").Specific.Value = Total_Weight;
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
                Adding_Items();
                addmode();
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Init Method Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
            }
            finally
            {
                objform.Freeze(false);
            }
        }
        #endregion

        #region addmode
        public void addmode()
        {
            try
            {
                objform.Items.Item("9").Enabled = false;
                oCombo = (ComboBox)objform.Items.Item("c_vehicle").Specific;
                oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    QryStr = "Select \"Code\",\"U_TVehicle\" from \"@AIS_OVEH\"";
                }
                else
                {
                    QryStr = "Select \"Code\",\"U_TVehicle\" from \"@AIS_OVEH\"";
                }
                    
                objSBOAPI.setComboBoxValue(oCombo, QryStr);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Add Mode Function Failure - " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Adding_Items
        public void Adding_Items()
        {
            //objSBOAPI.Adding_Items_Static("l_weight", 80, (objform.Items.Item("13").Top - 43), (objform.Items.Item("13").Left - 100), 14, "Total Weight", objform.UniqueID);
            //objSBOAPI.Adding_Items_Edit("t_weight", 100, objform.Items.Item("l_weight").Top, (objform.Items.Item("l_weight").Left + 100), 14, "Quotation No", objform.UniqueID);
            //objSBOAPI.Adding_Items_Static("St_QuotNo", oExtItem.Width, oExtItem.Top + 15, oExtItem.Left, oExtItem.Height, "Quotation No", objform.UniqueID);
            //objSBOAPI.Adding_Items_Static("St_QuotNo", oExtItem.Width, oExtItem.Top + 15, oExtItem.Left, oExtItem.Height, "Quotation No", objform.UniqueID);
            Item tem1 = objform.Items.Add("b_Del", BoFormItemTypes.it_BUTTON);
            tem1.Left = (objform.Items.Item("11").Left - 100);
            tem1.Width = 100;
            tem1.Height = objform.Items.Item("11").Height;
            tem1.Top = objform.Items.Item("11").Top;
            tem1.FromPane = 0;
            tem1.ToPane = 0;
            tem1.Visible = true;
            ((IButton)tem1.Specific).Caption = "Delivery";
            Item tem2 = objform.Items.Add("l_weight", BoFormItemTypes.it_STATIC);
            tem2.Left = (objform.Items.Item("13").Left - 100);
            tem2.Width = 80;
            tem2.Height = 14;
            tem2.Top = (objform.Items.Item("13").Top - 43);
            tem2.FromPane = 0;
            tem2.ToPane = 0;
            tem2.Visible = true;
            ((IStaticText)tem2.Specific).Caption = "Total Weight";
            Item tem3 = objform.Items.Add("t_weight", BoFormItemTypes.it_EDIT);
            tem3.Left = (objform.Items.Item("l_weight").Left + 100);
            tem3.Width = 100;
            tem3.Height = 14;
            tem3.Top = objform.Items.Item("l_weight").Top;
            tem3.FromPane = 0;
            tem3.ToPane = 0;
            tem3.Visible = true;
            tem3.Enabled = false;
            EditText editText = (EditText)tem3.Specific;
            Item tem4 = objform.Items.Add("l_vehicle", BoFormItemTypes.it_STATIC);
            tem4.Left = (objform.Items.Item("10000024").Left + 200);
            tem4.Width = 100;
            tem4.Height = objform.Items.Item("10000024").Height;
            tem4.Top = objform.Items.Item("10000024").Top;
            tem4.FromPane = 0;
            tem4.ToPane = 0;
            tem4.Visible = true;
            ((IStaticText)tem4.Specific).Caption = "Vehicle";
            Item tem5 = objform.Items.Add("c_vehicle", BoFormItemTypes.it_COMBO_BOX);
            tem5.Left = (objform.Items.Item("l_vehicle").Left + 100);
            tem5.Width = 100;
            tem5.Height = objform.Items.Item("l_vehicle").Height;
            tem5.Top = objform.Items.Item("l_vehicle").Top;
            tem5.FromPane = 0;
            tem5.ToPane = 0;
            tem5.Visible = true;
            ComboBox comboBox = (ComboBox)tem5.Specific;
            tem5.DisplayDesc = true;
        }
        #endregion

        #region ReleasePickList
        public bool ReleasePickList(string FormUID, int RowNo)
        {
            bool obj = false;
            try
            {
                double Tolerance = 0.0;
                oMatrix1 = (Matrix)objform.Items.Item("10").Specific;
                Recordset Orec = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string Vehicle_Code = ((IComboBox)objform.Items.Item("c_vehicle").Specific).Value.ToString().Trim();
                //Thiru - 19.03.2021
                //if (Vehicle_Code == string.Empty)
                //{

                //objSBOAPI.SBO_Appln.StatusBar.SetText("Select the Vehicle Type.....", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                //obj = false;

                //}
                //Thiru - 19.03.2021
                //Thiru - 19.03.2021
                //else
                //{
                //string QueryStr1 = null;
                //    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                //    {
                //        QueryStr1 = "select  \"U_TolerPer\" ,\"U_TolerKg\"  from \"@AIS_OVEH\" where \"Code\" ='" + Vehicle_Code + "'";
                //    }
                //    else
                //    {
                //        QueryStr1 = "select  U_TolerPer ,U_TolerKg  from [@AIS_OVEH] where Code ='" + Vehicle_Code + "'";
                //    }
                        
                //    Orec.DoQuery(QueryStr1);
                    double Tolerance_Percent = 0.0;
                    double Tolerance_Kg = 0.0;
                    //if (Orec.RecordCount > 0)
                    //{
                    //    Orec.MoveFirst();
                    //    Tolerance_Percent = Convert.ToDouble(Orec.Fields.Item("U_TolerPer").Value);
                    //    Tolerance_Kg = Convert.ToDouble(Orec.Fields.Item("U_TolerKg").Value);
                    //}
                    double Tolerance_Subtract = Tolerance_Kg - Tolerance_Percent;
                    double Tolerance_Add = Tolerance_Kg + Tolerance_Percent;
                    ArrayList arrayList = new ArrayList();
                    int visualRowCount1 = oMatrix1.VisualRowCount;
                    int Row = 1;
                    while (Row <= visualRowCount1)
                    {
                        string Document_No = Convert.ToString(oMatrix1.Columns.Item("11").Cells.Item(Row).Specific.Value);
                        string DeliveryorDueDate = Convert.ToString(oMatrix1.Columns.Item("5").Cells.Item(Row).Specific.Value);
                        if (((ICheckBox)oMatrix1.Columns.Item("1").Cells.Item(Row).Specific).Checked && Document_No != string.Empty)
                        {
                            string Presales_Entry = null;
                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                            {
                                Presales_Entry = Convert.ToString(objSBOAPI.Query_Execute("Select \"U_PreSalEnt\"  from \"ORDR\" Where \"DocNum\" ='" + Document_No + "' and \"DocDueDate\" ='" + DeliveryorDueDate + "'"));
                            }
                            else
                            {
                                Presales_Entry = Convert.ToString(objSBOAPI.Query_Execute("Select U_PreSalEnt  from ORDR Where DocNum ='" + Document_No + "' and DocDueDate ='" + DeliveryorDueDate + "'"));
                            }
                                
                            if (!arrayList.Contains(Presales_Entry))
                            {
                                string BP_Code = oMatrix1.Columns.Item("10").Cells.Item(Row).Specific.Value.ToString().Trim();
                                string Customer_Credit_Limit = oMatrix1.Columns.Item("1320000100").Cells.Item(Row).Specific.Value.ToString().Trim();
                                string Balance = null;
                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                {
                                    Balance = Convert.ToString(objSBOAPI.Query_Execute("Select \"Balance\" From \"OCRD\" Where \"CardCode\" ='" + BP_Code + "'"));
                                }
                                else
                                {
                                    Balance = Convert.ToString(objSBOAPI.Query_Execute("Select Balance From OCRD Where CardCode ='" + BP_Code + "'"));
                                }
                                    
                                double Customer_Credit_Limits = 0.0;
                                double Balances = 0.0;
                                if (Customer_Credit_Limit != string.Empty)
                                    Customer_Credit_Limits = Convert.ToDouble(Customer_Credit_Limit);
                                if (Balance != string.Empty)
                                    Balances = Convert.ToDouble(Balance);
                                string QueryStr2 = null;
                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                {
                                    //Thiru - Credit Check - 20211125 - Start
                                    //QueryStr2 = "CALL \"@AIS_BillDelivery_GetDocTotal\"('" + Document_No + "','" + DeliveryorDueDate + "') ";
                                    QueryStr2 = "CALL \"@AIS_BillDelivery_GetDocTotal_CreditCheck\"('" + Document_No + "','" + DeliveryorDueDate + "') ";
                                    //Thiru - Credit Check - 20211125 - End
                                }
                                else
                                {
                                    QueryStr2 = "Exec dbo.[@AIS_BillDelivery_GetDocTotal]'" + Document_No + "','" + DeliveryorDueDate + "' ";
                                }
                                    
                                Recordset Orec2 = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                Orec2.DoQuery(QueryStr2);
                                //Thiru - Credit Check - 20211125 - Start
                                if (Orec2.RecordCount > 0)
                                {
                                    
                                    //Orec2.MoveFirst();
                                    //Balances = Convert.ToDouble(Balances + Orec2.Fields.Item(0).Value);
                                    if (Orec2.Fields.Item("Credit Check").Value == "FALSE")
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("For Line No. " + Row.ToString() + ",Customer credit limit is exceeding..........", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        obj = false;
                                        return obj;
                                    }
                                    else
                                        arrayList.Add(Presales_Entry);
                                }
                                else
                                {
                                    objSBOAPI.SBO_Appln.StatusBar.SetText("No Records Found in @AIS_BillDelivery_GetDocTotal_CreditCheck SP...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    obj = false;
                                    return obj;
                                }
                                //if (Customer_Credit_Limits < Balances)
                                //{
                                //    objSBOAPI.SBO_Appln.StatusBar.SetText("For Line No. " + Row.ToString() + ",Customer credit limit is exceeding..........", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                //    obj = false;
                                //    return obj;
                                //}
                                //else
                                //    arrayList.Add(Presales_Entry);
                                //Thiru - Credit Check - 20211125 - End
                            }
                        }
                        { ++Row; }
                    }
                    bool flag = false;
                    int visualRowCount2 = oMatrix1.VisualRowCount;
                    int Row_Count = 1;
                    while (Row_Count <= visualRowCount2)
                    {
                        string Doc_No = Convert.ToString(oMatrix1.Columns.Item("11").Cells.Item(Row_Count).Specific.Value);
                        string DocDueDate = Convert.ToString(oMatrix1.Columns.Item("5").Cells.Item(Row_Count).Specific.Value);
                        if (((ICheckBox)oMatrix1.Columns.Item("1").Cells.Item(Row_Count).Specific).Checked)
                        {
                            if (Doc_No != string.Empty)
                            {
                                string Presales_Ent = null;
                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                {
                                    Presales_Ent = Convert.ToString(objSBOAPI.Query_Execute("Select \"U_PreSalEnt\"  from \"ORDR\" Where \"DocNum\" ='" + Doc_No + "' and \"DocDueDate\" ='" + DocDueDate + "'"));
                                }
                                else
                                {
                                    Presales_Ent = Convert.ToString(objSBOAPI.Query_Execute("Select U_PreSalEnt  from ORDR Where DocNum ='" + Doc_No + "' and DocDueDate ='" + DocDueDate + "'"));
                                }
                                    
                                if (!arrayList.Contains(Presales_Ent))
                                {
                                    string CardCode = oMatrix1.Columns.Item("10").Cells.Item(Row_Count).Specific.Value.ToString().Trim();
                                    string Customer_Credit_Limit = oMatrix1.Columns.Item("1320000100").Cells.Item(Row_Count).Specific.Value.ToString().Trim();
                                    string Balance = null;
                                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                    {
                                        Balance = Convert.ToString(objSBOAPI.Query_Execute("Select \"Balance\" From \"OCRD\" Where \"CardCode\" ='" + CardCode + "'"));
                                    }
                                    else
                                    {
                                        Balance = Convert.ToString(objSBOAPI.Query_Execute("Select Balance From OCRD Where CardCode ='" + CardCode + "'"));
                                    }
                                        
                                    double Customer_Credit_Limits = 0.0;
                                    double Balances = 0.0;
                                    if (Customer_Credit_Limit != string.Empty)
                                        Customer_Credit_Limits = Convert.ToDouble(Customer_Credit_Limit);
                                    if (Balance != string.Empty)
                                        Balances = Convert.ToDouble(Balance);
                                    string QueryStr2 = null;
                                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                    {
                                        //Thiru - Credit Check - 20211125 - Start
                                        //QueryStr2 = "CALL \"@AIS_BillDelivery_GetDocTotal\"('" + Doc_No + "','" + DocDueDate + "') ";
                                        QueryStr2 = "CALL \"@AIS_BillDelivery_GetDocTotal_CreditCheck\"('" + Doc_No + "','" + DocDueDate + "') ";
                                        //Thiru - Credit Check - 20211125 - End
                                    }
                                    else
                                    {
                                        QueryStr2 = "Exec dbo.[@AIS_BillDelivery_GetDocTotal]'" + Doc_No + "','" + DocDueDate + "' ";
                                    }
                                        
                                    Recordset Orec2 = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    Orec2.DoQuery(QueryStr2);
                                    //Thiru - Credit Check - 20211125 - Start
                                    if (Orec2.RecordCount > 0)
                                    {
                                        //Orec2.MoveFirst();
                                        //Balances = Convert.ToDouble(Balances + Orec2.Fields.Item(0).Value);
                                        if (Orec2.Fields.Item("Credit Check").Value == "FALSE")
                                        {
                                            objSBOAPI.SBO_Appln.StatusBar.SetText("For Line No. " + Row_Count.ToString() + ",Customer credit limit is exceeding..........", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                            obj = false;
                                            return obj;
                                        }
                                        else
                                            arrayList.Add(Presales_Ent);
                                    }
                                    else
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("No Records Found in @AIS_BillDelivery_GetDocTotal_CreditCheck SP...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        obj = false;
                                        return obj;
                                    }
                                    //if (Customer_Credit_Limits < Balances)
                                    //{
                                    //    objSBOAPI.SBO_Appln.StatusBar.SetText("For Line No. " + Row_Count.ToString() + ",Customer credit limit is exceeding..........", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    //    obj = false;
                                    //    return obj;
                                    //}
                                    //else
                                    //    arrayList.Add(Presales_Ent);
                                    //Thiru - Credit Check - 20211125 - End
                                }
                            }
                            if (oMatrix1.Columns.Item("8").Cells.Item(Row_Count).Specific.Value != string.Empty)
                            {
                                flag = true;
                                int Doc_Row = Convert.ToInt32(oMatrix1.Columns.Item("12").Cells.Item(Row_Count).Specific.Value);
                                int Rows = (Row_Count - Doc_Row);
                                if (!((ICheckBox)oMatrix1.Columns.Item("1").Cells.Item(Rows).Specific).Checked)
                                {
                                    objSBOAPI.SBO_Appln.StatusBar.SetText("For Line No. " + Rows.ToString() + ",Select the all Item for a Sale Order", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    obj = false;
                                    return obj;
                                }
                                else
                                {
                                    double Open_Qty = 0.0;
                                    if (oMatrix1.Columns.Item("4").Cells.Item(Row_Count).Specific.Value != string.Empty)
                                        Open_Qty = Convert.ToDouble(oMatrix1.Columns.Item("4").Cells.Item(Row_Count).Specific.Value);
                                    double Released_Quantity = 0.0;
                                    if (oMatrix1.Columns.Item("3").Cells.Item(Row_Count).Specific.Value != string.Empty)
                                        Released_Quantity = Convert.ToDouble(oMatrix1.Columns.Item("3").Cells.Item(Row_Count).Specific.Value);
                                    double Avail_To_Release = 0.0;
                                    if (oMatrix1.Columns.Item("2").Cells.Item(Row_Count).Specific.Value != string.Empty)
                                        Avail_To_Release = Convert.ToDouble(oMatrix1.Columns.Item("2").Cells.Item(Row_Count).Specific.Value);
                                    if (Avail_To_Release < Open_Qty)
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("For Line No. " + Row_Count.ToString() + " ,Stock isn't available", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        obj = false;
                                        return obj;
                                    }
                                    else if (Released_Quantity < Open_Qty)
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("For Line No. " + Row_Count.ToString() + " ,Released Quantity should be equal to Open Qty...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        obj = false;
                                        return obj;
                                    }
                                    else
                                    {
                                        //Doubt
                                        string Weight = Convert.ToString(oMatrix1.Columns.Item("1320000122").Cells.Item(Row_Count).Specific.Value);
                                        //Left3 = Left3.Replace(Convert.ToString(oMatrix1.Columns.Item("1320000122").Cells.Item(num9).Specific.Value), "T", "", 1, -1, 0);
                                        Weight = Weight.Replace("T", "");
                                        //Doubt
                                        if (Weight == string.Empty)
                                        {
                                            string Weights = Convert.ToString(0.0);
                                            Tolerance += Convert.ToDouble(Weights);
                                        }
                                        else
                                            Tolerance += Convert.ToDouble(Weight);
                                    }
                                }
                            }
                        }
                        else if (arrayList.Count > 0)
                        {
                            string PreSalEnt = null;
                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                            {
                                PreSalEnt = Convert.ToString(objSBOAPI.Query_Execute("Select \"U_PreSalEnt\"  from \"ORDR\" Where \"DocNum\" ='" + Doc_No + "' and \"DocDueDate\" ='" + DocDueDate + "'"));
                            }
                            else
                            {
                                PreSalEnt = Convert.ToString(objSBOAPI.Query_Execute("Select U_PreSalEnt  from ORDR Where DocNum ='" + Doc_No + "' and DocDueDate ='" + DocDueDate + "'"));
                            }
                            //Thiru - 19.03.2021
                            //if (arrayList.Contains(PreSalEnt))
                            //{
                            //    objSBOAPI.SBO_Appln.StatusBar.SetText("For Line No. " + Row_Count.ToString() + ",Select the Single Presales Order related all SO", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            //    obj = false;
                            //    return obj;
                            //}
                            //Thiru - 19.03.2021
                        }
                        { ++Row_Count; }
                    }
                    if (!flag)
                    {
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Expand the Sales Orders", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        obj = false;
                    }
                    else if (Tolerance >= Tolerance_Subtract && Tolerance <= Tolerance_Add)
                    {
                        obj = true;
                    }
                    //Thiru - 29.03.2021
                    else
                    {
                        //objSBOAPI.SBO_Appln.StatusBar.SetText("Overall Weight should be equal to Vehicle Weight", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        obj = true;
                    }
                    //Thiru - 29.03.2021
                //}
                //Thiru - 19.03.2021
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Released Pick ListFunction Failed: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                obj = false;
            }
            return obj;
        }
        #endregion

        #region LoadForm
        public void LoadForm()
        {
            try
            {
                objform = objSBOAPI.SBO_Appln.Forms.GetForm("81", -1);
                oDBDSHeader = objform.DataSources.DBDataSources.Item(0);
                oDBDSDetail = objform.DataSources.DBDataSources.Item(1);
                InitForm();
                objform.Items.Item("b_Del").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
                objform.Items.Item("t_weight").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 2, BoModeVisualBehavior.mvb_False);
                objform.Items.Item("t_weight").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("LoadForm Method Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
            }
        }
        #endregion
    }
}
