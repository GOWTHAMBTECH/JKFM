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
    class Cls_SalesOrderCreation
    {
        #region Declaration
        SAPbouiCOM.Form objform;
        ClsSBO objSBOAPI;
        SAPbobsCOM.Recordset oRS;
        SAPbobsCOM.Recordset oRS1;
        SAPbobsCOM.Recordset ORS2;
        SAPbobsCOM.Recordset ORS3;
        SAPbouiCOM.DBDataSource oDBDSDetail;
        SAPbouiCOM.DBDataSource oDBDSHeader;
        SAPbouiCOM.DBDataSource oDBDSDetail2;
        //public UIXML objUIXml;
        SAPbouiCOM.ComboBox oCombo;
        SAPbouiCOM.Matrix oMatrix;
        SAPbouiCOM.Matrix oMatrix2;
        SAPbouiCOM.CheckBox oCheck;
        string QryStr;
        SAPbobsCOM.Documents oDoc;
        string SQL;
        int chckFlag;
        string strPrivateDocEntry;
        #endregion        

        #region Constructor
        public Cls_SalesOrderCreation(ClsSBO objSBO)
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
                        case SAPbouiCOM.BoEventTypes.et_CLICK:
                            switch(pval.ItemUID)
                            {
                                case "Matrix":
                                    switch(pval.ColUID)
                                    {
                                        case "V_6":
                                            if (((SAPbouiCOM.ICheckBox)oMatrix.Columns.Item("V_6").Cells.Item(pval.Row).Specific).Checked && Convert.ToString(objSBOAPI.Query_Execute(" Select ISNULL(\"U_Return\",'N')  from \"@AIS_LOAD\" m left outer join \"@AIS_LOAD1\" d on m.\"DocEntry\" =d.\"DocEntry\" where \"DocNum\"='" + objform.Items.Item("9").Specific.Value.ToString().Trim() + "' and \"LineId\" ='" + Convert.ToString(pval.Row) + "'")).ToString().Trim() == "Y")
                                            {
                                                objSBOAPI.SBO_Appln.SetStatusBarMessage("You Cannot change....", BoMessageTime.bmt_Short, true);
                                                bubbleevent = false;
                                                return;
                                            }
                                            break;
                                    }
                                    break;
                                case "1":
                                    if (objform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || objform.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                    {
                                        if (!ValidateAll())
                                        {
                                            SystemSounds.Asterisk.Play();
                                            bubbleevent = false;
                                        }
                                        else
                                            AddCategory();
                                    }
                                    break;
                            }
                            break;

                        //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        //    switch (pval.ItemUID)
                        //    {
                        //        case "Matrix":
                        //            switch (pval.ColUID)
                        //            {
                        //                case "V_5":
                        //                    switch (pval.Row)
                        //                    {
                        //                        case 0:
                        //                            SAPbouiCOM.Matrix objmat;
                        //                            objmat = objform.Items.Item("Matrix").Specific;
                        //                            SAPbouiCOM.CheckBox oChk;
                        //                            for (int i = 1; i <= objmat.RowCount; i++)
                        //                            {
                        //                                oChk = objmat.Columns.Item("V_5").Cells.Item(i).Specific;
                        //                                if (((SAPbouiCOM.CheckBox)objmat.Columns.Item("V_5").Cells.Item(i).Specific). == true)
                        //                                {
                        //                                    oChk. = false;
                        //                                }
                        //                                else
                        //                                {
                        //                                    oChk. = true;
                        //                                }
                        //                            }
                        //                            break;
                        //                    }
                        //                    break;
                        //            }
                        //            break;
                        //    }
                        //    break;

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
                                    if (pval.Action_Success == true && objform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                    {
                                        InitForm();
                                    }
                                    break;
                                case "fld_ord":
                                    objform.PaneLevel = 1;
                                    objform.Settings.MatrixUID = "Matrix";
                                    objform.Items.Item(pval.ItemUID).AffectsFormMode = false;
                                    break;
                                case "fld_cat":
                                    objform.PaneLevel = 2;
                                    objform.Settings.MatrixUID = "Matrix2";
                                    objform.Items.Item(pval.ItemUID).AffectsFormMode = false;
                                    break;
                            }
                            break;

                        case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                            switch(pval.ItemUID)
                            {
                                case "c_series":
                                    if (objform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                    {
                                        oDBDSHeader.SetValue("DocNum", 0, Convert.ToString((long)objform.BusinessObject.GetNextSerialNumber(((SAPbouiCOM.IComboBox)objform.Items.Item("c_series").Specific).Value.ToString().Trim(), "AIS_LOAD")));
                                    }
                                    break;
                                case "c_BPGroup":
                                    if (objform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                    {
                                        if(!string.IsNullOrEmpty(oDBDSHeader.GetValue("U_BillRecDate", 0).ToString()))
                                        {
                                            LoadBills();
                                        }
                                    }
                                    break;
                            }
                            break;

                        case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
                            switch(pval.ItemUID)
                            {
                                case "Matrix":
                                    switch(pval.ColUID)
                                    {
                                        case "PreEnt":
                                            string str1 = oMatrix.Columns.Item("PreEnt").Cells.Item(pval.Row).Specific.Value.ToString().Trim();
                                            string str2 = null;
                                            string str3 = null;
                                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                            {
                                                str2 = Convert.ToString(objSBOAPI.Query_Execute("Select \"DocNum\" from \"@AIS_OPRE\" Where \"DocEntry\"='" + str1 + "'"));
                                            }
                                            else
                                            {
                                                str2 = Convert.ToString(objSBOAPI.Query_Execute("Select DocNum from [@AIS_OPRE] Where DocEntry='" + str1 + "'"));
                                            }
                                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                            {
                                                str3 = Convert.ToString(objSBOAPI.Query_Execute("Select \"Series\" from \"@AIS_OPRE\" Where \"DocEntry\"='" + str1 + "'"));
                                            }
                                            else
                                            {
                                                str3 = Convert.ToString(objSBOAPI.Query_Execute("Select Series from [@AIS_OPRE] Where DocEntry='" + str1 + "'"));
                                            }
                                                
                                            objSBOAPI.SBO_Appln.Menus.Item("OPRE").Activate();
                                            Form activeForm = objSBOAPI.SBO_Appln.Forms.ActiveForm;
                                            activeForm.Mode = BoFormMode.fm_FIND_MODE;
                                            activeForm.Items.Item("t_DocNum").Specific.Value = str2;
                                          
                                            ((SAPbouiCOM.IComboBox)activeForm.Items.Item("c_Series").Specific).Select(str3, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                            activeForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                            break;
                                    }
                                    break;
                            }
                            break;
                        case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                            switch(pval.ItemUID)
                            {
                                case "25":
                                    if (pval.ItemChanged == true)
                                    {
                                        LoadBills();
                                    }
                                    break;
                            }
                            break;

                        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                            DataTable selectedObjects = ((IChooseFromListEvent)pval).SelectedObjects;
                            if (selectedObjects != null)
                            {
                                switch(pval.ItemUID)
                                {
                                    case "Matrix":
                                        switch (pval.ColUID)
                                        {
                                            case "TaxCode":
                                                oMatrix.FlushToDataSource();
                                                oDBDSDetail.SetValue("U_TaxCode", pval.Row - 1, Convert.ToString(selectedObjects.GetValue("Code", 0)).Trim());
                                                oMatrix.LoadFromDataSource();
                                                if (objform.Mode == BoFormMode.fm_OK_MODE)
                                                {
                                                    objform.Mode = BoFormMode.fm_UPDATE_MODE;
                                                }
                                                break;
                                            case "WhsCode":
                                                oMatrix.FlushToDataSource();
                                                oDBDSDetail.SetValue("U_WhsCode", pval.Row - 1, Convert.ToString(selectedObjects.GetValue("WhsCode", 0)).Trim());
                                                oMatrix.LoadFromDataSource();
                                                if (objform.Mode == BoFormMode.fm_OK_MODE)
                                                {
                                                    objform.Mode = BoFormMode.fm_UPDATE_MODE;
                                                }
                                                break;
                                        }
                                        break;
                                }
                            }
                            break;

                        case SAPbouiCOM.BoEventTypes.et_CLICK:
                            switch (pval.ItemUID)
                            {
                                case "b_Del":
                                    if (objform.Items.Item(pval.ItemUID).Enabled && objform.Mode == BoFormMode.fm_OK_MODE)
                                    {
                                        string str = oDBDSHeader.GetValue("DocEntry", 0).ToString().Trim();
                                        Recordset businessObject = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        string QueryStr = null;
                                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                        {
                                            QueryStr = "SELECT  \"U_CusCode\" ,T1.\"U_TaxCode\"  FROM \"@AIS_LOAD1\" T1  Where \"DocEntry\" ='" + str + "' and IFNULL(\"U_TargetEntry\",'')='' and \"U_Select\" ='Y' group by  \"U_CusCode\" ,T1.\"U_TaxCode\"";
                                        }
                                        else
                                        {
                                            QueryStr = "SELECT  \"U_CusCode\" ,T1.U_TaxCode  FROM \"@AIS_LOAD1\" T1  Where \"DocEntry\" ='" + str + "' and isnull(U_TargetEntry,'')='' and \"U_Select\" ='Y' group by  \"U_CusCode\" ,T1.U_TaxCode";
                                        }
                                            
                                        businessObject.DoQuery(QueryStr);
                                        if (businessObject.RecordCount == 0)
                                        {
                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Document Already Posted...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        }
                                        else
                                        {
                                            if (!objSBOAPI.oCompany.InTransaction)
                                                objSBOAPI.oCompany.StartTransaction();
                                            if (PostingDelivery())
                                            {
                                                if (objSBOAPI.oCompany.InTransaction)
                                                    objSBOAPI.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                                    objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                                    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                            }
                                            else
                                            {
                                                int visualRowCount = oMatrix.VisualRowCount;
                                                int num = 1;
                                                while (num <= visualRowCount)
                                                {
                                                    if (((SAPbouiCOM.ICheckBox)oMatrix.Columns.Item("V_5").Cells.Item(num).Specific).Checked)
                                                    {
                                                        oMatrix.FlushToDataSource();
                                                        oDBDSDetail.SetValue("U_TargetEntry", num - 1, string.Empty);
                                                        oDBDSDetail.SetValue("U_TargetNum", num - 1, string.Empty);
                                                        oDBDSDetail.SetValue("U_TargetObject", num - 1, string.Empty);
                                                        oMatrix.LoadFromDataSource();
                                                    }
                                                    ++num;
                                                }
                                                if (objSBOAPI.oCompany.InTransaction)
                                                    objSBOAPI.oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                                                SystemSounds.Asterisk.Play();
                                            }
                                        }
                                    }
                                    break;

                                case "Matrix":
                                    switch (pval.ColUID)
                                    {
                                        case "V_5":
                                            if(pval.Row > 0)
                                            {
                                                objform.Freeze(true);

                                                if (((SAPbouiCOM.ICheckBox)oMatrix.Columns.Item("V_5").Cells.Item(pval.Row).Specific).Checked)
                                                {
                                                    double num1 = 0.0;
                                                    int visualRowCount = oMatrix.VisualRowCount;
                                                    int num2 = 1;
                                                    while (num2 <= visualRowCount)
                                                    {
                                                        if (((SAPbouiCOM.ICheckBox)oMatrix.Columns.Item("V_5").Cells.Item(num2).Specific).Checked)
                                                        {
                                                            string Left = Convert.ToString(oMatrix.Columns.Item("TonQty").Cells.Item(num2).Specific.Value);
                                                            if (Left == string.Empty)
                                                            {
                                                                string str = Convert.ToString(0.0);
                                                                num1 += Convert.ToDouble(str);
                                                            }
                                                            else
                                                            {
                                                                num1 += Convert.ToDouble(Left);
                                                            }
                                                            objform.Items.Item("t_tton").Specific.Value = num1;

                                                        }
                                                        ++num2;
                                                    }
                                                }
                                                else
                                                {
                                                    objform.Items.Item("t_tton").Specific.Value = (Convert.ToDouble(Convert.ToString(objform.Items.Item("t_tton").Specific.Value)) - Convert.ToDouble(Convert.ToString(oMatrix.Columns.Item("TonQty").Cells.Item(pval.Row).Specific.Value)));
                                                }
                                                objform.Freeze(false);
                                            }
                                            break;
                                    }
                                    break;

                            }
                            break;

                        case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                            switch (pval.ItemUID)
                            {
                                case "Matrix":
                                    switch (pval.ColUID)
                                    {
                                        case "V_5":
                                            switch (pval.Row)
                                            {
                                                case 0:
                                                    objform.Freeze(true);
                                                    SAPbouiCOM.Matrix objmat;
                                                    objmat = objform.Items.Item("Matrix").Specific;
                                                    SAPbouiCOM.CheckBox oChk;
                                                    for (int i = 1; i <= objmat.RowCount; i++)
                                                    {
                                                        oChk = objmat.Columns.Item("V_5").Cells.Item(i).Specific;
                                                        if (((SAPbouiCOM.CheckBox)objmat.Columns.Item("V_5").Cells.Item(i).Specific).Checked == true)
                                                        {
                                                            oChk.Checked = false;
                                                        }
                                                        else
                                                        {
                                                            oChk.Checked = true;
                                                        }
                                                    }
                                                    objform.Freeze(false);
                                                    break;
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
                        case "AIS_LOAD":
                            objform = objSBOAPI.LoadForm("SalesOrderCreation.xml", "AIS_LOAD");
                            objform.Freeze(true);
                            oDBDSHeader = objform.DataSources.DBDataSources.Item("@AIS_LOAD");
                            oDBDSDetail = objform.DataSources.DBDataSources.Item("@AIS_LOAD1");
                            oDBDSDetail2 = objform.DataSources.DBDataSources.Item("@AIS_LOAD2");
                            oMatrix = (SAPbouiCOM.Matrix)objform.Items.Item("Matrix").Specific;
                            oMatrix2 = (SAPbouiCOM.Matrix)objform.Items.Item("Matrix2").Specific;
                            SAPbouiCOM.ComboBox oCombo, oCombo1, oCombo2;
                            oCombo = objform.Items.Item("c_BPGroup").Specific;
                            oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                            oCombo1 = objform.Items.Item("c_series").Specific;
                            oCombo1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                            oCombo2 = objform.Items.Item("c_Status").Specific;
                            oCombo2.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                            if (objform.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            }
                            addmode();
                            objform.ActiveItem = "25";
                            DefineModeForFields();
                            oMatrix.Clear();
                            oDBDSDetail.Clear();
                            oMatrix2.Clear();
                            oDBDSDetail2.Clear();
                            objform.Items.Item("fld_ord").Click(BoCellClickType.ct_Regular);
                            objform.Freeze(false);
                            break;
                        //case "1281":
                        //    switch(pval.MenuUID)
                        //    {

                        //    }
                        //    if (Operators.CompareString(pVal.MenuUID, "1281", false) == 0 || Operators.CompareString(pVal.MenuUID, "1282", false) != 0)
                        //        return;
                        case "1282":
                            addmode();
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
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                            if (BusinessObjectInfo.ActionSuccess == true)                            {
                                strPrivateDocEntry = oDBDSHeader.GetValue("DocEntry", 0).ToString().Trim();
                            }                            break;

                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            if (BusinessObjectInfo.BeforeAction == true)
                                break;
                            string str = oDBDSHeader.GetValue("DocEntry", 0).ToString().Trim();
                            SAPbobsCOM.Recordset businessObject = (SAPbobsCOM.Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            string QueryStr = null;
                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                            {
                                QueryStr = "SELECT  \"U_CusCode\" ,T1.\"U_TaxCode\"  FROM \"@AIS_LOAD1\" T1  Where \"DocEntry\" ='" + str + "' and IFNULL(\"U_TargetEntry\",'')='' and \"U_Select\" ='Y' group by  \"U_CusCode\" ,T1.\"U_TaxCode\"";
                            }
                            else
                            {
                                QueryStr = "SELECT  \"U_CusCode\" ,T1.U_TaxCode  FROM \"@AIS_LOAD1\" T1  Where \"DocEntry\" ='" + str + "' and isnull(U_TargetEntry,'')='' and \"U_Select\" ='Y' group by  \"U_CusCode\" ,T1.U_TaxCode";
                            }
                                
                            businessObject.DoQuery(QueryStr);
                            if (businessObject.RecordCount == 0)
                            {
                                objform.Items.Item("Matrix").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
                                objform.Items.Item("b_Del").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
                            }
                            else
                            {
                                objform.Items.Item("Matrix").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_True);
                                objform.Items.Item("b_Del").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_True);
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

        #region LoadDocumentDate
        public bool LoadDocumentDate(SAPbouiCOM.EditText oEditText)
        {
            bool flag = false;
            try
            {
                oEditText.Active = true;
                oEditText.String = "A";
            }
            catch (Exception ex)
            {
                flag = false;
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("LoadDocumentDate Function Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            return flag;
        }
        #endregion

        #region LoadComboBoxSeries
        public bool LoadComboBoxSeries(SAPbouiCOM.ComboBox oComboBox, string UDOID)
        {
            bool flag = false;
            try
            {
                oComboBox.ValidValues.LoadSeries(UDOID, SAPbouiCOM.BoSeriesMode.sf_Add);
                oComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception ex)
            {
                flag = false;
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("LoadComboBoxSeries Function Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            return flag;
        }
        #endregion

        #region addmode
        public void addmode()
        {
            try
            {
                objform.Items.Item("9").Enabled = false;
                LoadDocumentDate((SAPbouiCOM.EditText)objform.Items.Item("10").Specific);
                LoadComboBoxSeries((SAPbouiCOM.ComboBox)objform.Items.Item("c_series").Specific, "AIS_LOAD");
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    objSBOAPI.setComboBoxValue((SAPbouiCOM.ComboBox)objform.Items.Item("c_BPGroup").Specific, "SELECT '-1' AS \"GroupCode\", 'ALL' AS \"GroupName\" FROM \"DUMMY\" UNION ALL SELECT \"GroupCode\", \"GroupName\" FROM \"OCRG\" WHERE \"GroupType\" = 'C'");
                }
                else
                {
                    objSBOAPI.setComboBoxValue((SAPbouiCOM.ComboBox)objform.Items.Item("c_BPGroup").Specific, "Select '-1' GroupCode ,'ALL' GroupName union all Select GroupCode ,GroupName  from OCRG Where GroupType ='C'");
                }
                    
                ((SAPbouiCOM.IComboBox)objform.Items.Item("c_BPGroup").Specific).Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                objform.Items.Item("fld_ord").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Add Mode Function Failure - " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region DefineModeForFields
        public void DefineModeForFields()
        {
            try
            {
                objform.Items.Item("c_series").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                objform.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                objform.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                objform.Items.Item("25").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                objform.Items.Item("25").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                objform.Items.Item("c_series").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                objform.Items.Item("25").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                objform.Items.Item("b_Del").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Define Mode For Fields Failed: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region InitForm
        public void InitForm()
        {
            try
            {
                objform.Freeze(true);
                string str1 = null;
                string str2 = null;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    str1 = Convert.ToString(objSBOAPI.Query_Execute("select \"DocNum\"  from \"@AIS_LOAD\" Where \"DocEntry\"='" + strPrivateDocEntry + "'"));
                }
                else
                {
                    str1 = Convert.ToString(objSBOAPI.Query_Execute("select DocNum  from \"@AIS_LOAD\" Where DocEntry='" + strPrivateDocEntry + "'"));
                }
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    str2 = Convert.ToString(objSBOAPI.Query_Execute("select \"Series\"  from \"@AIS_LOAD\" Where \"DocEntry\"='" + strPrivateDocEntry + "'"));
                }
                else
                {
                    str2 = Convert.ToString(objSBOAPI.Query_Execute("select Series  from \"@AIS_LOAD\" Where DocEntry='" + strPrivateDocEntry + "'"));
                }
                    
                string str3 = null;
                if (Convert.ToString(str1.Trim()) == "")
                {
                    str3 = Convert.ToString(0);
                }
                else
                {
                    str3 = Convert.ToString(str1);
                }
                //string str3 = Convert.ToString(Interactions.IIf(, , ));
                oMatrix.Clear();
                oDBDSDetail.Clear();
                oMatrix2.Clear();
                oDBDSDetail2.Clear();
                if (Convert.ToDouble(str3) != 0.0)
                {
                    objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    objform.Items.Item("9").Specific.Value = str3;
                    ((SAPbouiCOM.IComboBox)objform.Items.Item("c_series").Specific).Select(str2, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    objform.Items.Item("1").Click(BoCellClickType.ct_Regular);
                }
                objform.Items.Item("fld_ord").Click(BoCellClickType.ct_Regular);
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

        #region ValidateAll
        public bool ValidateAll()
        {
            bool flag1;
            try
            {
                oMatrix.FlushToDataSource();
                if (oMatrix.VisualRowCount == 0)
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Empty Document Shouldn't be Added...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    flag1 = false;
                    goto label_17;
                }
                else
                {
                    bool flag2 = false;
                    int visualRowCount = oMatrix.VisualRowCount;
                    int num1 = 1;
                    while (num1 <= visualRowCount)
                    {
                        if (((SAPbouiCOM.ICheckBox)oMatrix.Columns.Item("V_5").Cells.Item(num1).Specific).Checked)
                        {
                            flag2 = true;
                            break;
                        }
                        ++num1; 
                    }
                    if (!flag2)
                    {
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Select Atleast Anyone Sales Order...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        flag1 = false;
                        goto label_17;
                    }
                    else
                    {
                        DeleteUnselectedRow();
                        oMatrix.FlushToDataSource();
                        int num2 = oDBDSDetail.Size - 1;
                        int RecordNumber = 0;
                        while (RecordNumber <= num2)
                        {
                            oDBDSDetail.SetValue("LineID", RecordNumber, Convert.ToString(RecordNumber + 1));
                            ++RecordNumber; 
                        }
                        oMatrix.LoadFromDataSource();
                        flag1 = true;
                    }
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Validate Function Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                flag1 = false;
            }
            
        label_17:
            return flag1;
        }
        #endregion

        #region DeleteUnselectedRow

        private void DeleteUnselectedRow()
        {
            try
            {
                int visualRowCount1 = oMatrix.VisualRowCount;
                int RowNum = 1;
                int num1 = visualRowCount1;
                int num2 = 1;
                while (num2 <= num1)
                {
                    oCheck = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("V_5").Cells.Item(RowNum).Specific;
                    if (!oCheck.Checked)
                    {
                        oMatrix.DeleteRow(RowNum);
                        --RowNum;
                    }
                    ++RowNum;
                    int visualRowCount2 = oMatrix.VisualRowCount;
                    ++num2; 
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region AddCategory
        private void AddCategory()
        {
            ArrayList arrayList = new ArrayList();
            int visualRowCount1 = oMatrix.VisualRowCount;
            int num1 = 1;
            while (num1 <= visualRowCount1)
            {
                if (((SAPbouiCOM.ICheckBox)oMatrix.Columns.Item("V_5").Cells.Item(num1).Specific).Checked)
                {
                    string str = Convert.ToString(oMatrix.Columns.Item("V_2").Cells.Item(num1).Specific.Value);
                    if (!arrayList.Contains(str))
                        arrayList.Add(str);
                }
                ++num1; 
            }
            oMatrix2.Clear();
            oDBDSDetail2.Clear();
            oMatrix2.FlushToDataSource();
            int num2 = 1;
            int count = arrayList.Count;
            int num3 = 1;
            while (num3 <= count)
            {
                string newVal = Convert.ToString(arrayList[num3 - 1]);
                string empty = string.Empty;
                double num4 = 0.0;
                double num5 = 0.0;
                double num6 = 0.0;
                double num7 = 0.0;
                double num8 = 0.0;
                double num9 = 0.0;
                int visualRowCount2 = oMatrix.VisualRowCount;
                int num10 = 1;
                while (num10 <= visualRowCount2)
                {
                    if (((SAPbouiCOM.ICheckBox)oMatrix.Columns.Item("V_5").Cells.Item(num10).Specific).Checked)
                    {
                        empty = Convert.ToString(oMatrix.Columns.Item("V_1").Cells.Item(num10).Specific.Value);
                        if (oMatrix.Columns.Item("ItmCatgy").Cells.Item(num10).Specific.Value == "M1")
                        {
                            string str = oMatrix.Columns.Item("TonQty").Cells.Item(num10).Specific.Value.ToString();
                            num4 += Convert.ToDouble(str);
                        }
                        else if (oMatrix.Columns.Item("ItmCatgy").Cells.Item(num10).Specific.Value == "M2")
                        {
                            string str = Convert.ToString(oMatrix.Columns.Item("TonQty").Cells.Item(num10).Specific.Value);
                            num5 += Convert.ToDouble(str);
                        }
                        else if (oMatrix.Columns.Item("ItmCatgy").Cells.Item(num10).Specific.Value == "M3")
                        {
                            string str = Convert.ToString(oMatrix.Columns.Item("TonQty").Cells.Item(num10).Specific.Value);
                            num6 += Convert.ToDouble(str);
                        }
                        else if (oMatrix.Columns.Item("ItmCatgy").Cells.Item(num10).Specific.Value == "SUJI")
                        {
                            string str = Convert.ToString(oMatrix.Columns.Item("TonQty").Cells.Item(num10).Specific.Value);
                            num7 += Convert.ToDouble(str);
                        }
                        else if (oMatrix.Columns.Item("ItmCatgy").Cells.Item(num10).Specific.Value == "ATTA")
                        {
                            string str = Convert.ToString(oMatrix.Columns.Item("TonQty").Cells.Item(num10).Specific.Value);
                            num8 += Convert.ToDouble(str);
                        }
                        else if (oMatrix.Columns.Item("ItmCatgy").Cells.Item(num10).Specific.Value == "BRAN")
                        {
                            string str = Convert.ToString(oMatrix.Columns.Item("TonQty").Cells.Item(num10).Specific.Value);
                            num9 += Convert.ToDouble(str);
                        }
                    }
                    ++num10;
                }
                oDBDSDetail2.InsertRecord(num2 - 1);
                oDBDSDetail2.SetValue("LineId", num2 - 1, Convert.ToString(num2));
                oDBDSDetail2.SetValue("U_CardCode", num2 - 1, newVal);
                oDBDSDetail2.SetValue("U_CardName", num2 - 1, empty);
                oDBDSDetail2.SetValue("U_M1", num2 - 1, Convert.ToString(num4));
                oDBDSDetail2.SetValue("U_M2", num2 - 1, Convert.ToString(num5));
                oDBDSDetail2.SetValue("U_M3", num2 - 1, Convert.ToString(num6));
                oDBDSDetail2.SetValue("U_SUJI", num2 - 1, Convert.ToString(num7));
                oDBDSDetail2.SetValue("U_ATTA", num2 - 1, Convert.ToString(num8));
                oDBDSDetail2.SetValue("U_BRAN", num2 - 1, Convert.ToString(num9));
                ++num2; 
                ++num3; 
            }
            oMatrix2.LoadFromDataSource();
            //THIRU
            if (oMatrix2.Columns.Item("V_2").Cells.Item(oMatrix2.VisualRowCount).Specific.Value == "")
            {
                oMatrix2.DeleteRow(oMatrix2.VisualRowCount);
            }
            //THIRU
        }
        #endregion

        #region LoadBills
        private void LoadBills()
        {
            object obj;
            try
            {
                objform.Freeze(true);
                oRS = (SAPbobsCOM.Recordset)objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS1 = (SAPbobsCOM.Recordset)objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                ORS2 = (SAPbobsCOM.Recordset)objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                ArrayList arrayList = new ArrayList();
                string str1 = objSBOAPI.GetDateFromField(oDBDSHeader.GetValue("U_BillRecDate", 0).ToString().Trim());
                
                ComboBox comboBox = (SAPbouiCOM.ComboBox)objform.Items.Item("c_BPGroup").Specific;
                string empty1 = string.Empty;
                if (comboBox.Selected != null)
                    empty1 = comboBox.Selected.Value;

                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    oRS.DoQuery("CALL \"@AIS_BillDelivery_LoadPreSalesOrder\"('" + str1 + "', '" + empty1 + "')");
                }
                else
                {
                    oRS.DoQuery("Exec dbo.[@AIS_BillDelivery_LoadPreSalesOrder]'" + str1 + "','" + empty1 + "'");
                }
                
                oMatrix.Clear();
                oDBDSDetail.Clear();
                oMatrix.FlushToDataSource();
                int size = oDBDSDetail.Size;
                if (oRS.RecordCount > 0)
                {
                    oRS.MoveFirst();
                    int recordCount = oRS.RecordCount;
                    int num1 = 1;
                    while (num1 <= recordCount)//thiru (num1 <= recordCount)
                    {
                        oDBDSDetail.InsertRecord(oDBDSDetail.Size);
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait Loading - " + Convert.ToString(size), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                        oDBDSDetail.SetValue("LineId", size - 1, Convert.ToString(size));
                        oDBDSDetail.SetValue("U_PreSaleEnt", size - 1, Convert.ToString(oRS.Fields.Item("DocEntry").Value));
                        oDBDSDetail.SetValue("U_PreSalesNo", size - 1, Convert.ToString(oRS.Fields.Item("DocNum").Value));
                        oDBDSDetail.SetValue("U_CusCode", size - 1, Convert.ToString(oRS.Fields.Item("U_CardCode").Value));
                        oDBDSDetail.SetValue("U_CusName", size - 1, Convert.ToString(oRS.Fields.Item("U_CardName").Value));
                        oDBDSDetail.SetValue("U_ItemCode", size - 1, Convert.ToString(oRS.Fields.Item("U_ItemCode").Value));
                        oDBDSDetail.SetValue("U_ItemName", size - 1, Convert.ToString(oRS.Fields.Item("U_ItemName").Value));
                        oDBDSDetail.SetValue("U_ItmCatgy", size - 1, Convert.ToString(oRS.Fields.Item("U_ItemCategory").Value));
                        oDBDSDetail.SetValue("U_DefUnit", size - 1, Convert.ToString(oRS.Fields.Item("U_DefUnit").Value));
                        if (!arrayList.Contains(RuntimeHelpers.GetObjectValue(oRS.Fields.Item("U_CardCode").Value)))
                            arrayList.Add(RuntimeHelpers.GetObjectValue(oRS.Fields.Item("U_CardCode").Value));
                        string str2 = Convert.ToString(oRS.Fields.Item("U_ItemCode").Value);
                        double num2 = 0.0;
                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        {
                            num2 = Convert.ToDouble(Convert.ToDouble(oRS.Fields.Item("BagQuantity").Value) * Convert.ToDouble(objSBOAPI.Query_Execute("Select \"SalPackUn\" from \"OITM\" Where  \"ItemCode\" = '" + str2 + "'  "))) / 1000.0;
                        }
                        else
                        {
                            num2 = Convert.ToDouble(Convert.ToDouble(oRS.Fields.Item("BagQuantity").Value) * Convert.ToDouble(objSBOAPI.Query_Execute("Select SalPackUn from OITM Where  ItemCode = '" + str2 + "'  "))) / 1000.0;
                        }
                            
                        oDBDSDetail.SetValue("U_TonQty", size - 1, Convert.ToString(num2));
                        oDBDSDetail.SetValue("U_BagQty", size - 1, Convert.ToString(oRS.Fields.Item("BagQuantity").Value));
                        oDBDSDetail.SetValue("U_Price", size - 1, Convert.ToString(oRS.Fields.Item("U_UnitPrice").Value));
                        oDBDSDetail.SetValue("U_TaxCode", size - 1, Convert.ToString(oRS.Fields.Item("U_TaxCode").Value));
                        oDBDSDetail.SetValue("U_WhsCode", size - 1, Convert.ToString(oRS.Fields.Item("U_WhsCode").Value));
                        ++size; 
                        oRS.MoveNext();
                        ++num1; 
                    }
                    oMatrix.LoadFromDataSource();
                    //THIRU
                    if(oMatrix.Columns.Item("PreEnt").Cells.Item(oMatrix.VisualRowCount).Specific.Value == "")
                    {
                        oMatrix.DeleteRow(oMatrix.VisualRowCount);
                    }
                    //THIRU
                    oMatrix2.Clear();
                    oDBDSDetail2.Clear();
                    oMatrix2.FlushToDataSource();
                    int num3 = 1;
                    int count = arrayList.Count;
                    int num4 = 1;
                    while (num4 <= count)//thiru(num4 <= count)
                    {
                        string newVal = Convert.ToString(arrayList[num4 - 1]);
                        string empty2 = string.Empty;
                        double num2 = 0.0;
                        double num5 = 0.0;
                        double num6 = 0.0;
                        double num7 = 0.0;
                        double num8 = 0.0;
                        double num9 = 0.0;
                        int visualRowCount = oMatrix.VisualRowCount;//thiru oMatrix.VisualRowCount
                        int num10 = 1;
                        while (num10 <= visualRowCount)//thiru (num10 <= visualRowCount)
                        {
                            if (oMatrix.Columns.Item("V_2").Cells.Item(num10).Specific.Value == newVal)
                            {
                                empty2 = Convert.ToString(oMatrix.Columns.Item("V_1").Cells.Item(num10).Specific.Value);
                                if (oMatrix.Columns.Item("ItmCatgy").Cells.Item(num10).Specific.Value == "M1")
                                {
                                    string str2 = Convert.ToString(oMatrix.Columns.Item("TonQty").Cells.Item(num10).Specific.Value);
                                    num2 += Convert.ToDouble(str2);
                                }
                                else if (oMatrix.Columns.Item("ItmCatgy").Cells.Item(num10).Specific.Value == "M2")
                                {
                                    string str2 = Convert.ToString(oMatrix.Columns.Item("TonQty").Cells.Item(num10).Specific.Value);
                                    num5 += Convert.ToDouble(str2);
                                }
                                else if (oMatrix.Columns.Item("ItmCatgy").Cells.Item(num10).Specific.Value == "M3")
                                {
                                    string str2 = Convert.ToString(oMatrix.Columns.Item("TonQty").Cells.Item(num10).Specific.Value);
                                    num6 += Convert.ToDouble(str2);
                                }
                                else if (oMatrix.Columns.Item("ItmCatgy").Cells.Item(num10).Specific.Value == "SUJI")
                                {
                                    string str2 = Convert.ToString(oMatrix.Columns.Item("TonQty").Cells.Item(num10).Specific.Value);
                                    num7 += Convert.ToDouble(str2);
                                }
                                else if (oMatrix.Columns.Item("ItmCatgy").Cells.Item(num10).Specific.Value == "ATTA")
                                {
                                    string str2 = Convert.ToString(oMatrix.Columns.Item("TonQty").Cells.Item(num10).Specific.Value);
                                    num8 += Convert.ToDouble(str2);
                                }
                                else if (oMatrix.Columns.Item("ItmCatgy").Cells.Item(num10).Specific.Value == "BRAN")
                                {
                                    string str2 = Convert.ToString(oMatrix.Columns.Item("TonQty").Cells.Item(num10).Specific.Value);
                                    num9 += Convert.ToDouble(str2);
                                }
                            }
                            ++num10; 
                        }
                        oDBDSDetail2.InsertRecord(num3 - 1);
                        oDBDSDetail2.SetValue("LineId", num3 - 1, Convert.ToString(num3));
                        oDBDSDetail2.SetValue("U_CardCode", num3 - 1, newVal);
                        oDBDSDetail2.SetValue("U_CardName", num3 - 1, empty2);
                        oDBDSDetail2.SetValue("U_M1", num3 - 1, Convert.ToString(num2));
                        oDBDSDetail2.SetValue("U_M2", num3 - 1, Convert.ToString(num5));
                        oDBDSDetail2.SetValue("U_M3", num3 - 1, Convert.ToString(num6));
                        oDBDSDetail2.SetValue("U_SUJI", num3 - 1, Convert.ToString(num7));
                        oDBDSDetail2.SetValue("U_ATTA", num3 - 1, Convert.ToString(num8));
                        oDBDSDetail2.SetValue("U_BRAN", num3 - 1, Convert.ToString(num9));
                        ++num3; 
                        ++num4; 
                    }
                    oMatrix2.LoadFromDataSource();
                    //THIRU
                    if (oMatrix2.Columns.Item("V_2").Cells.Item(oMatrix2.VisualRowCount).Specific.Value == "")
                    {
                        oMatrix2.DeleteRow(oMatrix2.VisualRowCount);
                    }
                    //THIRU
                }
                else
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Billing Details Not available for this Bill Receipt Date...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    obj = false;
                    return;
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Load Bills Method Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                objform.Freeze(false);
            }
            return;
        }
        #endregion

        #region PostingDelivery
        public bool PostingDelivery()
        {
            bool flag1;
            try
            {
                objform.Freeze(true);
                string empty1 = string.Empty;
                string empty2 = string.Empty;
                string str1 = string.Empty;
                string str2 = oDBDSHeader.GetValue("DocEntry", 0).ToString().Trim();
                string str3 = oDBDSHeader.GetValue("U_BillRecDate", 0).ToString().Trim();
                string str4 = null;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    str4 = Convert.ToString(objSBOAPI.Query_Execute("SELECT CAST(CAST('" + str3 + "' AS date) AS varchar(10)) FROM DUMMY"));
                }
                else
                {
                    str4 = Convert.ToString(objSBOAPI.Query_Execute("Select Convert(Varchar(10),Convert(Date,'" + str3 + "') ,121) "));
                }
                    
                string str5 = oDBDSHeader.GetValue("DocNum", 0).ToString().Trim();
                SAPbouiCOM.ComboBox comboBox = (SAPbouiCOM.ComboBox)objform.Items.Item("c_series").Specific;
                if (comboBox.Selected != null)
                    str1 = comboBox.Selected.Value.ToString().Trim();
                SAPbobsCOM.Recordset businessObject1 = (SAPbobsCOM.Recordset)objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset businessObject2 = (SAPbobsCOM.Recordset)objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string QueryStr1 = null;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    QueryStr1 = "SELECT  T1.\"U_CusCode\" ,T1.\"U_TaxCode\",T1.\"U_PreSalesNo\",T1.\"U_WhsCode\"   FROM \"@AIS_LOAD1\" T1  Where T1.\"DocEntry\" ='" + str2 + "' and IFNULL(T1.\"U_TargetEntry\",'')='' and T1.\"U_Select\" ='Y' group by  T1.\"U_CusCode\" ,T1.\"U_PreSalesNo\" ,T1.\"U_TaxCode\",T1.\"U_WhsCode\"";
                }
                else
                {
                    QueryStr1 = "SELECT  T1.U_CusCode ,T1.U_TaxCode,T1.U_PreSalesNo,T1.U_WhsCode   FROM \"@AIS_LOAD1\" T1  Where T1.DocEntry ='" + str2 + "' and isnull(T1.U_TargetEntry,'')='' and T1.U_Select ='Y' group by  T1.U_CusCode ,T1.U_PreSalesNo ,T1.U_TaxCode,T1.U_WhsCode";
                }
                    
                businessObject1.DoQuery(QueryStr1);
                if (businessObject1.RecordCount > 0)
                    businessObject1.MoveFirst();
                int recordCount1 = businessObject1.RecordCount;
                int num1 = 1;
                while (num1 <= recordCount1)
                {
                    
                    bool flag2 = false;
                    string Left1 = businessObject1.Fields.Item("U_TaxCode").Value.ToString().Trim();
                    string str6 = businessObject1.Fields.Item("U_CusCode").Value.ToString().Trim();
                    string Left2 = businessObject1.Fields.Item("U_WhsCode").Value.ToString().Trim();
                    string Left3 = businessObject1.Fields.Item("U_PreSalesNo").Value.ToString().Trim();
                    string str7 = null;
                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        str7 = Convert.ToString(objSBOAPI.Query_Execute("select \"U_BranchCode\"  from \"@AIS_BRN1\"  where \"U_WhsCode\" ='" + Left2 + "'"));
                    }
                    else
                    {
                        str7 = Convert.ToString(objSBOAPI.Query_Execute("select U_BranchCode  from [@AIS_BRN1]  where U_WhsCode ='" + Left2 + "'"));
                    }
                        
                    Documents oSalesOrder = (Documents)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.oOrders);
                    string str8 = null;
                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        str8 = Convert.ToString(objSBOAPI.Query_Execute("select \"Series\"  from \"NNM1\" T0 inner join \"@AIS_OUNT\" T1 on T1.\"U_SOR\" =T0.\"SeriesName\" where \"ObjectCode\" ='17' and T1.\"Code\"='" + str7 + "'"));
                    }
                    else
                    {
                        str8 = Convert.ToString(objSBOAPI.Query_Execute("select Series  from NNM1 T0 inner join [@AIS_OUNT] T1 on T1.U_SOR =T0.SeriesName where ObjectCode ='17' and T1.Code='" + str7 + "'"));
                    }
                        
                    oSalesOrder.CardCode = businessObject1.Fields.Item("U_CusCode").Value.ToString().Trim();
                    oSalesOrder.Series = Convert.ToInt32(str8);
                    oSalesOrder.TaxDate = DateTime.Now;
                    oSalesOrder.DocDueDate = Convert.ToDateTime(str4);
                    oSalesOrder.DocDate = DateTime.Now;
                    string Left4 = null;
                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        Left4 = Convert.ToString(objSBOAPI.Query_Execute("Select \"BPLId\" from \"OBPL\" Where \"BPLName\" ='" + objSBOAPI.strBranchForDebitMemo + "'"));
                    }
                    else
                    {
                        Left4 = Convert.ToString(objSBOAPI.Query_Execute("Select \"BPLId\" from \"OBPL\" Where \"BPLName\" ='" + objSBOAPI.strBranchForDebitMemo + "'"));
                    }
                        
                    if (Left4 != string.Empty)
                        oSalesOrder.BPL_IDAssignedToInvoice = Convert.ToInt32(Left4);
                    oSalesOrder.UserFields.Fields.Item("U_UNIT").Value = str7;
                    oSalesOrder.UserFields.Fields.Item("U_BaseEntry").Value = str2;
                    oSalesOrder.UserFields.Fields.Item("U_BaseNum").Value = str5;
                    oSalesOrder.UserFields.Fields.Item("U_BaseObject").Value = "AIS_LOAD";
                    oSalesOrder.Comments = "Sales Order Document Posted from Customized Sales Order Doc No. " + str5 + "/" + str1;
                    string QueryStr2 = null;
                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        QueryStr2 = "Select \"U_Select\" ,\"U_TargetEntry\" ,\"U_PreSaleEnt\",\"U_CusCode\" ,\"U_TaxCode\" ,\"U_WhsCode\" ,\"U_PreSalesNo\",\"U_ItemCode\",\"U_BagQty\" ,\"U_Price\" from \"@AIS_LOAD1\" Where \"DocEntry\" ='" + str2 + "' and IFNULL(\"U_Select\",'')='Y' and \"U_PreSalesNo\"='" + Left3 + "' and \"U_WhsCode\" ='" + Left2 + "' and \"U_TaxCode\" ='" + Left1 + "'  and \"U_CusCode\" ='" + str6 + "'";
                    }
                    else
                    {
                        QueryStr2 = "Select U_Select ,U_TargetEntry ,U_PreSaleEnt,U_CusCode ,U_TaxCode ,U_WhsCode ,U_PreSalesNo,U_ItemCode,U_BagQty ,U_Price from \"@AIS_LOAD1\" Where DocEntry ='" + str2 + "' and isnull(U_Select,'')='Y' and U_PreSalesNo='" + Left3 + "' and U_WhsCode ='" + Left2 + "' and U_TaxCode ='" + Left1 + "'  and U_CusCode ='" + str6 + "'";
                    }
                        
                    businessObject2.DoQuery(QueryStr2);
                    if (businessObject2.RecordCount > 0)
                        businessObject2.MoveFirst();
                    int recordCount2 = businessObject2.RecordCount;
                    int num2 = 1;
                    while (num2 <= recordCount2)
                    {
                        string Left5 = businessObject2.Fields.Item("U_Select").Value.ToString().Trim();
                        string Left6 = businessObject2.Fields.Item("U_TargetEntry").Value.ToString().Trim();
                        string str9 = businessObject2.Fields.Item("U_PreSaleEnt").Value.ToString().Trim();
                        string QueryStr3 = null;
                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        {
                            QueryStr3 = "Select CASE When IFNULL(\"U_TrkReq\",'N')='Y' then 'YES' When IFNULL(\"U_TrkReq\",'N')='N' then 'NO' end  As \"U_TrkReq\",\"U_NoOfTrk\"  from \"@AIS_OPRE\" Where \"DocEntry\" ='" + str9 + "'";
                        }
                        else
                        {
                            QueryStr3 = "Select CASE When isnull(U_TrkReq,'N')='Y' then 'YES' When isnull(U_TrkReq,'N')='N' then 'NO' end  As U_TrkReq,U_NoOfTrk  from [@AIS_OPRE] Where DocEntry ='" + str9 + "'";
                        }
                            
                        Recordset businessObject4 = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        businessObject4.DoQuery(QueryStr3);
                        if (businessObject4.RecordCount > 0)
                        {
                            businessObject4.MoveFirst();
                            oSalesOrder.UserFields.Fields.Item("U_TPlanning").Value = businessObject4.Fields.Item("U_TrkReq").Value.ToString().Trim();
                            oSalesOrder.UserFields.Fields.Item("U_VNo").Value = businessObject4.Fields.Item("U_NoOfTrk").Value.ToString().Trim();
                        }
                        oSalesOrder.UserFields.Fields.Item("U_PreSalEnt").Value = str9;
                        oSalesOrder.UserFields.Fields.Item("U_PreSalNo").Value = Left3;
                        string Right1 = businessObject2.Fields.Item("U_CusCode").Value.ToString().Trim();
                        string Right2 = businessObject2.Fields.Item("U_TaxCode").Value.ToString().Trim();
                        string Right3 = businessObject2.Fields.Item("U_WhsCode").Value.ToString().Trim();
                        string Right4 = businessObject2.Fields.Item("U_PreSalesNo").Value.ToString().Trim();
                        if (Left5 == "Y" && Left6 == string.Empty && str6 == Right1 && Left1 == Right2 && Left3 == Right4 && Left2 == Right3)
                        {
                            string str10 = businessObject2.Fields.Item("U_ItemCode").Value.ToString().Trim();
                            string str11 = null;
                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                            {
                                str11 = Convert.ToString(objSBOAPI.Query_Execute("SELECT IFNULL(\"ChapterID\",'0')  FROM \"OITM\" Where \"ItemCode\" ='" + str10 + "'"));
                            }
                            else
                            {
                                str11 = Convert.ToString(objSBOAPI.Query_Execute("SELECT isnull(ChapterID,'')  FROM OITM Where ItemCode ='" + str10 + "'"));
                            }
                                
                            oSalesOrder.Lines.ItemCode = str10;
                            oSalesOrder.Lines.HSNEntry = Convert.ToInt32(str11);
                            oSalesOrder.Lines.TaxCode = businessObject2.Fields.Item("U_TaxCode").Value.ToString().Trim();
                            oSalesOrder.Lines.WarehouseCode = businessObject2.Fields.Item("U_WhsCode").Value.ToString().Trim();
                            double num3 = Convert.ToDouble(businessObject2.Fields.Item("U_BagQty").Value.ToString().Trim());
                            Convert.ToDouble(businessObject2.Fields.Item("U_Price").Value.ToString().Trim());
                            oSalesOrder.Lines.Quantity = num3;
                            oSalesOrder.Lines.UnitPrice = 0.0;
                            oSalesOrder.Lines.Add();
                            flag2 = true;
                        }
                        businessObject2.MoveNext();
                         { ++num2; }
                    }
                    string errMsg = string.Empty;
                    int errCode=0;
                    if (flag2)
                        errCode = oSalesOrder.Add();
                    if (errCode != 0)
                    {
                        objSBOAPI.oCompany.GetLastError(out errCode, out errMsg);
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Error : (" + errMsg + ") , Customer : " + str6 + " Tax Code : " + Left1, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        flag1 = false;
                        return flag1;
                    }
                    else
                    {
                        //SAPbobsCOM.Company objCompany = objSBOAPI.oCompany;
                        int integer1=0;
                        string str9 = Convert.ToString(integer1);
                        ref string local = ref str9;
                        objSBOAPI.oCompany.GetNewObjectCode(out local);
                        integer1 = Convert.ToInt32(str9);
                        Recordset businessObject4 = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        {
                            businessObject4.DoQuery("Select \"DocNum\" from \"ORDR\" Where \"DocEntry\"='" + Convert.ToString(integer1) + "'");
                        }
                        else
                        {
                            businessObject4.DoQuery("Select \"DocNum\" from ORDR Where \"DocEntry\"='" + Convert.ToString(integer1) + "'");
                        }
                            
                        int integer2=0;
                        if (businessObject4.RecordCount > 0)
                        {
                            businessObject4.MoveFirst();
                            integer2 = Convert.ToInt32(businessObject4.Fields.Item("DocNum").Value.ToString().Trim());
                        }
                        Recordset businessObject5 = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        string QueryStr3 = null;
                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        {
                            QueryStr3 = "Select \"LineId\", \"U_Select\" ,\"U_TargetEntry\" ,\"U_PreSaleEnt\",\"U_CusCode\" ,\"U_TaxCode\" ,\"U_WhsCode\" ,\"U_PreSalesNo\",\"U_ItemCode\" from \"@AIS_LOAD1\" Where \"DocEntry\" ='" + str2 + "' and IFNULL(\"U_Select\",'')='Y' and \"U_PreSalesNo\"='" + Left3 + "' and \"U_WhsCode\" ='" + Left2 + "' and \"U_TaxCode\" ='" + Left1 + "'  and \"U_CusCode\" ='" + str6 + "'";
                        }
                        else
                        {
                            QueryStr3 = "Select LineId, U_Select ,U_TargetEntry ,U_PreSaleEnt,U_CusCode ,U_TaxCode ,U_WhsCode ,U_PreSalesNo,U_ItemCode from \"@AIS_LOAD1\" Where DocEntry ='" + str2 + "' and isnull(U_Select,'')='Y' and U_PreSalesNo='" + Left3 + "' and U_WhsCode ='" + Left2 + "' and U_TaxCode ='" + Left1 + "'  and U_CusCode ='" + str6 + "'";
                        }
                            
                        businessObject5.DoQuery(QueryStr3);
                        if (businessObject5.RecordCount > 0)
                            businessObject5.MoveFirst();
                        int recordCount3 = businessObject5.RecordCount;
                        int num3 = 1;
                        while (num3 <= recordCount3)
                        {
                            string Left5 = businessObject5.Fields.Item("U_Select").Value.ToString().Trim();
                            string Left6 = businessObject5.Fields.Item("U_TargetEntry").Value.ToString().Trim();
                            string Right1 = businessObject5.Fields.Item("U_TaxCode").Value.ToString().Trim();
                            string Left7 = businessObject5.Fields.Item("U_CusCode").Value.ToString().Trim();
                            string Right2 = businessObject5.Fields.Item("U_WhsCode").Value.ToString().Trim();
                            string Right3 = businessObject5.Fields.Item("U_PreSalesNo").Value.ToString().Trim();
                            int integer3 = Convert.ToInt32(businessObject5.Fields.Item("LineId").Value.ToString().Trim());
                            if (Left5 == "Y" && Left6 == string.Empty && Left7 == str6 && Left1 == Right1 && Left3 == Right3 && Left2 == Right2)
                            {
                                oMatrix.FlushToDataSource();
                                oDBDSDetail.SetValue("U_TargetEntry", (integer3 - 1), Convert.ToString(integer1));
                                oDBDSDetail.SetValue("U_TargetNum", (integer3 - 1), Convert.ToString(integer2));
                                oDBDSDetail.SetValue("U_TargetObject", (integer3 - 1), "17");
                                oMatrix.LoadFromDataSource();
                            }
                            businessObject5.MoveNext();
                             { ++num3; }
                        }
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Sales Order Posting Successfully , Customer : " + str6 + " Tax Code : " + Left1, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                        businessObject1.MoveNext();
                         { ++num1; }
                    }
                }
                flag1 = true;
            }
            catch (Exception ex)
            {
                if (objSBOAPI.oCompany.InTransaction)
                    objSBOAPI.oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Delivery Posting Method Failed: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                flag1 = false;
            }
            finally
            {
                objform.Freeze(false);
            }
            return flag1;
        }
        #endregion
    }
}
