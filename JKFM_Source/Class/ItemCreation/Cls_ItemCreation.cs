using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Sap.Data.Hana;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace JKFM_Source
{
    class Cls_ItemCreation
    {

        #region Declaration
        SAPbouiCOM.Form objform;
        ClsSBO objSBOAPI;

        int RetCode;
        int PriceListCode;

        SAPbouiCOM.ComboBox oCombo;
        #endregion        

        #region Constructor
        public Cls_ItemCreation(ClsSBO objSBO)
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
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                            if (pval.ItemUID == "BtnCreate")
                            {
                                string Type1 = objform.DataSources.UserDataSources.Item("Type").Value;
                                //if (Type1 == "3")
                                //{
                                //    if (((SAPbouiCOM.EditText)objform.Items.Item("EtFrom").Specific).Value == "")
                                //    {
                                //        objSBOAPI.SBO_Appln.StatusBar.SetText("From Date is Mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                //        bubbleevent = false;
                                //        return;
                                //    }
                                //    else if (((SAPbouiCOM.EditText)objform.Items.Item("EtTo").Specific).Value != "")
                                //    {
                                //        if (Convert.ToInt32(objform.Items.Item("EtTo").Specific.Value) < Convert.ToInt32(objform.Items.Item("EtFrom").Specific.Value))
                                //        {
                                //            objSBOAPI.SBO_Appln.StatusBar.SetText("To Date must be greater than From Date...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                //            bubbleevent = false;
                                //            objform.Items.Item("EtTo").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                //            return;
                                //        }
                                //    }
                                //}
                                if (Type1 == "1")
                                {
                                    string FileName = objform.DataSources.UserDataSources.Item("FileName").Value;
                                    if (System.IO.File.Exists(FileName))
                                    {
                                        XSSFWorkbook hssfworkbook = new XSSFWorkbook(FileName);
                                        ISheet sheet = hssfworkbook.GetSheetAt(0);
                                        IRow headerRow = sheet.GetRow(0);
                                        int colCount = headerRow.LastCellNum;

                                        bool oDateValidation = false;

                                        for (int c = 0; c < colCount; c++)
                                        {
                                            if (headerRow.GetCell(c).ToString().StartsWith("PR_"))
                                            {
                                                oDateValidation = true;
                                                break;
                                            }

                                        }
                                        hssfworkbook.Close();

                                        if (oDateValidation)
                                        {
                                            if (((SAPbouiCOM.EditText)objform.Items.Item("EtFrom").Specific).Value == "")
                                            {
                                                objSBOAPI.SBO_Appln.StatusBar.SetText("From Date is Mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                                bubbleevent = false;
                                                return;
                                            }
                                            else if (((SAPbouiCOM.EditText)objform.Items.Item("EtTo").Specific).Value != "")
                                            {
                                                if (Convert.ToInt32(objform.Items.Item("EtTo").Specific.Value) < Convert.ToInt32(objform.Items.Item("EtFrom").Specific.Value))
                                                {
                                                    objSBOAPI.SBO_Appln.StatusBar.SetText("To Date must be greater than From Date...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                                    bubbleevent = false;
                                                    objform.Items.Item("EtTo").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                                    return;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            break;
                    }
                }
                else
                {
                    switch (pval.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                            oCombo = objform.Items.Item("CbType").Specific;
                            string Type1 = objform.DataSources.UserDataSources.Item("Type").Value;
                            if (!string.IsNullOrEmpty(Type1))
                            {
                                if (pval.ItemUID == "BtnBrowse")
                                {
                                    string MacroType = objform.DataSources.UserDataSources.Item("Macro").Value;

                                    if (MacroType == "Y")
                                    {
                                        objSBOAPI.Browse1("Open", objform.UniqueID, "macro", "EtFile");
                                    }
                                    else
                                    {
                                        objSBOAPI.Browse1("Open", objform.UniqueID, "excel", "EtFile");
                                    }
                                }
                                else if (pval.ItemUID == "BtnExport")
                                {
                                    string MacroType = objform.DataSources.UserDataSources.Item("Macro").Value;

                                    string FileFormat = ".xlsx";
                                    string FileType = "excel";
                                    if (MacroType == "Y")
                                    {
                                        FileFormat = ".xlsm";
                                        FileType = "macro";
                                    }

                                    switch (Type1)
                                    {
                                        case "1":
                                            objSBOAPI.Browse1("Save", "", FileType, "", "SampleTemplate_ItemCreation" + FileFormat, false);
                                            break;
                                        //case "2":
                                        //    objSBOAPI.Browse1("Save", "", FileType, "", "SampleTemplate_MinMaxUpload" + FileFormat, true);
                                        //    break;
                                        //case "3":
                                        //    objSBOAPI.Browse1("Save", "", FileType, "", "SampleTemplate_PriceUpload" + FileFormat, true);
                                        //    break;
                                        //case "4":
                                        //    objSBOAPI.Browse1("Save", "", FileType, "", "SampleTemplate_PreferredVendor" + FileFormat, true);
                                        //    break;
                                        //case "5":
                                        //    objSBOAPI.Browse1("Save", "", FileType, "", "SampleTemplate_DeletePreferredVendor" + FileFormat, true);
                                        //    break;
                                    }
                                }
                                else if (pval.ItemUID == "BtnCreate")
                                {
                                    //var timer = new System.Diagnostics.Stopwatch();
                                    //timer.Start();

                                    System_Threading(Type1);

                                    //timer.Stop();
                                    //TimeSpan timeTaken = timer.Elapsed;
                                    //string foo = "Time taken: " + timeTaken.ToString(@"m\:ss\.fff");
                                }
                            }
                            else
                            {
                                objSBOAPI.SBO_Appln.StatusBar.SetText("Please Select Type...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            }

                            break;

                        case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:

                            if (pval.ItemUID == "CbType" && pval.ItemChanged == true)
                            {
                                Type1 = objform.DataSources.UserDataSources.Item("Type").Value;
                                //if (Type1 == "3" || Type1 == "1")
                                if (Type1 == "1")
                                {
                                    objform.Items.Item("StFrom").Visible = true;
                                    objform.Items.Item("EtFrom").Visible = true;
                                    objform.Items.Item("StTo").Visible = true;
                                    objform.Items.Item("EtTo").Visible = true;
                                }
                                //else
                                //{
                                //    objform.Items.Item("StFrom").Visible = false;
                                //    objform.Items.Item("EtFrom").Visible = false;
                                //    objform.Items.Item("StTo").Visible = false;
                                //    objform.Items.Item("EtTo").Visible = false;
                                //}
                                objform.DataSources.UserDataSources.Item("FromDate").Value = "";
                                objform.DataSources.UserDataSources.Item("ToDate").Value = "";
                                objform.DataSources.UserDataSources.Item("FileName").Value = "";
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
                        case "AV_OITMM":
                            objform = objSBOAPI.LoadForm("ItemCreation.xml", "AVA_OITMF");
                            SAPbouiCOM.ComboBox oCombo = objform.Items.Item("CbType").Specific;
                            oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
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

        #region System Threading
        public bool System_Threading(string Type)
        {
            try
            {
                System.Threading.Timer tm = new System.Threading.Timer(new TimerCallback(KeepUIAlive));
                tm.Change(1000 * 60, 0);

                Thread t = null;
                switch (Type)
                {
                    case "1":
                        ItemMasterCreation();
                        break;
                    //case "2":
                    //    MinMaxUpload();
                    //    break;
                    //case "3":
                    //    PricelistImport();
                    //    break;
                    //case "4":
                    //    PreferredVendor();
                    //    break;
                    //case "5":
                    //    PreferredVendorDeletion();
                    //    break;
                }
                t.TrySetApartmentState(System.Threading.ApartmentState.STA);
                t.Start();
                t.Join();
                tm.Dispose();
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
        public void KeepUIAlive(object info)
        {
            System.Threading.Timer t = (System.Threading.Timer)info;
            objSBOAPI.SBO_Appln.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, true);
            t.Change(1000 * 60, 0);
        }
        #endregion

        #region Item Master Creation
        public void ItemMasterCreation()
        {
            try
            {
                string FileName = objform.DataSources.UserDataSources.Item("FileName").Value;
                if (System.IO.File.Exists(FileName))
                {
                    XSSFWorkbook hssfworkbook = new XSSFWorkbook(FileName);
                    ISheet sheet = hssfworkbook.GetSheetAt(0);

                    DataTable dt = new DataTable();
                    IRow headerRow = sheet.GetRow(0);
                    IEnumerator rows = sheet.GetRowEnumerator();

                    int colCount = headerRow.LastCellNum;
                    int rowCount = sheet.LastRowNum;

                    for (int c = 0; c < colCount; c++)
                        dt.Columns.Add(headerRow.GetCell(c).ToString());

                    while (rows.MoveNext())
                    {
                        IRow row = (XSSFRow)rows.Current;

                        ICell cell = row.GetCell(0);
                        if (cell != null)
                        {
                            if (!string.IsNullOrEmpty(cell.ToString()) && cell.ToString().ToUpper() != "ITEMCODE")
                            {
                                DataRow dr = dt.NewRow();

                                for (int i = 0; i < colCount; i++)
                                {
                                    cell = row.GetCell(i);
                                    if (cell != null)
                                        dr[i] = cell.ToString();
                                }
                                dt.Rows.Add(dr);
                            }
                        }
                    }

                    hssfworkbook.Close();

                    int DocEntry = Convert.ToInt32(objSBOAPI.Query_Execute("SELECT IFNULL(Max(\"U_AVA_DocEntry\"),0) + 1 FROM \"@AVA_ITLOG\""));
                    //int PriceListDocEntry = Convert.ToInt32(objSBOAPI.Query_Execute("SELECT IFNULL(Max(\"U_AVA_DocEntry\"),0) + 1 FROM \"@AVA_PLLOG\""));
                    //Item_Master_Data_Creation(dt, FileName, DocEntry, PriceListDocEntry);
                    Item_Master_Data_Creation(dt, FileName, DocEntry);

                    Uploaded_Result("ItemMaster", DocEntry, FileName);
                    //Uploaded_Result("PriceList", PriceListDocEntry, FileName, "I");

                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Item Master Data Creation
        public void Item_Master_Data_Creation(DataTable oDt, string FileName, int DocEntry)
        {
            HanaConnection oDBConnection = null;
            try
            {
                string oConn = string.Format("Server = {0}; UserID = {1}; Password = {2}", objSBOAPI.objMain.HANA_ServerName, objSBOAPI.objMain.HANA_UserID, objSBOAPI.objMain.HANA_Pwd);
                oDBConnection = new HanaConnection(oConn);
                oDBConnection.Open();

                HanaCommand da;
                da = new HanaCommand("SET SCHEMA " + objSBOAPI.oCompany.CompanyDB, oDBConnection);
                da.ExecuteNonQuery();

                SAPbobsCOM.Items oItem;
                int SuccessCount = 0;
                int ErrorCount = 0;

                int Code = Convert.ToInt32(objSBOAPI.Query_Execute("SELECT IFNULL(Max(\"Code\"),0) + 1 FROM \"@AVA_ITLOG\""));

                //PriceListCode = Convert.ToInt32(objSBOAPI.Query_Execute("SELECT IFNULL(Max(\"Code\"),0) + 1 FROM \"@AVA_PLLOG\""));

                int oProcessLineNo = 1;

                List<DataColumn> oPriceData = (from DataColumn dc in oDt.Columns where dc.ColumnName.ToUpper().Contains("PR_") select dc).ToList();
                IDictionary<string, string> oPriceLists = new Dictionary<string, string>();
                foreach (DataColumn oColumn in oPriceData)
                {
                    string ListNum = objSBOAPI.HANA_GetValue("Select \"ListNum\" from \"OPLN\" where \"ListName\" = '" + oColumn.ToString().Replace("PR_", "") + "'", oDBConnection);
                    oPriceLists.Add(oColumn.ToString(), ListNum);
                }

                string FromDate = ((SAPbouiCOM.EditText)objform.Items.Item("EtFrom").Specific).Value;
                string ToDate = ((SAPbouiCOM.EditText)objform.Items.Item("EtTo").Specific).Value;

                foreach (DataRow oRow in oDt.Rows)
                {
                    string ItemCode = "";
                    string ItemName = "";
                    int ItemGroup = 0;
                    string ProdDesc = "";
                    int FirmCode = 0;
                    string CharDesc = "";
                    string CapDesc = "";
                    string Status = "F";
                    string RespMsg = "";

                    bool oExistItem = false;
                    try
                    {
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait Item Master is Uploading for Line No : " + oProcessLineNo.ToString() + " Out of (" + oDt.Rows.Count.ToString() + ")", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                        oItem = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);

                        oExistItem = oItem.GetByKey(oRow["ItemCode"].ToString()) ? true : false;

                        List<DataColumn> oData = (from DataColumn dc in oDt.Columns where !dc.ColumnName.ToUpper().Contains("PR_") select dc).ToList();

                        foreach (DataColumn oColumn in oData)
                        {
                            switch (oColumn.ToString())
                            {
                                case "Series":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.Series = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "ItemCode":
                                    oItem.ItemCode = oRow[oColumn].ToString();
                                    break;
                                case "ItemName":
                                    oItem.ItemName = oRow[oColumn].ToString();
                                    ItemName = oRow[oColumn].ToString();
                                    break;
                                case "FrgnName":
                                    oItem.ForeignName = oRow[oColumn].ToString();
                                    break;
                                case "ItemType":
                                    if (oRow[oColumn].ToString() == "I")
                                    {
                                        oItem.ItemType = SAPbobsCOM.ItemTypeEnum.itItems;
                                    }
                                    else if (oRow[oColumn].ToString() == "L")
                                    {
                                        oItem.ItemType = SAPbobsCOM.ItemTypeEnum.itLabor;
                                    }
                                    else if (oRow[oColumn].ToString() == "T")
                                    {
                                        oItem.ItemType = SAPbobsCOM.ItemTypeEnum.itTravel;
                                    }
                                    break;
                                case "ItmsGrpCod":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.ItemsGroupCode = Convert.ToInt32(oRow[oColumn].ToString());

                                        ItemGroup = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "ItemClass":
                                    if (oRow[oColumn].ToString() == "2")
                                    {
                                        oItem.ItemClass = SAPbobsCOM.ItemClassEnum.itcMaterial;
                                    }
                                    else if (oRow[oColumn].ToString() == "1")
                                    {
                                        oItem.ItemClass = SAPbobsCOM.ItemClassEnum.itcService;
                                    }
                                    break;
                                case "FirmCode":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        string oFirmCode = objSBOAPI.Query_Execute("Select \"FirmCode\" from \"OMRC\" Where \"FirmName\" = '" + oRow[oColumn].ToString() + "'");
                                        if (!string.IsNullOrEmpty(oFirmCode))
                                        {
                                            oItem.Manufacturer = Convert.ToInt32(oFirmCode);
                                            FirmCode = Convert.ToInt32(oFirmCode);
                                        }
                                    }
                                    break;
                                case "SWW":
                                    oItem.SWW = oRow[oColumn].ToString();
                                    break;
                                case "ShipType":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.ShipType = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "ManSerNum":
                                    if (oRow[oColumn].ToString() == "Y")
                                    {
                                        oItem.ManageSerialNumbers = SAPbobsCOM.BoYesNoEnum.tYES;
                                    }
                                    else
                                    {
                                        oItem.ManageSerialNumbers = SAPbobsCOM.BoYesNoEnum.tNO;
                                    }
                                    break;
                                case "MngMethod":
                                    if (oRow[oColumn].ToString() == "R")
                                    {
                                        oItem.SRIAndBatchManageMethod = SAPbobsCOM.BoManageMethod.bomm_OnReleaseOnly;
                                    }
                                    else if (oRow[oColumn].ToString() == "A")
                                    {
                                        oItem.SRIAndBatchManageMethod = SAPbobsCOM.BoManageMethod.bomm_OnEveryTransaction;
                                    }
                                    break;
                                case "ManOutOnly":
                                    if (oRow[oColumn].ToString() == "Y")
                                    {
                                        oItem.ManageSerialNumbersOnReleaseOnly = SAPbobsCOM.BoYesNoEnum.tYES;
                                    }
                                    else
                                    {
                                        oItem.ManageSerialNumbersOnReleaseOnly = SAPbobsCOM.BoYesNoEnum.tNO;
                                    }
                                    break;
                                case "IssuePriBy":
                                    if (oRow[oColumn].ToString() == "0")
                                    {
                                        oItem.IssuePrimarilyBy = SAPbobsCOM.IssuePrimarilyByEnum.ipbSerialAndBatchNumbers;
                                    }
                                    else if (oRow[oColumn].ToString() == "1")
                                    {
                                        oItem.IssuePrimarilyBy = SAPbobsCOM.IssuePrimarilyByEnum.ipbBinLocations;
                                    }
                                    break;
                                case "WarrntTmpl":
                                    oItem.WarrantyTemplate = oRow[oColumn].ToString();
                                    break;

                                case "ManBtchNum":
                                    if (oRow[oColumn].ToString() == "Y")
                                    {
                                        oItem.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tYES;
                                    }
                                    else
                                    {
                                        oItem.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tNO;
                                    }
                                    break;

                                case "UgpEntry":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.UoMGroupEntry = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;

                                case "CodeBars":
                                    oItem.BarCode = oRow[oColumn].ToString();
                                    break;

                                case "InvntItem":
                                    if (oRow[oColumn].ToString() == "Y")
                                    {
                                        oItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tYES;
                                    }
                                    else
                                    {
                                        oItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tNO;
                                    }
                                    break;

                                case "SellItem":
                                    if (oRow[oColumn].ToString() == "Y")
                                    {
                                        oItem.SalesItem = SAPbobsCOM.BoYesNoEnum.tYES;
                                    }
                                    else
                                    {
                                        oItem.SalesItem = SAPbobsCOM.BoYesNoEnum.tNO;
                                    }
                                    break;

                                case "PrchseItem":
                                    if (oRow[oColumn].ToString() == "Y")
                                    {
                                        oItem.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tYES;
                                    }
                                    else
                                    {
                                        oItem.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tNO;
                                    }
                                    break;

                                case "validFor":
                                    if (oRow[oColumn].ToString() == "Y")
                                    {
                                        oItem.Valid = SAPbobsCOM.BoYesNoEnum.tYES;
                                    }
                                    else
                                    {
                                        oItem.Valid = SAPbobsCOM.BoYesNoEnum.tNO;
                                    }
                                    break;

                                case "validFrom":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.ValidFrom = objSBOAPI.GetDateTimeValue(oRow[oColumn].ToString());
                                    }
                                    break;

                                case "validTo":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.ValidTo = objSBOAPI.GetDateTimeValue(oRow[oColumn].ToString());
                                    }
                                    break;

                                case "ValidComm":
                                    oItem.ValidRemarks = oRow[oColumn].ToString();
                                    break;

                                case "frozenFor":
                                    if (oRow[oColumn].ToString() == "Y")
                                    {
                                        oItem.Frozen = SAPbobsCOM.BoYesNoEnum.tYES;
                                    }
                                    else
                                    {
                                        oItem.Frozen = SAPbobsCOM.BoYesNoEnum.tNO;
                                    }
                                    break;

                                case "frozenFrom":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.FrozenFrom = objSBOAPI.GetDateTimeValue(oRow[oColumn].ToString());
                                    }
                                    break;

                                case "frozenTo":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.FrozenTo = objSBOAPI.GetDateTimeValue(oRow[oColumn].ToString());
                                    }
                                    break;

                                case "FrozenComm":
                                    oItem.FrozenRemarks = oRow[oColumn].ToString();
                                    break;

                                case "CardCode":
                                    oItem.Mainsupplier = oRow[oColumn].ToString();
                                    break;

                                case "SuppCatNum":
                                    oItem.SupplierCatalogNo = oRow[oColumn].ToString();
                                    break;

                                case "BuyUnitMsr":
                                    oItem.PurchaseUnit = oRow[oColumn].ToString();
                                    break;

                                case "NumInBuy":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseItemsPerUnit = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;

                                case "PurPackMsr":
                                    oItem.PurchasePackagingUnit = oRow[oColumn].ToString();
                                    break;

                                case "PurPackUn":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseQtyPerPackUnit = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;

                                case "PurFactor1":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseFactor1 = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;

                                case "PurFactor2":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseFactor2 = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;

                                case "PurFactor3":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseFactor3 = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;

                                case "PurFactor4":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseFactor4 = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;

                                case "BWght1Unit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseWeightUnit = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "BWght2Unit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseWeightUnit1 = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;

                                case "BWeight1":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseUnitWeight = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "BWeight2":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseUnitWeight1 = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "BLen1Unit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseLengthUnit = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;

                                case "BLen2Unit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseLengthUnit1 = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "BLength1":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseUnitLength = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "Blength2":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseUnitLength1 = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "BHght1Unit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseHeightUnit = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "BHght2Unit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseHeightUnit1 = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "BHeight1":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseUnitHeight = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "BHeight2":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseUnitHeight1 = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "BWdth1Unit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseWidthUnit = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "BWdth2Unit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseWidthUnit1 = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "BWidth1":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseUnitWidth = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "BWidth2":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseUnitWidth1 = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "BVolUnit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseVolumeUnit = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "BVolume":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.PurchaseUnitVolume = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "CstGrpCode":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.CustomsGroupCode = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;

                                case "VatGroupPu":
                                    oItem.PurchaseVATGroup = oRow[oColumn].ToString();
                                    break;

                                case "VatGourpSa":

                                    oItem.SalesVATGroup = oRow[oColumn].ToString();
                                    break;
                                case "CommisGrp":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.CommissionGroup = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "CommisPcnt":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.CommissionPercent = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SalUnitMsr":

                                    oItem.SalesUnit = oRow[oColumn].ToString();
                                    break;

                                case "NumInSale":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesItemsPerUnit = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SalPackMsr":
                                    oItem.SalesPackagingUnit = oRow[oColumn].ToString();
                                    break;
                                case "SalPackUn":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesQtyPerPackUnit = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SalFactor1":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesFactor1 = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SalFactor2":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesFactor2 = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SalFactor3":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesFactor3 = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SalFactor4":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesFactor4 = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SWght1Unit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesWeightUnit = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SWght2Unit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesWeightUnit1 = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SWeight1":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesUnitWeight = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SWeight2":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesUnitWeight1 = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SLen1Unit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesLengthUnit = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SLen2Unit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesLengthUnit1 = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SLength1":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesUnitLength = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "Slength2":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesUnitLength1 = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SHght1Unit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesHeightUnit = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SHght2Unit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesHeightUnit1 = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SHeight1":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesUnitHeight = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SHeight2":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesUnitHeight1 = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SWdth1Unit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesWidthUnit = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SWdth2Unit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesWidthUnit1 = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SWidth1":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesUnitWidth = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SWidth2":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesUnitWidth1 = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SVolUnit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesVolumeUnit = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "SVolume":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.SalesUnitVolume = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;

                                case "GLMethod":
                                    if (oRow[oColumn].ToString() == "W")
                                    {
                                        oItem.GLMethod = SAPbobsCOM.BoGLMethods.glm_WH;
                                    }
                                    else if (oRow[oColumn].ToString() == "C")
                                    {
                                        oItem.GLMethod = SAPbobsCOM.BoGLMethods.glm_ItemClass;
                                    }
                                    else if (oRow[oColumn].ToString() == "L")
                                    {
                                        oItem.GLMethod = SAPbobsCOM.BoGLMethods.glm_ItemLevel;
                                    }
                                    break;
                                case "InvntryUom":

                                    oItem.InventoryUOM = oRow[oColumn].ToString();
                                    break;

                                case "IWght1Unit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.InventoryWeightUnit = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "IWght2Unit":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.InventoryWeightUnit1 = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "IWeight1":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.InventoryWeight = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "IWeight2":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.InventoryWeight1 = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "EvalSystem":
                                    if (oRow[oColumn].ToString() == "A")
                                    {
                                        oItem.CostAccountingMethod = SAPbobsCOM.BoInventorySystem.bis_MovingAverage;
                                    }
                                    else if (oRow[oColumn].ToString() == "S")
                                    {
                                        oItem.CostAccountingMethod = SAPbobsCOM.BoInventorySystem.bis_Standard;
                                    }
                                    else if (oRow[oColumn].ToString() == "F")
                                    {
                                        oItem.CostAccountingMethod = SAPbobsCOM.BoInventorySystem.bis_FIFO;
                                    }
                                    else if (oRow[oColumn].ToString() == "B")
                                    {
                                        oItem.CostAccountingMethod = SAPbobsCOM.BoInventorySystem.bis_SNB;
                                    }
                                    break;
                                case "AvgPrice":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.AvgStdPrice = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "ByWh":
                                    if (oRow[oColumn].ToString() == "Y")
                                    {
                                        oItem.ManageStockByWarehouse = SAPbobsCOM.BoYesNoEnum.tYES;
                                    }
                                    else
                                    {
                                        oItem.ManageStockByWarehouse = SAPbobsCOM.BoYesNoEnum.tNO;
                                    }
                                    break;
                                case "ReorderQty":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.DesiredInventory = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "MinLevel":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.MinInventory = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "MaxLevel":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.MaxInventory = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;
                                //------------------Planning Data-------------------------------------------
                                case "PlaningSys":
                                    if (oRow[oColumn].ToString() == "M")
                                    {
                                        oItem.PlanningSystem = SAPbobsCOM.BoPlanningSystem.bop_MRP;
                                    }
                                    else
                                    {
                                        oItem.PlanningSystem = SAPbobsCOM.BoPlanningSystem.bop_None;
                                    }
                                    break;
                                case "PrcrmntMtd":
                                    if (oRow[oColumn].ToString() == "M")
                                    {
                                        oItem.ProcurementMethod = SAPbobsCOM.BoProcurementMethod.bom_Make;
                                    }
                                    else
                                    {
                                        oItem.ProcurementMethod = SAPbobsCOM.BoProcurementMethod.bom_Buy;
                                    }
                                    break;
                                case "CompoWH":
                                    if (oRow[oColumn].ToString() == "B")
                                    {
                                        oItem.ComponentWarehouse = SAPbobsCOM.BoMRPComponentWarehouse.bomcw_BOM;
                                    }
                                    else if (oRow[oColumn].ToString() == "P")
                                    {
                                        oItem.ComponentWarehouse = SAPbobsCOM.BoMRPComponentWarehouse.bomcw_Parent;
                                    }
                                    break;

                                case "OrdrIntrvl":
                                    string OrderCode = objSBOAPI.Query_Execute("Select \"Code\" from OCYC where \"Code\" = '" + oRow[oColumn].ToString() + "'");
                                    if (!string.IsNullOrEmpty(OrderCode))
                                    {
                                        oItem.OrderIntervals = OrderCode;
                                    }
                                    break;

                                case "OrdrMulti":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.OrderMultiple = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;

                                case "MinOrdrQty":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.MinOrderQuantity = Convert.ToDouble(oRow[oColumn].ToString());
                                    }
                                    break;

                                case "LeadTime":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.LeadTime = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;
                                case "ToleranDay":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        oItem.ToleranceDays = Convert.ToInt32(oRow[oColumn].ToString());
                                    }
                                    break;

                                //--------------Production Data---------------------
                                case "Phantom":
                                    if (oRow[oColumn].ToString() == "Y")
                                    {
                                        oItem.IsPhantom = SAPbobsCOM.BoYesNoEnum.tYES;
                                    }
                                    else
                                    {
                                        oItem.IsPhantom = SAPbobsCOM.BoYesNoEnum.tNO;
                                    }
                                    break;
                                case "IssueMthd":
                                    if (oRow[oColumn].ToString() == "B")
                                    {
                                        oItem.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Backflush;
                                    }
                                    else if (oRow[oColumn].ToString() == "M")
                                    {
                                        oItem.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Manual;
                                    }
                                    break;
                                case "UserText":
                                    oItem.User_Text = oRow[oColumn].ToString();
                                    break;

                                case "GSTRelevnt":
                                    if (oRow[oColumn].ToString() == "Y")
                                    {
                                        oItem.GSTRelevnt = SAPbobsCOM.BoYesNoEnum.tYES;
                                    }
                                    else
                                    {
                                        oItem.GSTRelevnt = SAPbobsCOM.BoYesNoEnum.tNO;
                                    }
                                    break;

                                case "ChapterID":
                                    if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                    {
                                        string AbsEntry = objSBOAPI.Query_Execute("Select \"AbsEntry\" from \"OCHP\" Where \"ChapterID\" = '" + oRow[oColumn].ToString() + "'");
                                        if (!string.IsNullOrEmpty(AbsEntry))
                                        {
                                            oItem.ChapterID = Convert.ToInt32(AbsEntry);
                                        }
                                    }
                                    break;

                                case "GstTaxCtg":

                                    if (oRow[oColumn].ToString() == "R")
                                    {
                                        oItem.GSTTaxCategory = SAPbobsCOM.GSTTaxCategoryEnum.gtc_Regular;
                                    }
                                    else if (oRow[oColumn].ToString() == "N")
                                    {
                                        oItem.GSTTaxCategory = SAPbobsCOM.GSTTaxCategoryEnum.gtc_NilRated;
                                    }
                                    else if (oRow[oColumn].ToString() == "E")
                                    {
                                        oItem.GSTTaxCategory = SAPbobsCOM.GSTTaxCategoryEnum.gtc_Exempt;
                                    }
                                    break;

                                case "MatType":
                                    if (oRow[oColumn].ToString() == "1")
                                    {
                                        oItem.MaterialType = SAPbobsCOM.BoMaterialTypes.mt_FinishedGoods;
                                    }
                                    else if (oRow[oColumn].ToString() == "2")
                                    {
                                        oItem.MaterialType = SAPbobsCOM.BoMaterialTypes.mt_GoodsInProcess;
                                    }
                                    else if (oRow[oColumn].ToString() == "3")
                                    {
                                        oItem.MaterialType = SAPbobsCOM.BoMaterialTypes.mt_RawMaterial;
                                    }
                                    break;

                                default:
                                    bool UserFields = oColumn.ToString().StartsWith("U_");
                                    if (UserFields == true)
                                    {
                                        int UDF_Available = Convert.ToInt32(objSBOAPI.Query_Execute("Select Count(*) from CUFD where \"TableID\" = 'OITM' and \"AliasID\" = '" + oColumn.ToString().Replace("U_", "") + "'"));
                                        if (UDF_Available > 0)
                                        {
                                            oItem.UserFields.Fields.Item("" + oColumn.ToString() + "").Value = oRow[oColumn].ToString();

                                            switch (oColumn.ToString())
                                            {
                                                case "U_AVA_PRODDESC":
                                                    ProdDesc = oRow[oColumn].ToString();
                                                    break;

                                                case "U_AVA_CHARDESC":
                                                    CharDesc = oRow[oColumn].ToString();
                                                    break;

                                                case "U_AVA_CAPCITY":
                                                    CapDesc = oRow[oColumn].ToString();
                                                    break;
                                            }
                                        }
                                    }

                                    bool Property = oColumn.ToString().Contains("Property");
                                    if (Property == true)
                                    {
                                        int PropCode = Convert.ToInt32(oColumn.ToString().Replace("Property", ""));

                                        if (oRow[oColumn].ToString() == "Y")
                                        {
                                            oItem.Properties[PropCode] = SAPbobsCOM.BoYesNoEnum.tYES;
                                        }
                                        else
                                        {
                                            oItem.Properties[PropCode] = SAPbobsCOM.BoYesNoEnum.tNO;
                                        }
                                    }
                                    break;
                            }
                        }

                        RetCode = oExistItem ? oItem.Update() : oItem.Add();

                        if (RetCode != 0)
                        {
                            RespMsg = objSBOAPI.oCompany.GetLastErrorCode().ToString() + " - " + objSBOAPI.oCompany.GetLastErrorDescription();
                            ErrorCount = ErrorCount + 1;
                        }
                        else
                        {
                            Status = "S";
                            ItemCode = objSBOAPI.oCompany.GetNewObjectKey();
                            RespMsg = oExistItem ? "Updated Successfully" : "Added Successfully";
                            SuccessCount = SuccessCount + 1;

                            if (oPriceLists.Count > 0)
                            {
                                //ItemMaster_PriceImport(PriceListDocEntry, FileName, ItemCode, oPriceLists, oRow, FromDate, ToDate, oDBConnection);
                                ItemMaster_PriceImport(FileName, ItemCode, oPriceLists, oRow, FromDate, ToDate, oDBConnection);
                            }

                        }
                    }
                    catch (Exception ex)
                    {
                        RespMsg = ex.Message;
                        ErrorCount = ErrorCount + 1;
                    }
                    finally
                    {
                        string InsertQuery = "INSERT INTO \"@AVA_ITLOG\"(\"Code\",\"Name\",\"U_AVA_CreateDate\",\"U_AVA_FileName\",\"U_AVA_DocEntry\",\"U_AVA_ItmCod\",\"U_AVA_ItmNam\",\"U_AVA_ItmGrp\",\"U_AVA_PrdDesc\",\"U_AVA_FirmCod\",\"U_AVA_CharDesc\",\"U_AVA_CapDesc\",\"U_AVA_Status\",\"U_AVA_RespMsg\") VALUES ('" + Code + "','" + objSBOAPI.oCompany.UserName + "','" + DateTime.Now.ToString("dd MMM yyyy HH:mm:ss") + "','" + FileName + "','" + DocEntry + "','" + ItemCode + "','" + ItemName + "','" + ItemGroup + "','" + ProdDesc + "','" + FirmCode + "','" + CharDesc + "','" + CapDesc + "','" + Status + "','" + RespMsg + "')";
                        objSBOAPI.Query_Execute(InsertQuery);

                        Code = Code + 1;

                        oProcessLineNo = oProcessLineNo + 1;
                        GC.Collect();
                    }
                }

                oDBConnection.Close();
                oDBConnection.Dispose();

                objSBOAPI.SBO_Appln.StatusBar.SetText("Item Master Uploaded Status : Total - " + oDt.Rows.Count.ToString() + " / Success - " + SuccessCount.ToString() + " / Failed - " + ErrorCount.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Item Master - Price List Import

        public void ItemMaster_PriceImport(string FileName, string ItemCode, IDictionary<string, string> oPriceLists, DataRow oRow, string FromDate, string ToDate, HanaConnection oDBConnection)
        {
            try
            {
                foreach (var oField in oPriceLists)
                {
                    string ListName = oField.Key;
                    string ListNum = oField.Value;
                    string Price = oRow[ListName].ToString();
                    ListName = ListName.Replace("PR_", "");
                    string RespMsg = "Failed";
                    string Status = "F";
                    try
                    {
                        if (!string.IsNullOrEmpty(ListNum))
                        {
                            objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait PriceList is Uploading for Item Code : " + ItemCode + " and ListName : " + ListName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                            string UpdateQuery = "Update \"ITM1\" Set \"Price\" = '" + Price + "' Where \"ItemCode\" = '" + ItemCode + "' And \"PriceList\" = '" + ListNum + "'";
                            if (objSBOAPI.HANA_ExecuteNonQuery(UpdateQuery, oDBConnection) == 1)
                            {
                                Status = "S";
                                RespMsg = "Successfully Updated";
                            }
                            else
                            {
                                RespMsg = "Update Failed";
                            }
                        }
                        else
                        {
                            RespMsg = "Invalid PriceList";
                        }
                    }
                    catch (Exception ex)
                    {
                        RespMsg = ex.Message;
                    }
                    finally
                    {
                        //string InsertQuery = "INSERT INTO \"@AVA_PLLOG\"(\"Code\",\"Name\",\"U_AVA_CreateDate\",\"U_AVA_FileName\",\"U_AVA_DocEntry\",\"U_AVA_Type\",\"U_AVA_ItmCod\",\"U_AVA_ListName\",\"U_AVA_FromDate\",\"U_AVA_ToDate\",\"U_AVA_Price\",\"U_AVA_Status\",\"U_AVA_RespMsg\") VALUES ('" + PriceListCode + "','" + objSBOAPI.oCompany.UserName + "','" + DateTime.Now.ToString("dd MMM yyyy HH:mm:ss") + "','" + FileName + "','" + DocEntry + "','I','" + ItemCode + "','" + ListName + "','" + FromDate + "','" + ToDate + "','" + Price + "','" + Status + "','" + RespMsg + "')";
                        //PriceListCode = PriceListCode + 1;
                        //objSBOAPI.HANA_ExecuteNonQuery(InsertQuery, oDBConnection);
                    }
                }
            }
            catch (Exception ex)
            {
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region MinMaxUpload
        public void MinMaxUpload()
        {
            try
            {
                objform.Freeze(true);
                string FileName = objform.DataSources.UserDataSources.Item("FileName").Value;
                if (System.IO.File.Exists(FileName))
                {
                    XSSFWorkbook hssfworkbook = new XSSFWorkbook(FileName);
                    ISheet sheet = hssfworkbook.GetSheetAt(0);

                    DataTable dt = new DataTable();
                    IRow headerRow = sheet.GetRow(0);
                    IEnumerator rows = sheet.GetRowEnumerator();

                    int colCount = headerRow.LastCellNum;
                    int rowCount = sheet.LastRowNum;

                    for (int c = 0; c < colCount; c++)
                        dt.Columns.Add(headerRow.GetCell(c).ToString());

                    while (rows.MoveNext())
                    {
                        IRow row = (XSSFRow)rows.Current;

                        ICell cell = row.GetCell(0);
                        if (cell != null)
                        {
                            if (!string.IsNullOrEmpty(cell.ToString()) && cell.ToString().ToUpper() != "ITEMCODE")
                            {
                                DataRow dr = dt.NewRow();

                                for (int i = 0; i < colCount; i++)
                                {
                                    cell = row.GetCell(i);
                                    if (cell != null)
                                        dr[i] = cell.ToString();
                                }
                                dt.Rows.Add(dr);
                            }
                        }
                    }

                    dt.Columns.Add("RespMsg");

                    hssfworkbook.Close();

                    int DocEntry = Convert.ToInt32(objSBOAPI.Query_Execute("SELECT IFNULL(Max(\"U_AVA_DocEntry\"),0) + 1 FROM \"@AVA_MULOG\""));
                    Item_Quantity_Update(dt, FileName, DocEntry);

                    Uploaded_Result("MinMax", DocEntry, FileName);

                }
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Item Quantity Update
        public void Item_Quantity_Update(DataTable oDt, string FileName, int DocEntry)
        {
            try
            {
                SAPbobsCOM.Items oItem;
                SAPbobsCOM.ItemWarehouseInfo oItemWhs;
                int SuccessCount = 0;
                int ErrorCount = 0;

                int Code = Convert.ToInt32(objSBOAPI.Query_Execute("SELECT IFNULL(Max(\"Code\"),0) + 1 FROM \"@AVA_MULOG\""));

                string MinMaxQuery = "";
                MinMaxQuery = objSBOAPI.Query_Execute("SELECT \"QString\" FROM OUQR WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM OQCN WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='MinMaxUpload'");
                if (MinMaxQuery == "")
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("MinMaxUpload Query is Missing, Kindly do the Query Manager Update", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return;
                }
                else
                {
                    HanaConnection oDBConnection = null;
                    string oConn = string.Format("Server = {0}; UserID = {1}; Password = {2}", objSBOAPI.objMain.HANA_ServerName, objSBOAPI.objMain.HANA_UserID, objSBOAPI.objMain.HANA_Pwd);
                    oDBConnection = new HanaConnection(oConn);
                    oDBConnection.Open();
                    HanaCommand da;
                    da = new HanaCommand("SET SCHEMA " + objSBOAPI.oCompany.CompanyDB, oDBConnection);
                    da.ExecuteNonQuery();

                    IDictionary<string, string> oWhsLine = new Dictionary<string, string>();

                    DataTable oWhsListData = objSBOAPI.HANA_ExecuteQuery(MinMaxQuery, oDBConnection);

                    if (oWhsListData != null)
                    {
                        int i = 0;
                        while (i < oWhsListData.Rows.Count)
                        {
                            DataRow oRow = oWhsListData.Rows[i];

                            string WhsName = oRow["WhsCode"].ToString();
                            string WhsLine = oRow["LineNum"].ToString();

                            oWhsLine.Add(WhsName, WhsLine);
                            i++;
                        }
                    }
                    if (oDBConnection.State == ConnectionState.Open)
                    {
                        oDBConnection.Close();
                        oDBConnection.Dispose();
                    }

                    int Total_Items = 0;
                    int oProcessLineNo = 1;
                    List<DataRow> oData = null;
                    if (oDt.Columns.Contains("ItemCode") && oDt.Columns.Contains("Warehouse"))
                    {
                        var Item_List = oDt.AsEnumerable().Select(d => new { ItemCode = d.Field<string>("ItemCode") }).Distinct().ToList();

                        Total_Items = Item_List.Count();

                        foreach (var oItm in Item_List)
                        {
                            string ItemCode = oItm.ItemCode;
                            string Status = "F";
                            string RespMsg = "";

                            try
                            {
                                objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait Min/Max is Uploading for Line No : " + oProcessLineNo.ToString() + " Out of (" + Item_List.Count.ToString() + ")", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                                oData = (from e in oDt.AsEnumerable()
                                         where e.Field<string>("ItemCode").Equals(ItemCode)
                                         select e).ToList();
                                if (oData != null)
                                {
                                    if (oData.Count() > 0)
                                    {
                                        oItem = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);

                                        if (oItem.GetByKey(ItemCode))
                                        {
                                            foreach (DataRow oRow in oData)
                                            {
                                                string WhsCode = oRow["Warehouse"].ToString();
                                                string LineNo = oWhsLine.Where(a => a.Key == WhsCode).Select(p => p.Value).FirstOrDefault();
                                                if (!string.IsNullOrEmpty(LineNo))
                                                {
                                                    oItemWhs = oItem.WhsInfo;
                                                    oItemWhs.SetCurrentLine(Convert.ToInt32(LineNo));

                                                    foreach (DataColumn oColumn in oDt.Columns)
                                                    {
                                                        switch (oColumn.ToString())
                                                        {
                                                            case "MinStock":
                                                                if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                                                {
                                                                    oItemWhs.MinimalStock = Convert.ToDouble(oRow[oColumn].ToString());
                                                                }
                                                                break;
                                                            case "MaxStock":
                                                                if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                                                {
                                                                    oItemWhs.MaximalStock = Convert.ToDouble(oRow[oColumn].ToString());
                                                                }
                                                                break;
                                                            case "PreferredVendor":
                                                                oItemWhs.UserFields.Fields.Item("U_AVA_APCreditVendor").Value = oRow[oColumn].ToString();
                                                                break;

                                                            case "ClubPO":
                                                                oItemWhs.UserFields.Fields.Item("U_AVA_ClubPO").Value = oRow[oColumn].ToString();
                                                                break;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    oRow["RespMsg"] = "No Warehouse or Inactive Warehouse";
                                                    ErrorCount = ErrorCount + 1;
                                                }
                                            }

                                            RetCode = oItem.Update();
                                            if (RetCode != 0)
                                            {
                                                RespMsg = objSBOAPI.oCompany.GetLastErrorCode().ToString() + " - " + objSBOAPI.oCompany.GetLastErrorDescription();
                                                RespMsg = Regex.Replace(RespMsg, @"[^0-9a-zA-Z ]+", "");
                                                ErrorCount = ErrorCount + 1;
                                            }
                                            else
                                            {
                                                Status = "S";
                                                RespMsg = "Updated Successfully";
                                                SuccessCount = SuccessCount + 1;
                                            }
                                        }
                                        else
                                        {
                                            RespMsg = "Invalid ItemCode";
                                            ErrorCount = ErrorCount + 1;
                                        }
                                    }
                                    else
                                    {
                                        RespMsg = "No data found against the ItemCode";
                                        ErrorCount = ErrorCount + 1;
                                    }
                                }
                                else
                                {
                                    RespMsg = "No data found against the ItemCode";
                                    ErrorCount = ErrorCount + 1;
                                }
                            }
                            catch (Exception ex)
                            {
                                RespMsg = ex.Message;
                                ErrorCount = ErrorCount + 1;
                            }
                            finally
                            {
                                foreach (DataRow oRow in oData)
                                {
                                    string WhsCode = oRow["Warehouse"].ToString();
                                    double MaxQty = 0.0;
                                    double MinQty = 0.0;
                                    string Vendor = "";
                                    string ClubPO = "";

                                    foreach (DataColumn oColumn in oDt.Columns)
                                    {
                                        switch (oColumn.ToString())
                                        {
                                            case "MinStock":
                                                if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                                {
                                                    MinQty = Convert.ToDouble(oRow[oColumn].ToString());
                                                }
                                                break;
                                            case "MaxStock":
                                                if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
                                                {
                                                    MaxQty = Convert.ToDouble(oRow[oColumn].ToString());
                                                }
                                                break;
                                            case "PreferredVendor":
                                                Vendor = oRow[oColumn].ToString();
                                                break;

                                            case "ClubPO":
                                                ClubPO = oRow[oColumn].ToString();
                                                break;
                                        }
                                    }

                                    string InsertQuery = "";
                                    if (!string.IsNullOrEmpty(oRow["RespMsg"].ToString()))
                                    {
                                        InsertQuery = "INSERT INTO \"@AVA_MULOG\"(\"Code\",\"Name\",\"U_AVA_CreateDate\",\"U_AVA_FileName\",\"U_AVA_DocEntry\",\"U_AVA_ItmCod\",\"U_AVA_WhsCod\",\"U_AVA_MinQty\",\"U_AVA_MaxQty\",\"U_AVA_VenCod\",\"U_AVA_ClubPO\",\"U_AVA_Status\",\"U_AVA_RespMsg\") VALUES ('" + Code + "','" + objSBOAPI.oCompany.UserName + "','" + DateTime.Now.ToString("dd MMM yyyy HH:mm:ss") + "','" + FileName + "','" + DocEntry + "','" + ItemCode + "','" + WhsCode + "','" + MinQty + "','" + MaxQty + "','" + Vendor + "','" + ClubPO + "','F','" + oRow["RespMsg"].ToString() + "')";
                                    }
                                    else
                                    {
                                        InsertQuery = "INSERT INTO \"@AVA_MULOG\"(\"Code\",\"Name\",\"U_AVA_CreateDate\",\"U_AVA_FileName\",\"U_AVA_DocEntry\",\"U_AVA_ItmCod\",\"U_AVA_WhsCod\",\"U_AVA_MinQty\",\"U_AVA_MaxQty\",\"U_AVA_VenCod\",\"U_AVA_ClubPO\",\"U_AVA_Status\",\"U_AVA_RespMsg\") VALUES ('" + Code + "','" + objSBOAPI.oCompany.UserName + "','" + DateTime.Now.ToString("dd MMM yyyy HH:mm:ss") + "','" + FileName + "','" + DocEntry + "','" + ItemCode + "','" + WhsCode + "','" + MinQty + "','" + MaxQty + "','" + Vendor + "','" + ClubPO + "','" + Status + "','" + RespMsg + "')";
                                    }

                                    objSBOAPI.Query_Execute(InsertQuery);

                                    Code = Code + 1;
                                }
                                oProcessLineNo = oProcessLineNo + 1;
                            }

                        }
                    }
                    else
                    {
                        objSBOAPI.SBO_Appln.StatusBar.SetText("ItemCode and Warehouse columns are mandatory in excel", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                objSBOAPI.SBO_Appln.StatusBar.SetText("Min/Max Uploaded Status : Total - " + oDt.Rows.Count.ToString() + " / Success - " + SuccessCount.ToString() + " / Failed - " + ErrorCount.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        //#region Item Quantity Update
        //public void Item_Quantity_Update(DataTable oDt,string FileName,int DocEntry)
        //{
        //    try
        //    {
        //        SAPbobsCOM.Items oItem;
        //        SAPbobsCOM.ItemWarehouseInfo oItemWhs;
        //        int SuccessCount = 0;
        //        int ErrorCount = 0;

        //        int Code = Convert.ToInt32(objSBOAPI.Query_Execute("SELECT COUNT(*) + 1 FROM \"@AVA_MULOG\""));

        //        string MinMaxQuery = "";
        //        MinMaxQuery = objSBOAPI.Query_Execute("SELECT \"QString\" FROM OUQR WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM OQCN WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='MinMaxUpload'");
        //        if (MinMaxQuery == "")
        //        {
        //            objSBOAPI.SBO_Appln.StatusBar.SetText("MinMaxUpload Query is Missing, Kindly do the Query Manager Update", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //            return;
        //        }
        //        else
        //        {
        //            if(oDt.Columns.Contains("ItemCode") && oDt.Columns.Contains("Warehouse"))
        //            {
        //                int oProcessLineNo = 1;
        //                foreach (DataRow oRow in oDt.Rows)
        //                {
        //                    string ItemCode = "";
        //                    string WhsCode = "";
        //                    double MaxQty = 0.0;
        //                    double MinQty = 0.0;
        //                    string Vendor = "";
        //                    string ClubPO = "";
        //                    string Status = "F";
        //                    string RespMsg = "";
        //                    try
        //                    {
        //                        objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait Min/Max is Uploading for Line No : " + oProcessLineNo.ToString() + " Out of (" + oDt.Rows.Count.ToString() + ")", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

        //                        ItemCode = oRow["ItemCode"].ToString();
        //                        WhsCode = oRow["Warehouse"].ToString();

        //                        string Str = MinMaxQuery;
        //                        Str = Str.Replace("[%1]", oRow["ItemCode"].ToString());
        //                        Str = Str.Replace("[%2]", WhsCode);

        //                        string LineNo = objSBOAPI.Query_Execute(Str);

        //                        if (!string.IsNullOrEmpty(LineNo))
        //                        {
        //                            oItem = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);

        //                            if (oItem.GetByKey(ItemCode) == true)
        //                            {
        //                                int LineId = Convert.ToInt32(LineNo);

        //                                oItem.WhsInfo.SetCurrentLine(LineId);
        //                                oItemWhs = oItem.WhsInfo;

        //                                foreach (DataColumn oColumn in oDt.Columns)
        //                                {
        //                                    switch (oColumn.ToString())
        //                                    {
        //                                        case "MinStock":
        //                                            if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
        //                                            {
        //                                                MinQty = Convert.ToDouble(oRow[oColumn].ToString());
        //                                                oItemWhs.MinimalStock = MinQty;
        //                                            }
        //                                            break;
        //                                        case "MaxStock":
        //                                            if (!string.IsNullOrEmpty(oRow[oColumn].ToString()))
        //                                            {
        //                                                MaxQty = Convert.ToDouble(oRow[oColumn].ToString());
        //                                                oItemWhs.MaximalStock = MaxQty;
        //                                            }
        //                                            break;
        //                                        case "PreferredVendor":
        //                                            Vendor = oRow[oColumn].ToString();
        //                                            oItemWhs.UserFields.Fields.Item("U_AVA_APCreditVendor").Value = Vendor;
        //                                            break;

        //                                        case "ClubPO":
        //                                            ClubPO = oRow[oColumn].ToString();
        //                                            oItemWhs.UserFields.Fields.Item("U_AVA_ClubPO").Value = ClubPO;
        //                                            break;
        //                                    }
        //                                }

        //                                RetCode = oItem.Update();
        //                                if (RetCode != 0)
        //                                {
        //                                    RespMsg = objSBOAPI.oCompany.GetLastErrorCode().ToString() + " - " + objSBOAPI.oCompany.GetLastErrorDescription();
        //                                    ErrorCount = ErrorCount + 1;
        //                                }
        //                                else
        //                                {
        //                                    Status = "S";
        //                                    RespMsg = "Updated Successfully";
        //                                    SuccessCount = SuccessCount + 1;
        //                                }

        //                            }
        //                            else
        //                            {
        //                                RespMsg = "Invalid ItemCode";
        //                                ErrorCount = ErrorCount + 1;
        //                            }
        //                        }
        //                        else
        //                        {
        //                            RespMsg = "Query Output is Empty";
        //                            ErrorCount = ErrorCount + 1;
        //                        }
        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        RespMsg = ex.Message;
        //                        ErrorCount = ErrorCount + 1;
        //                    }
        //                    finally
        //                    {
        //                        string InsertQuery = "INSERT INTO \"@AVA_MULOG\"(\"Code\",\"Name\",\"U_AVA_CreateDate\",\"U_AVA_FileName\",\"U_AVA_DocEntry\",\"U_AVA_ItmCod\",\"U_AVA_WhsCod\",\"U_AVA_MinQty\",\"U_AVA_MaxQty\",\"U_AVA_VenCod\",\"U_AVA_ClubPO\",\"U_AVA_Status\",\"U_AVA_RespMsg\") VALUES ('" + Code + "','" + objSBOAPI.oCompany.UserName + "','"+ DateTime.Now.ToString("dd MMM yyyy HH:mm:ss") +"','" + FileName + "','" + DocEntry + "','" + ItemCode + "','"+ WhsCode +"','"+ MinQty +"','"+ MaxQty +"','"+ Vendor +"','"+ClubPO+"','"+ Status +"','"+ RespMsg+"')";
        //                        objSBOAPI.Query_Execute(InsertQuery);

        //                        Code = Code + 1;

        //                        oProcessLineNo = oProcessLineNo + 1;
        //                        GC.Collect();
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                objSBOAPI.SBO_Appln.StatusBar.SetText("ItemCode and Warehouse columns are mandatory in excel", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //            }
        //        }
        //        objSBOAPI.SBO_Appln.StatusBar.SetText("Min/Max Uploaded Status : Total - " + oDt.Rows.Count.ToString() + " / Success - " + SuccessCount.ToString() + " / Failed - " + ErrorCount.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        //    }
        //    catch (Exception ex)
        //    {
        //        objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //    }
        //}
        //#endregion

        #region PricelistImport
        public void PricelistImport()
        {
            try
            {
                string FileName = objform.DataSources.UserDataSources.Item("FileName").Value;
                if (System.IO.File.Exists(FileName))
                {
                    XSSFWorkbook hssfworkbook = new XSSFWorkbook(FileName);
                    ISheet sheet = hssfworkbook.GetSheetAt(0);

                    DataTable dt = new DataTable();
                    IRow headerRow = sheet.GetRow(0);
                    IEnumerator rows = sheet.GetRowEnumerator();

                    int colCount = headerRow.LastCellNum;
                    int rowCount = sheet.LastRowNum;

                    int DocEntry = Convert.ToInt32(objSBOAPI.Query_Execute("SELECT IFNULL(Max(\"U_AVA_DocEntry\"),0) + 1 FROM \"@AVA_PLLOG\""));
                    int Code = Convert.ToInt32(objSBOAPI.Query_Execute("SELECT IFNULL(Max(\"Code\"),0) + 1 FROM \"@AVA_PLLOG\""));

                    objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait pricelist data is reading from excel file...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);


                    dt.Columns.Add("Code").AutoIncrement = true;
                    dt.Columns["Code"].AutoIncrementSeed = Code;
                    dt.Columns.Add("ItemCode");
                    dt.Columns.Add("PriceList");
                    dt.Columns.Add("Price", typeof(decimal));
                    dt.Columns.Add("DocEntry", typeof(Int32)).DefaultValue = DocEntry;
                    dt.Columns.Add("Name").DefaultValue = objSBOAPI.oCompany.UserName;
                    dt.Columns.Add("FileName").DefaultValue = FileName;
                    dt.Columns.Add("Type").DefaultValue = "P";
                    dt.Columns.Add("CreateDate").DefaultValue = DateTime.Now.ToString("dd MMM yyyy HH:mm:ss");

                    //for (int c = 0; c < colCount; c++)
                    //    dt.Columns.Add(headerRow.GetCell(c).ToString());

                    while (rows.MoveNext())
                    {
                        IRow row = (XSSFRow)rows.Current;

                        ICell cell = row.GetCell(0);
                        if (cell != null)
                        {
                            if (!string.IsNullOrEmpty(cell.ToString()) && cell.ToString().ToUpper() != "ITEMCODE")
                            {
                                string ItemCode = cell.ToString();
                                for (int i = 1; i < colCount; i++)
                                {
                                    string PriceList = headerRow.GetCell(i).ToString().Trim();
                                    DataRow dr = dt.NewRow();

                                    dr[1] = ItemCode;
                                    dr[2] = PriceList;
                                    cell = row.GetCell(i);
                                    if (cell != null)
                                        dr[3] = cell.ToString().Trim();

                                    dt.Rows.Add(dr);
                                }
                            }
                        }
                    }

                    hssfworkbook.Close();

                    Pricelist_Update(dt, DocEntry);

                    objform.DataSources.UserDataSources.Item("FromDate").Value = "";
                    objform.DataSources.UserDataSources.Item("ToDate").Value = "";

                    Uploaded_Result("PriceList", DocEntry, FileName, "P");
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Pricelist Update
        public void Pricelist_Update(DataTable oDt, int DocEntry)
        {
            HanaConnection oDBConnection = null;
            try
            {
                objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait pricelist data is inserting into log table...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                string oConn = string.Format("Server = {0}; UserID = {1}; Password = {2}", objSBOAPI.objMain.HANA_ServerName, objSBOAPI.objMain.HANA_UserID, objSBOAPI.objMain.HANA_Pwd);
                oDBConnection = new HanaConnection(oConn);
                oDBConnection.Open();
                HanaCommand da;
                da = new HanaCommand("SET SCHEMA " + objSBOAPI.oCompany.CompanyDB, oDBConnection);
                da.ExecuteNonQuery();

                HanaBulkCopy blkcmd = new HanaBulkCopy(oDBConnection);
                blkcmd.BulkCopyTimeout = 1000;
                blkcmd.DestinationTableName = "@AVA_PLLOG";
                blkcmd.ColumnMappings.Add("Code", "Code");
                blkcmd.ColumnMappings.Add("Name", "Name");
                blkcmd.ColumnMappings.Add("ItemCode", "U_AVA_ItmCod");
                blkcmd.ColumnMappings.Add("PriceList", "U_AVA_ListName");
                blkcmd.ColumnMappings.Add("Price", "U_AVA_Price");
                blkcmd.ColumnMappings.Add("FileName", "U_AVA_FileName");
                blkcmd.ColumnMappings.Add("CreateDate", "U_AVA_CreateDate");
                blkcmd.ColumnMappings.Add("DocEntry", "U_AVA_DocEntry");
                blkcmd.ColumnMappings.Add("Type", "U_AVA_Type");

                blkcmd.BatchSize = 1000;
                blkcmd.WriteToServer(oDt);
                blkcmd.Close();
                blkcmd.Dispose();


                oDt = new DataTable();
                string Str = "SELECT A.\"Code\",A.\"U_AVA_ItmCod\" AS \"ItemCode\",C.\"ListNum\",A.\"U_AVA_Price\" AS \"Price\",CASE WHEN IFNULL(B.\"ItemCode\",'')='' THEN 'Invalid ItemCode' WHEN IFNULL(CAST(C.\"ListNum\" AS NVARCHAR(11)),'')='' THEN 'Invalid PriceList' WHEN IFNULL(CAST(A.\"U_AVA_Price\" AS NVARCHAR(11)),'')='' THEN 'No Price Found' ELSE 'Create' END AS \"RespMsg\" FROM \"@AVA_PLLOG\" A LEFT JOIN \"OITM\" B ON B.\"ItemCode\"=A.\"U_AVA_ItmCod\" LEFT JOIN \"OPLN\" C ON C.\"ListName\"=A.\"U_AVA_ListName\" WHERE \"U_AVA_DocEntry\"='" + DocEntry + "'";

                oDt = objSBOAPI.HANA_ExecuteQuery(Str, oDBConnection);

                if (oDt != null)
                {
                    string FromDate = ((SAPbouiCOM.EditText)objform.Items.Item("EtFrom").Specific).Value;
                    string ToDate = ((SAPbouiCOM.EditText)objform.Items.Item("EtTo").Specific).Value;

                    int oProcessLineNo = 1;

                    int i = 0;
                    while (i < oDt.Rows.Count)
                    {
                        DataRow oRow = oDt.Rows[i];

                        string Code = oRow["Code"].ToString();
                        string ItemCode = oRow["ItemCode"].ToString();
                        string ListNum = oRow["ListNum"].ToString();
                        string Price = oRow["Price"].ToString();
                        string RespMsg = oRow["RespMsg"].ToString();
                        string Status = "F";
                        try
                        {
                            objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait PriceList is Uploading for Line No : " + oProcessLineNo.ToString() + " Out of (" + oDt.Rows.Count.ToString() + ")", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                            if (RespMsg == "Create")
                            {
                                string UpdateQuery = "Update \"ITM1\" Set \"Price\" = '" + Price + "' Where \"ItemCode\" = '" + ItemCode + "' And \"PriceList\" = '" + ListNum + "'";

                                if (objSBOAPI.HANA_ExecuteNonQuery(UpdateQuery, oDBConnection) == 1)
                                {
                                    Status = "S";
                                    RespMsg = "Successfully Updated";
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            RespMsg = ex.Message;
                        }
                        finally
                        {
                            string InsertQuery = "UPDATE \"@AVA_PLLOG\" SET \"U_AVA_Status\"='" + Status + "',\"U_AVA_CreateDate\"='" + DateTime.Now.ToString("dd MMM yyyy HH:mm:ss") + "',\"U_AVA_RespMsg\"='" + RespMsg + "',\"U_AVA_FromDate\"='" + FromDate + "',\"U_AVA_ToDate\"='" + ToDate + "' WHERE \"Code\"='" + Code + "' AND \"U_AVA_DocEntry\"='" + DocEntry + "'";
                            objSBOAPI.HANA_ExecuteNonQuery(InsertQuery, oDBConnection);
                        }

                        oProcessLineNo = oProcessLineNo + 1;

                        i++;
                    }

                    objSBOAPI.SBO_Appln.StatusBar.SetText("PriceList Uploaded...Please check the Status", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }

                oDBConnection.Close();
                oDBConnection.Dispose();
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Preferred Vendor
        public void PreferredVendor()
        {
            try
            {
                string FileName = objform.DataSources.UserDataSources.Item("FileName").Value;
                if (System.IO.File.Exists(FileName))
                {
                    XSSFWorkbook hssfworkbook = new XSSFWorkbook(FileName);
                    ISheet sheet = hssfworkbook.GetSheetAt(0);

                    DataTable dt = new DataTable();
                    IRow headerRow = sheet.GetRow(0);
                    IEnumerator rows = sheet.GetRowEnumerator();

                    int colCount = headerRow.LastCellNum;
                    int rowCount = sheet.LastRowNum;

                    for (int c = 0; c < colCount; c++)
                        dt.Columns.Add(headerRow.GetCell(c).ToString());

                    while (rows.MoveNext())
                    {
                        IRow row = (XSSFRow)rows.Current;

                        ICell cell = row.GetCell(0);
                        if (cell != null)
                        {
                            if (!string.IsNullOrEmpty(cell.ToString()) && cell.ToString().ToUpper() != "ITEMCODE")
                            {
                                DataRow dr = dt.NewRow();

                                for (int i = 0; i < colCount; i++)
                                {
                                    cell = row.GetCell(i);
                                    if (cell != null)
                                        dr[i] = cell.ToString().Trim();
                                }
                                dt.Rows.Add(dr);
                            }
                        }
                    }

                    hssfworkbook.Close();

                    int DocEntry = Convert.ToInt32(objSBOAPI.Query_Execute("SELECT IFNULL(Max(\"U_AVA_DocEntry\"),0) + 1 FROM \"@AVA_PVLOG\""));
                    PreferredVendor_Update(dt, FileName, DocEntry);

                    Uploaded_Result("PreVendor", DocEntry, FileName);
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Preferred Vendor Update
        public void PreferredVendor_Update(DataTable oDt, string FileName, int DocEntry)
        {
            int Total_Items = 0;
            try
            {
                List<DataRow> oData = null;
                int oProcessLineNo = 1;
                int SuccessCount = 0;
                int ErrorCount = 0;

                int Code = Convert.ToInt32(objSBOAPI.Query_Execute("SELECT IFNULL(Max(\"Code\"),0) + 1 FROM \"@AVA_PVLOG\""));

                //var Items = oDt.AsEnumerable().Select(d => d.Field<string>("ItemCode")).Distinct().ToList();
                var Item_List = oDt.AsEnumerable().Select(d => new { ItemCode = d.Field<string>("ItemCode") }).Distinct().ToList();

                Total_Items = Item_List.Count();
                foreach (var oItm in Item_List)
                {
                    string ItemCode = oItm.ItemCode;
                    string Status = "F";
                    string RespMsg = "";

                    try
                    {
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait Preferred Vendor is Uploading for Line No : " + oProcessLineNo.ToString() + " Out of (" + Item_List.Count.ToString() + ")", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                        oData = (from e in oDt.AsEnumerable()
                                 where e.Field<string>("ItemCode").Equals(oItm.ItemCode) && e.Field<string>("VendorCode") != ""
                                 select e).ToList();
                        if (oData != null)
                        {
                            if (oData.Count() > 0)
                            {
                                SAPbobsCOM.Items oItem = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);

                                if (oItem.GetByKey(oItm.ItemCode) == true)
                                {
                                    foreach (DataRow oRow in oData)
                                    {
                                        oItem.PreferredVendors.BPCode = oRow["VendorCode"].ToString();
                                        oItem.PreferredVendors.UserFields.Fields.Item("U_AVA_BRANCH").Value = oRow["Division"].ToString();
                                        oItem.PreferredVendors.Add();
                                    }

                                    RetCode = oItem.Update();
                                    if (RetCode != 0)
                                    {
                                        RespMsg = objSBOAPI.oCompany.GetLastErrorCode().ToString() + " - " + objSBOAPI.oCompany.GetLastErrorDescription();
                                        ErrorCount = ErrorCount + 1;
                                    }
                                    else
                                    {
                                        SuccessCount = SuccessCount + 1;
                                        Status = "S";
                                    }
                                }
                                else
                                {
                                    RespMsg = "Invalid ItemCode";
                                    ErrorCount = ErrorCount + 1;
                                }
                            }
                            else
                            {
                                RespMsg = "No Vendor details found against the ItemCode";
                                ErrorCount = ErrorCount + 1;
                            }
                        }
                        else
                        {
                            RespMsg = "No Vendor details found against the ItemCode";
                            ErrorCount = ErrorCount + 1;
                        }
                    }
                    catch (Exception ex)
                    {
                        RespMsg = ex.Message;
                        ErrorCount = ErrorCount + 1;
                    }
                    finally
                    {
                        foreach (DataRow oRow in oData)
                        {
                            string InsertQuery = "INSERT INTO \"@AVA_PVLOG\"(\"Code\",\"Name\",\"U_AVA_CreateDate\",\"U_AVA_FileName\",\"U_AVA_DocEntry\",\"U_AVA_ItmCod\",\"U_AVA_VenCod\",\"U_AVA_Division\",\"U_AVA_Status\",\"U_AVA_RespMsg\") VALUES ('" + Code + "','" + objSBOAPI.oCompany.UserName + "','" + DateTime.Now.ToString("dd MMM yyyy HH:mm:ss") + "','" + FileName + "','" + DocEntry + "','" + ItemCode + "','" + oRow["VendorCode"].ToString() + "','" + oRow["Division"].ToString() + "','" + Status + "','" + RespMsg + "')";
                            objSBOAPI.Query_Execute(InsertQuery);

                            Code = Code + 1;
                        }

                        oProcessLineNo = oProcessLineNo + 1;
                    }
                }
                objSBOAPI.SBO_Appln.StatusBar.SetText("Preferred Vendor Uploaded Status : Total - " + Total_Items.ToString() + " / Success - " + SuccessCount.ToString() + " / Failed - " + ErrorCount.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Preferred Vendor Deletion
        public void PreferredVendorDeletion()
        {
            try
            {
                string FileName = objform.DataSources.UserDataSources.Item("FileName").Value;
                if (System.IO.File.Exists(FileName))
                {
                    XSSFWorkbook hssfworkbook = new XSSFWorkbook(FileName);
                    ISheet sheet = hssfworkbook.GetSheetAt(0);

                    DataTable dt = new DataTable();
                    IRow headerRow = sheet.GetRow(0);
                    IEnumerator rows = sheet.GetRowEnumerator();

                    int colCount = headerRow.LastCellNum;
                    int rowCount = sheet.LastRowNum;

                    for (int c = 0; c < colCount; c++)
                        dt.Columns.Add(headerRow.GetCell(c).ToString());

                    dt.Columns.Add("Status");
                    dt.Columns.Add("RespMsg");

                    while (rows.MoveNext())
                    {
                        IRow row = (XSSFRow)rows.Current;

                        ICell cell = row.GetCell(0);
                        if (cell != null)
                        {
                            if (!string.IsNullOrEmpty(cell.ToString()) && cell.ToString().ToUpper() != "ITEMCODE")
                            {
                                DataRow dr = dt.NewRow();

                                for (int i = 0; i < colCount; i++)
                                {
                                    cell = row.GetCell(i);
                                    if (cell != null)
                                        dr[i] = cell.ToString().Trim();
                                }
                                dt.Rows.Add(dr);
                            }
                        }
                    }

                    hssfworkbook.Close();

                    int DocEntry = Convert.ToInt32(objSBOAPI.Query_Execute("SELECT IFNULL(Max(\"U_AVA_DocEntry\"),0) + 1 FROM \"@AVA_PDLOG\""));
                    PreferredVendor_Delete(dt, FileName, DocEntry);

                    Uploaded_Result("Delete_PreVendor", DocEntry, FileName);
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Preferred Vendor - Delete
        public void PreferredVendor_Delete(DataTable oDt, string FileName, int DocEntry)
        {
            int Total_Items = 0;
            try
            {
                List<DataRow> oData = null;
                int oProcessLineNo = 1;
                int SuccessCount = 0;
                int ErrorCount = 0;

                int Code = Convert.ToInt32(objSBOAPI.Query_Execute("SELECT IFNULL(Max(\"Code\"),0) + 1 FROM \"@AVA_PDLOG\""));

                //var Items = oDt.AsEnumerable().Select(d => d.Field<string>("ItemCode")).Distinct().ToList();
                var Item_List = oDt.AsEnumerable().Select(d => new { ItemCode = d.Field<string>("ItemCode") }).Distinct().ToList();

                Total_Items = Item_List.Count();
                foreach (var oItm in Item_List)
                {
                    string ItemCode = oItm.ItemCode;
                    string Status = "F";
                    string RespMsg = "";

                    try
                    {
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Please wait Preferred Vendor is Uploading for Line No : " + oProcessLineNo.ToString() + " Out of (" + Item_List.Count.ToString() + ")", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                        oData = (from e in oDt.AsEnumerable()
                                 where e.Field<string>("ItemCode").Equals(oItm.ItemCode) //&& e.Field<string>("VendorCode") != ""
                                 select e).ToList();
                        if (oData != null)
                        {
                            if (oData.Count() > 0)
                            {
                                oData.ForEach(x => { x["Status"] = Status; x["RespMsg"] = RespMsg; });

                                SAPbobsCOM.Items oItem = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);

                                if (oItem.GetByKey(oItm.ItemCode) == true)
                                {
                                    bool Update = false;
                                    foreach (DataRow oRow in oData)
                                    {
                                        bool oVendorExist = false;
                                        if (!string.IsNullOrEmpty(oRow["VendorCode"].ToString()))
                                        {

                                            SAPbobsCOM.Items_PreferredVendors oPV_List = oItem.PreferredVendors;

                                            int i = 0;

                                            while (i < oPV_List.Count)
                                            {
                                                oItem.PreferredVendors.SetCurrentLine(i);

                                                if (oItem.PreferredVendors.BPCode == oRow["VendorCode"].ToString())
                                                {
                                                    oItem.PreferredVendors.Delete();


                                                    if (oItem.Mainsupplier == oRow["VendorCode"].ToString())
                                                    {
                                                        oItem.Mainsupplier = "";
                                                    }

                                                    Update = true;
                                                    oVendorExist = true;
                                                    break;
                                                }
                                                i++;
                                            }
                                            if (!oVendorExist)
                                            {
                                                oRow["Status"] = "F";
                                                oRow["RespMsg"] = "Vendor is not found in DB";
                                            }
                                        }
                                        else
                                        {
                                            oRow["Status"] = "F";
                                            oRow["RespMsg"] = "Vendor is Missing";
                                        }
                                    }

                                    if (Update)
                                    {
                                        RetCode = oItem.Update();
                                        if (RetCode != 0)
                                        {
                                            RespMsg = objSBOAPI.oCompany.GetLastErrorCode().ToString() + " - " + objSBOAPI.oCompany.GetLastErrorDescription();
                                            ErrorCount = ErrorCount + 1;

                                            oData.Where(x => x["RespMsg"].ToString() == "").ToList().ForEach(x => { x["Status"] = Status; x["RespMsg"] = RespMsg; });
                                        }
                                        else
                                        {
                                            SuccessCount = SuccessCount + 1;
                                            Status = "S";

                                            oData.Where(x => x["RespMsg"].ToString() == "").ToList().ForEach(x => { x["Status"] = Status; x["RespMsg"] = "Deleted Successfully"; });
                                        }
                                    }
                                    else
                                    {
                                        RespMsg = "No Vendor details found against the ItemCode";
                                        ErrorCount = ErrorCount + 1;

                                        oData.Where(x => x["RespMsg"].ToString() == "").ToList().ForEach(x => { x["Status"] = "F"; x["RespMsg"] = RespMsg; });
                                    }
                                }
                                else
                                {
                                    RespMsg = "Invalid ItemCode";
                                    ErrorCount = ErrorCount + 1;

                                    oData.Where(x => x["RespMsg"].ToString() == "").ToList().ForEach(x => { x["Status"] = Status; x["RespMsg"] = RespMsg; });
                                }
                            }
                            else
                            {
                                RespMsg = "No Vendor details found against the ItemCode";
                                ErrorCount = ErrorCount + 1;
                            }
                        }
                        else
                        {
                            RespMsg = "No Vendor details found against the ItemCode";
                            ErrorCount = ErrorCount + 1;
                        }
                    }
                    catch (Exception ex)
                    {
                        RespMsg = ex.Message;
                        ErrorCount = ErrorCount + 1;

                        oData.Where(x => x["RespMsg"].ToString() == "").ToList().ForEach(x => { x["Status"] = Status; x["RespMsg"] = RespMsg; });
                    }
                    finally
                    {
                        foreach (DataRow oRow in oData)
                        {
                            string InsertQuery = "INSERT INTO \"@AVA_PDLOG\"(\"Code\",\"Name\",\"U_AVA_CreateDate\",\"U_AVA_FileName\",\"U_AVA_DocEntry\",\"U_AVA_ItmCod\",\"U_AVA_VenCod\",\"U_AVA_Status\",\"U_AVA_RespMsg\") VALUES ('" + Code + "','" + objSBOAPI.oCompany.UserName + "','" + DateTime.Now.ToString("dd MMM yyyy HH:mm:ss") + "','" + FileName + "','" + DocEntry + "','" + ItemCode + "','" + oRow["VendorCode"].ToString() + "','" + oRow["Status"].ToString() + "','" + oRow["RespMsg"].ToString() + "')";
                            objSBOAPI.Query_Execute(InsertQuery);

                            Code = Code + 1;
                        }

                        oProcessLineNo = oProcessLineNo + 1;
                    }
                }
                objSBOAPI.SBO_Appln.StatusBar.SetText("Preferred Vendor Deleted Status : Total - " + Total_Items.ToString() + " / Success - " + SuccessCount.ToString() + " / Failed - " + ErrorCount.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Uploaded Result
        public void Uploaded_Result(string Type, int DocEntry, string Source_File, string PLType = "")
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Grid oGrid = null;
            try
            {
                string ResultQuery = objSBOAPI.Query_Execute("SELECT \"QString\" FROM OUQR WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM OQCN WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='UploadedResults'");
                if (ResultQuery == "")
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("UploadedResults Query is Missing, Kindly do the Query Manager Update", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return;
                }
                else
                {
                    oForm = objSBOAPI.LoadForm("Result.xml", "AVA_RSSCF");

                    oForm.Freeze(true);

                    oGrid = oForm.Items.Item("GdList").Specific;
                    oGrid.DataTable.Clear();

                    ResultQuery = ResultQuery.Replace("[%C]", Type);
                    ResultQuery = ResultQuery.Replace("[%1]", DocEntry.ToString());
                    ResultQuery = ResultQuery.Replace("[%2]", PLType);

                    oGrid.DataTable.ExecuteQuery(ResultQuery);
                    SAPbouiCOM.DataTable oDt = oGrid.DataTable;
                    if (oGrid.DataTable.IsEmpty == false)
                    {
                        try
                        {
                            SAPbouiCOM.EditTextColumn oEdit = null;
                            oGrid.Columns.Item("RowsHeader").Visible = false;
                            for (int i = 0; i <= oDt.Columns.Count - 1; i++)
                            {
                                string columncheck = "";
                                columncheck = oGrid.Columns.Item(i).TitleObject.Caption;

                                string ObjType = columncheck.Contains("_") ? columncheck.ToString().Split('_')[0].ToString() : "";

                                if (!string.IsNullOrEmpty(ObjType))
                                {
                                    oEdit = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item(columncheck);
                                    oEdit.LinkedObjectType = ObjType.Replace("_", "");
                                    oGrid.Columns.Item(columncheck).TitleObject.Caption = columncheck.ToString().Replace(ObjType + "_", "");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            objform.Freeze(false);
                            objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                        oGrid.AutoResizeColumns();
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Data are Loaded Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                    else
                    {
                        oGrid.DataTable.Clear();
                        objSBOAPI.SBO_Appln.StatusBar.SetText("No Records Found", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }
                    oForm.Freeze(false);

                    objform.DataSources.UserDataSources.Item("FileName").Value = "";

                    string Path = objSBOAPI.Query_Execute("Select \"U_AV_FPATH\" from \"OUSR\" where \"USER_CODE\"='" + objSBOAPI.oCompany.UserName + "'");
                    switch (Type)
                    {
                        //case "MinMax":
                        //    oForm.Title = "Uploaded Result - Min/Max";
                        //    Path = Path + "\\" + "MinMax_Uploaded Files";
                        //    if (!System.IO.Directory.Exists(Path))
                        //    {
                        //        System.IO.Directory.CreateDirectory(Path);
                        //    }
                        //    break;

                        case "ItemMaster":
                            oForm.Title = "Uploaded Result - Item Master Data";
                            Path = Path + "\\" + "ItemMaster_Uploaded Files";
                            if (!System.IO.Directory.Exists(Path))
                            {
                                System.IO.Directory.CreateDirectory(Path);
                            }
                            break;

                        //case "PreVendor":
                        //    oForm.Title = "Uploaded Result - Preferred Vendor";
                        //    Path = Path + "\\" + "PreferredVendor_Uploaded Files";
                        //    if (!System.IO.Directory.Exists(Path))
                        //    {
                        //        System.IO.Directory.CreateDirectory(Path);
                        //    }
                        //    break;


                        //case "Delete_PreVendor":
                        //    oForm.Title = "Deleted Result - Preferred Vendor";
                        //    Path = Path + "\\" + "PreferredVendor_Deleted Files";
                        //    if (!System.IO.Directory.Exists(Path))
                        //    {
                        //        System.IO.Directory.CreateDirectory(Path);
                        //    }
                        //    break;

                        //case "PriceList":
                        //    oForm.Title = "Uploaded Result - Price List";
                        //    Path = Path + "\\" + "PriceList_Uploaded Files";
                        //    if (!System.IO.Directory.Exists(Path))
                        //    {
                        //        System.IO.Directory.CreateDirectory(Path);
                        //    }
                        //    break;
                    }

                    if (PLType != "I")
                    {
                        string FileName = System.IO.Path.GetFileNameWithoutExtension(Source_File) + "_" + DateTime.Now.ToString("yyyyMMddhhmmss");
                        string Extension = System.IO.Path.GetExtension(Source_File);
                        Path = System.IO.Path.Combine(Path, FileName + Extension);
                        System.IO.File.Copy(Source_File, Path, true);
                    }
                }
            }
            catch (Exception ex)
            {
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

    }
}
