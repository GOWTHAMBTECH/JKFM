using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Reflection;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;
using System.Configuration;
using Sap.Data.Hana;
using System.Data;

namespace JKFM_Source
{
    public class ClsSBO
    {

        #region Declaration
        public SAPbouiCOM.Application SBO_Appln;
        public SAPbobsCOM.Company oCompany;
        private SAPbouiCOM.Form objform = null;
        public ClsMain objMain;
        public Thread ShowFolderBrowserThread;

        public string Sys_DtFormat;
        public string DateSep;
        //Error Handling Parameters
        public string sErrMsg;
        public int lErrCode;
        public int lRetCode;
        public string strBranchForDebitMemo;

        //Add-on License
        public string AddonName = ConfigurationManager.AppSettings["AddonName"];
        Cls_LicenseVerify obj_LicenseVerify;
        Cls_MyLicense obj_MyLicense;
        Cls_LicenseUpload obj_LicenseUpload;

        //Database Update
        Cls_DatabaseUpdate obj_DatabaseUpdate;

        //Query Manager Update
        Cls_QueryManagerUpdate obj_QueryManagerUpdate;

        //Pre Sales Order
        Cls_PreSalesOrder obj_PreSalesOrder;

        //Purchase Details
        Cls_PurchaseDetails obj_PurchaseDetails;

        //Customer Wise Item Allocation
        Cls_CustomerWiseItemAllocation obj_CustomerWiseItemAllocation;

        //Sales Order Creation
        Cls_SalesOrderCreation obj_SalesOrderCreation;

        //Unit Master
        Cls_UnitMaster obj_UnitMaster;

        //Bill Delivery
        Cls_BillDelivery obj_BillDelivery;

        //Pick List Manager
        Cls_PickListManager obj_PickListManager;

        //Pick List Wizard Data
        Cls_PickListWizardData obj_PickListWizardData;

        //Pick List Data
        Cls_PickListData obj_PickListData;

        //Yield Master
        Cls_YieldMaster obj_YieldMaster;

        //Shift Master
        Cls_ShiftMaster obj_ShiftMaster;

        //Packing Capacity
        Cls_PackingCapacity obj_PackingCapacity;

        //Grinding Capacity
        Cls_GrindingCapacity obj_GrindingCapacity;

        //Pre Sale Approval
        Cls_PreSaleApproval obj_PreSaleApproval;

        //Sales Plan
        //Cls_SalesPlan obj_SalesPlan;

        //APInvoice
        public Cls_APInvoice obj_APInvoice;

        //Insurance Details
        Cls_InsuranceDetails obj_InsuranceDetails;

        //GRPO
        Cls_GRPO obj_GRPO;

        //List of GRPO
        Cls_ListofGRPO obj_ListofGRPO;

        //Email Setup
        Cls_EmailSetup obj_EmailSetup;

        //Customer Statement
        Cls_CustomerStatement obj_CustomerStatement;

        //Item Creation
        Cls_ItemCreation obj_ItemCreation;

        //PickList
        Cls_PickList obj_PickList;

        #endregion

        #region Constructor
        public ClsSBO(ClsMain objBP)
        {
            objMain = objBP;
        }
        #endregion

        #region Connect to SAP B1, Set Filters and Create Menus

        #region Connect to SAP B1, Set Filters and Create Menus
        public bool Connect()
        {
            if (!initialiseApplication())
            {
                return false;
            }
            SetFilters();
            if (!ConnectCompany())
            {
                return false;
            }
            

            SBO_Appln.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            SBO_Appln.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
            SBO_Appln.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            SBO_Appln.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Appln_FormDataEvent);
            SBO_Appln.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Appln_RightClickEvent);

            obj_LicenseVerify = new Cls_LicenseVerify(this);
            obj_MyLicense = new Cls_MyLicense();
            obj_LicenseUpload = new Cls_LicenseUpload(this);

            if (!obj_LicenseVerify.Validate_License())
            {
                return false;
            }
            AddMenus();

            return true;
        }
        #endregion

        #region Connect to SAP B1 Application
        public bool initialiseApplication()
        {
            try
            {
                //Connection String = 0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056
                SAPbouiCOM.SboGuiApi SboGuiApi = null;
                string sConnectionString = null;

                SboGuiApi = new SAPbouiCOM.SboGuiApi();

                sConnectionString = Environment.GetCommandLineArgs().GetValue(1).ToString();

                SboGuiApi.Connect(sConnectionString);

                SBO_Appln = SboGuiApi.GetApplication();

                //SBO_Appln = SAPbouiCOM.Framework.Application.SBO_Application;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        #endregion

        #region Connect to SAP B1 Company
        private bool ConnectCompany()
        {
            try
            {
                oCompany = new SAPbobsCOM.Company();
                string ocookies = null;
                string ocookiecontext = null;
                ocookies = oCompany.GetContextCookie();
                ocookiecontext = SBO_Appln.Company.GetConnectionContext(ocookies);
                oCompany.SetSboLoginContext(ocookiecontext);
                if (oCompany.Connect() != 0)
                {
                    oCompany.GetLastError(out lErrCode, out sErrMsg);
                    SBO_Appln.MessageBox(lErrCode + "-" + sErrMsg);
                    return false;
                }
                else
                {
                    Sys_DtFormat = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
                    if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        DateSep = Query_Execute("Select \"DateSep\" from OADM");
                    }
                    else
                    {
                        DateSep = Query_Execute("Select DateSep from OADM");
                    }
                    Intialize();
                }
                oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode;
                //oCompany = (SAPbobsCOM.Company)SBO_Appln.Company.GetDICompany();
                return true;
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
                return false;
            }
        }
        #endregion

        #region Set Filters
        private void SetFilters()
        {
            SAPbouiCOM.EventFilters oFilters;
            SAPbouiCOM.EventFilter oFilter;

            oFilters = new SAPbouiCOM.EventFilters();

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
            oFilter.AddEx("AV_DBUPF");
            oFilter.AddEx("OPRE");
            oFilter.AddEx("OCIA");
            oFilter.AddEx("AIS_LOAD");
            oFilter.AddEx("OBRN");
            oFilter.AddEx("OPDE");
            oFilter.AddEx("OREM");
            oFilter.AddEx("OSFT");
            oFilter.AddEx("ORMC");
            oFilter.AddEx("OCAC");
            oFilter.AddEx("OPSA");

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK);
            oFilter.AddEx("ORMC");

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);
            oFilter.AddEx("PRE2");
            oFilter.AddEx("81");
            oFilter.AddEx("80");
            oFilter.AddEx("60020");
            oFilter.AddEx("141");

            //oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE);

            //oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE);

            //oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_RESIZE);

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD);
            oFilter.AddEx("AIS_LOAD");
            oFilter.AddEx("60020");
            oFilter.AddEx("141");

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD);
            oFilter.AddEx("AIS_LOAD");
            oFilter.AddEx("OBRN");
            oFilter.AddEx("OREM");
            oFilter.AddEx("141");
            oFilter.AddEx("AVA_INSDET");
            oFilter.AddEx("OPRE");
            oFilter.AddEx("OPDE");

            //oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE);

            //oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE);

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);
            oFilter.AddEx("AV_DBUPF");
            oFilter.AddEx("AV_QRUPF");
            oFilter.AddEx("AV_ADLUF");
            oFilter.AddEx("OPRE");
            oFilter.AddEx("OCIA");
            oFilter.AddEx("AIS_LOAD");
            oFilter.AddEx("OBRN");
            oFilter.AddEx("OPDE");
            oFilter.AddEx("OREM");
            oFilter.AddEx("OSFT");
            oFilter.AddEx("ORMC");
            oFilter.AddEx("OCAC");
            //oFilter.AddEx("OPSA");
            oFilter.AddEx("141");
            oFilter.AddEx("AVA_INSDET");
            oFilter.AddEx("10019");
            oFilter.AddEx("AV_CUSTF");
            oFilter.AddEx("AV_EMSTF");
            oFilter.AddEx("AVA_OITMF");
            oFilter.AddEx("85");

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
            oFilter.AddEx("OPRE");
            oFilter.AddEx("PRE2");
            oFilter.AddEx("OCIA");
            oFilter.AddEx("AIS_LOAD");
            oFilter.AddEx("OBRN");
            oFilter.AddEx("81");
            oFilter.AddEx("80");
            oFilter.AddEx("OPDE");
            oFilter.AddEx("OREM");
            oFilter.AddEx("OSFT");
            oFilter.AddEx("ORMC");
            oFilter.AddEx("OCAC");
            //oFilter.AddEx("OPSA");
            oFilter.AddEx("AV_CUSTF");

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK);
            oFilter.AddEx("AV_QRUPF");
            oFilter.AddEx("OPRE");
            oFilter.AddEx("AIS_LOAD");
            oFilter.AddEx("81"); 
            oFilter.AddEx("AV_ADLUF");
            oFilter.AddEx("10019");
            oFilter.AddEx("AV_CUSTF");

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS);
            oFilter.AddEx("OPRE");
            oFilter.AddEx("PRE2");
            oFilter.AddEx("OCIA");
            oFilter.AddEx("OBRN");
            oFilter.AddEx("OREM");
            oFilter.AddEx("ORMC");
            oFilter.AddEx("OCAC");
            oFilter.AddEx("141");

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS);
            //oFilter.AddEx("OPRE");

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);
            oFilter.AddEx("OPRE");
            oFilter.AddEx("AIS_LOAD");
            oFilter.AddEx("OPDE");
            oFilter.AddEx("141");
            oFilter.AddEx("AVA_INSDET");
            oFilter.AddEx("143");
            oFilter.AddEx("AV_CUSTF");
            oFilter.AddEx("AVA_OITMF");

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST);
            oFilter.AddEx("OPRE");
            oFilter.AddEx("PRE2");
            oFilter.AddEx("OCIA");
            oFilter.AddEx("AIS_LOAD");
            oFilter.AddEx("OBRN");
            oFilter.AddEx("OPDE");
            oFilter.AddEx("OREM");
            oFilter.AddEx("ORMC");
            oFilter.AddEx("OCAC");
            oFilter.AddEx("141");
            oFilter.AddEx("AV_CUSTF");

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED);
            oFilter.AddEx("AIS_LOAD");
            oFilter.AddEx("AVA_INSDET");

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN);
            oFilter.AddEx("AV_CUSTF");

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE);
            oFilter.AddEx("OPRE");
            oFilter.AddEx("AIS_LOAD");
            oFilter.AddEx("OPDE");
            oFilter.AddEx("OSFT");
            oFilter.AddEx("ORMC");
            oFilter.AddEx("OCAC");
            oFilter.AddEx("141");

            SBO_Appln.SetFilter(oFilters);

        }
        #endregion

        #region Create Menus

        #region Create Menus using XML - Method 1
        private void AddMenus()
        {
            try
            {
                XmlDocument JCMenus = null;
                JCMenus = new XmlDocument();

                string strResource;
                strResource = Assembly.GetExecutingAssembly().GetName().Name + ".XML." + "addMenus.xml";
                JCMenus.Load(Assembly.GetExecutingAssembly().GetManifestResourceStream(strResource));
                SBO_Appln.LoadBatchActions(JCMenus.InnerXml);
                //SBO_Appln.Menus.Item("AV_AVNKM").Image = System.Windows.Forms.Application.StartupPath + "\\Avaniko_Icon.jpg";
                SBO_Appln.Menus.Item("AV_AVNKM").Image = (System.Windows.Forms.Application.StartupPath).Replace("\\bin\\Debug","") + "\\TextFiles\\Avaniko_Icon.jpg";

            }
            catch (Exception ex)
            {

            }
        }
        #endregion

        #region Create Menus using Code - Method 2

        #region Load Menu
        private void loadMenu()
        {
            SAPbouiCOM.MenuItem objRptMenu;
            objRptMenu = SBO_Appln.Menus.Item("43520").SubMenus.Item("3328").SubMenus.Item("43525");
            int MenuCount = objRptMenu.SubMenus.Count;
            CreateMenu("", 1, "NanaDesi", SAPbouiCOM.BoMenuType.mt_POPUP, "NM", objRptMenu);
            CreateMenu("", 1, "Color", SAPbouiCOM.BoMenuType.mt_STRING, "NM1", objRptMenu.SubMenus.Item("NM"));
            CreateMenu("", 2, "Human Type", SAPbouiCOM.BoMenuType.mt_STRING, "NM2", objRptMenu.SubMenus.Item("NM"));
            CreateMenu("", 3, "Parent Form", SAPbouiCOM.BoMenuType.mt_STRING, "NM3", objRptMenu.SubMenus.Item("NM"));
            CreateMenu("", 4, "Size Type", SAPbouiCOM.BoMenuType.mt_STRING, "NM4", objRptMenu.SubMenus.Item("NM"));
            CreateMenu("", 5, "Account Type", SAPbouiCOM.BoMenuType.mt_STRING, "NM6", objRptMenu.SubMenus.Item("NM"));
            //CreateMenu("", 4, "Resize Customization", SAPbouiCOM.BoMenuType.mt_STRING, "NM5", SBO_Appln.Menus.Item("43520").SubMenus.Item("3072").SubMenus.Item("43540"));
            CreateMenu("", 1, "Sales Update", SAPbouiCOM.BoMenuType.mt_STRING, "NU", SBO_Appln.Menus.Item("43520").SubMenus.Item("2048"));
            CreateMenu("", 3, "Database Update", SAPbouiCOM.BoMenuType.mt_STRING, "AV_DBUPM", SBO_Appln.Menus.Item("43523"));
            CreateMenu("", 6, "Driver Admin Jobs Import", SAPbouiCOM.BoMenuType.mt_STRING, "NM7", objRptMenu.SubMenus.Item("NM"));
            CreateMenu("", 7, "Driver Admin Jobs", SAPbouiCOM.BoMenuType.mt_STRING, "NM8", objRptMenu.SubMenus.Item("NM"));
            CreateMenu("", 0, "Cash Flow Projections", SAPbouiCOM.BoMenuType.mt_STRING, "AV_CFPJM", SBO_Appln.Menus.Item("43531"));
        }
        #endregion

        #region Create Menu
        private SAPbouiCOM.MenuItem CreateMenu(string ImagePath, int Position, string DisplayName, SAPbouiCOM.BoMenuType MenuType, string UniqueID, SAPbouiCOM.MenuItem ParentMenu)
        {
            try
            {
                SAPbouiCOM.MenuCreationParams oMenuPackage;
                oMenuPackage = (SAPbouiCOM.MenuCreationParams)SBO_Appln.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oMenuPackage.Image = ImagePath;
                oMenuPackage.Position = Position;
                oMenuPackage.Type = MenuType;
                oMenuPackage.UniqueID = UniqueID;
                oMenuPackage.String = DisplayName;
                ParentMenu.SubMenus.AddEx(oMenuPackage);
            }
            catch (Exception)
            {
                SBO_Appln.StatusBar.SetText("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
            }
            return ParentMenu.SubMenus.Item(UniqueID);
        }
        #endregion

        #endregion        

        #endregion

        #endregion

        #region Create Objects
        public void CreateObjects()
        {
            obj_DatabaseUpdate = new Cls_DatabaseUpdate(this);
            obj_QueryManagerUpdate = new Cls_QueryManagerUpdate(this);
            obj_PreSalesOrder = new Cls_PreSalesOrder(this);
            obj_PurchaseDetails = new Cls_PurchaseDetails(this);
            obj_CustomerWiseItemAllocation = new Cls_CustomerWiseItemAllocation(this);
            obj_SalesOrderCreation = new Cls_SalesOrderCreation(this);
            obj_UnitMaster = new Cls_UnitMaster(this);
            obj_BillDelivery = new Cls_BillDelivery(this);
            obj_PickListManager = new Cls_PickListManager(this);
            obj_PickListWizardData = new Cls_PickListWizardData(this);
            obj_PickListData = new Cls_PickListData(this);
            obj_YieldMaster = new Cls_YieldMaster(this);
            obj_ShiftMaster = new Cls_ShiftMaster(this);
            obj_PackingCapacity = new Cls_PackingCapacity(this);
            obj_GrindingCapacity = new Cls_GrindingCapacity(this);
            obj_PreSaleApproval = new Cls_PreSaleApproval(this);
            //obj_SalesPlan = new Cls_SalesPlan(this);
            obj_APInvoice = new Cls_APInvoice(this);
            obj_InsuranceDetails = new Cls_InsuranceDetails(this);
            obj_GRPO = new Cls_GRPO(this);
            obj_ListofGRPO = new Cls_ListofGRPO(this);
            obj_EmailSetup = new Cls_EmailSetup(this);
            obj_CustomerStatement = new Cls_CustomerStatement(this);
            obj_ItemCreation = new Cls_ItemCreation(this);
            obj_PickList = new Cls_PickList(this);
        }
        #endregion

        #region Load Form
        public SAPbouiCOM.Form LoadForm(string XMLFile, string FormType, bool Modality = false)
        {
            return LoadForm(XMLFile, FormType.ToString(), FormType + "_" + SBO_Appln.Forms.Count.ToString(), Modality);
        }

        public SAPbouiCOM.Form LoadForm(string XMLFile, string FormType, string FormUID, bool Modality)
        {
            XmlDocument oXML = null;
            SAPbouiCOM.FormCreationParams objFormCreationParams = null;
            string strResource;
            try
            {
                oXML = new XmlDocument();
                strResource = Assembly.GetExecutingAssembly().GetName().Name + ".XML." + XMLFile;
                oXML.Load(System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(strResource));
                objFormCreationParams = (SAPbouiCOM.FormCreationParams)SBO_Appln.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                objFormCreationParams.XmlData = oXML.InnerXml;
                objFormCreationParams.FormType = FormType;
                if (Modality == false)
                {
                    objFormCreationParams.Modality = SAPbouiCOM.BoFormModality.fm_None;
                }
                else if (Modality == true)
                {
                    objFormCreationParams.Modality = SAPbouiCOM.BoFormModality.fm_Modal;
                }
                return SBO_Appln.Forms.AddEx(objFormCreationParams);
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
                return SBO_Appln.Forms.AddEx(objFormCreationParams);
            }
        }

        public void LoadFromXML(string FileName)
        {
            XmlDocument oXmlDoc;
            oXmlDoc = new XmlDocument();
            string strResource;
            try
            {
                strResource = Assembly.GetExecutingAssembly().GetName().Name + ".XML." + FileName;
                oXmlDoc.Load(Assembly.GetExecutingAssembly().GetManifestResourceStream(strResource));
                SBO_Appln.LoadBatchActions(oXmlDoc.InnerXml);
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
            }
        }

        public void LoadMenu(string XMLFile)
        {
            XmlDocument oXML;
            string strXML;
            try
            {
                oXML = new XmlDocument();
                oXML.Load(XMLFile);
                strXML = oXML.InnerXml;
                SBO_Appln.LoadBatchActions(strXML);
            }
            catch (Exception)
            {

            }
        }

        #endregion

        #region Events

        #region Application Event
        private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    oCompany.Disconnect();
                    Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    oCompany.Disconnect();
                    Application.Exit();
                    break;
                    //case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    //    oCompany.Disconnect();
                    //    break;
            }
        }
        #endregion

        #region Menu Event
        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        //For Default Forms
                        //case "menuid":
                        //DefaultForm_Menu("udocode - udoname");
                        //break;
                        case "AV_DBUPM":
                            obj_DatabaseUpdate.MenuEvent(ref pVal, ref BubbleEvent);
                            break;

                        //FIND
                        case "1281":
                            //if (SBO_Appln.Forms.ActiveForm.TypeEx == "139")
                            //{
                            //    obj_SalesOrder.MenuEvent(ref pVal, ref BubbleEvent);
                            //}   
                            break;
                        //ADD
                        case "1282":
                            //if (SBO_Appln.Forms.ActiveForm.TypeEx == "139")
                            //{
                            //    obj_SalesOrder.MenuEvent(ref pVal, ref BubbleEvent);
                            //}
                            break;
                        //REMOVE
                        case "1283":
                            break;
                        //CANCEL
                        case "1284":
                            switch (SBO_Appln.Forms.ActiveForm.TypeEx)
                            {
                                case "141":
                                    obj_APInvoice.MenuEvent(ref pVal, ref BubbleEvent);
                                    break;
                            }
                            break;
                        //RESTORE
                        case "1285":
                            break;
                        //CLOSE
                        case "1286":
                            break;
                        //DUPLICATE
                        case "1287":
                            break;
                        //ADD ROW
                        case "1292":
                            break;
                        //DELETE ROW
                        case "1293":
                            break;
                    }
                }
                else
                {
                    switch (pVal.MenuUID)
                    {
                        //For Default Forms
                        //case "menuid":
                        //DefaultForm_Menu("udocode - udoname");
                        //break;
                        case "AV_DBUPM":
                            obj_DatabaseUpdate.MenuEvent(ref pVal, ref BubbleEvent);
                            break;

                        case "OPRE":
                            obj_PreSalesOrder.MenuEvent(ref pVal, ref BubbleEvent);
                            break;

                        case "OCIA":
                            obj_CustomerWiseItemAllocation.MenuEvent(ref pVal, ref BubbleEvent);
                            break;

                        case "AIS_LOAD":
                            obj_SalesOrderCreation.MenuEvent(ref pVal, ref BubbleEvent);
                            break;

                        case "OBRN":
                            obj_UnitMaster.MenuEvent(ref pVal, ref BubbleEvent);
                            break;

                        case "OPDE":
                            obj_BillDelivery.MenuEvent(ref pVal, ref BubbleEvent);
                            break;

                        case "OREM":
                            obj_YieldMaster.MenuEvent(ref pVal, ref BubbleEvent);
                            break;

                        case "OSFT":
                            obj_ShiftMaster.MenuEvent(ref pVal, ref BubbleEvent);
                            break;

                        case "ORMC":
                            obj_PackingCapacity.MenuEvent(ref pVal, ref BubbleEvent);
                            break;

                        case "OCAC":
                            obj_GrindingCapacity.MenuEvent(ref pVal, ref BubbleEvent);
                            break;

                        case "OPSA":
                            obj_PreSaleApproval.MenuEvent(ref pVal, ref BubbleEvent);
                            break;

                        //case "OSAP":
                        //    obj_SalesPlan.MenuEvent(ref pVal, ref BubbleEvent);
                        //    break;

                        case "AV_INSDTM":
                            obj_InsuranceDetails.MenuEvent(ref pVal, ref BubbleEvent);
                            break;

                        case "AV_CUSTM":
                            obj_CustomerStatement.MenuEvent(ref pVal, ref BubbleEvent);
                            break;

                        case "AV_EMSTM":
                            obj_EmailSetup.MenuEvent(ref pVal, ref BubbleEvent);
                            break;

                        case "AV_OITMM":
                            obj_ItemCreation.MenuEvent(ref pVal, ref BubbleEvent);
                            break;

                        //FIND
                        case "1281":
                            //if (SBO_Appln.Forms.ActiveForm.TypeEx == "139")
                            //{
                            //    obj_SalesOrder.MenuEvent(ref pVal, ref BubbleEvent);
                            //}   
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "OPRE")
                            {
                                obj_PreSalesOrder.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "OPDE")
                            {
                                obj_BillDelivery.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "OSFT")
                            {
                                obj_ShiftMaster.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "ORMC")
                            {
                                obj_PackingCapacity.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "OCAC")
                            {
                                obj_GrindingCapacity.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "AV_INSDTM")
                            {
                                obj_InsuranceDetails.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            break;
                        //ADD
                        case "1282":
                            //if (SBO_Appln.Forms.ActiveForm.TypeEx == "139")
                            //{
                            //    obj_SalesOrder.MenuEvent(ref pVal, ref BubbleEvent);
                            //}
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "OPRE")
                            {
                                obj_PreSalesOrder.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "OCIA")
                            {
                                obj_CustomerWiseItemAllocation.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "AIS_LOAD")
                            {
                                obj_SalesOrderCreation.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "OPDE")
                            {
                                obj_BillDelivery.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "OREM")
                            {
                                obj_YieldMaster.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "OSFT")
                            {
                                obj_ShiftMaster.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "ORMC")
                            {
                                obj_PackingCapacity.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "OCAC")
                            {
                                obj_GrindingCapacity.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "141")
                            {
                                obj_APInvoice.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "AV_INSDTM")
                            {
                                obj_InsuranceDetails.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            break;
                        //REMOVE
                        case "1283":
                            break;
                        //CANCEL
                        case "1284":
                            break;
                        //RESTORE
                        case "1285":
                            break;
                        //CLOSE
                        case "1286":
                            break;
                        //DUPLICATE
                        case "1287":
                            break;
                        //ADD ROW
                        case "1292":
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "OPRE")
                            {
                                obj_PreSalesOrder.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "OCIA")
                            {
                                obj_CustomerWiseItemAllocation.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "ORMC")
                            {
                                obj_PackingCapacity.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            break;
                        //DELETE ROW
                        case "1293":
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "OPRE")
                            {
                                obj_PreSalesOrder.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "OCIA")
                            {
                                obj_CustomerWiseItemAllocation.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "ORMC")
                            {
                                obj_PackingCapacity.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            break;
                        case "4870":
                            if (SBO_Appln.Forms.ActiveForm.TypeEx == "OPRE")
                            {
                                obj_PreSalesOrder.MenuEvent(ref pVal, ref BubbleEvent);
                            }
                            break;
                    }
                    // 1288 - Next Record ; 1289 - Previous Record ; 1290 - First Data Record ; 1291 - Last Data Record
                    if (pVal.MenuUID == "1288" || pVal.MenuUID == "1289" || pVal.MenuUID == "1290" || pVal.MenuUID == "1291")
                    {
                        //if (SBO_Appln.Forms.ActiveForm.TypeEx == "139")
                        //{
                        //    obj_SalesOrder.MenuEvent(ref pVal, ref BubbleEvent);
                        //}
                    }
                }
            }
            catch (Exception)
            {
            }
        }
        #endregion

        #region Item Event
        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (pVal.FormTypeEx)
                {
                    case "AV_ADLUF":
                        obj_LicenseUpload.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "AV_DBUPF":
                        obj_DatabaseUpdate.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "AV_QRUPF":
                        obj_QueryManagerUpdate.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "OPRE":
                        obj_PreSalesOrder.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "PRE2":
                        obj_PurchaseDetails.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "OCIA":
                        obj_CustomerWiseItemAllocation.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "AIS_LOAD":
                        obj_SalesOrderCreation.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "OBRN":
                        obj_UnitMaster.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "81":
                        obj_PickListManager.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "80":
                        obj_PickListWizardData.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "60020":
                        obj_PickListData.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "OPDE":
                        obj_BillDelivery.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "OREM":
                        obj_YieldMaster.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "OSFT":
                        obj_ShiftMaster.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "ORMC":
                        obj_PackingCapacity.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "OCAC":
                        obj_GrindingCapacity.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    //case "OPSA":
                    //    obj_PreSaleApproval.itemevent(FormUID, ref pVal, ref BubbleEvent);
                    //    break;
                    //case "OSAP":
                    //    obj_SalesPlan.itemevent(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    case "141":
                        obj_APInvoice.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "AVA_INSDET":
                        obj_InsuranceDetails.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "143":
                        obj_GRPO.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "10019":
                        obj_ListofGRPO.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "AV_EMSTF":
                        obj_EmailSetup.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "AV_CUSTF":
                        obj_CustomerStatement.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "AVA_OITMF":
                        obj_ItemCreation.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                    case "85":
                        obj_PickList.itemevent(FormUID, ref pVal, ref BubbleEvent);
                        break;
                }
            }
            catch (Exception)
            {
            }
        }
        #endregion

        #region Form Data Event
        private void SBO_Appln_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    switch (BusinessObjectInfo.FormTypeEx)
                    {
                        case "OPRE":
                            obj_PreSalesOrder.FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                            break;
                    }
                }
                if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.FormTypeEx)
                    {
                        case "OPRE":
                            obj_PreSalesOrder.FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                            break;
                        case "AIS_LOAD":
                            obj_SalesOrderCreation.FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                            break;
                        case "OBRN":
                            obj_UnitMaster.FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                            break;
                        case "60020":
                            obj_PickListData.FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                            break;
                        case "OPDE":
                            obj_BillDelivery.FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                            break;
                        case "OREM":
                            obj_YieldMaster.FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                            break;
                        case "141":
                            obj_APInvoice.FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                            break;
                        case "AVA_INSDET":
                            obj_InsuranceDetails.FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                            break;
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
        #endregion

        #region Right Click
        private void SBO_Appln_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                //if(eventInfo.BeforeAction == false)
                //{
                //    switch(eventInfo.FormUID)
                //    {
                //        case "ORMC":

                //            break;
                //    }
                //}
                if (SBO_Appln.Forms.ActiveForm.TypeEx == "ORMC")
                {
                    obj_PackingCapacity.Right_Click(ref eventInfo, out BubbleEvent);
                }
                //if (SBO_Appln.Forms.ActiveForm.TypeEx == "AV_PRMPF")
                //{
                //    if (SBO_Appln.Menus.Exists("1283") == true)
                //    {
                //        SBO_Appln.Menus.RemoveEx("1283");
                //    }
                //}
                //if (SBO_Appln.Forms.ActiveForm.TypeEx == "139")
                //{
                //    obj_SalesOrder.Right_Click(eventInfo, BubbleEvent);
                //}
                //else
                //{
                //    if (SBO_Appln.Menus.Exists("AV_PICMM") == true)
                //    {
                //        SBO_Appln.Menus.RemoveEx("AV_PICMM");
                //    }                       
                //}
            }
            catch (Exception)
            {
            }
        }
        #endregion

       

        #endregion

        #region Database Update

        #region Create UDO
        public void createUDO(string tablename, string udocode, string udoname, SAPbobsCOM.BoUDOObjType type, string chldtable, bool DfltForm = false, string Style = "M", string Form_Fields_Name = "", string Form_Fields_Descrip = "", string Form_Fields_Editable = "", string Menu = "N", string MenuCaption = "", string FatherMenuID = "", string Position = "", string MenuUID = "")
        {
            SBO_Appln.StatusBar.SetText("Creating UDO Please Wait...........", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            SAPbobsCOM.UserObjectsMD objUserobjectMD = null;
            objUserobjectMD = (SAPbobsCOM.UserObjectsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            try
            {
                if (!objUserobjectMD.GetByKey(udocode))
                {
                    objUserobjectMD.Code = udocode;
                    objUserobjectMD.Name = udoname;
                    objUserobjectMD.TableName = tablename;
                    objUserobjectMD.ObjectType = type;
                    //objUserobjectMD.LogTableName = "AMPL_BOM1";

                    objUserobjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                    objUserobjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO;
                    objUserobjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO;
                    objUserobjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES;
                    objUserobjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO;

                    if (type == SAPbobsCOM.BoUDOObjType.boud_MasterData)
                    {
                        objUserobjectMD.FindColumns.ColumnAlias = "Code";
                        objUserobjectMD.FindColumns.ColumnDescription = "Code";
                        objUserobjectMD.FindColumns.Add();

                        objUserobjectMD.FindColumns.ColumnAlias = "Name";
                        objUserobjectMD.FindColumns.ColumnDescription = "Name";
                        objUserobjectMD.FindColumns.Add();
                    }
                    else
                    {
                        objUserobjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
                        objUserobjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO;

                        objUserobjectMD.FindColumns.ColumnAlias = "DocEntry";
                        objUserobjectMD.FindColumns.ColumnDescription = "DocEntry";
                        objUserobjectMD.FindColumns.Add();

                        objUserobjectMD.FindColumns.ColumnAlias = "DocNum";
                        objUserobjectMD.FindColumns.ColumnDescription = "DocNum";
                        objUserobjectMD.FindColumns.Add();
                    }

                    string[] str_chldtable;
                    str_chldtable = chldtable.Split(Convert.ToChar(","));
                    if (str_chldtable[0] != "")
                    {
                        for (int intLoop = 0; intLoop <= str_chldtable.GetLength(0) - 1; intLoop++)
                        {
                            objUserobjectMD.ChildTables.SetCurrentLine(intLoop);
                            objUserobjectMD.ChildTables.TableName = str_chldtable[intLoop];
                            objUserobjectMD.ChildTables.Add();
                        }
                    }

                    if (DfltForm)
                    {
                        objUserobjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;
                        objUserobjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;
                        if (Style == "M")
                        {
                            objUserobjectMD.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO;
                        }
                        else
                        {
                            objUserobjectMD.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tYES;
                        }
                        if (type == SAPbobsCOM.BoUDOObjType.boud_MasterData)
                        {
                            objUserobjectMD.FormColumns.FormColumnAlias = "Code";
                            objUserobjectMD.FormColumns.FormColumnDescription = "Code";
                            objUserobjectMD.FormColumns.Add();

                            objUserobjectMD.FormColumns.FormColumnAlias = "Name";
                            objUserobjectMD.FormColumns.FormColumnDescription = "Name";
                            objUserobjectMD.FormColumns.Add();
                        }
                        else
                        {
                            objUserobjectMD.FormColumns.FormColumnAlias = "DocEntry";
                            objUserobjectMD.FormColumns.FormColumnDescription = "DocEntry";
                            objUserobjectMD.FormColumns.Add();

                            objUserobjectMD.FormColumns.FormColumnAlias = "DocNum";
                            objUserobjectMD.FormColumns.FormColumnDescription = "DocNum";
                            objUserobjectMD.FormColumns.Add();
                        }

                        if (Form_Fields_Name != "")
                        {
                            string[] Str_Form_Fields_Name, Str_Form_Fields_Descrip, Str_Form_Fields_Editable;
                            Str_Form_Fields_Name = Form_Fields_Name.Split(Convert.ToChar(","));
                            Str_Form_Fields_Descrip = Form_Fields_Descrip.Split(Convert.ToChar(","));
                            Str_Form_Fields_Editable = Form_Fields_Editable.Split(Convert.ToChar(","));

                            if (Str_Form_Fields_Name.GetLength(0) != Str_Form_Fields_Descrip.GetLength(0))
                            {
                                throw new Exception("Invalid Field or Description of the Default Form for UDO - " + udocode + "");
                            }

                            if (Str_Form_Fields_Name[0] != "")
                            {
                                for (int intLoop = 0; intLoop <= Str_Form_Fields_Name.GetLength(0) - 1; intLoop++)
                                {
                                    objUserobjectMD.FormColumns.FormColumnAlias = Str_Form_Fields_Name[intLoop];
                                    objUserobjectMD.FormColumns.FormColumnDescription = Str_Form_Fields_Descrip[intLoop];

                                    if (Str_Form_Fields_Editable[intLoop] == "Y")
                                    {
                                        objUserobjectMD.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                                    }
                                    else
                                    {
                                        objUserobjectMD.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tNO;
                                    }
                                    objUserobjectMD.FormColumns.Add();
                                }
                            }
                        }
                        if (Menu == "Y")
                        {
                            objUserobjectMD.MenuItem = SAPbobsCOM.BoYesNoEnum.tYES;
                            objUserobjectMD.MenuCaption = MenuCaption;
                            objUserobjectMD.FatherMenuID = Convert.ToInt32(FatherMenuID);
                            objUserobjectMD.Position = Convert.ToInt32(Position);
                            objUserobjectMD.MenuUID = MenuUID;
                        }
                    }

                    lRetCode = objUserobjectMD.Add();
                    if (lRetCode != 0)
                    {
                        oCompany.GetLastError(out lErrCode, out sErrMsg);
                        SBO_Appln.MessageBox("Error in While Creating UDO - " + lErrCode + " - " + sErrMsg);
                    }
                }
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserobjectMD);
                GC.Collect();
            }
        }
        #endregion

        #region Create Table
        public bool CreateTable(string TableName,string TableDescription, SAPbobsCOM.BoUTBTableType TableType)
        {
            SBO_Appln.StatusBar.SetText("Creating Tables Please Wait...........", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            SAPbobsCOM.UserTablesMD objUserTableMD = null;
            objUserTableMD = (SAPbobsCOM.UserTablesMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
            if (!objUserTableMD.GetByKey(TableName))
            {
                try
                {
                    objUserTableMD.TableName = TableName;
                    objUserTableMD.TableDescription = TableDescription;
                    objUserTableMD.TableType = TableType;
                    lRetCode = objUserTableMD.Add();
                    if (lRetCode == 0)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    SBO_Appln.MessageBox(ex.Message);
                    return false;
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserTableMD);
                    GC.Collect();
                }
            }
            else
            {
                return false;
            }
        }
        #endregion

        #region Create Field

        #region Create Field
        public void addField(string TableName, string ColumnName, string ColDescription, SAPbobsCOM.BoFieldTypes FieldType, int Size, SAPbobsCOM.BoFldSubTypes SubType, string ValidValues, string ValidDescriptions, string SetValidValue, string LinkTable = "", string SysObj = "")
        {
            if (!isColumnExist(TableName, ColumnName))
            {
                int intLoop;
                string[] strValue, strDesc;
                SAPbobsCOM.UserFieldsMD objUserFieldMD;
                objUserFieldMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                try
                {
                    strValue = ValidValues.Split(Convert.ToChar(","));
                    strDesc = ValidDescriptions.Split(Convert.ToChar(","));
                    if (strValue.GetLength(0) != strDesc.GetLength(0))
                    {
                        throw new Exception("Invalid Valid Values");
                    }

                    objUserFieldMD.TableName = TableName;
                    objUserFieldMD.Name = ColumnName;
                    objUserFieldMD.Description = ColDescription;
                    objUserFieldMD.Type = FieldType;
                    if (Size > 0)
                    {
                        if (FieldType != SAPbobsCOM.BoFieldTypes.db_Numeric)
                        {
                            objUserFieldMD.Size = Size;
                        }
                        else
                        {
                            objUserFieldMD.EditSize = Size;
                        }
                    }
                    objUserFieldMD.SubType = SubType;
                    if (ValidValues != "")
                    {
                        if (SetValidValue != "")
                        {
                            objUserFieldMD.DefaultValue = SetValidValue;
                            for (intLoop = 0; intLoop <= strValue.GetLength(0) - 1; intLoop++)
                            {
                                objUserFieldMD.ValidValues.Value = strValue[intLoop];
                                objUserFieldMD.ValidValues.Description = strDesc[intLoop];
                                objUserFieldMD.ValidValues.Add();
                            }
                        }
                        else
                        {
                            for (intLoop = 0; intLoop <= strValue.GetLength(0) - 1; intLoop++)
                            {
                                objUserFieldMD.ValidValues.Value = strValue[intLoop];
                                objUserFieldMD.ValidValues.Description = strDesc[intLoop];
                                objUserFieldMD.ValidValues.Add();
                            }
                        }
                    }
                    if (LinkTable != "")
                    {
                        objUserFieldMD.LinkedTable = LinkTable;
                    }
                    if (SysObj != "")
                    {
                        switch (SysObj)
                        {
                            case "64":
                                objUserFieldMD.LinkedSystemObject = SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulWarehouses;
                                break;

                            case "2":
                                objUserFieldMD.LinkedSystemObject = SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulBusinessPartners;
                                break;

                            case "22":
                                objUserFieldMD.LinkedSystemObject = SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulPurchaseOrders;
                                break;

                            case "17":
                                objUserFieldMD.LinkedSystemObject = SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulOrders;
                                break;

                            case "3":
                                objUserFieldMD.LinkedSystemObject = SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulBanks;
                                break;

                            case "1":
                                objUserFieldMD.LinkedSystemObject = SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulChartOfAccounts;
                                break;
                        }
                    }
                    if (objUserFieldMD.Add() != 0)
                    {
                        oCompany.GetLastError(out lErrCode, out sErrMsg);
                        SBO_Appln.MessageBox(lErrCode + " - " + sErrMsg);
                    }
                }
                catch (Exception)
                {
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD);
                    GC.Collect();
                }
            }
        }
        #endregion

        #region Check Field
        private bool isColumnExist(string TableName, string ColumnName)
        {
            SAPbobsCOM.Recordset objRecordSet;
            objRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string Str = "";
                if (TableName.Length > 4)
                {
                    if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        Str = "SELECT * FROM CUFD WHERE \"TableID\" = '@" + TableName + "' AND \"AliasID\" = '" + ColumnName + "'";
                    }
                    else
                    {
                        Str = "SELECT * FROM CUFD WHERE TableID = '@" + TableName + "' AND AliasID = '" + ColumnName + "'";
                    }
                }
                else
                {
                    if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        Str = "SELECT * FROM CUFD WHERE \"TableID\" = '" + TableName + "' AND \"AliasID\" = '" + ColumnName + "'";
                    }
                    else
                    {
                        Str = "SELECT * FROM CUFD WHERE TableID = '" + TableName + "' AND AliasID = '" + ColumnName + "'";
                    }
                }
                objRecordSet.DoQuery(Str);
                if (objRecordSet.RecordCount > 0)
                {
                    return true;
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecordSet);
                GC.Collect();
            }
            return false;
        }
        #endregion

        #region AlphaNumeric Field
        public void AddAlphaField(string TableName, string ColumnName, string ColDescription, int Size)
        {
            try
            {
                addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "");
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #region Alpha Field With Types
        public void AddAlphaField(string TableName, string ColumnName, string ColDescription, int Size, SAPbobsCOM.BoFldSubTypes SubType)
        {
            try
            {
                addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SubType, "", "", "");
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #region AplhaNumeric Field with ValidValues
        public void AddAlphaField(string TableName, string ColumnName, string ColDescription, int Size, string ValidValues, string ValidDescriptions, string SetValidValue)
        {
            try
            {
                addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SAPbobsCOM.BoFldSubTypes.st_None, ValidValues, ValidDescriptions, SetValidValue);
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #region AlphaMemo Field
        public void AddAlphaMemoField(string TableName, string ColumnName, string ColDescription, int Size)
        {
            try
            {
                addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Memo, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "");
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #region Numeric Field
        public void AddNumericField(string TableName, string ColumnName, string ColDescription, int Size)
        {
            try
            {
                addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Numeric, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "");
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #region Numeric Field with ValidValues
        public void AddNumericField(string TableName, string ColumnName, string ColDescription, int Size, string ValidValues, string ValidDescriptions, string DefultValue)
        {
            try
            {
                addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Numeric, Size, SAPbobsCOM.BoFldSubTypes.st_None, ValidValues, ValidDescriptions, DefultValue);
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #region Float Field
        public void AddFloatField(string TableName, string ColumnName, string ColDescription, SAPbobsCOM.BoFldSubTypes SubType)
        {
            try
            {
                addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Float, 0, SubType, "", "", "");
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #region Date Field
        public void AddDateField(string TableName,string ColumnName, string ColDescription, SAPbobsCOM.BoFldSubTypes SubType)
        {
            try
            {
                addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Date, 0, SubType, "", "", "");
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #region Attachment Field
        public void AddAttachField(string TableName, string ColumnName, string ColDescription, SAPbobsCOM.BoFldSubTypes SubType)
        {
            try
            {
                addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, 0, SubType, "", "", "");
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #region Linked Table Field
        public void AddFieldwithLinkTable(string TableName, string ColumnName, string ColDescription, int Size, string LinkTable)
        {
            try
            {
                addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", LinkTable);
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #region Link UDO Field
        public void AddFieldwithSysObj(string TableName, string ColumnName, string ColDescription, string SysObj)
        {
            try
            {
                addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", "", SysObj);
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #endregion

        #endregion

        #region Query Manager Update

        #region Create Query Category
        public void create_queryCat(string QueryCatName, string Querycatpermission)
        {
            try
            {
                if (!isQueryCatExist(QueryCatName))
                {
                    SAPbobsCOM.QueryCategories objquerycat;
                    objquerycat = (SAPbobsCOM.QueryCategories)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQueryCategories);
                    try
                    {
                        objquerycat.Name = QueryCatName;
                        objquerycat.Permissions = Querycatpermission;
                        lRetCode = objquerycat.Add();
                        if (lRetCode != 0)
                        {
                            oCompany.GetLastError(out lErrCode, out sErrMsg);
                            SBO_Appln.MessageBox(lErrCode + "-" + sErrMsg);
                        }
                    }
                    catch (Exception ex)
                    {
                        SBO_Appln.MessageBox(ex.Message);
                    }
                    finally
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objquerycat);
                        GC.Collect();
                    }
                }
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #region Check Query Category
        public bool isQueryCatExist(string QueryCatName)
        {
            SAPbobsCOM.Recordset objRecordSet1;
            objRecordSet1 = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string Str = "";
            if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
            {
                Str = "SELECT \"CategoryId\" FROM OQCN WHERE \"CatName\" = '" + QueryCatName + "'";
            }
            else
            {
                Str = "SELECT CategoryId FROM OQCN WHERE CatName = '" + QueryCatName + "'";
            }
            try
            {
                objRecordSet1.DoQuery(Str);
                if (objRecordSet1.RecordCount > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
                return false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecordSet1);
                GC.Collect();
            }
        }
        #endregion

        #region Create Query
        public void createQuery(string QueryCatName, string QName, string QueryTextFile)
        {
            SAPbobsCOM.UserQueries objquery = null;
            objquery = (SAPbobsCOM.UserQueries)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries);
            SBO_Appln.StatusBar.SetText("Creating Query Categories and Queries Please Wait...........", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            try
            {
                int QueryCatId = 0;
                string Qstring;
                SAPbobsCOM.Recordset objRecordSet1 = null;
                objRecordSet1 = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string Str = "";
                if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    Str = "SELECT \"CategoryId\" FROM OQCN WHERE \"CatName\" = '" + QueryCatName + "'";
                }
                else
                {
                    Str = "SELECT CategoryId FROM OQCN WHERE CatName = '" + QueryCatName + "'";
                }
                 
                objRecordSet1.DoQuery(Str);
                if (objRecordSet1.RecordCount > 0)
                {
                    QueryCatId = Convert.ToInt32(objRecordSet1.Fields.Item("CategoryId").Value);
                }

                Qstring = Reading_Addon_Query_NotePad(QueryTextFile);

                if (!isQueryExist(QueryCatName, QName))
                {
                    objquery.QueryCategory = QueryCatId;
                    objquery.QueryDescription = QName;
                    objquery.Query = Qstring;
                    lRetCode = objquery.Add();
                    if (lRetCode != 0)
                    {
                        oCompany.GetLastError(out lErrCode, out sErrMsg);
                        SBO_Appln.MessageBox(lErrCode + "-" + sErrMsg);
                    }
                }
                else
                {
                    SAPbobsCOM.Recordset objRecordSet2 = null;
                    objRecordSet2 = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string Str1 = "";
                    if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        Str1 = "SELECT * FROM OUQR A INNER JOIN OQCN B ON A.\"QCategory\" = B.\"CategoryId\" WHERE B.\"CatName\" = '" + QueryCatName + "' AND A.\"QName\" = '" + QName + "'";
                    }
                    else
                    {
                        Str1 = "SELECT * FROM OUQR A INNER JOIN OQCN B ON A.QCategory = B.CategoryId WHERE B.CatName = '" + QueryCatName + "' AND A.QName = '" + QName + "'";
                    }
                                                           
                    objRecordSet2.DoQuery(Str1);
                    if (Convert.ToString(objRecordSet2.Fields.Item("U_AV_Sel").Value) == "Y")
                    {
                        if (Query_Backup(Convert.ToString(objRecordSet2.Fields.Item("QName").Value), Convert.ToString(objRecordSet2.Fields.Item("QString").Value)))
                        {
                            int InternalKey = 0;
                            if (objRecordSet2.RecordCount > 0)
                            {
                                InternalKey = Convert.ToInt32(objRecordSet2.Fields.Item("IntrnalKey").Value);
                            }
                            if (objquery.GetByKey(InternalKey, QueryCatId))
                            {
                                objquery.Query = Qstring;
                                lRetCode = objquery.Update();
                                if (lRetCode != 0)
                                {
                                    oCompany.GetLastError(out lErrCode, out sErrMsg);
                                    SBO_Appln.MessageBox(lErrCode + "-" + sErrMsg);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objquery);
                GC.Collect();
            }
        }
        #endregion

        #region Check Query
        private bool isQueryExist(string QueryCatName, string QName)
        {
            SAPbobsCOM.Recordset objRecordSet1 = null;
            objRecordSet1 = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string Str = "";
            if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
            {
                Str = "SELECT * FROM OUQR A INNER JOIN OQCN B ON A.\"QCategory\" = B.\"CategoryId\" WHERE B.\"CatName\" = '" + QueryCatName + "' AND A.\"QName\" = '" + QName + "'";
            }
            else
            {
                Str = "SELECT * FROM OUQR A INNER JOIN OQCN B ON A.QCategory = B.CategoryId WHERE B.CatName = '" + QueryCatName + "' AND A.QName = '" + QName + "'";
            }
            try
            {
                objRecordSet1.DoQuery(Str);
                if (objRecordSet1.RecordCount > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
                return false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecordSet1);
                GC.Collect();
            }
        }
        #endregion

        #region Reading Addon Query for Embedded TextFile
        public string Reading_Addon_Query_NotePad(string QueryTextFile)
        {
            try
            {
                string strResource;
                strResource = Assembly.GetExecutingAssembly().GetName().Name + ".TextFiles." + QueryTextFile;
                Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(strResource);
                using (StreamReader NotepadPath = new System.IO.StreamReader(stream))
                {
                    string NotePadSingleLine = "";
                    NotePadSingleLine = NotepadPath.ReadToEnd();
                    NotepadPath.Close();
                    return NotePadSingleLine;
                }
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
                return "";
            }
        }
        #endregion

        #region Reading Addon Query - Idle
        public string Reading_Addon_Query_NotePad1(string QueryTextFile)
        {
            try
            {
                string fileName = AppDomain.CurrentDomain.BaseDirectory + "\\" + QueryTextFile;
                if (System.IO.File.Exists(fileName))
                {
                    System.IO.StreamReader NotepadPath = new System.IO.StreamReader(fileName);
                    string NotePadSingleLine = "";

                    NotePadSingleLine = NotepadPath.ReadToEnd();
                    NotepadPath.Close();
                    return NotePadSingleLine;
                }
                else
                {
                    SBO_Appln.MessageBox("File is not found in this Path - '" + fileName + "'");
                    return "";
                }
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
                return "";
            }
        }
        #endregion

        #region Query Backup
        public bool Query_Backup(string QName, string QString)
        {
            try
            {
                SAPbobsCOM.UserTable User_Table;
                User_Table = oCompany.UserTables.Item("AV_QRBPT");
                try
                {
                    string Dat = DateTime.Now.ToString("dd MMM yyyy HH:mm:ss:fff");
                    User_Table.Code = Dat;
                    User_Table.Name = Dat;
                    User_Table.UserFields.Fields.Item("U_AV_QName").Value = QName;
                    User_Table.UserFields.Fields.Item("U_AV_QString").Value = QString;
                    User_Table.UserFields.Fields.Item("U_AV_User").Value = oCompany.UserName;
                    lRetCode = User_Table.Add();
                    if (lRetCode != 0)
                    {
                        oCompany.GetLastError(out lErrCode, out sErrMsg);
                        SBO_Appln.MessageBox(QName + " is not updated in Addon Query Backup Table" + lErrCode + "-" + sErrMsg);
                        return false;
                    }
                    else
                    {
                        return true;
                    }       
                }
                catch (Exception ex)
                {
                    SBO_Appln.MessageBox(ex.Message + "Addon Query Backup Table is not updated");
                    return false;
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(User_Table);
                    GC.Collect();
                }
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
                return false;
            }
        }
        #endregion

        #endregion

        #region Create UI Controls in SAP B1 Forms

        #region Static
        public void Adding_Items_Static(string UID, int Width, int Top, int Left, int Height, string Caption, string FormUID, int FromPane = 0, int ToPane = 0, bool Visible = true)
        {
            SAPbouiCOM.StaticText oStatic = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Form objform = null;
            objform = SBO_Appln.Forms.Item(FormUID);
            oItem = objform.Items.Add(UID, SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Width = Width;
            oItem.Top = Top;
            oItem.Left = Left;
            oItem.Height = Height;
            oItem.FromPane = FromPane;
            oItem.ToPane = ToPane;
            oItem.Visible = Visible;
            //oItem.LinkTo = Editstr;
            oStatic = (SAPbouiCOM.StaticText)oItem.Specific;
            oStatic.Caption = Caption;
        }
        #endregion

        #region Edit
        public void Adding_Items_Edit(string UID, int Width, int Top, int Left, int Height, string tablename, string fieldname, string FormUID, int FromPane = 0, int ToPane = 0, bool Enabled = true, bool Visible = true)
        {
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Form objform = null;
            objform = SBO_Appln.Forms.Item(FormUID);
            oItem = objform.Items.Add(UID, SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Width = Width;
            oItem.Top = Top;
            oItem.Left = Left;
            oItem.Height = Height;
            oItem.FromPane = FromPane;
            oItem.ToPane = ToPane;
            oItem.Enabled = Enabled;
            oItem.Visible = Visible;
            //oItem.LinkTo = Editstr;
            oEdit = (SAPbouiCOM.EditText)oItem.Specific;
            //oEdit.TabOrder = i;
            oEdit.DataBind.SetBound(true, tablename, fieldname);
        }
        #endregion

        #region Combo
        public void Adding_Items_Combo(string UID, int Width, int Top, int Left, int Height, string tablename, string fieldname, string FormUID, int FromPane = 0, int ToPane = 0, bool Enabled = true)
        {
            SAPbouiCOM.ComboBox oCombo = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Form objform = null;
            objform = SBO_Appln.Forms.Item(FormUID);
            oItem = objform.Items.Add(UID, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oItem.Width = Width;
            oItem.Top = Top;
            oItem.Left = Left;
            oItem.Height = Height;
            oItem.FromPane = FromPane;
            oItem.ToPane = ToPane;
            oItem.Enabled = Enabled;
            oItem.DisplayDesc = true;
            oCombo = (SAPbouiCOM.ComboBox)oItem.Specific;
            oCombo.DataBind.SetBound(true, tablename, fieldname);
        }
        #endregion

        #region Check
        public void Adding_Items_Check(string UID, int Width, int Top, int Left, int Height, string Caption, string tablename, string fieldname, string FormUID, int FromPane = 0, int ToPane = 0, bool Enabled = true)
        {
            SAPbouiCOM.CheckBox oCheck = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Form objform = null;
            objform = SBO_Appln.Forms.Item(FormUID);
            oItem = objform.Items.Add(UID, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oItem.Width = Width;
            oItem.Top = Top;
            oItem.Left = Left;
            oItem.Height = Height;
            oItem.FromPane = FromPane;
            oItem.ToPane = ToPane;
            oItem.Enabled = Enabled;
            oCheck = (SAPbouiCOM.CheckBox)oItem.Specific;
            oCheck.Caption = Caption;
            oCheck.DataBind.SetBound(true, tablename, fieldname);
        }
        #endregion

        #region Button
        public void Adding_Items_Button(string UID, int Width, int Top, int Left, int Height, string Caption, string FormUID, int FromPane = 0, int ToPane = 0, bool Enabled = true)
        {
            SAPbouiCOM.Button oButton = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Form objform = null;
            objform = SBO_Appln.Forms.Item(FormUID);
            oItem = objform.Items.Add(UID, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Width = Width;
            oItem.Top = Top;
            oItem.Left = Left;
            oItem.Height = Height;
            oItem.FromPane = FromPane;
            oItem.ToPane = ToPane;
            oItem.Enabled = Enabled;
            //oItem.LinkTo = Editstr;
            oButton = (SAPbouiCOM.Button)oItem.Specific;
            oButton.Caption = Caption;
        }
        #endregion

        #region ButtonCombo
        public void Adding_Items_ButtonCombo(string UID, int Width, int Top, int Left, int Height, string Caption, string FormUID, int FromPane = 0, int ToPane = 0, bool Enabled = true)
        {
            SAPbouiCOM.ButtonCombo oButCombo = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Form objform = null;
            objform = SBO_Appln.Forms.Item(FormUID);
            oItem = objform.Items.Add(UID, SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO);
            oItem.Width = Width;
            oItem.Top = Top;
            oItem.Left = Left;
            oItem.Height = Height;
            oItem.FromPane = FromPane;
            oItem.ToPane = ToPane;
            oItem.Enabled = Enabled;
            oButCombo = (SAPbouiCOM.ButtonCombo)oItem.Specific;
            oButCombo.Caption = Caption;            
        }
        #endregion

        #region Grid
        public void Adding_Items_Grid(string UID, int Width, int Top, int Left, int Height, string FormUID, int FromPane = 0, int ToPane = 0)
        {
            SAPbouiCOM.Grid oGrid = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Form objform = null;
            objform = SBO_Appln.Forms.Item(FormUID);
            oItem = objform.Items.Add(UID, SAPbouiCOM.BoFormItemTypes.it_GRID);
            oItem.Width = Width;
            oItem.Top = Top;
            oItem.Left = Left;
            oItem.Height = Height;
            oItem.FromPane = FromPane;
            oItem.ToPane = ToPane;
            oGrid = (SAPbouiCOM.Grid)oItem.Specific;
        }
        #endregion

        #region LinkButton
        public void Adding_Items_Link_Button(string UID, int Width, int Top, int Left, int Height, string FormUID, string LinkTo, int FromPane = 0, int ToPane = 0)
        {
            SAPbouiCOM.LinkedButton oLink = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Form objform = null;
            objform = SBO_Appln.Forms.Item(FormUID);
            oItem = objform.Items.Add(UID, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oItem.Width = Width;
            oItem.Top = Top;
            oItem.Left = Left;
            oItem.Height = Height;
            oItem.FromPane = FromPane;
            oItem.ToPane = ToPane;
            oLink = (SAPbouiCOM.LinkedButton)oItem.Specific;
            oLink.Item.LinkTo = LinkTo;
        }
        #endregion

        #region Picture
        public void Adding_Items_Picture(string UID, int Width, int Top, int Left, int Height, string tablename, string fieldname, string FormUID, int FromPane = 0, int ToPane = 0, bool Enabled = true)
        {
            SAPbouiCOM.PictureBox oPicbox = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Form objform = null;
            objform = SBO_Appln.Forms.Item(FormUID);
            oItem = objform.Items.Add(UID, SAPbouiCOM.BoFormItemTypes.it_PICTURE);
            oItem.Width = Width;
            oItem.Top = Top;
            oItem.Left = Left;
            oItem.Height = Height;
            oItem.FromPane = FromPane;
            oItem.ToPane = ToPane;
            oPicbox = (SAPbouiCOM.PictureBox)oItem.Specific;
            oPicbox.DataBind.SetBound(true, tablename, fieldname);
        }
        #endregion

        #region Folder
        public void Adding_Items_Folder(string UID, int Width, int Top, int Left, int Height, int Pane, string GroupWith, string Caption, string fieldname, string FormUID, bool AutoPaneSelection = false)
        {
            SAPbouiCOM.Folder oPanel = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Form objform = null;
            objform = SBO_Appln.Forms.Item(FormUID);
            oItem = objform.Items.Add(UID, SAPbouiCOM.BoFormItemTypes.it_FOLDER);
            oItem.Width = Width;
            oItem.Top = Top;
            oItem.Left = Left;
            oItem.Height = Height;
            oPanel = (SAPbouiCOM.Folder)oItem.Specific;
            oPanel.Caption = Caption;
            oPanel.Pane = Pane;
            oPanel.AutoPaneSelection = false;
            oPanel.DataBind.SetBound(true, "", fieldname);
            oPanel.GroupWith(GroupWith);
        }
        #endregion

        #endregion

        #region Execute Query
        public string Query_Execute(string Str)
        {
            try
            {
                string Output_Str = "";
                SAPbobsCOM.Recordset Orec;
                Orec = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                Orec.DoQuery(Str);
                if (Orec.RecordCount > 0)
                {
                    Output_Str = Convert.ToString(Orec.Fields.Item(0).Value);
                }
                return Output_Str;
            }
            catch (Exception ex)
            {
                return "";
            }
        }
        #endregion

        #region Get Embedded Excel File Path
        public string Get_Embedded_Excel_FilePath(string FileName)
        {
            try
            {
                return Assembly.GetExecutingAssembly().GetName().Name + ".ExcelTemplates." + FileName;
            }
            catch (Exception)
            {
                return "";
            }
        }
        #endregion

        #region Open & Save Dialog

        #region Browse Thread
        public void Browse(string BrowseType, string FormUID = "", string ComboUID = "", string EditUID = "", string FileName = "")
        {
            try
            {
                if ((BrowseType == "Open"))
                {
                    ShowFolderBrowserThread = new Thread(() => ShowFolderBrowser_Text_Open(FormUID, ComboUID, EditUID));
                }
                else if ((BrowseType == "Save"))
                {
                    ShowFolderBrowserThread = new Thread(() => ShowFolderBrowser_Text_Save(FileName));
                }
                else
                {
                    return;
                }

                if ((ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted))
                {
                    ShowFolderBrowserThread.SetApartmentState(ApartmentState.STA);
                    ShowFolderBrowserThread.Start();
                    ShowFolderBrowserThread.Join();
                }
                else if ((ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Stopped))
                {
                    ShowFolderBrowserThread.Start();
                    ShowFolderBrowserThread.Join();
                }

                while ((ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Running))
                {
                    Application.DoEvents();
                }

            }
            catch (Exception ex)
            {
                SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Open Dialog
        private void ShowFolderBrowser_Text_Open(string FormUID, string ComboUID, string EditUID)
        {
            OpenFileDialog oDialogBox = new OpenFileDialog();
            try
            {
                Process[] MyProcs = null;
                SAPbouiCOM.Form oForm;
                oForm = SBO_Appln.Forms.Item(FormUID);

                if (ComboUID == "text")
                {
                    oDialogBox.Filter = "NotePad(*.txt)|*.txt";
                }
                else if (ComboUID == "excel")
                {
                    oDialogBox.Filter = "Excel Worksheets(*.xls;*.xlsx)|*.xls;*.xlsx";
                }
                else if (ComboUID == "license")
                {
                    oDialogBox.Filter = "License files (*.lic)|*.lic";
                }
                else
                {
                    SAPbouiCOM.ComboBox oCombo;
                    oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item(ComboUID).Specific;
                    if ((oCombo.Selected.Value == "txt"))
                    {
                        oDialogBox.Filter = "NotePad(*.txt)|*.txt";
                    }
                    else
                    {
                        oDialogBox.Filter = "Excel Worksheets(*.xls;*.xlsx)|*.xls;*.xlsx";
                    }
                }

                string OrgTitle = SBO_Appln.Desktop.Title;
                SBO_Appln.Desktop.Title = ("SBO under " + oCompany.UserName);
                MyProcs = Process.GetProcessesByName("SAP Business One");
                for (int i = 0; i <= MyProcs.Length - 1; i++)
                {
                    if ((MyProcs[i].MainWindowTitle == SBO_Appln.Desktop.Title))
                    {
                        WindowWrapper MyWindow = new WindowWrapper(MyProcs[i].MainWindowHandle);
                        if ((oDialogBox.ShowDialog(MyWindow) == DialogResult.OK))
                        {
                            ((SAPbouiCOM.EditText)oForm.Items.Item(EditUID).Specific).Value = oDialogBox.FileName;
                        }

                    }

                }

                SBO_Appln.Desktop.Title = OrgTitle;
            }
            catch (Exception ex)
            {
                SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                ShowFolderBrowserThread.Abort();
            }
        }
        #endregion

        #region Save Dialog
        private void ShowFolderBrowser_Text_Save(string FileName)
        {
            SaveFileDialog oDialogBox = new SaveFileDialog();
            try
            {
                Process[] MyProcs = null;
                oDialogBox.Filter = "Excel WorkBook(*.xlsx;*.xls)|*.xlsx;*.xls";
                string OrgTitle = SBO_Appln.Desktop.Title;
                SBO_Appln.Desktop.Title = ("SBO under " + oCompany.UserName);
                MyProcs = Process.GetProcessesByName("SAP Business One");
                for (int i = 0; i <= MyProcs.Length - 1; i++)
                {
                    if ((MyProcs[i].MainWindowTitle == SBO_Appln.Desktop.Title))
                    {
                        WindowWrapper MyWindow = new WindowWrapper(MyProcs[i].MainWindowHandle);
                        if ((oDialogBox.ShowDialog(MyWindow) == System.Windows.Forms.DialogResult.OK))
                        {

                            string SourcePath = Assembly.GetExecutingAssembly().GetName().Name + ".ExcelTemplates." + FileName;
                            string DestinationPath = oDialogBox.FileName;
                            using (FileStream fileStream = new FileStream(DestinationPath, FileMode.Append))
                            {
                                Assembly.GetExecutingAssembly().GetManifestResourceStream(SourcePath).CopyTo(fileStream);
                            }
                            SBO_Appln.MessageBox(("ExcelFile is Created in this Path- " + (DestinationPath + "")));
                        }

                    }

                }

                SBO_Appln.Desktop.Title = OrgTitle;
            }
            catch (Exception ex)
            {
                SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                ShowFolderBrowserThread.Abort();
            }
        }
        #endregion

        #endregion

        #region Open Link Button
        public void Link_Button_Function(string ObjectType, string DocEntry)
        {
            try
            {
                SAPbobsCOM.Recordset orec = null;
                orec = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string Str;
                if (ObjectType == "13")
                {
                    SBO_Appln.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_Invoice, ObjectType, DocEntry);
                }   
                else if (ObjectType == "15")
                {
                    SBO_Appln.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_DeliveryNotes, ObjectType, DocEntry);
                }
                else if (ObjectType == "16")
                {
                    SBO_Appln.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_DeliveryNotesReturns, ObjectType, DocEntry);
                }
                else if (ObjectType == "24")
                {
                    SBO_Appln.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_Receipt, ObjectType, DocEntry);
                }
                else if (ObjectType == "17")
                {
                    SBO_Appln.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_Order, ObjectType, DocEntry);
                }
                else if (ObjectType == "30")
                {
                    Str = "SELECT TransId FROM OJDT WHERE Number='" + DocEntry + "'";
                    orec.DoQuery(Str);
                    SBO_Appln.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_JournalPosting, ObjectType, Convert.ToString(orec.Fields.Item("TransId").Value));
                }
                else if (ObjectType == "14")
                {
                    SBO_Appln.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_InvoiceCreditMemo, ObjectType, DocEntry);
                }
                else if (ObjectType == "18")
                {
                    SBO_Appln.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_PurchaseInvoice, ObjectType, DocEntry);
                }
                else if (ObjectType == "22")
                {
                    SBO_Appln.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_PurchaseOrder, ObjectType, DocEntry);
                }
                else if (ObjectType == "19")
                {
                    SBO_Appln.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_PurchaseInvoiceCreditMemo, ObjectType, DocEntry);
                }
                else if (ObjectType == "25")
                {
                    SBO_Appln.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_Deposit, ObjectType, DocEntry);
                }
                else if (ObjectType == "67")
                {
                    SBO_Appln.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_StockTransfers, ObjectType, DocEntry);
                }
                else if (ObjectType == "203")
                {
                    SBO_Appln.ActivateMenuItem("2071");
                    SBO_Appln.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    Str = "select DocNum from ODPI where DocEntry='" + DocEntry + "'";
                    orec.DoQuery(Str);
                    ((SAPbouiCOM.EditText)SBO_Appln.Forms.ActiveForm.Items.Item("8").Specific).Value = Convert.ToString(orec.Fields.Item("DocNum").Value);
                    SBO_Appln.Forms.ActiveForm.Items.Item("1").Click();
                }
                else if (ObjectType == "46")
                {
                    SBO_Appln.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_VendorPayment, ObjectType, DocEntry);
                }
                else if (ObjectType == "4")
                {
                    SBO_Appln.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_Items, ObjectType, DocEntry);
                }
            }
            catch (Exception ex)
            {
                SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Find DefaultForm MenuId
        public void DefaultForm_Menu(string MenuString)
        {
            try
            {
                SAPbouiCOM.MenuItem oMenuItem;
                string menuuid = "";
                oMenuItem = SBO_Appln.Menus.Item("47616");
                for (int i = 0; i <= oMenuItem.SubMenus.Count - 1; i++)
                {
                    if (oMenuItem.SubMenus.Item(i).String == MenuString)
                    {
                        menuuid = oMenuItem.SubMenus.Item(i).UID;
                        break;
                    }
                }
                if (menuuid != "")
                {
                    SBO_Appln.ActivateMenuItem(menuuid);
                }
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #region AddRow & DeleteRow

        #region Row No Arrange
        public void Row_No_Arrange(string Mat_UID, string Col_UID, string LineTable, string FormUID, string Add_Str)
        {
            try
            {
                SAPbouiCOM.Form oForm;
                oForm = SBO_Appln.Forms.Item(FormUID);

                SAPbouiCOM.Matrix objmat;
                objmat = (SAPbouiCOM.Matrix)oForm.Items.Item(Mat_UID).Specific;

                int i = objmat.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder);

                Table_Update_For_Add_Row(Mat_UID, Col_UID, LineTable, FormUID);

                if (Add_Str == "ADD")
                {
                    objmat.AddRow(1, i);
                    objmat.ClearRowData(i + 1);
                    oForm.Update();
                }

                if (Add_Str != "Close")
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }
                }

                for (int k = 1; k <= objmat.VisualRowCount;k++)
                {
                    ((SAPbouiCOM.EditText)objmat.Columns.Item("#").Cells.Item(k).Specific).Value = Convert.ToString(k);
                }
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #region Table Update for Add Row
        public void Table_Update_For_Add_Row(string Mat_UID, string Col_UID, string Line_Table, string FormUID)
        {
            try
            {
                SAPbouiCOM.Form oform;
                oform = SBO_Appln.Forms.Item(FormUID);

                SAPbouiCOM.Matrix oMatrix;
                oMatrix = (SAPbouiCOM.Matrix)oform.Items.Item(Mat_UID).Specific;

                oMatrix.FlushToDataSource();
                oMatrix.LoadFromDataSource();

                oform.DataSources.DBDataSources.Item(Line_Table).Clear();

                if (oMatrix.Columns.Item(Col_UID).Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                {
                    SAPbouiCOM.ComboBox oCombo;
                    for (int index = oMatrix.VisualRowCount; index >= 1; index--)
                    {
                        oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(Col_UID).Cells.Item(index).Specific;
                        if (oCombo.Value == "")
                        {
                            oMatrix.DeleteRow(index);
                        }
                    }
                }
                else
                {
                    for (int index = oMatrix.VisualRowCount; index >= 1; index--)
                    {
                        if (((SAPbouiCOM.EditText)oMatrix.Columns.Item(Col_UID).Cells.Item(index).Specific).Value == "")
                        {
                            oMatrix.DeleteRow(index);
                        }
                    }
                }

                oMatrix.AddRow();
                oMatrix.ClearRowData(oMatrix.VisualRowCount);
                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("#").Cells.Item(oMatrix.VisualRowCount).Specific).Value = Convert.ToString(oMatrix.VisualRowCount);

                oMatrix.FlushToDataSource();
            }
            catch (Exception ex)
            {
                SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #endregion

        #region DateTime

        #region Date_Changing
        //Date Dte_Test = GetDateTimeValue(DateTime.Now.ToString())   ------- Example
        public DateTime GetDateTimeValue(string SBODaAvanikoAGNTMASring)
        {
            SAPbobsCOM.SBObob objBridge;
            objBridge = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
            SBODaAvanikoAGNTMASring = Convert.ToDateTime(SBODaAvanikoAGNTMASring).ToString(Sys_DtFormat);
            SBODaAvanikoAGNTMASring = Convert.ToDateTime(SBODaAvanikoAGNTMASring).ToString("yyyyMMdd");
            return objBridge.Format_StringToDate(SBODaAvanikoAGNTMASring).Fields.Item(0).Value;
        }

        // String  Dte_T= GetSBODateString(Dte_Test)    ------------------------------------Example
        public string GetSBODateString(System.DateTime DateVal)
        {
            SAPbobsCOM.SBObob objBridge;
            objBridge = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
            DateVal = Convert.ToDateTime(DateVal.ToString(Sys_DtFormat));
            return objBridge.Format_DateToString(DateVal).Fields.Item(0).Value;
        }
        //Edit Field Value : objform.items.item("1").Specific.Value  ----------------------------Example
        public string GetDateFromField(string DateVal)
        {
            DateVal = DateVal.Substring(0,4) + DateSep + DateVal.Substring(4, 2) + DateSep + DateVal.Substring(6, 2);
            return DateVal;
        }
        //Insert Query Format : YearMonthDate
        public string GetDateToInsert(string SBODaAvanikoAGNTMASring)
        {
            SBODaAvanikoAGNTMASring = Convert.ToDateTime(SBODaAvanikoAGNTMASring).ToString(Sys_DtFormat);
            SBODaAvanikoAGNTMASring = Convert.ToDateTime(SBODaAvanikoAGNTMASring).ToString("yyyyMMdd");
            return SBODaAvanikoAGNTMASring;
        }

        #endregion

        #endregion

        #region CFL Function - Idle
        public void CFL_Function(string FormUID, ref SAPbouiCOM.ItemEvent pval, string MatrixUID, string No_oF_ColUID, string No_oF_SetValue, string No_oF_ColUID_UserCalc = "", string No_oF_SetValue_UserCalc = "", string UserCalc = "")
        {
            try
            {
                SAPbouiCOM.DataTable dt;
                SAPbouiCOM.ChooseFromListEvent cfl;
                cfl = (SAPbouiCOM.ChooseFromListEvent)pval;
                dt = cfl.SelectedObjects;
                if (dt != null)
                {
                    SAPbouiCOM.Form oForm;
                    oForm = SBO_Appln.Forms.Item(FormUID);

                    SAPbouiCOM.Matrix oMat;
                    oMat = (SAPbouiCOM.Matrix)oForm.Items.Item(MatrixUID).Specific;

                    string[] Col_UID, SetValue;
                    string[] Col_UID_UserColumn, SetValue_User_Column;
                    Col_UID = No_oF_ColUID.Split(Convert.ToChar(","));
                    SetValue = No_oF_SetValue.Split(Convert.ToChar(","));
                    Col_UID_UserColumn = No_oF_ColUID_UserCalc.Split(Convert.ToChar(","));
                    SetValue_User_Column = No_oF_SetValue_UserCalc.Split(Convert.ToChar(","));

                    int RowNo = pval.Row;

                    for (int i = 0; i <= cfl.SelectedObjects.Rows.Count - 1; i++)
                    {
                        double Retail = 0;
                        double Wholesale = 0;

                        if (UserCalc == "Y")
                        {
                            Retail = Convert.ToDouble(Query_Execute("SELECT ISNULL(AvgPrice,0)*(SELECT ISNULL(U_AV_Percen,0) FROM OPLN WHERE U_AV_Prtype='R') FROM OITM WHERE ItemCode='" + dt.GetValue("ItemCode", i) + "'"));
                            Wholesale = Convert.ToDouble(Query_Execute("SELECT ISNULL(AvgPrice,0)*(SELECT ISNULL(U_AV_Percen,0) FROM OPLN WHERE U_AV_Prtype='W') FROM OITM WHERE ItemCode='" + dt.GetValue("ItemCode", i) + "'"));
                        }

                        for (int intLoop = 0; intLoop <= Col_UID.GetLength(0) - 1; intLoop++)
                        {
                            try
                            {
                                oMat.SetCellWithoutValidation(RowNo, Col_UID[intLoop], Convert.ToString(dt.GetValue(SetValue[intLoop], i)));
                            }
                            catch (Exception)
                            {
                                oMat.AddRow();
                                oMat.ClearRowData(oMat.VisualRowCount);
                                ((SAPbouiCOM.EditText)oMat.Columns.Item("#").Cells.Item(oMat.VisualRowCount).Specific).Value = Convert.ToString(oMat.VisualRowCount);

                                oMat.SetCellWithoutValidation(RowNo, Col_UID[intLoop], Convert.ToString(dt.GetValue(SetValue[intLoop], i)));
                            }

                        }

                        //USER COLUMNS VALUE SETUP

                        if (No_oF_ColUID_UserCalc != "")
                        {
                            for (int intLoop = 0; intLoop <= Col_UID_UserColumn.GetLength(0) - 1; intLoop++)
                            {
                                try
                                {
                                    oMat.SetCellWithoutValidation(RowNo, Col_UID_UserColumn[intLoop], SetValue_User_Column[intLoop]);
                                }
                                catch (Exception)
                                {
                                    oMat.AddRow();
                                    oMat.ClearRowData(oMat.VisualRowCount);
                                    ((SAPbouiCOM.EditText)oMat.Columns.Item("#").Cells.Item(oMat.VisualRowCount).Specific).Value = Convert.ToString(oMat.VisualRowCount);
                                    oMat.SetCellWithoutValidation(RowNo, Col_UID_UserColumn[intLoop], SetValue_User_Column[intLoop]);
                                }
                            }
                        }

                        if (UserCalc == "Y")
                        {
                            oMat.SetCellWithoutValidation(RowNo, "Col_7", Convert.ToString(Retail));

                            oMat.SetCellWithoutValidation(RowNo, "Col_8", Convert.ToString(Wholesale));
                        }

                    RowNo = RowNo + 1;

                    }
                }                
            }
            catch (Exception ex)
            {
                SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Add New Line
        public void AddNewLine(SAPbouiCOM.Matrix oMatrix,SAPbouiCOM.DBDataSource oDBDSDetail,int RowID = 1,string ColumnUID = "")
        {
            try
            {
                if (!ColumnUID.Equals(""))
                {
                    if (oMatrix.VisualRowCount > 0)
                    {
                        if (!string.IsNullOrEmpty(oMatrix.Columns.Item(ColumnUID).Cells.Item(RowID).Specific.Value) && RowID == oMatrix.VisualRowCount)
                        {
                            oMatrix.FlushToDataSource();
                            oDBDSDetail.InsertRecord(oDBDSDetail.Size);
                            oDBDSDetail.Offset = oMatrix.VisualRowCount;
                            oDBDSDetail.SetValue("LineID", oDBDSDetail.Offset, (oDBDSDetail.Offset + 1).ToString());
                            oMatrix.LoadFromDataSource();
                        }
                    }
                    else
                    {
                        oMatrix.FlushToDataSource();
                        oMatrix.AddRow(1, -1);
                        oDBDSDetail.InsertRecord(oDBDSDetail.Size);
                        oDBDSDetail.Offset = oMatrix.VisualRowCount - 1;
                        oDBDSDetail.SetValue("LineID", oDBDSDetail.Offset, oMatrix.VisualRowCount.ToString());
                        oMatrix.SetLineData(oMatrix.VisualRowCount);
                        oMatrix.FlushToDataSource();
                    }
                }
                else
                {
                    oMatrix.FlushToDataSource();
                    oMatrix.AddRow(1, -1);
                    oDBDSDetail.InsertRecord(oDBDSDetail.Size);
                    oDBDSDetail.Offset = oMatrix.VisualRowCount - 1;
                    oDBDSDetail.SetValue("LineID", oDBDSDetail.Offset, oMatrix.VisualRowCount.ToString());
                    oMatrix.SetLineData(oMatrix.VisualRowCount);
                    oMatrix.FlushToDataSource();
                }
            }
            catch (Exception ex)
            {
                SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            
        }
        #endregion

        #region SetNewLineSubGrid
        public void SetNewLineSubGrid(
      int UniqID,
      SAPbouiCOM.Matrix oMatSubGrid,
      SAPbouiCOM.DBDataSource oDBDSSubGrid,
      int RowID = 1,
      string ColumnUID = "",
      string[,] DefaulFields = null)
        {
            try
            {
                if (!ColumnUID.Equals(""))
                {
                    if (!string.IsNullOrEmpty(oMatSubGrid.Columns.Item(ColumnUID).Cells.Item(RowID).Specific.Value) && RowID == oMatSubGrid.VisualRowCount)
                    {
                        oMatSubGrid.FlushToDataSource();
                        oMatSubGrid.AddRow(1, -1);
                        oDBDSSubGrid.InsertRecord(oDBDSSubGrid.Size);
                        oDBDSSubGrid.Offset = oMatSubGrid.VisualRowCount - 1;
                        oDBDSSubGrid.SetValue("LineID", oDBDSSubGrid.Offset, Convert.ToString(oMatSubGrid.VisualRowCount));
                        oDBDSSubGrid.SetValue("U_UniqID", oDBDSSubGrid.Offset, Convert.ToString(UniqID));
                        if (DefaulFields != null)
                        {
                            short num1 = (short)(DefaulFields.GetLength(0) - 1);
                            short num2 = 0;
                            while ((int)num2 <= (int)num1)
                            {
                                oDBDSSubGrid.SetValue(DefaulFields[(int)num2, 0], oDBDSSubGrid.Offset, DefaulFields[(int)num2, 1]);
                                ++num2; 
                            }
                        }
                        oMatSubGrid.SetLineData(oMatSubGrid.VisualRowCount);
                        oMatSubGrid.FlushToDataSource();
                        oMatSubGrid.LoadFromDataSource();
                    }
                }
                else
                {
                    oMatSubGrid.FlushToDataSource();
                    oMatSubGrid.AddRow(1, -1);
                    oDBDSSubGrid.InsertRecord(oDBDSSubGrid.Size);
                    oDBDSSubGrid.Offset = oMatSubGrid.VisualRowCount - 1;
                    oDBDSSubGrid.SetValue("LineID", oDBDSSubGrid.Offset, Convert.ToString(oMatSubGrid.VisualRowCount));
                    oDBDSSubGrid.SetValue("U_UniqID", oDBDSSubGrid.Offset, Convert.ToString(UniqID));
                    if (DefaulFields != null)
                    {
                        short num1 = (short)(DefaulFields.GetLength(0) - 1);
                        short num2 = 0;
                        while ((int)num2 <= (int)num1)
                        {
                            oDBDSSubGrid.SetValue(DefaulFields[(int)num2, 0], oDBDSSubGrid.Offset, DefaulFields[(int)num2, 1]);
                            ++num2; 
                        }
                    }
                    oMatSubGrid.SetLineData(oMatSubGrid.VisualRowCount);
                    oMatSubGrid.FlushToDataSource();
                    oMatSubGrid.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            
        }
        #endregion

        #region SetNewLineSubGridForSubgrid
        public void SetNewLineSubGridForSubgrid(
      int UniqID,
      SAPbouiCOM.Matrix oMatSubGrid,
      SAPbouiCOM.DBDataSource oDBDSSubGrid,
      int RowID = 1,
      string DayId = "",
      string ColumnUID = "",
      string[,] DefaulFields = null)
        {
            try
            {
                if (!ColumnUID.Equals(""))
                {
                    if (!string.IsNullOrEmpty(oMatSubGrid.Columns.Item(ColumnUID).Cells.Item(RowID).Specific.Value) && RowID == oMatSubGrid.VisualRowCount)
                    {
                        oMatSubGrid.FlushToDataSource();
                        oMatSubGrid.AddRow(1, -1);
                        oDBDSSubGrid.InsertRecord(oDBDSSubGrid.Size);
                        oDBDSSubGrid.Offset = oMatSubGrid.VisualRowCount - 1;
                        oDBDSSubGrid.SetValue("LineID", oDBDSSubGrid.Offset, Convert.ToString(oMatSubGrid.VisualRowCount));
                        oDBDSSubGrid.SetValue("U_UniqID", oDBDSSubGrid.Offset, Convert.ToString(UniqID));
                        oMatSubGrid.SetLineData(oMatSubGrid.VisualRowCount);
                        oMatSubGrid.FlushToDataSource();
                        oMatSubGrid.LoadFromDataSource();
                    }
                    
                }
                else
                {
                    oMatSubGrid.FlushToDataSource();
                    oMatSubGrid.AddRow(1, -1);
                    oDBDSSubGrid.InsertRecord(oDBDSSubGrid.Size);
                    oDBDSSubGrid.Offset = oMatSubGrid.VisualRowCount - 1;
                    oDBDSSubGrid.SetValue("LineID", oDBDSSubGrid.Offset, Convert.ToString(oMatSubGrid.VisualRowCount));
                    oDBDSSubGrid.SetValue("U_UniqID", oDBDSSubGrid.Offset, Convert.ToString(UniqID));
                    oMatSubGrid.SetLineData(oMatSubGrid.VisualRowCount);
                    oMatSubGrid.FlushToDataSource();
                    oMatSubGrid.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region GetCodeGeneration
        public int GetCodeGeneration(string TableName)
        {
            int num;
            try
            {
                SAPbobsCOM.Recordset Orec = null;
                string QueryStr = null;
                Orec = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                if (oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    QueryStr = "Select IFNULL(Max(IFNULL(\"DocEntry\",0)),0) +1 AS \"Code\" From \"" + TableName.Trim() + "\"";
                }
                else
                {
                    QueryStr = "Select ISNULL(Max(ISNULL(DocEntry,0)),0) + 1 Code From " + TableName.Trim() + "";
                }
                Orec.DoQuery(QueryStr);
                num = int.Parse(Orec.Fields.Item("Code").Value.ToString());
            }
            catch (Exception ex)
            {
                string str_GetCodeGeneration_Function_Failed = "GetCodeGeneration Function Failed";
                Msg(str_GetCodeGeneration_Function_Failed, "S", "W");
                num = 0;
                SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            return num;
        }
        #endregion

        #region Msg
        public void Msg(string strMsg, string msgTime = "S", string errType = "W")
        {
            string upper1 = errType.ToUpper();
            SAPbouiCOM.BoStatusBarMessageType Type = upper1 != "E" ? (upper1 != "W" ? (upper1 != "N" ? (upper1 != "S" ? SAPbouiCOM.BoStatusBarMessageType.smt_Warning : SAPbouiCOM.BoStatusBarMessageType.smt_Success) : SAPbouiCOM.BoStatusBarMessageType.smt_None) : SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : SAPbouiCOM.BoStatusBarMessageType.smt_Error;
            string upper2 = msgTime.ToUpper();
            SAPbouiCOM.BoMessageTime Seconds = upper2  != "M" ? (upper2 != "S" ? (upper2 != "L" ? SAPbouiCOM.BoMessageTime.bmt_Medium : SAPbouiCOM.BoMessageTime.bmt_Long) : SAPbouiCOM.BoMessageTime.bmt_Short) : SAPbouiCOM.BoMessageTime.bmt_Medium;
            SBO_Appln.StatusBar.SetText(strMsg, Seconds, Type);
        }
        #endregion

        #region setComboBoxValue
        public bool setComboBoxValue(SAPbouiCOM.ComboBox oComboBox, string strQry)
        {
            bool flag = false;
            try
            {
                if (oComboBox.ValidValues.Count == 0)
                {
                    SAPbobsCOM.Recordset recordset = null;
                    recordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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
            }
            catch (Exception ex)
            {
                SBO_Appln.StatusBar.SetText("setComboBoxValue Function Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                flag = true;
                return flag;
            }
            return flag;
        }
        #endregion

        #region Intialize
        public void Intialize()
        {
            try
            {
                strBranchForDebitMemo = string.Empty;
                SBO_Appln.ActivateMenuItem("11010");
                SAPbouiCOM.Form activeForm = SBO_Appln.Forms.ActiveForm;
                activeForm.Freeze(true);
                SAPbouiCOM.Matrix matrix = (SAPbouiCOM.Matrix)activeForm.Items.Item("1320000003").Specific;
                int visualRowCount = matrix.VisualRowCount;
                int RowNum = 1;
                while (RowNum <= visualRowCount)
                {
                    if (matrix.IsRowSelected(RowNum))
                    {
                        strBranchForDebitMemo = matrix.Columns.Item("1320000005").Cells.Item(RowNum).Specific.Value.ToString().Trim();
                        break;
                    }
                    ++RowNum; 
                }
                activeForm.Freeze(false);
                activeForm.Items.Item("1320000002").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            }
            catch (Exception ex)
            {
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
                SBO_Appln.StatusBar.SetText("LoadDocumentDate Function Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                flag = false;
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
                SBO_Appln.StatusBar.SetText("LoadComboBoxSeries Function Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                flag = false;
            }
            return flag;
        }
        #endregion

        #region SetComboBoxValueRefresh
        public void SetComboBoxValueRefresh(SAPbouiCOM.ComboBox oComboBox, string strQry)
        {
            try
            {
                SAPbobsCOM.Recordset businessObject = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                Clear_Combo(oComboBox);
                businessObject.DoQuery(strQry);
                businessObject.MoveFirst();
                int num1 = (businessObject.RecordCount - 1);
                int num2 = 0;
                while (num2 <= num1)
                {
                    oComboBox.ValidValues.Add(Convert.ToString(businessObject.Fields.Item(0).Value), Convert.ToString(businessObject.Fields.Item(1).Value));
                    businessObject.MoveNext();
                    { ++num2; }
                }
            }
            catch (Exception ex)
            {
                SBO_Appln.StatusBar.SetText("SetComboBoxValueRefresh Function Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }
        #endregion

        #region Clear_Combo
        public void Clear_Combo(SAPbouiCOM.ComboBox oComboBox)
        {
            int count = oComboBox.ValidValues.Count;
            if (count <= 0)
                return;
            while (count > 0)
            {
                try
                {
                    oComboBox.ValidValues.Remove((count - 1), SAPbouiCOM.BoSearchKey.psk_Index);
                }
                catch (Exception ex)
                {
                    SBO_Appln.StatusBar.SetText("Clear_Combo Function Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
                finally
                {
                    { --count; }
                }
            }
        }
        #endregion

        #region DeleteEmptyRowInFormDataEvent
        public void DeleteEmptyRowInFormDataEvent(SAPbouiCOM.Matrix oMatrix,string ColumnUID,SAPbouiCOM.DBDataSource oDBDSDetail)
        {
            try
            {
                if (oMatrix.VisualRowCount <= 0 || !oMatrix.Columns.Item(ColumnUID).Cells.Item(oMatrix.VisualRowCount).Specific.Value.Equals(""))
                    return;
                oMatrix.DeleteRow(oMatrix.VisualRowCount);
                oDBDSDetail.RemoveRecord(oDBDSDetail.Size - 1);
                oMatrix.FlushToDataSource();
            }
            catch (Exception ex)
            {
                SBO_Appln.StatusBar.SetText("Delete Empty RowIn Function Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
        }
        #endregion

        #region Hana - Query Execution
        public string HANA_GetValue(string QueryStr, HanaConnection oConn)
        {
            string Value = "";
            try
            {
                using (HanaCommand cmd = new HanaCommand(QueryStr, oConn))
                {
                    cmd.CommandType = CommandType.Text;
                    using (HanaDataAdapter sda = new HanaDataAdapter(cmd))
                    {
                        using (DataTable dt = new DataTable())
                        {
                            sda.Fill(dt);

                            if (dt.Rows.Count > 0)
                            {
                                Value = dt.Rows[0][0].ToString();
                            }

                        }
                    }
                }
                return Value;
            }
            catch (Exception)
            {
                return Value;
            }
        }

        public DataTable HANA_ExecuteQuery(string QueryStr, HanaConnection oConn)
        {
            DataTable oDt = new DataTable();
            try
            {
                using (HanaCommand cmd = new HanaCommand(QueryStr, oConn))
                {
                    cmd.CommandType = CommandType.Text;
                    using (HanaDataAdapter sda = new HanaDataAdapter(cmd))
                    {
                        sda.Fill(oDt);
                        return oDt;
                    }
                }
            }
            catch (Exception)
            {
                return oDt;
            }
        }

        public bool HANA_ExecuteNonQuery(string Price_UDT, string Price_SAP, HanaConnection oConn, out string RespMsg)
        {
            RespMsg = "";
            HanaTransaction oTransaction = oConn.BeginTransaction(HanaIsolationLevel.RepeatableRead);
            try
            {
                HanaCommand cmd = new HanaCommand();
                cmd.Connection = oConn;
                cmd.Transaction = oTransaction;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Price_UDT;
                cmd.ExecuteNonQuery();
                cmd.CommandText = Price_SAP;
                cmd.ExecuteNonQuery();

                oTransaction.Commit();
                return true;
            }
            catch (Exception ex)
            {
                RespMsg = ex.Message;
                oTransaction.Rollback();
                return false;
            }
        }

        public int HANA_ExecuteNonQuery(string Query, HanaConnection oConn)
        {
            int Added = 0;
            try
            {
                HanaCommand cmd = new HanaCommand();
                cmd.Connection = oConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = Query;
                Added = cmd.ExecuteNonQuery();
                return Added;
            }
            catch (HanaException)
            {
                return Added;
            }
        }
        #endregion

        #region Open & Save Dialog

        #region Browse Thread
        public void Browse1(string BrowseType, string FormUID = "", string ComboUID = "", string EditUID = "", string FileName = "", bool SharedFolder = false)
        {
            try
            {
                if (BrowseType == "Open")
                {
                    ShowFolderBrowserThread = new Thread(() => ShowFolderBrowser_Text_Open1(FormUID, ComboUID, EditUID));
                }
                else if (BrowseType == "Save")
                {
                    ShowFolderBrowserThread = new Thread(() => ShowFolderBrowser_Text_Save1(FileName, ComboUID, SharedFolder));
                }
                else
                {
                    return;
                }

                if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted)
                {
                    ShowFolderBrowserThread.SetApartmentState(ApartmentState.STA);
                    ShowFolderBrowserThread.Start();
                    ShowFolderBrowserThread.Join();
                }
                else if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Stopped)
                {
                    ShowFolderBrowserThread.Start();
                    ShowFolderBrowserThread.Join();
                }
                while (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Running)
                {
                    Application.DoEvents();
                }

            }
            catch (Exception ex)
            {
                SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Open Dialog
        private void ShowFolderBrowser_Text_Open1(string FormUID, string ComboUID, string EditUID)
        {
            OpenFileDialog oDialogBox = new OpenFileDialog();
            try
            {
                Process[] MyProcs = null;
                SAPbouiCOM.Form oForm;
                oForm = SBO_Appln.Forms.Item(FormUID);

                if (ComboUID == "text")
                {
                    oDialogBox.Filter = "NotePad(*.txt)|*.txt";
                }
                else if (ComboUID == "excel")
                {
                    oDialogBox.Filter = "Excel Worksheets(*.xls;*.xlsx;)|*.xls;*.xlsx;";
                }
                else if (ComboUID == "macro")
                {
                    oDialogBox.Filter = "Excel Worksheets(*.xlsm)|*.xlsm";
                }
                else if (ComboUID == "license")
                {
                    oDialogBox.Filter = "License files (*.lic)|*.lic";
                }
                else if (ComboUID == "csv")
                {
                    oDialogBox.Filter = "CSV Files(*.csv)| *.csv";
                }
                else
                {
                    SAPbouiCOM.ComboBox oCombo;
                    oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item(ComboUID).Specific;
                    if ((oCombo.Selected.Value == "txt"))
                    {
                        oDialogBox.Filter = "NotePad(*.txt)|*.txt";
                    }
                    else
                    {
                        oDialogBox.Filter = "Excel Worksheets(*.xls;*.xlsx;)|*.xls;*.xlsx;";
                    }
                }

                string OrgTitle = SBO_Appln.Desktop.Title;
                SBO_Appln.Desktop.Title = ("SBO under " + oCompany.UserName);
                MyProcs = Process.GetProcessesByName("SAP Business One");
                for (int i = 0; i <= MyProcs.Length - 1; i++)
                {
                    if ((MyProcs[i].MainWindowTitle == SBO_Appln.Desktop.Title))
                    {
                        WindowWrapper MyWindow = new WindowWrapper(MyProcs[i].MainWindowHandle);
                        if ((oDialogBox.ShowDialog(MyWindow) == DialogResult.OK))
                        {
                            ((SAPbouiCOM.EditText)oForm.Items.Item(EditUID).Specific).Value = oDialogBox.FileName;
                        }

                    }

                }

                SBO_Appln.Desktop.Title = OrgTitle;
            }
            catch (Exception ex)
            {
                SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                ShowFolderBrowserThread.Abort();
            }
        }
        #endregion

        #region Save Dialog
        private void ShowFolderBrowser_Text_Save1(string FileName, string Type, bool SharedFolder)
        {
            SaveFileDialog oDialogBox = new SaveFileDialog();
            try
            {
                Process[] MyProcs = null;

                if (Type == "macro")
                {
                    oDialogBox.Filter = "Excel WorkBook(*.xlsm)|*.xlsm";
                }
                else
                {
                    oDialogBox.Filter = "Excel WorkBook(*.xlsx;*.xls;)|*.xlsx;*.xls;";
                }

                string OrgTitle = SBO_Appln.Desktop.Title;
                SBO_Appln.Desktop.Title = ("SBO under " + oCompany.UserName);
                MyProcs = Process.GetProcessesByName("SAP Business One");
                for (int i = 0; i <= MyProcs.Length - 1; i++)
                {
                    if ((MyProcs[i].MainWindowTitle == SBO_Appln.Desktop.Title))
                    {
                        WindowWrapper MyWindow = new WindowWrapper(MyProcs[i].MainWindowHandle);
                        if ((oDialogBox.ShowDialog(MyWindow) == System.Windows.Forms.DialogResult.OK))
                        {
                            string SourcePath = "";
                            if (!SharedFolder)
                            {
                                SourcePath = Assembly.GetExecutingAssembly().GetName().Name + ".ExcelTemplates." + FileName;
                            }
                            else
                            {
                                string Path = Query_Execute("Select \"U_AV_FPATH\" from \"OUSR\" where \"USER_CODE\"='" + oCompany.UserName + "'");

                                Path = Path + "\\" + "ExcelTemplates";
                                if (!System.IO.Directory.Exists(Path))
                                {
                                    System.IO.Directory.CreateDirectory(Path);
                                }
                                SourcePath = Path + "\\" + FileName;
                            }

                            string DestinationPath = oDialogBox.FileName;
                            if (!SharedFolder)
                            {
                                using (FileStream fileStream = new FileStream(DestinationPath, FileMode.Append))
                                {
                                    Assembly.GetExecutingAssembly().GetManifestResourceStream(SourcePath).CopyTo(fileStream);
                                }
                            }
                            else
                            {
                                System.IO.File.Copy(SourcePath, DestinationPath, true);
                            }
                            SBO_Appln.MessageBox(("ExcelFile is Created in this Path - " + (DestinationPath + "")));
                        }

                    }

                }

                SBO_Appln.Desktop.Title = OrgTitle;
            }
            catch (Exception ex)
            {
                SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                ShowFolderBrowserThread.Abort();
            }
        }
        #endregion

        #endregion

    }
}