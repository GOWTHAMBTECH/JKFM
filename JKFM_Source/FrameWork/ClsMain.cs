using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JKFM_Source
{
    public class ClsMain
    {

        #region Declaration
        public ClsSBO objSBOAPI;
        public string updatequery = "";

        public string HANA_ServerName = "";
        public string HANA_UserID = "";
        public string HANA_Pwd = "";
        #endregion

        #region Initialise

        public ClsMain()
        {
            objSBOAPI = new ClsSBO(this);
        }

        public bool Initialise()
        {            
            if (!objSBOAPI.Connect())
            {
                return false;
            }
            objSBOAPI.CreateObjects();
            return true;
        }

        #endregion

        #region Create Tables
        public void CreateTables()
        {
            try
            {
                // ***************************** SAMPLES ************************************ //

                //TABLES

                //objSBOAPI.CreateTable("AVA_A1", "Color", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                //objSBOAPI.CreateTable("AV_Test", "Test", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);
                //objSBOAPI.CreateTable("AV_PROC", "Process", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                //objSBOAPI.CreateTable("AV_PROC1", "Process1", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
                //objSBOAPI.CreateTable("AV_CTEM", "CostingTemplate", SAPbobsCOM.BoUTBTableType.bott_Document);
                //objSBOAPI.CreateTable("AV_CTEM1", "CostingTemplate1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);

                //FIELDS

                //objSBOAPI.AddAlphaField("AV_PROC1", "AV_Procs1", "Process1", 150);
                //objSBOAPI.AddAlphaField("AV_PROC", "AV_Active", "Activate", 10, "Y,N", "Yes,No", "N");
                //objSBOAPI.AddAttachField("OITM", "AV_Attmt", "Attachment", SAPbobsCOM.BoFldSubTypes.st_Link);
                //objSBOAPI.AddAttachField("OITM", "AV_Pictr", "Picture", SAPbobsCOM.BoFldSubTypes.st_Image);
                //objSBOAPI.AddFieldwithLinkTable("OITM", "HType", "HType", 25, "AVA_A4");
                //objSBOAPI.AddDateField("OWOR", "AV_ReTim", "Required Time", SAPbobsCOM.BoFldSubTypes.st_Time);
                //objSBOAPI.AddDateField("OWOR", "AV_ReTim", "Required Time", SAPbobsCOM.BoFldSubTypes.st_None);
                //objSBOAPI.AddFloatField("AVA_A5", "AV_Quant", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddFloatField("AVA_A5", "AV_Quant", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Price);
                //objSBOAPI.AddNumericField("AV_STAT", "AV_DocNm", "DocNum", 11);
                //objSBOAPI.AddAlphaMemoField("AV_STAT", "AV_DocSt", "Doc Status", 254);

                // ***************************** SAMPLES ************************************ //


                // ******************* AVANIKO TABLES & FIELDS CREATION ********************* //

                objSBOAPI.SBO_Appln.StatusBar.SetText("Creating Tables Please Wait...........", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                // ***************** MANDATORY TABLES FOR QUERY MANAGER ********************* //

                // For Query Backup - In Query Manager

                objSBOAPI.AddAlphaField("OUQR", "AV_Sel", "Selected", 1, "Y,N", "Yes,No", "N");

                // Query Backup Table

                objSBOAPI.CreateTable("AV_QRBPT", "Addon Query Backup", SAPbobsCOM.BoUTBTableType.bott_NoObject);

                objSBOAPI.AddAlphaField("AV_QRBPT", "AV_QName", "QueryName", 100);

                objSBOAPI.AddAlphaMemoField("AV_QRBPT", "AV_QString", "QueryString", 256000);

                objSBOAPI.AddAlphaField("AV_QRBPT", "AV_User", "UserName", 25);

                // ***************** MANDATORY TABLES FOR QUERY MANAGER ********************* //


                // ************************** NO OBJECT TABLES ****************************** //

                objSBOAPI.SBO_Appln.StatusBar.SetText("Creating Tables Please Wait...........", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                objSBOAPI.CreateTable("AIS_OUSR", "User Detail", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                objSBOAPI.AddAlphaField("AIS_OUSR", "USERID", "USERID", 50);
                objSBOAPI.AddAlphaField("AIS_OUSR", "UNAME", "USERNAME", 50);

                objSBOAPI.CreateTable("AIS_OUNT", "Unit Master", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                objSBOAPI.AddAlphaField("AIS_OUNT", "SOR", "SaleOrder", 50);
                objSBOAPI.AddAlphaField("AIS_OUNT", "ARTax", "ARTax", 50);
                objSBOAPI.AddAlphaField("AIS_OUNT", "ARNTax", "ARNonTax", 50);
                objSBOAPI.AddAlphaField("AIS_OUNT", "ARExmpt", "ARExcempt", 50);

                objSBOAPI.CreateTable("AIS_OPRL", "Print Layout", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                objSBOAPI.AddAlphaField("AIS_OPRL", "Object", "Object", 50);
                objSBOAPI.AddAlphaField("AIS_OPRL", "ReportName", "ReportName", 50);
                objSBOAPI.AddAlphaField("AIS_OPRL", "ReportPath", "ReportPath", 100);
                objSBOAPI.AddAlphaField("AIS_OPRL", "PDFSavePath", "PDFSavePath", 100);
                objSBOAPI.AddAlphaField("AIS_OPRL", "IsDefault", "IsDefault", 1);

                objSBOAPI.CreateTable("AIS_OWBT", "WeighBridge Type", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                objSBOAPI.AddAlphaField("AIS_OWBT", "Unit", "Unit", 50);

                objSBOAPI.CreateTable("AIS_TCS", "TCS Service Setup", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                objSBOAPI.AddAlphaField("AIS_TCS", "TCSPercent", "TCS Percentage", 50);

                objSBOAPI.CreateTable("AIS_TRAN", "Driver Detailes", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                objSBOAPI.AddAlphaField("AIS_TRAN", "TRNO", "TRUCK NO", 50);
                objSBOAPI.AddAlphaField("AIS_TRAN", "DRVNME", "DRIVER NAME", 50);
                objSBOAPI.AddAlphaField("AIS_TRAN", "FRETY", "FREIGHT TYPE", 50, "H,O,C", "HIRED,OWN,CUTOMER","");
                objSBOAPI.AddAlphaField("AIS_TRAN", "DRVMOB", "DRIVER MOBILE NO", 10, SAPbobsCOM.BoFldSubTypes.st_Phone);

                objSBOAPI.CreateTable("BATCHTEMPTAB", "BATCH TEMP TABLE", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                objSBOAPI.AddAlphaField("BATCHTEMPTAB", "ItemCode", "ItemCode", 50);
                objSBOAPI.AddAlphaField("BATCHTEMPTAB", "WhsCode", "WhsCode", 50);
                objSBOAPI.AddAlphaField("BATCHTEMPTAB", "BatchNum", "BatchNum", 50);
                objSBOAPI.AddFloatField("BATCHTEMPTAB", "BatchQty", "BatchQty", SAPbobsCOM.BoFldSubTypes.st_Quantity);

                objSBOAPI.CreateTable("AIS_OVEH", "Vehicle Type", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);
                objSBOAPI.AddAlphaField("AIS_OVEH", "TVehicle", "Type of Vehicle", 50);
                objSBOAPI.AddFloatField("AIS_OVEH", "TolerPer", "Tolerance", SAPbobsCOM.BoFldSubTypes.st_Percentage);
                objSBOAPI.AddFloatField("AIS_OVEH", "TolerKg", "Tolerance Kg", SAPbobsCOM.BoFldSubTypes.st_Quantity);

                objSBOAPI.CreateTable("AIS_OLOC", "Location", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);
                objSBOAPI.AddAlphaField("AIS_OLOC", "Location", "Location", 100);
                objSBOAPI.AddFloatField("AIS_OLOC", "Km", "Km", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_OLOC", "Rate", "Rate", SAPbobsCOM.BoFldSubTypes.st_Rate);

                //Lab Deduction Details
                objSBOAPI.CreateTable("AVA_DED1", "Lab Deduction Details", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);
                objSBOAPI.AddAlphaField("AVA_DED1", "AVA_AnlysNme", "Analysis Name", 100);
                objSBOAPI.AddAlphaField("AVA_DED1", "AVA_DocEntry", "DocEntry", 10);
                objSBOAPI.AddFloatField("AVA_DED1", "AVA_ActVal", "Actual Value", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddAlphaField("AVA_DED1", "AVA_Unit", "Unit", 10);
                objSBOAPI.AddFloatField("AVA_DED1", "AVA_DedAmt", "Deduction Amount", SAPbobsCOM.BoFldSubTypes.st_Price);
                objSBOAPI.AddFloatField("AVA_DED1", "AVA_Qty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AVA_DED1", "AVA_StdPar", "Standard Parameter", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddAlphaField("AVA_DED1", "AVA_Remarks", "Remarks", 254);
                objSBOAPI.AddNumericField("AVA_DED1", "AVA_Sample", "Sample", 10);
                objSBOAPI.AddNumericField("AVA_DED1", "AVA_QCLevel", "QC Level", 10);

                //Other Deduction Details
                objSBOAPI.CreateTable("AVA_DED2", "Other Deduction Details", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);
                objSBOAPI.AddAlphaField("AVA_DED2", "AVA_DedName", "Deduction Name", 100);
                objSBOAPI.AddAlphaField("AVA_DED2", "AVA_DocEntry", "DocEntry", 10);
                objSBOAPI.AddFloatField("AVA_DED2", "AVA_Debit", "Debit Amount", SAPbobsCOM.BoFldSubTypes.st_Price);
                objSBOAPI.AddFloatField("AVA_DED2", "AVA_Credit", "Credit Amount", SAPbobsCOM.BoFldSubTypes.st_Price);
                objSBOAPI.AddFloatField("AVA_DED2", "AVA_Total", "Total Value", SAPbobsCOM.BoFldSubTypes.st_Price);
                objSBOAPI.AddAlphaField("AVA_DED2", "AVA_Ledger", "Ledger Name", 254);

                //Other Deduction Details - Temp Table
                objSBOAPI.CreateTable("AVA_TEMP2", "Temp - Other Deduction", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);
                objSBOAPI.AddAlphaField("AVA_TEMP2", "AVA_DedName", "Deduction Name", 100);
                objSBOAPI.AddAlphaField("AVA_TEMP2", "AVA_Select", "Selection", 10);
                objSBOAPI.AddAlphaField("AVA_TEMP2", "AVA_DocEntry", "DocEntry", 10);
                objSBOAPI.AddFloatField("AVA_TEMP2", "AVA_Debit", "Debit Amount", SAPbobsCOM.BoFldSubTypes.st_Price);
                objSBOAPI.AddFloatField("AVA_TEMP2", "AVA_Credit", "Credit Amount", SAPbobsCOM.BoFldSubTypes.st_Price);
                objSBOAPI.AddFloatField("AVA_TEMP2", "AVA_Total", "Total Value", SAPbobsCOM.BoFldSubTypes.st_Price);
                objSBOAPI.AddAlphaField("AVA_TEMP2", "AVA_Ledger", "Ledger Name", 254);

                //objSBOAPI.CreateTable("AVA_OSTM", "Transport Master Data", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);
                //objSBOAPI.AddAlphaField("AVA_OSTM", "AVA_TRANSPORTER", "Transporter", 254);
                //objSBOAPI.AddAlphaField("AVA_OSTM", "AVA_VEHICLENO", "Vehicle No", 10);
                //objSBOAPI.AddAlphaField("AVA_OSTM", "AVA_DRIVERMOBILENO", "Driver Mobile No", 10);
                //objSBOAPI.AddAlphaField("AVA_OSTM", "AVA_REGION", "Region", 30);
                //objSBOAPI.AddAlphaField("AVA_OSTM", "AVA_DRIVERLNO", "Driver License No", 20);
                //objSBOAPI.AddAlphaField("AVA_OSTM", "AVA_REMARKS", "Remarks", 254);
                //objSBOAPI.AddAlphaField("AVA_OSTM", "AVA_OWNTRUCK", "Own Truck", 10, "Y,N", "Yes,No", "N");

                objSBOAPI.CreateTable("AVA_OSIP", "Sales Insurance Policy", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);
                objSBOAPI.AddAlphaField("AVA_OSIP", "AVA_INSPNO", "Insurance Policy No", 50);
                objSBOAPI.AddFloatField("AVA_OSIP", "AVA_INSAMT", "Insurance Amount", SAPbobsCOM.BoFldSubTypes.st_Price);
                objSBOAPI.AddFloatField("AVA_OSIP", "AVA_INSCHRGS", "Insurance Charges", SAPbobsCOM.BoFldSubTypes.st_Price);
                objSBOAPI.AddFloatField("AVA_OSIP", "AVA_INSBAL", "Insurance Balance", SAPbobsCOM.BoFldSubTypes.st_Price);
                objSBOAPI.AddDateField("AVA_OSIP", "AVA_INSVAL", "Insurance Validity", SAPbobsCOM.BoFldSubTypes.st_None);

                //Customer Statement
                objSBOAPI.CreateTable("AV_LPST", "Layout Path Setup", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                objSBOAPI.AddAlphaField("AV_LPST", "AV_Type", "Type", 3, "AR,PO,CN,GP,OP,CS", "A/R Invoice,PO Mailer,Credit Memo,GRPO,Outgoing Payments,Customer Statement", "");
                objSBOAPI.AddAlphaMemoField("AV_LPST", "AV_LayPt", "Layout Path", 254);
                objSBOAPI.AddAlphaField("AV_LPST", "AV_Deflt", "Default", 10, "Y,N", "Yes,No", "");

                objSBOAPI.CreateTable("AV_GRPTMP", "Grouping Temp", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);
                objSBOAPI.AddDateField("AV_GRPTMP", "AV_FrmDt", "FromDate", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddDateField("AV_GRPTMP", "AV_ToDate", "ToDate", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddAlphaField("AV_GRPTMP", "AV_RpTyp", "ReportType", 10, "AR,PO,CN,GP,OP,CS", "A/R Invoice,PO Mailer,Credit Memo,GRPO,Outgoing Payments,Customer Scheme", "");
                objSBOAPI.AddAlphaField("AV_GRPTMP", "AV_LtTyp", "LayoutType", 50);
                objSBOAPI.AddAlphaField("AV_GRPTMP", "AV_Type", "Type", 10, "AL,PF,PE", "ALL,PDF,PDF_EMAIL", "PF");
                objSBOAPI.AddDateField("AV_GRPTMP", "AV_Date", "Date", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddNumericField("AV_GRPTMP", "AV_DocEnt", "DocEntry", 11);
                objSBOAPI.AddNumericField("AV_GRPTMP", "AV_DocNm", "DocNum", 11);
                objSBOAPI.AddAlphaField("AV_GRPTMP", "AV_CrdCd", "CardCode", 50);
                objSBOAPI.AddAlphaField("AV_GRPTMP", "AV_CrdNm", "CardName", 100);
                objSBOAPI.AddAlphaField("AV_GRPTMP", "AV_CntPn", "ContactPerson", 100);
                objSBOAPI.AddAlphaField("AV_GRPTMP", "AV_GrpNm", "GroupName", 50);
                objSBOAPI.AddFloatField("AV_GRPTMP", "AV_DocTot", "DocTotal", SAPbobsCOM.BoFldSubTypes.st_Sum);
                objSBOAPI.AddDateField("AV_GRPTMP", "AV_PosDt", "PostingDate", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddDateField("AV_GRPTMP", "AV_DueDt", "DueDate", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddAlphaField("AV_GRPTMP", "AV_EmlId", "EmailId", 100);
                objSBOAPI.AddAlphaField("AV_GRPTMP", "AV_CCmlId", "CC_MailId", 254);
                objSBOAPI.AddAlphaField("AV_GRPTMP", "AV_Stats", "Status", 254);
                objSBOAPI.AddAlphaField("AV_GRPTMP", "AV_PDFPath", "PDFPath", 254);
                objSBOAPI.AddAlphaField("AV_GRPTMP", "AV_Status", "Status", 10, "S,F", "Success,Failed", "");
                objSBOAPI.AddNumericField("AV_GRPTMP", "AV_GrdRw", "GridRow", 11);

                //Thiru - 20210805
                //Log Table

                //objSBOAPI.CreateTable("AVA_MULOG", "Min/Max Upload Log", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);
                //objSBOAPI.AddAlphaField("AVA_MULOG", "AVA_CreateDate", "Create Date", 20);
                //objSBOAPI.AddAlphaField("AVA_MULOG", "AVA_FileName", "FileName", 254);
                //objSBOAPI.AddNumericField("AVA_MULOG", "AVA_DocEntry", "DocEntry", 11);
                //objSBOAPI.AddAlphaField("AVA_MULOG", "AVA_ItmCod", "ItemCode", 50);
                //objSBOAPI.AddAlphaField("AVA_MULOG", "AVA_WhsCod", "Warehouse Code", 8);
                //objSBOAPI.AddFloatField("AVA_MULOG", "AVA_MinQty", "Minimum Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddFloatField("AVA_MULOG", "AVA_MaxQty", "Maximum Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddAlphaField("AVA_MULOG", "AVA_VenCod", "Vendor Code", 15);
                //objSBOAPI.AddAlphaField("AVA_MULOG", "AVA_ClubPO", "Club PO", 1);
                //objSBOAPI.AddAlphaField("AVA_MULOG", "AVA_Status", "Status", 1);
                //objSBOAPI.AddAlphaMemoField("AVA_MULOG", "AVA_RespMsg", "Response Message", 25600);

                objSBOAPI.CreateTable("AVA_ITLOG", "ItemMaster Upload Log", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);
                objSBOAPI.AddAlphaField("AVA_ITLOG", "AVA_CreateDate", "Create Date", 20);
                objSBOAPI.AddAlphaField("AVA_ITLOG", "AVA_FileName", "FileName", 254);
                objSBOAPI.AddNumericField("AVA_ITLOG", "AVA_DocEntry", "DocEntry", 11);
                objSBOAPI.AddAlphaField("AVA_ITLOG", "AVA_ItmCod", "ItemCode", 50);
                objSBOAPI.AddAlphaField("AVA_ITLOG", "AVA_ItmNam", "Item Description", 200);
                objSBOAPI.AddNumericField("AVA_ITLOG", "AVA_ItmGrp", "Item Group Code", 11);
                objSBOAPI.AddAlphaField("AVA_ITLOG", "AVA_PrdDesc", "Product Description", 100);
                objSBOAPI.AddNumericField("AVA_ITLOG", "AVA_FirmCod", "Manufacturer", 11);
                objSBOAPI.AddAlphaField("AVA_ITLOG", "AVA_CharDesc", "Character Description", 100);
                objSBOAPI.AddAlphaField("AVA_ITLOG", "AVA_CapDesc", "Capacity Description", 100);
                objSBOAPI.AddAlphaField("AVA_ITLOG", "AVA_Status", "Status", 1);
                objSBOAPI.AddAlphaMemoField("AVA_ITLOG", "AVA_RespMsg", "Response Message", 25600);

                //objSBOAPI.CreateTable("AVA_PVLOG", "Preferred Vendor Upload Log", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);
                //objSBOAPI.AddAlphaField("AVA_PVLOG", "AVA_CreateDate", "Create Date", 20);
                //objSBOAPI.AddAlphaField("AVA_PVLOG", "AVA_FileName", "FileName", 254);
                //objSBOAPI.AddNumericField("AVA_PVLOG", "AVA_DocEntry", "DocEntry", 11);
                //objSBOAPI.AddAlphaField("AVA_PVLOG", "AVA_ItmCod", "ItemCode", 50);
                //objSBOAPI.AddAlphaField("AVA_PVLOG", "AVA_VenCod", "Vendor Code", 15);
                //objSBOAPI.AddAlphaField("AVA_PVLOG", "AVA_Division", "Division", 100);
                //objSBOAPI.AddAlphaField("AVA_PVLOG", "AVA_Status", "Status", 1);
                //objSBOAPI.AddAlphaMemoField("AVA_PVLOG", "AVA_RespMsg", "Response Message", 25600);

                //objSBOAPI.CreateTable("AVA_PDLOG", "Preferred Vendor Delete Log", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);
                //objSBOAPI.AddAlphaField("AVA_PDLOG", "AVA_CreateDate", "Create Date", 20);
                //objSBOAPI.AddAlphaField("AVA_PDLOG", "AVA_FileName", "FileName", 254);
                //objSBOAPI.AddNumericField("AVA_PDLOG", "AVA_DocEntry", "DocEntry", 11);
                //objSBOAPI.AddAlphaField("AVA_PDLOG", "AVA_ItmCod", "ItemCode", 50);
                //objSBOAPI.AddAlphaField("AVA_PDLOG", "AVA_VenCod", "Vendor Code", 15);
                //objSBOAPI.AddAlphaField("AVA_PDLOG", "AVA_Status", "Status", 1);
                //objSBOAPI.AddAlphaMemoField("AVA_PDLOG", "AVA_RespMsg", "Response Message", 25600);

                //objSBOAPI.CreateTable("AVA_PLLOG", "Price List Log", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);
                //objSBOAPI.AddAlphaField("AVA_PLLOG", "AVA_CreateDate", "Create Date", 20);
                //objSBOAPI.AddAlphaField("AVA_PLLOG", "AVA_FileName", "FileName", 254);
                //objSBOAPI.AddNumericField("AVA_PLLOG", "AVA_DocEntry", "DocEntry", 11);
                //objSBOAPI.AddAlphaField("AVA_PLLOG", "AVA_Type", "Type", 1, "P,I,S", "Price List Upload,Item Master Upload,Standard Item Master", "");
                //objSBOAPI.AddAlphaField("AVA_PLLOG", "AVA_ItmCod", "ItemCode", 50);
                //objSBOAPI.AddAlphaField("AVA_PLLOG", "AVA_ListName", "Price List Name", 32);
                //objSBOAPI.AddDateField("AVA_PLLOG", "AVA_FromDate", "FromDate", SAPbobsCOM.BoFldSubTypes.st_None);
                //objSBOAPI.AddDateField("AVA_PLLOG", "AVA_ToDate", "ToDate", SAPbobsCOM.BoFldSubTypes.st_None);
                //objSBOAPI.AddFloatField("AVA_PLLOG", "AVA_Price", "Price", SAPbobsCOM.BoFldSubTypes.st_Price);
                //objSBOAPI.AddAlphaField("AVA_PLLOG", "AVA_Status", "Status", 1);
                //objSBOAPI.AddAlphaMemoField("AVA_PLLOG", "AVA_RespMsg", "Response Message", 25600);

                //Thiru - 20210805

                objSBOAPI.CreateTable("AVA_ENEST", "Edit/Non-Edit Settings Table", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);

                // ************************** MASTER DATA TABLES **************************** // 

                objSBOAPI.SBO_Appln.StatusBar.SetText("Creating Tables Please Wait...........", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                objSBOAPI.CreateTable("AIS_OBRN", "Branch Master Header", SAPbobsCOM.BoUTBTableType.bott_MasterData);

                objSBOAPI.CreateTable("AIS_BRN1", "Branch Master Detail", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
                objSBOAPI.AddAlphaField("AIS_BRN1", "BranchCode", "Branch Code", 50);
                objSBOAPI.AddAlphaField("AIS_BRN1", "BranchName", "Branch Name", 100);
                objSBOAPI.AddAlphaField("AIS_BRN1", "WhsCode", "Whs Name", 8);
                objSBOAPI.AddAlphaField("AIS_BRN1", "Active", "Active", 1);

                objSBOAPI.CreateTable("AIS_OCIA", "Custome Wise ItemAllocation", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                objSBOAPI.AddAlphaField("AIS_OCIA", "CardCode", "CardCode", 50);
                objSBOAPI.AddAlphaField("AIS_OCIA", "CardName", "CardName", 100);

                //Thiru Addon Changes - Start - 20.04.2021
                objSBOAPI.AddAlphaField("AIS_OCIA", "AVA_BrName", "Broker Name", 100);
                //Thiru Addon Changes - End - 20.04.2021


                objSBOAPI.CreateTable("AIS_CIA1", "Custome Wise ItemAllocation", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
                objSBOAPI.AddAlphaField("AIS_CIA1", "ItemCode", "ItemCode", 50);
                objSBOAPI.AddAlphaField("AIS_CIA1", "ItemName", "ItemName", 100);

                //Thiru Addon Changes - Start - 20.04.2021
                objSBOAPI.AddFloatField("AIS_CIA1", "AVA_Cdpt", "Cash discount per ton", SAPbobsCOM.BoFldSubTypes.st_Sum);
                objSBOAPI.AddFloatField("AIS_CIA1", "AVA_Icpt", "Incentive per ton", SAPbobsCOM.BoFldSubTypes.st_Sum);
                objSBOAPI.AddFloatField("AIS_CIA1", "AVA_Brpt", "Brokerage per ton", SAPbobsCOM.BoFldSubTypes.st_Sum);
                //Thiru Addon Changes - End - 20.04.2021
                objSBOAPI.AddFloatField("AIS_CIA1", "AVA_SplPrice", "SpecialPrice", SAPbobsCOM.BoFldSubTypes.st_Price);

                objSBOAPI.CreateTable("AIS_OREM", "Yield Master Header", SAPbobsCOM.BoUTBTableType.bott_MasterData);

                objSBOAPI.CreateTable("AIS_REM1", "Yield Master Detail", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
                objSBOAPI.AddAlphaField("AIS_REM1", "Category", "Category", 50);
                objSBOAPI.AddAlphaField("AIS_REM1", "Unit", "Unit", 50);
                objSBOAPI.AddFloatField("AIS_REM1", "Perntage", "Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage);
                objSBOAPI.AddFloatField("AIS_REM1", "Tolarence", "Tolarence", SAPbobsCOM.BoFldSubTypes.st_Quantity);

                objSBOAPI.CreateTable("AIS_OSFT", "Shift Master Header", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                objSBOAPI.AddDateField("AIS_OSFT", "SStartTime", "Shift Start Time", SAPbobsCOM.BoFldSubTypes.st_Time);
                objSBOAPI.AddDateField("AIS_OSFT", "SEndTime", "Shift End Time", SAPbobsCOM.BoFldSubTypes.st_Time);
                objSBOAPI.AddDateField("AIS_OSFT", "BStartTime", "Break Start Time", SAPbobsCOM.BoFldSubTypes.st_Time);
                objSBOAPI.AddDateField("AIS_OSFT", "BEndTime", "Break End Time", SAPbobsCOM.BoFldSubTypes.st_Time);
                objSBOAPI.AddAlphaField("AIS_OSFT", "ShiftHours", "Shift Hours", 100);
                objSBOAPI.AddAlphaField("AIS_OSFT", "WorkHours", "Working Hours", 100);
                objSBOAPI.AddAlphaField("AIS_OSFT", "FlexHours", "Flex Hours", 1);
                objSBOAPI.AddAlphaField("AIS_OSFT", "OverLap", "OverLap", 1);
                objSBOAPI.AddAlphaField("AIS_OSFT", "OT", "Over Time", 1);
                objSBOAPI.AddAlphaField("AIS_OSFT", "Active", "Active", 1);
                objSBOAPI.AddAlphaField("AIS_OSFT", "ILunch", "Include Lunch", 1);

                objSBOAPI.CreateTable("AIS_ORMC", "Packing Capacity  Master", SAPbobsCOM.BoUTBTableType.bott_MasterData);

                objSBOAPI.CreateTable("AIS_RMC1", "Packing Capacity Detail1", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
                objSBOAPI.AddAlphaField("AIS_RMC1", "ItemCode", "Item Code", 50);
                objSBOAPI.AddAlphaField("AIS_RMC1", "ItemName", "Item Name", 100);
                objSBOAPI.AddFloatField("AIS_RMC1", "Unt1QtyBag", "Unit-1 Quantity(Bag)", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_RMC1", "Unt1QtyTon", "Unit-1 Quantity(Ton)", SAPbobsCOM.BoFldSubTypes.st_Quantity);

                objSBOAPI.CreateTable("AIS_RMC2", "Packing Capacity Detail2", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
                objSBOAPI.AddAlphaField("AIS_RMC2", "ItemCode", "Item Code", 50);
                objSBOAPI.AddAlphaField("AIS_RMC2", "ItemName", "Item Name", 100);
                objSBOAPI.AddFloatField("AIS_RMC2", "Unt2QtyBag", "Unit-2 Quantity(Bag)", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_RMC2", "Unt2QtyTon", "Unit-2 Quantity(Ton)", SAPbobsCOM.BoFldSubTypes.st_Quantity);

                //Insurance Details Header

                objSBOAPI.CreateTable("AVA_INSURANCEH", "Insurance Details Header", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                objSBOAPI.AddFloatField("AVA_INSURANCEH", "AVA_InsAmt", "Insurance Amount", SAPbobsCOM.BoFldSubTypes.st_Price);
                objSBOAPI.AddFloatField("AVA_INSURANCEH", "AVA_BalAmt", "Balance Amount", SAPbobsCOM.BoFldSubTypes.st_Price);
                objSBOAPI.AddDateField("AVA_INSURANCEH", "AVA_DocDate", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddDateField("AVA_INSURANCEH", "AVA_FromDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddDateField("AVA_INSURANCEH", "AVA_ToDate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None);

                //Insurance Details Lines

                objSBOAPI.CreateTable("AVA_INSURANCEL", "Insurance Details Child", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
                objSBOAPI.AddAlphaField("AVA_INSURANCEL", "AVA_CardCode", "CardCode", 100);
                objSBOAPI.AddAlphaField("AVA_INSURANCEL", "AVA_CardName", "CardName", 150);
                objSBOAPI.AddAlphaField("AVA_INSURANCEL", "AVA_Type", "Type", 10);
                objSBOAPI.AddAlphaField("AVA_INSURANCEL", "AVA_DocEntry", "DocEntry", 10);
                objSBOAPI.AddAlphaField("AVA_INSURANCEL", "AVA_DocNum", "DocNum", 10);
                objSBOAPI.AddAlphaField("AVA_INSURANCEL", "AVA_DocStatus", "DocStatus", 10);
                objSBOAPI.AddDateField("AVA_INSURANCEL", "AVA_Date", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddFloatField("AVA_INSURANCEL", "AVA_Debit", "Debit", SAPbobsCOM.BoFldSubTypes.st_Price);
                objSBOAPI.AddFloatField("AVA_INSURANCEL", "AVA_Credit", "Credit", SAPbobsCOM.BoFldSubTypes.st_Price);

                //Customer Statement//
                objSBOAPI.CreateTable("AV_CSLGT", "Customer Statement Log", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                objSBOAPI.AddDateField("AV_CSLGT", "AV_FrmDt", "FromDate", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddDateField("AV_CSLGT", "AV_ToDate", "ToDate", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddAlphaField("AV_CSLGT", "AV_RpTyp", "ReportType", 10, "AR,PO,CN,GP,OP,CS", "A/R Invoice,PO Mailer,Credit Memo,GRPO,Outgoing Payment,Customer Scheme", "");
                objSBOAPI.AddAlphaField("AV_CSLGT", "AV_LtTyp", "LayoutType", 50);
                objSBOAPI.AddAlphaField("AV_CSLGT", "AV_Type", "Type", 10, "AL,PF,PE", "ALL,PDF,PDF_EMAIL", "PF");
                objSBOAPI.AddDateField("AV_CSLGT", "AV_Date", "Date", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddNumericField("AV_CSLGT", "AV_DocEnt", "DocEntry", 11);
                objSBOAPI.AddAlphaField("AV_CSLGT", "AV_DocNm", "DocNum", 25);
                objSBOAPI.AddAlphaField("AV_CSLGT", "AV_CrdCd", "CardCode", 50);
                objSBOAPI.AddAlphaField("AV_CSLGT", "AV_CrdNm", "CardName", 100);
                objSBOAPI.AddAlphaField("AV_CSLGT", "AV_CntPn", "ContactPerson", 100);
                objSBOAPI.AddAlphaField("AV_CSLGT", "AV_GrpNm", "GroupName", 50);
                objSBOAPI.AddFloatField("AV_CSLGT", "AV_DocTot", "DocTotal", SAPbobsCOM.BoFldSubTypes.st_Sum);
                objSBOAPI.AddDateField("AV_CSLGT", "AV_PosDt", "PostingDate", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddDateField("AV_CSLGT", "AV_DueDt", "DueDate", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddAlphaField("AV_CSLGT", "AV_EmlId", "EmailId", 100);
                objSBOAPI.AddAlphaField("AV_CSLGT", "AV_CCmlId", "CC_MailId", 254);
                objSBOAPI.AddAlphaField("AV_CSLGT", "AV_Stats", "Status", 254);
                objSBOAPI.AddAlphaField("AV_CSLGT", "AV_PDFPath", "PDFPath", 254);
                objSBOAPI.AddAlphaField("AV_CSLGT", "AV_Status", "Status", 10, "S,F", "Success,Failed", "");
                objSBOAPI.AddAlphaField("AV_CSLGT", "AV_MailType", "MailType", 25, "Auto,Man", "Automated,Manual", "Man");
                objSBOAPI.AddFloatField("AV_CSLGT", "AV_OpnBl", "OpeningBalance", SAPbobsCOM.BoFldSubTypes.st_Sum);
                objSBOAPI.AddFloatField("AV_CSLGT", "AV_ClsBl", "ClosingBalance", SAPbobsCOM.BoFldSubTypes.st_Sum);
                objSBOAPI.AddFloatField("AV_CSLGT", "AV_DebBl", "DebitBalance", SAPbobsCOM.BoFldSubTypes.st_Sum);
                objSBOAPI.AddFloatField("AV_CSLGT", "AV_CrdBl", "CreditBalance", SAPbobsCOM.BoFldSubTypes.st_Sum);

                // ************************* DOCUMENT DATA TABLES *************************** //

                objSBOAPI.SBO_Appln.StatusBar.SetText("Creating Tables Please Wait...........", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                objSBOAPI.CreateTable("AIS_OPRE", "Pre Sales Order", SAPbobsCOM.BoUTBTableType.bott_Document);
                objSBOAPI.AddAlphaField("AIS_OPRE", "CardCode", "CardCode", 50);
                objSBOAPI.AddAlphaField("AIS_OPRE", "CardName", "CardName", 100);
                objSBOAPI.AddDateField("AIS_OPRE", "DocDate", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddFloatField("AIS_OPRE", "Discount", "Discount", SAPbobsCOM.BoFldSubTypes.st_Sum);
                objSBOAPI.AddFloatField("AIS_OPRE", "DocTotal", "DocTotal", SAPbobsCOM.BoFldSubTypes.st_Sum);
                objSBOAPI.AddFloatField("AIS_OPRE", "BDiscount", "BDiscount", SAPbobsCOM.BoFldSubTypes.st_Sum);
                objSBOAPI.AddFloatField("AIS_OPRE", "TotalBag", "Total Bag", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_OPRE", "TotalTon", "Total Ton", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddDateField("AIS_OPRE", "FromDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddDateField("AIS_OPRE", "ToDate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddAlphaField("AIS_OPRE", "BaseEntry", "Base Entry", 11);
                objSBOAPI.AddAlphaField("AIS_OPRE", "BaseNum", "BaseNum", 11);
                objSBOAPI.AddAlphaField("AIS_OPRE", "BaseSeries", "Base Series", 11);
                objSBOAPI.AddAlphaField("AIS_OPRE", "BaseObject", "Base Object", 11);
                objSBOAPI.AddAlphaField("AIS_OPRE", "Status", "Status", 50);
                objSBOAPI.AddAlphaField("AIS_OPRE", "Cancelled", "Cancelled", 1, "Y,N", "Yes,No", "N");
                objSBOAPI.AddAlphaField("AIS_OPRE", "TrkReq", "TrkReq", 1, "Y,N", "Yes,No", "N");
                objSBOAPI.AddAlphaField("AIS_OPRE", "Type", "Type", 1, "M,S", "Monthly Order,Special Order", "M");
                objSBOAPI.AddFloatField("AIS_OPRE", "NoOfTrk", "NoOfTrk", SAPbobsCOM.BoFldSubTypes.st_Quantity);

                objSBOAPI.CreateTable("AIS_PRE1", "Pre Sales Order Detail", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                objSBOAPI.AddAlphaField("AIS_PRE1", "ItemCode", "Item Code", 50);
                objSBOAPI.AddAlphaField("AIS_PRE1", "ItemName", "Item Name", 100);
                objSBOAPI.AddFloatField("AIS_PRE1", "ReqQty", "ReqQty", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_PRE1", "UnitPrice", "UnitPrice", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_PRE1", "ReqDate", "ReqDate", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddAlphaField("AIS_PRE1", "TaxCode", "TaxCode", 50);
                objSBOAPI.AddFloatField("AIS_PRE1", "DisCount", "DisCount", SAPbobsCOM.BoFldSubTypes.st_Rate);
                objSBOAPI.AddAlphaField("AIS_PRE1", "PlanUnit", "Plan Unit", 50);
                objSBOAPI.AddAlphaField("AIS_PRE1", "UomCode", "UomCode", 50);
                objSBOAPI.AddFloatField("AIS_PRE1", "LineTotal", "LineTotal", SAPbobsCOM.BoFldSubTypes.st_Sum);
                objSBOAPI.AddAlphaField("AIS_PRE1", "Status", "Status", 1, "C,O", "Closed,Open", "O");
                objSBOAPI.AddAlphaField("AIS_PRE1", "PStatus", "PStatus", 1, "C,O", "Closed,Open", "O");
                objSBOAPI.AddAlphaField("AIS_PRE1", "MobStatus", "MobStatus", 1, "N,E,D", "New,Edit,Delete", "N");
                objSBOAPI.AddAlphaField("AIS_PRE1", "Approval", "Approval", 1);
                objSBOAPI.AddAlphaField("AIS_PRE1", "DefUnit", "Default Unit", 50);

                objSBOAPI.CreateTable("AIS_PRE2", "Pre Sales Order Detail", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                objSBOAPI.AddAlphaField("AIS_PRE2", "UniqID", "UniqID", 11);
                objSBOAPI.AddDateField("AIS_PRE2", "PlanDate", "PlanDate", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddFloatField("AIS_PRE2", "Qty", "Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddAlphaField("AIS_PRE2", "DefUnit", "Default Unit", 50);
                objSBOAPI.AddAlphaField("AIS_PRE2", "WhsCode", "WhsCode", 50);
                objSBOAPI.AddAlphaField("AIS_PRE2", "WhsName", "Warehouse Name", 100);
                objSBOAPI.AddAlphaField("AIS_PRE2", "Approve", "Approval", 1);
                objSBOAPI.AddAlphaField("AIS_PRE2", "ReApprove", "Re-Approval", 1);
                objSBOAPI.AddAlphaField("AIS_PRE2", "MobStatus", "MobStatus", 5, "-1,C", ",Cancel", "-1");
                objSBOAPI.AddAlphaField("AIS_PRE2", "CompStatus", "CompStatus", 5, "-1,C", ",Cancel", "-1");
                objSBOAPI.AddAlphaField("AIS_PRE2", "Approval", "Approval", 1, "Y,N", "Yes,No", "N");

                objSBOAPI.CreateTable("AIS_LOAD", "Delivery Header", SAPbobsCOM.BoUTBTableType.bott_Document);
                objSBOAPI.AddDateField("AIS_LOAD", "DocDate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddAlphaField("AIS_LOAD", "Remarks", "Remarks", 200);
                objSBOAPI.AddFloatField("AIS_LOAD", "TotAmt", "Total Amount", SAPbobsCOM.BoFldSubTypes.st_Sum);
                objSBOAPI.AddDateField("AIS_LOAD", "BillRecDate", "Bill Rec Date", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddFloatField("AIS_LOAD", "TotTon", "Total Tonnage ", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddAlphaField("AIS_LOAD", "BPGroup", "BPGroup", 100);
                objSBOAPI.addField("AIS_LOAD", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "O,C", "Open,Close", "O");

                objSBOAPI.CreateTable("AIS_LOAD1", "Delivery Deatils", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                objSBOAPI.AddAlphaField("AIS_LOAD1", "Select", "Select", 1);
                objSBOAPI.AddAlphaField("AIS_LOAD1", "PreSalesNo", "PreSales No.", 20);
                objSBOAPI.AddAlphaField("AIS_LOAD1", "PreSaleEnt", "PreSales Entry", 20);
                objSBOAPI.AddAlphaField("AIS_LOAD1", "CusCode", "Customer Code", 50);
                objSBOAPI.AddAlphaField("AIS_LOAD1", "CusName", "Customer Name", 200);
                objSBOAPI.AddAlphaField("AIS_LOAD1", "ItemCode", "Item Code", 50);
                objSBOAPI.AddAlphaField("AIS_LOAD1", "ItemName", "Item Name", 100);
                objSBOAPI.AddAlphaField("AIS_LOAD1", "ItmCatgy", "Item Category", 100);
                objSBOAPI.AddAlphaField("AIS_LOAD1", "WhsCode", "WhsCode", 50);
                objSBOAPI.AddAlphaField("AIS_LOAD1", "WhsName", "WhsName", 50);
                objSBOAPI.AddFloatField("AIS_LOAD1", "BagQty", "Bag Quantity ", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_LOAD1", "TonQty", "Ton Quantity ", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_LOAD1", "Price", "Price", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddAlphaField("AIS_LOAD1", "TargetEntry", "Target Entry", 11);
                objSBOAPI.AddAlphaField("AIS_LOAD1", "TargetNum", "Target Num", 11);
                objSBOAPI.AddAlphaField("AIS_LOAD1", "TargetObject", "Target Object", 11);
                objSBOAPI.AddAlphaField("AIS_LOAD1", "TaxCode", "TaxCode", 50);
                objSBOAPI.AddAlphaField("AIS_LOAD1", "DefUnit", "Default Unit", 50);

                objSBOAPI.CreateTable("AIS_LOAD2", "Delivery Deatils2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                objSBOAPI.AddFloatField("AIS_LOAD2", "M1", "M1", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_LOAD2", "M2", "M2", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_LOAD2", "M3", "M3", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_LOAD2", "SUJI", "SUJI", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_LOAD2", "ATTA", "ATTA", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_LOAD2", "BRAN", "BRAN", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddAlphaField("AIS_LOAD2", "CardCode", "Customer Code", 50);
                objSBOAPI.AddAlphaField("AIS_LOAD2", "CardName", "Customer Name", 200);

                objSBOAPI.CreateTable("AIS_OPDE", "Pick Delivery Header", SAPbobsCOM.BoUTBTableType.bott_Document);
                objSBOAPI.AddDateField("AIS_OPDE", "DocDate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddAlphaField("AIS_OPDE", "Remarks", "Remarks", 200);
                objSBOAPI.AddDateField("AIS_OPDE", "BillRecDate", "Bill Rec Date", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.addField("AIS_OPDE", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "O,C,R", "Open,Close,Release", "O");
                objSBOAPI.AddFloatField("AIS_OPDE", "DocTotal", "Doc Total", SAPbobsCOM.BoFldSubTypes.st_Sum);
                objSBOAPI.AddAlphaField("AIS_OPDE", "VehicleNo", "Vehicl No", 50);
                objSBOAPI.AddAlphaField("AIS_OPDE", "Location", "Location", 50);
                objSBOAPI.AddAlphaField("AIS_OPDE", "TFVehicle", "TypeOfVehicle", 50);
                objSBOAPI.AddAlphaField("AIS_OPDE", "DriverName", "Driver Name", 100);
                objSBOAPI.AddDateField("AIS_OPDE", "LorryIn", "Lorry In time", SAPbobsCOM.BoFldSubTypes.st_Time);
                objSBOAPI.AddDateField("AIS_OPDE", "LorryOut", "Lorry Out time", SAPbobsCOM.BoFldSubTypes.st_Time);
                objSBOAPI.AddAlphaField("AIS_OPDE", "Unit", "Unit", 50);
                objSBOAPI.AddAlphaField("AIS_OPDE", "WBType", "WB Type", 50);
                objSBOAPI.AddAlphaField("AIS_OPDE", "Mobile", "Mobile", 50);
                objSBOAPI.addField("AIS_OPDE", "Approved", "Approved", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N");
                objSBOAPI.AddFloatField("AIS_OPDE", "ActWeight", "Actual Weight", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_OPDE", "TNetWeight", "Tem Net Weight", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_OPDE", "WBNetWt", "WB Net Weight", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_OPDE", "DiffWeight", "Diff Weight", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_OPDE", "Coolie", "Coolie", SAPbobsCOM.BoFldSubTypes.st_Rate);
                objSBOAPI.AddFloatField("AIS_OPDE", "LFrieght", "Lorry Frieght", SAPbobsCOM.BoFldSubTypes.st_Sum);
                objSBOAPI.AddFloatField("AIS_OPDE", "FWeight", "First Weight", SAPbobsCOM.BoFldSubTypes.st_Rate);
                objSBOAPI.AddFloatField("AIS_OPDE", "SWeight", "Second Weight", SAPbobsCOM.BoFldSubTypes.st_Rate);
                objSBOAPI.AddFloatField("AIS_OPDE", "TotalTon", "Total Ton", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_OPDE", "TotalBag", "Total Bag", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddAlphaField("AIS_OPDE", "FrightType", "Fright Type", 1, "H,O,C", "Hired,Own,Customer", "H");

                objSBOAPI.CreateTable("AIS_PDE1", "Pick Delivery Deatils", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                objSBOAPI.AddAlphaField("AIS_PDE1", "Select", "Select", 1);
                objSBOAPI.AddAlphaField("AIS_PDE1", "PListNo", "PListNo", 11);
                objSBOAPI.AddAlphaField("AIS_PDE1", "BillNo", "Bill Number", 20);
                objSBOAPI.AddAlphaField("AIS_PDE1", "BDocEntry", "Bill Entry", 20);
                objSBOAPI.AddDateField("AIS_PDE1", "BDate", "Bill Date", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddAlphaField("AIS_PDE1", "CusCode", "Customer Code", 15);
                objSBOAPI.AddAlphaField("AIS_PDE1", "CusName", "Customer Name", 100);
                objSBOAPI.AddAlphaField("AIS_PDE1", "BaseLine", "Base Line", 11);
                objSBOAPI.AddAlphaField("AIS_PDE1", "ItemCode", "Item Code", 50);
                objSBOAPI.AddAlphaField("AIS_PDE1", "WhsCode", "WhsCode", 50);
                objSBOAPI.AddAlphaField("AIS_PDE1", "ItemName", "Item Name", 100);
                objSBOAPI.AddAlphaField("AIS_PDE1", "LoadingChg", "LoadingChrg", 50);
                objSBOAPI.AddAlphaField("AIS_PDE1", "DefUnit", "Default Unit", 20);
                objSBOAPI.AddAlphaField("AIS_PDE1", "TreeType", "TreeType", 50);
                objSBOAPI.AddNumericField("AIS_PDE1", "DocEntry", "DocEntry", 9);
                objSBOAPI.AddNumericField("AIS_PDE1", "UomEntry", "UomEntry", 9);
                objSBOAPI.AddAlphaField("AIS_PDE1", "UomCode", "UomCode", 50);
                objSBOAPI.AddFloatField("AIS_PDE1", "SOWeight", "SOWeight", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_PDE1", "ItemPerKg", "Inv Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_PDE1", "Price", "Price", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_PDE1", "Qty", "Inv Quantity ", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_PDE1", "PlanQty", "Plan Quantity ", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_PDE1", "LineTotal", "Line Total", SAPbobsCOM.BoFldSubTypes.st_Sum);
                objSBOAPI.AddAlphaField("AIS_PDE1", "TaxCode", "TaxCode", 50);
                objSBOAPI.AddAlphaField("AIS_PDE1", "TargetEntry", "TargetEntry", 50);
                objSBOAPI.AddAlphaField("AIS_PDE1", "TargetNum", "TargetNum", 50);
                objSBOAPI.AddAlphaField("AIS_PDE1", "TargetObject", "TargetObject", 50);

                //Thiru Addon Changes - Start - 02.07.2021
                objSBOAPI.AddFloatField("AIS_PDE1", "AVA_SalVal", "Sales value for this FY", SAPbobsCOM.BoFldSubTypes.st_Sum);
                objSBOAPI.AddFloatField("AIS_PDE1", "AVA_TaxAmt", "Tax Amount", SAPbobsCOM.BoFldSubTypes.st_Sum);
                objSBOAPI.AddFloatField("AIS_PDE1", "AVA_DocTot", "Document Total", SAPbobsCOM.BoFldSubTypes.st_Sum);
                //Thiru Addon Changes - End - 02.07.2021

                objSBOAPI.CreateTable("AIS_OCAC", "Grinding Capacity Header", SAPbobsCOM.BoUTBTableType.bott_Document);
                objSBOAPI.AddDateField("AIS_OCAC", "FromDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddDateField("AIS_OCAC", "ToDate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddDateField("AIS_OCAC", "DocDate", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddAlphaField("AIS_OCAC", "Unit", "Unit", 50);

                objSBOAPI.CreateTable("AIS_CAC1", "Grinding Capacity Detail", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                objSBOAPI.AddDateField("AIS_CAC1", "PlanDate", "PlanDate", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddFloatField("AIS_CAC1", "DayPlan", "DayPlan", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_CAC1", "NightPlan", "NightPlan", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_CAC1", "Available", "Available", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_CAC1", "Allocated", "Allocated", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddFloatField("AIS_CAC1", "Balance", "Balance", SAPbobsCOM.BoFldSubTypes.st_Quantity);

                //objSBOAPI.CreateTable("AIS_OSAP", "Sales Planning Header", SAPbobsCOM.BoUTBTableType.bott_Document);
                //objSBOAPI.AddAlphaField("AIS_OSAP", "Unit", "Unit", 50);
                //objSBOAPI.AddAlphaField("AIS_OSAP", "ShiftCode", "ShiftCode", 100);
                //objSBOAPI.AddDateField("AIS_OSAP", "CPeriodFrom", "Capacity Period From", SAPbobsCOM.BoFldSubTypes.st_None);
                //objSBOAPI.AddDateField("AIS_OSAP", "CPeriodTo", "Capacity Period To", SAPbobsCOM.BoFldSubTypes.st_None);
                //objSBOAPI.AddDateField("AIS_OSAP", "DocDate", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None);
                //objSBOAPI.AddAlphaField("AIS_OSAP", "Status", "Status", 1, "A,C,O,W", "Active,Sales Order Created,Open,Work Order Created", "O");
                //objSBOAPI.AddAlphaField("AIS_OSAP", "ShiftCode", "ShiftCode", 100);
                //objSBOAPI.AddAlphaField("AIS_OSAP", "Type", "Type", 1, "M,W", "Monthly Plan,Daily Plan", "M");
                //objSBOAPI.AddAlphaField("AIS_OSAP", "QtyType", "Qty Type", 1, "B,T", "Bag,Tonnage", "B");
                //objSBOAPI.AddAlphaField("AIS_OSAP", "CatgryType", "CatgryType", 1, "I,C", "Item,Category", "I");

                //objSBOAPI.CreateTable("AIS_SAP1", "Sales Planning Detai1l", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //objSBOAPI.AddAlphaField("AIS_SAP1", "Selected", "Selected", 1);
                //objSBOAPI.AddDateField("AIS_SAP1", "PlanDate", "Planned Date", SAPbobsCOM.BoFldSubTypes.st_None);
                //objSBOAPI.AddAlphaField("AIS_SAP1", "ItemCode", "Item Code", 50);
                //objSBOAPI.AddAlphaField("AIS_SAP1", "ItemName", "Item Name", 100);
                //objSBOAPI.AddFloatField("AIS_SAP1", "Instock", "Instock", SAPbobsCOM.BoFldSubTypes.st_Quantity);

                //objSBOAPI.CreateTable("AIS_SAP2", "Sales Planning Detail2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //objSBOAPI.AddDateField("AIS_SAP2", "Date", "Date", SAPbobsCOM.BoFldSubTypes.st_None);
                //objSBOAPI.AddAlphaField("AIS_SAP2", "ItemCode", "Item Code", 50);
                //objSBOAPI.AddAlphaField("AIS_SAP2", "Unit", "Unit", 50);
                //objSBOAPI.AddFloatField("AIS_SAP2", "Capacity", "Capacity", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddFloatField("AIS_SAP2", "Qty", "Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddFloatField("AIS_SAP2", "M1", "M1", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddFloatField("AIS_SAP2", "M2", "M2", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddFloatField("AIS_SAP2", "M3", "M3", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddFloatField("AIS_SAP2", "Atta", "Atta", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddFloatField("AIS_SAP2", "Suji", "Suji", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddFloatField("AIS_SAP2", "Bran", "Bran", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddFloatField("AIS_SAP2", "Total", "Total", SAPbobsCOM.BoFldSubTypes.st_Sum);

                //objSBOAPI.CreateTable("AIS_SAP3", "Sales Planning Detail3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //objSBOAPI.AddAlphaField("AIS_SAP3", "UniqID", "UniqID", 11);
                //objSBOAPI.AddAlphaField("AIS_SAP3", "BaseNum", "BaseNum", 11);
                //objSBOAPI.AddAlphaField("AIS_SAP3", "BaseEntry", "BaseEntry", 11);
                //objSBOAPI.AddAlphaField("AIS_SAP3", "Selected", "Selected", 1);
                //objSBOAPI.AddFloatField("AIS_SAP3", "Qty", "Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddFloatField("AIS_SAP3", "ActualQty", "Actual Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddAlphaField("AIS_SAP3", "DayID", "DayID", 50);
                //objSBOAPI.AddAlphaField("AIS_SAP3", "CardCode", "CardCode", 50);
                //objSBOAPI.AddAlphaField("AIS_SAP3", "CardName", "CardName", 100);
                //objSBOAPI.AddAlphaField("AIS_SAP3", "Unit", "Unit", 50);
                //objSBOAPI.AddAlphaField("AIS_SAP3", "UpdUnit", "Updated Unit", 50);


                //objSBOAPI.CreateTable("AIS_SAP4", "Sales Planning Detail4", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //objSBOAPI.AddAlphaField("AIS_SAP4", "UniqID", "Unique Id", 50);
                //objSBOAPI.AddAlphaField("AIS_SAP4", "ItemCode", "Item Code", 50);
                //objSBOAPI.AddFloatField("AIS_SAP4", "DCapcity", "Daily Capacity", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddFloatField("AIS_SAP4", "SOQty", "SO Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddFloatField("AIS_SAP4", "InStock", "InStock", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddFloatField("AIS_SAP4", "Feasibi", "Feasibi", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddFloatField("AIS_SAP4", "ActPacking", "ActPacking", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddFloatField("AIS_SAP4", "ShrtExcess", "Shortage/Excess", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddAlphaField("AIS_SAP4", "Slob", "Slob", 100);

                //objSBOAPI.CreateTable("AIS_SAP5", "Sales Planning Detail5", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //objSBOAPI.AddAlphaField("AIS_SAP5", "Selected", "Selected", 1);
                //objSBOAPI.AddDateField("AIS_SAP5", "PlanDate", "Planned Date", SAPbobsCOM.BoFldSubTypes.st_None);
                //objSBOAPI.AddAlphaField("AIS_SAP5", "ItemCode", "Item Code", 50);
                //objSBOAPI.AddAlphaField("AIS_SAP5", "ItemName", "Item Name", 100);
                //objSBOAPI.AddFloatField("AIS_SAP5", "DCapacity", "Daily Capacity", SAPbobsCOM.BoFldSubTypes.st_Quantity);

                //objSBOAPI.CreateTable("AIS_SAP6", "Sales Planning Detail6", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //objSBOAPI.AddAlphaField("AIS_SAP6", "Selected", "Selected", 1);
                //objSBOAPI.AddDateField("AIS_SAP6", "PlanDate", "Planned Date", SAPbobsCOM.BoFldSubTypes.st_None);
                //objSBOAPI.AddAlphaField("AIS_SAP6", "ItemCode", "Item Code", 50);
                //objSBOAPI.AddAlphaField("AIS_SAP6", "ItemName", "Item Name", 100);
                //objSBOAPI.AddFloatField("AIS_SAP6", "SOQuantity", "SO Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity);

                //objSBOAPI.CreateTable("AIS_SAP7", "Sales Planning Detail7", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //objSBOAPI.AddAlphaField("AIS_SAP7", "Selected", "Selected", 1);
                //objSBOAPI.AddDateField("AIS_SAP7", "PlanDate", "Planned Date", SAPbobsCOM.BoFldSubTypes.st_None);
                //objSBOAPI.AddAlphaField("AIS_SAP7", "ItemCode", "Item Code", 50);
                //objSBOAPI.AddAlphaField("AIS_SAP7", "ItemName", "Item Name", 100);
                //objSBOAPI.AddFloatField("AIS_SAP7", "Feasbility", "Feasbility", SAPbobsCOM.BoFldSubTypes.st_Quantity);

                //objSBOAPI.CreateTable("AIS_SAP8", "Sales Planning Detail8", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //objSBOAPI.AddAlphaField("AIS_SAP8", "Selected", "Selected", 1);
                //objSBOAPI.AddDateField("AIS_SAP8", "PlanDate", "Planned Date", SAPbobsCOM.BoFldSubTypes.st_None);
                //objSBOAPI.AddAlphaField("AIS_SAP8", "ItemCode", "Item Code", 50);
                //objSBOAPI.AddAlphaField("AIS_SAP8", "ItemName", "Item Name", 100);
                //objSBOAPI.AddFloatField("AIS_SAP8", "ActPackscd", "Actual Packing schedule", SAPbobsCOM.BoFldSubTypes.st_Quantity);

                //objSBOAPI.CreateTable("AIS_SAP9", "Sales Planning Detail9", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //objSBOAPI.AddAlphaField("AIS_SAP9", "Selected", "Selected", 1);
                //objSBOAPI.AddDateField("AIS_SAP9", "PlanDate", "Planned Date", SAPbobsCOM.BoFldSubTypes.st_None);
                //objSBOAPI.AddAlphaField("AIS_SAP9", "ItemCode", "Item Code", 50);
                //objSBOAPI.AddAlphaField("AIS_SAP9", "ItemName", "Item Name", 100);
                //objSBOAPI.AddFloatField("AIS_SAP9", "ShrtExcess", "Shortage/Excess", SAPbobsCOM.BoFldSubTypes.st_Quantity);

                //objSBOAPI.CreateTable("AIS_SAP10", "Sales Planning Detail10", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //objSBOAPI.AddDateField("AIS_SAP10", "PlanDate", "Planned Date", SAPbobsCOM.BoFldSubTypes.st_None);
                //objSBOAPI.AddAlphaField("AIS_SAP10", "ItemCode", "Item Code", 50);
                //objSBOAPI.AddAlphaField("AIS_SAP10", "ItemName", "Item Name", 100);
                //objSBOAPI.AddAlphaField("AIS_SAP10", "SlobDays", "Slob Days", 50);

                //objSBOAPI.CreateTable("AIS_SAP11", "Sales Planning Detail10", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //objSBOAPI.AddAlphaField("AIS_SAP11", "PlanDate", "Planned Date", 50);
                //objSBOAPI.AddFloatField("AIS_SAP11", "Stock", "Stock", SAPbobsCOM.BoFldSubTypes.st_Quantity);

                // ************************* EXISTING SAP TABLES **************************** //

                objSBOAPI.SBO_Appln.StatusBar.SetText("Creating Tables Please Wait...........", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                objSBOAPI.AddAlphaField("OITM", "ItemCategory", "Item Category", 50);
                objSBOAPI.AddAlphaField("OITM", "DefUnit", "Default Unit", 50);
                objSBOAPI.AddAlphaField("OITM", "TCategory", "Tax Category", 50);
                objSBOAPI.AddAlphaField("OITM", "UNIT", "UNIT", 50, "01,02,03", "01,02,03", "");
                objSBOAPI.AddFieldwithLinkTable("OITM", "UNIT", "UNIT", 50, "AIS_OUNT");
                objSBOAPI.AddAlphaField("OITM", "IsProd", "Production",  1, "Y,N", "Yes,No", "N");
                objSBOAPI.AddAlphaField("OITM", "SplItem", "Special Item", 1, "Y,N", "Yes,No", "N");
                objSBOAPI.AddAlphaField("OITM", "PlanUnit", "Plan Unit", 50);
                objSBOAPI.AddAlphaField("OITM", "BWeight", "Bag Weight", 10);
                objSBOAPI.AddAlphaField("OITM", "LCHARGES", "LCHARGES", 50);
                objSBOAPI.AddFloatField("OITM", "Daylimit", "Day Limit", SAPbobsCOM.BoFldSubTypes.st_Rate);

                objSBOAPI.AddAlphaField("INV1", "BaseEntry", "BaseEntry", 11);
                objSBOAPI.AddAlphaField("INV1", "BaseNum", "BaseNum", 11);
                objSBOAPI.AddAlphaField("INV1", "BaseObject", "BaseObject", 20);

                objSBOAPI.AddAlphaField("OPKL", "TruckNo", "TruckNo", 50);
                objSBOAPI.AddAlphaField("OPKL", "MobileNum", "MobileNum", 10);
                objSBOAPI.AddAlphaField("OPKL", "Drivname", "Drivname", 50);
                objSBOAPI.AddAlphaField("OPKL", "Vehicle", "VehicleType", 50);
                objSBOAPI.AddFloatField("OPKL", "Percentage", "Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage);
                objSBOAPI.AddAlphaField("OPKL", "FRTY", "Freight Type", 50);
                //Thiru - 20210728
                objSBOAPI.AddAlphaField("OPKL", "AVA_UNIT", "UNIT", 10);
                objSBOAPI.AddAlphaField("OPKL", "AVA_LOCN", "LOCATION", 50);
                //Thiru - 20210728

                objSBOAPI.AddAlphaField("OWOR", "BaseNum", "BaseNum", 50);
                objSBOAPI.AddAlphaField("OWOR", "BaseEntry", "BaseEntry", 50);
                objSBOAPI.AddAlphaField("OWOR", "Series", "Series", 50);
                objSBOAPI.AddAlphaField("OWOR", "BaseObject", "BaseObject", 50);

                objSBOAPI.AddAlphaField("ORDR", "UNIT", "UNIT", 10);
                objSBOAPI.AddAlphaField("ORDR", "BaseEntry", "Base Entry", 11);
                objSBOAPI.AddAlphaField("ORDR", "BaseObject", "Base Object", 11);
                objSBOAPI.AddAlphaField("ORDR", "BaseNum", "Base Num", 11);
                objSBOAPI.AddAlphaField("ORDR", "BaseSeries", "Base Series", 11);
                objSBOAPI.AddAlphaField("ORDR", "TPlanning", "Truck Planning", 3, "YES,NO", "YES,NO", "NO");
                objSBOAPI.AddAlphaField("ORDR", "VNo", "Truck No", 20);
                objSBOAPI.AddAlphaField("ORDR", "PreSalEnt", "PreSalEntry", 11,"","","");
                objSBOAPI.AddAlphaField("ORDR", "PreSalNo", "PreSalNo", 11,"","","");

                objSBOAPI.AddAlphaField("OCRD", "UserName", "UserName", 50);
                objSBOAPI.AddAlphaField("OCRD", "Password", "Password", 50);
                objSBOAPI.AddAlphaField("OCRD", "UNIT", "UNIT", 10);
                objSBOAPI.AddDateField("OCRD", "MOFromDate", "MOFromDate",SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddDateField("OCRD", "MOToDate", "MOToDate", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddDateField("OCRD", "SOFromDate", "SOFromDate", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddDateField("OCRD", "SOToDate", "SOToDate", SAPbobsCOM.BoFldSubTypes.st_None);
                objSBOAPI.AddAlphaField("OCRD", "ItemType", "ItemType", 3, "S,O", "Same Category,Otehr Category", "");
                objSBOAPI.AddAlphaField("OCRD", "OthCatgory", "Other Category", 1, "Y,N", "Yes,No", "N");

                objSBOAPI.AddAlphaField("OEXD", "IsTCS", "Is TCS", 1);

                //APInvoice 
                //objSBOAPI.AddAlphaField("OPCH", "AVA_LOADINGPERSON", "Loading Person", 50);
                //objSBOAPI.AddAlphaField("OPCH", "AVA_DRIVERMOBILENO", "Driver Mobile No", 15);
                //objSBOAPI.AddAlphaField("OPCH", "AVA_DRIVERLNO", "Driver License No", 20);
                //objSBOAPI.AddAlphaField("OPCH", "AVA_OWNTRUCK", "Own Truck", 10, "Y,N", "Yes,No", "N");
                //objSBOAPI.AddAlphaField("OPCH", "AVA_SUPERVISOR", "Supervisor", 100);
                objSBOAPI.AddAlphaField("OPCH", "AVA_PURCHASETOKENNO", "Purchase Token No", 50);
                //objSBOAPI.AddAlphaField("OPCH", "AVA_VEHICLENO", "Vehicle No", 50);
                objSBOAPI.AddFloatField("OPCH", "AVA_INSBAL", "Insurance Balance", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                objSBOAPI.AddAlphaField("OPCH", "AVA_INSPNO", "Insurance Policy No", 100);
                //objSBOAPI.AddAlphaField("OPCH", "AVA_FRTCONDITION", "Freight Condition", 20, "FOR,EX-MILL,DIRECT", "FOR,EX-MILL,DIRECT", "DIRECT");
                //objSBOAPI.AddAlphaField("OPCH", "AVA_TRANSNAME", "Transporter Name", 100);
                //objSBOAPI.AddFloatField("OPCH", "AVA_FREIGHTBALANCE", "Freight Balance", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddFloatField("OPCH", "AVA_FREIGHTADVANCE", "Freight Advance", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddFloatField("OPCH", "AVA_FREIGHTAMT", "Freight Amount", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //objSBOAPI.AddFloatField("OPCH", "AVA_FREIGHTRATE", "Freight Rate", SAPbobsCOM.BoFldSubTypes.st_Quantity);

                objSBOAPI.AddAlphaField("OEXD", "AVA_INSURANCE", "Insurance", 5, "Y,N", "Yes,No", "N");
                //objSBOAPI.AddAlphaField("OEXD", "AVA_FREIGHT", "Freight", 5, "Y,N", "Yes,No", "N");
                objSBOAPI.AddAlphaField("OSTC", "AVA_INSURANCE", "Insurance", 5, "Y,N", "Yes,No", "N");
                //objSBOAPI.AddAlphaField("OSTC", "AVA_FREIGHT", "Freight", 5, "Y,N", "Yes,No", "N");
                //objSBOAPI.AddAlphaField("OACT", "AVA_FREIGHT", "Freight", 5, "Y,N", "Yes,No", "N");
                //objSBOAPI.AddAlphaField("OSAC", "AVA_FREIGHT", "Freight", 5, "Y,N", "Yes,No", "N");

                //Email Setup --> UDF's
                objSBOAPI.AddAlphaField("OADM", "AV_SMTPS", "SMTP Server", 254);
                objSBOAPI.AddAlphaField("OADM", "AV_Port", "Port", 10);
                objSBOAPI.AddAlphaField("OADM", "AV_Essl", "Enable SSL", 1, "Y,N", "Yes,No", "N");
                objSBOAPI.AddAlphaField("OADM", "AV_EMID", "Email Id", 100);
                objSBOAPI.AddAlphaField("OADM", "AV_EPswd", "Email Password", 100);
                objSBOAPI.AddAlphaField("OADM", "AV_ToEMID", "To Email", 100);

                //Server Credentials
                objSBOAPI.AddAlphaField("OADM", "AV_Domain", "Domain", 150);
                objSBOAPI.AddAlphaField("OADM", "AV_UserName", "UserName", 150);
                objSBOAPI.AddAlphaField("OADM", "AV_Password", "Password", 150);
                objSBOAPI.AddAlphaField("OADM", "AV_WorkStation", "Enable WorkStation", 1, "Y,N", "Yes,No", "N");

                objSBOAPI.AddAlphaField("OADM", "AVA_EmailSbj", "EmailSubject", 254);
                objSBOAPI.AddAlphaMemoField("OADM", "AVA_EmailBody", "Email Body", 256000);

                objSBOAPI.AddAlphaMemoField("OADM", "AV_LPATH", "Layout Path", 200);
                objSBOAPI.AddAlphaMemoField("OUSR", "AV_FPATH", "Folder Path", 200);
                //Thiru Addon Changes - Start - 05.08.2021
                objSBOAPI.AddAlphaField("OUSR", "AVA_BDPEUR", "Bill Delivery Price Edit User", 1, "Y,N", "Yes,No", "N");
                //Thiru Addon Changes - End - 05.08.2021

                //Thiru Addon Changes - Start - 01.07.2021

                objSBOAPI.AddAlphaField("OCRD", "AVA_TurnOver", "Last FY Turn Over Greater than 10 Cr", 10, "Y,N,Unknown", "Yes,No,Unknown", "Unknown");
                objSBOAPI.AddAlphaField("OCRD", "AVA_ITReturnFiled", "Last 2 FY IT Return Filed", 10, "Y,N,Unknown", "Yes,No,Unknown", "Unknown");
                objSBOAPI.AddAlphaField("OCRD", "AVA_SumTDSTCS", "Sum Of TDS+TCS Greater than 50000 for Last 2 FY", 10, "Y,N,Unknown", "Yes,No,Unknown", "Unknown");

                objSBOAPI.AddAlphaField("OCRD", "AVA_CusTDSDed", "Customer Deducting TDS", 10, "Y,N,Unknown", "Yes,No,Unknown", "Unknown");

                objSBOAPI.AddAlphaField("OADM", "AVA_CusTCSVal", "Customer TCS Value", 15);
                objSBOAPI.AddFloatField("OADM", "AVA_TCSRSI", "TCS Rate for Scrap Items", SAPbobsCOM.BoFldSubTypes.st_Quantity);
                //Thiru Addon Changes - End - 01.07.2021

                //QR Code Generation
                objSBOAPI.AddAlphaMemoField("OADM", "AVA_FPATH", "File Path", 254);
                objSBOAPI.AddAlphaMemoField("OADM", "AVA_SSIC", "Server Image Connection", 254);

                objSBOAPI.AddAlphaMemoField("OPKL", "AV_IMG", "QR Image", 254);
                objSBOAPI.AddAlphaMemoField("OPKL", "AV_QPATH", "QR Path", 254);

            }
            catch (Exception)
            {

            }
            CreateUDOs();
        }
        #endregion

        #region Create UDOs
        public void CreateUDOs()
        {
            try
            {
                // ***************************** SAMPLES ************************************ //

                //Default forms like No-Objects table view

                //objSBOAPI.createUDO("AVA_A2", "AVA_A2", "HumanType", SAPbobsCOM.BoUDOObjType.boud_MasterData, "", true);

                //Default forms with UDF visible like No - Objects table view

                //objSBOAPI.createUDO("AVA_A3", "AVA_A3", "ParentForm", SAPbobsCOM.BoUDOObjType.boud_MasterData, "", true, "M", "U_Size,U_Sizecode,U_Gender,U_ST", "Size,Size Code,Gender,Size Type", "Y,Y,Y,Y");

                //Master UDO Creation

                //objSBOAPI.createUDO("AV_PROC", "AV_PROC", "Process", SAPbobsCOM.BoUDOObjType.boud_MasterData, "AV_PROC1");

                //Document UDO Creation

                //objSBOAPI.createUDO("AV_CDES", "AV_CDES", "CostingDesign", SAPbobsCOM.BoUDOObjType.boud_Document, "AV_CDES1");

                // ***************************** SAMPLES ************************************ //


                // ***************************** DEFAULT FORMS ****************************** //

                //Write here

                // ***************************** MASTER UDO ********************************* //

                objSBOAPI.createUDO("AIS_OBRN", "OBRN", "Branch Master", SAPbobsCOM.BoUDOObjType.boud_MasterData, "AIS_BRN1");
                objSBOAPI.createUDO("AIS_OCIA", "OCIA", "Customer Wise ItemAllocation", SAPbobsCOM.BoUDOObjType.boud_MasterData, "AIS_CIA1");
                objSBOAPI.createUDO("AIS_OREM", "OREM", "Yield Master", SAPbobsCOM.BoUDOObjType.boud_MasterData, "AIS_REM1");
                objSBOAPI.createUDO("AIS_OSFT", "OSFT", "Shift Master", SAPbobsCOM.BoUDOObjType.boud_MasterData, "");
                objSBOAPI.createUDO("AIS_ORMC", "ORMC", "Packing Capacity", SAPbobsCOM.BoUDOObjType.boud_MasterData, "AIS_RMC1,AIS_RMC2");

                objSBOAPI.createUDO("AVA_INSURANCEH", "AVA_INSURANCE", "Insurance Details", SAPbobsCOM.BoUDOObjType.boud_MasterData, "AVA_INSURANCEL");
                objSBOAPI.createUDO("AV_CSLGT", "AV_CSLGT", "Customer Statement Log", SAPbobsCOM.BoUDOObjType.boud_MasterData, "");

                // ***************************** DOCUMENT UDO ******************************* //

                objSBOAPI.createUDO("AIS_OPRE", "OPRE", "Pre Sales Order", SAPbobsCOM.BoUDOObjType.boud_Document, "AIS_PRE1,AIS_PRE2");
                objSBOAPI.createUDO("AIS_LOAD", "AIS_LOAD", "Bill Delivery", SAPbobsCOM.BoUDOObjType.boud_Document, "AIS_LOAD1,AIS_LOAD2");
                objSBOAPI.createUDO("AIS_OPDE", "OPDE", "PickList Delivery", SAPbobsCOM.BoUDOObjType.boud_Document, "AIS_PDE1");
                objSBOAPI.createUDO("AIS_OCAC", "OCAC", "Grinding Capacity", SAPbobsCOM.BoUDOObjType.boud_Document, "AIS_CAC1");
                //objSBOAPI.createUDO("AIS_OSAP", "OSAP", "Sales Planning", SAPbobsCOM.BoUDOObjType.boud_Document, "AIS_SAP1, AIS_SAP2, AIS_SAP3, AIS_SAP4, AIS_SAP5, AIS_SAP6, AIS_SAP7, AIS_SAP8, AIS_SAP9, AIS_SAP10, AIS_SAP11");

            }
            catch (Exception)
            {

            }
        }
        #endregion

        #region Create Queries
        public void CreateQueries()
        {
            try
            {
                // ***************************** SAMPLES ************************************ //

                //Create Query Category

                //objSBOAPI.create_queryCat("Addon Query", "YYYYYYYYYYYYYYYYYYYY");

                //Create Query

                //objSBOAPI.createQuery("Addon Query", "Test_CardCode", "Test_CardCode.txt");

                // ***************************** SAMPLES ************************************ //


                // ************************** Query Manager Update ************************** //

                objSBOAPI.SBO_Appln.StatusBar.SetText("Creating Query Categories and Queries Please Wait...........", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    objSBOAPI.create_queryCat("AddonQuery_Hana", "YYYYYYYYYYYYYYYYYYYY");
                    objSBOAPI.createQuery("AddonQuery_Hana", "ItemCode_CFL_Loading_HANA", "ItemCode_CFL_Loading_HANA.txt");
                    objSBOAPI.createQuery("AddonQuery_Hana", "PreSalesOrder_Update_Fields_HANA", "PreSalesOrder_Update_Fields_HANA.txt");

                    objSBOAPI.createQuery("AddonQuery_Hana", "Lab_Deduction", "Lab_Deduction.txt");
                    objSBOAPI.createQuery("AddonQuery_Hana", "Deduction_Details", "Deduction_Details.txt");
                    objSBOAPI.createQuery("AddonQuery_Hana", "APCreditMemo", "APCreditMemo.txt");
                    objSBOAPI.createQuery("AddonQuery_Hana", "APInvoice", "APInvoice.txt");
                    objSBOAPI.createQuery("AddonQuery_Hana", "Insurance_Update", "Insurance_Update.txt");

                    objSBOAPI.createQuery("AddonQuery_Hana", "AR_Invoice_HANA", "AR_Invoice_HANA.txt");
                    objSBOAPI.createQuery("AddonQuery_Hana", "AR_CreditMemo_HANA", "AR_CreditMemo_HANA.txt");
                    objSBOAPI.createQuery("AddonQuery_Hana", "CSGrouping_HANA", "CSGrouping_HANA.txt");
                    objSBOAPI.createQuery("AddonQuery_Hana", "GRPO_HANA", "GRPO_HANA.txt");
                    objSBOAPI.createQuery("AddonQuery_Hana", "OutgoingPayment_HANA", "OutgoingPayment_HANA.txt");
                    objSBOAPI.createQuery("AddonQuery_Hana", "CustomerStatement_HANA", "CustomerStatement_HANA.txt");
                    objSBOAPI.createQuery("AddonQuery_Hana", "MinMaxUpload", "MinMaxUpload.txt");
                    objSBOAPI.createQuery("AddonQuery_Hana", "UploadedResults", "UploadedResults.txt");
                    objSBOAPI.createQuery("AddonQuery_Hana", "PickListQRCode", "PickListQRCode.txt");
                }
                else
                {
                    objSBOAPI.create_queryCat("AddonQuery_SQL", "YYYYYYYYYYYYYYYYYYYY");
                    objSBOAPI.createQuery("AddonQuery_SQL", "ItemCode_CFL_Loading_SQL", "ItemCode_CFL_Loading_SQL.txt");
                    objSBOAPI.createQuery("AddonQuery_SQL", "PreSalesOrder_Update_Fields_SQL", "PreSalesOrder_Update_Fields_SQL.txt");
                }
                                

                objSBOAPI.SBO_Appln.StatusBar.SetText("Query Manager will be updated successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            }
            catch (Exception)
            {

            }
        }
        #endregion

    }
}
