using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using SAPbobsCOM;
using SAPbouiCOM;
using iTextSharp.text.xml.xmp;
using iTextSharp.text.pdf;
using iTextSharp.text;
using ThoughtWorks.QRCode.Codec;
using Image = iTextSharp.text.Image;
using Sap.Data.Hana;

namespace JKFM_Source
{
    class Cls_PickList
    {
        #region Declaration
        SAPbouiCOM.Form objform;
        ClsSBO objSBOAPI;
        string AbsEntry = string.Empty;
        int QRCODE_Customer_Count = 0;
        int Invoice_LineNum_Count = 0;
        int U_IMG_NULL_Count = 0;
        int SuccessCount = 0;

        string ObjType = string.Empty;
        string MyXml = string.Empty;
        #endregion

        #region Constructor
        public Cls_PickList(ClsSBO objSBO)
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

                    if ((pval.EventType == SAPbouiCOM.BoEventTypes.et_CLICK) | (pval.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) | (pval.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE) | (pval.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE) | (pval.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE))
                    {
                        try
                        {
                            SAPbouiCOM.StaticText oST;
                            oST = objform.Items.Item("ST1").Specific;
                            SAPbouiCOM.Form oform;
                            oform = objSBOAPI.SBO_Appln.Forms.Item(oST.Caption);
                            oform.Select();
                            bubbleevent = false;
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                }
                else
                {
                    switch (pval.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            
                            if (pval.ItemUID == "3" && (objform.Mode == BoFormMode.fm_UPDATE_MODE || objform.Mode == BoFormMode.fm_OK_MODE))
                            {
                                AbsEntry = objform.DataSources.DBDataSources.Item("OPKL").GetValue("AbsEntry", 0);

                                if (objSBOAPI.Query_Execute("Select \"U_AV_IMG\" where \"AbsEntry\" = '"+ AbsEntry +"'") == "")
                                {
                                    Generate_QRCode("PickList", AbsEntry);
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
                    switch (pval.MenuUID)
                    {
                        case "1281":
                            objform.Items.Item("Btn_RGPOut").Visible = false;
                            objform.Items.Item("Btn_QRCode").Visible = false;
                            break;

                        case "1282":
                            objform.Items.Item("Btn_QRCode").Visible = false;
                            break;

                        case "1287":
                            objform.Items.Item("Btn_QRCode").Visible = false;
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

        #region Generate_QRCode
        private void Generate_QRCode(string Type, string AbsEntry)
        {
            try
            {
                SAPbobsCOM.Recordset objRecordSet = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                //Thiru
                string Load = string.Empty;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    Load = objSBOAPI.Query_Execute("SELECT \"QString\" FROM \"OUQR\" WHERE \"QCategory\"=(SELECT \"CategoryId\" FROM \"OQCN\" WHERE \"CatName\"='AddonQuery_Hana') AND \"QName\"='PickListQRCode'");
                }
                else
                {
                    Load = objSBOAPI.Query_Execute("SELECT QString FROM OUQR WHERE QCategory=(SELECT CategoryId FROM OQCN WHERE CatName='AddonQuery_SQL') AND QName='PickListQRCode'");
                }
                if (Load == "")
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("PickListQRCode Query is Missing, Kindly do the Query Manager Update...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    objform.Freeze(false);
                    return;
                }
                else
                {
                    Load = Load.Replace("[%C]", Type);
                    Load = Load.Replace("[%D]", AbsEntry);

                    objRecordSet.DoQuery(Load);
                    if (objRecordSet.RecordCount > 0)
                    {
                        string str2 = "";
                        string text = "";
                        string conInfo1_1 = "";
                        string conInfo1_2 = "";
                        string Str1 = "";
                        string path = "";
                        string str3 = "";

                        SAPbobsCOM.Recordset Orec;
                        Orec = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        {
                            Str1 = "Select \"U_AVA_FPATH\",\"U_AVA_SSIC\" from \"OADM\"";
                        }
                        else
                        {
                            Str1 = "Select U_AVA_FPATH,U_AVA_SSIC from OADM";
                        }
                        Orec.DoQuery(Str1);
                        if (Orec.RecordCount > 0)
                        {
                            if (Orec.Fields.Item("U_AVA_FPATH").Value != "" & Orec.Fields.Item("U_AVA_SSIC").Value != "")
                            {
                                conInfo1_1 = Orec.Fields.Item("U_AVA_FPATH").Value;
                                conInfo1_2 = Orec.Fields.Item("U_AVA_SSIC").Value;

                                if (System.IO.Directory.Exists(conInfo1_1) == false)
                                {
                                    objSBOAPI.SBO_Appln.MessageBox("File Path is Missing");
                                    return;
                                }
                            }
                            else
                            {
                                objSBOAPI.SBO_Appln.MessageBox("Server and Folder Details are Missing. Please provide it in Company Details");
                                return;
                            }
                        }
                        else
                        {
                            objSBOAPI.SBO_Appln.MessageBox("Server and Folder Details are Missing. Please provide it in Company Details");
                            return;
                        }
                        
                        str2 = conInfo1_1;

                        while (!objRecordSet.EoF)
                        {
                            text = Convert.ToString(objRecordSet.Fields.Item("QRData").Value);

                            str3 += text + "##";

                            objRecordSet.MoveNext();
                        }

                        str3 = str3.Substring(0, str3.Length - 2);

                        Barcode128 barcode128 = new Barcode128();
                        QRCodeEncoder qrCodeEncoder = new QRCodeEncoder();
                        qrCodeEncoder.QRCodeEncodeMode = ((QRCodeEncoder.ENCODE_MODE)4);//((QRCodeEncoder.ENCODE_MODE)2);
                        qrCodeEncoder.QRCodeScale = (4);//(2);
                        qrCodeEncoder.QRCodeVersion = (10);//(5);
                        qrCodeEncoder.QRCodeErrorCorrect = ((QRCodeEncoder.ERROR_CORRECTION)0);
                        new Bitmap((System.Drawing.Image)qrCodeEncoder.Encode(str3), 120, 120).Save(str2 + "PickList_" + AbsEntry + "_QR.jpg");
                        HanaConnection connection = new HanaConnection(conInfo1_2);
                        connection.Open();
                        path = str2 + "PickList_" + AbsEntry + "_QR.jpg";
                        string withoutExtension = Path.GetFileNameWithoutExtension(path);
                        byte[] numArray = File.ReadAllBytes(path);

                        HanaCommand updateCmd = new HanaCommand("UPDATE \"OPKL\" SET \"U_AV_IMG\" = ? WHERE \"AbsEntry\" = ?", connection);

                        HanaParameter picParameter = new HanaParameter();
                        picParameter.HanaDbType = HanaDbType.AlphaNum;
                        picParameter.Value = numArray;
                        updateCmd.Parameters.Add(picParameter);

                        HanaParameter absEntryParameter = new HanaParameter();
                        absEntryParameter.HanaDbType = HanaDbType.Integer;
                        absEntryParameter.Value = AbsEntry;
                        updateCmd.Parameters.Add(absEntryParameter);

                        int recordsAffected = updateCmd.ExecuteNonQuery();

                        //new HanaCommand("update \"OPKL\" set \"U_AV_IMG\" = :pic where \"AbsEntry\" = :AbsEntry", connection)
                        //{
                        //    Parameters = { { ":pic", (object)numArray }, { ":AbsEntry", (object)AbsEntry } }
                        //}
                        //.ExecuteNonQuery();
                        objSBOAPI.SBO_Appln.StatusBar.SetText("QR Code Created Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        objSBOAPI.SBO_Appln.MessageBox("QR Code Created Successfully");
                        objSBOAPI.Query_Execute("Update OPKL Set \"U_AV_QPATH\" = '" + path + "' Where \"AbsEntry\" = '" + AbsEntry + "'");
                    }
                    else
                    {
                        objSBOAPI.SBO_Appln.StatusBar.SetText("No record found in PickListQRCode", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                    Marshal.ReleaseComObject(objRecordSet);
                }
            }
            catch (Exception ex)
            {
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                SAPbouiCOM.Menus oMenus = objSBOAPI.SBO_Appln.Menus;
                if (oMenus.Item("1304").Enabled == true)
                {
                    objSBOAPI.SBO_Appln.Menus.Item("1304").Activate();
                }
            }
            //Thiru
        }
        #endregion
    }
}
