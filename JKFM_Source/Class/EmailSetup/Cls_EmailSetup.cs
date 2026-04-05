using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace JKFM_Source
{
    class Cls_EmailSetup
    {
        #region Declaration
        SAPbouiCOM.Form objform;
        SAPbouiCOM.CheckBox oChk2;
        private ClsSBO objSBOAPI;
        #endregion

        #region Constructor
        public Cls_EmailSetup(ClsSBO objSBO)
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
                            switch (pval.ItemUID)
                            {
                                case "Item_58":
                                    if (Validate_EmailCredentials() == false)
                                    {
                                        bubbleevent = false;
                                    }
                                    break;

                                case "Item_59":
                                    if (Validate_ServerCredentials() == false)
                                    {
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
                            switch (pval.ItemUID)
                            {
                                case "1":
                                    load_data();
                                    break;

                                case "Item_58":
                                    Test_Mail_Service();
                                    break;

                                case "Item_59":
                                    objform.Freeze(true);
                                    oChk2 = objform.Items.Item("Item_23").Specific;
                                    if(oChk2.Checked == true)
                                    {
                                        try
                                        {
                                            NetworkCredential cred = new NetworkCredential(objform.Items.Item("Item_53").Specific.value, objform.Items.Item("Item_56").Specific.value);
                                            //NetworkConnection con = new NetworkConnection(tbUNCPath.Text, cred);

                                            CredentialCache theNetCache = new CredentialCache();
                                            theNetCache.Add(new Uri(@"\\computer"), "Basic", cred);

                                            var dirs = Directory.GetFiles(objform.Items.Item("Item_31").Specific.value);

                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Server Credentials Succeeded", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                        }
                                        catch (Exception ex)
                                        {
                                            objform.Freeze(false);
                                            objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        }
                                    }
                                    else
                                    {
                                        UNCAccessWithCredentials unc = new UNCAccessWithCredentials();
                                        if (unc.NetUseWithCredentials(objform.Items.Item("Item_31").Specific.value, objform.Items.Item("Item_53").Specific.value, objform.Items.Item("Item_39").Specific.value, objform.Items.Item("Item_56").Specific.value))
                                        {
                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Server Credentials Succeeded", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                        }
                                        else
                                        {
                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Layout Path - Server Credentials Failed", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        }
                                    }
                                   
                                    objform.Freeze(false);
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
                switch (pval.MenuUID)
                {
                    case "AV_EMSTM":
                        objform = objSBOAPI.LoadForm("EmailSetup.xml", "AV_EMSTF");
                        objform.EnableMenu("1281", false);
                        objform.EnableMenu("1282", false);
                        Form_Mode();
                        break;
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

            }
        }
        #endregion

        #region Form Mode
        public void Form_Mode()
        {
            try
            {
                objform.Freeze(true);
                SAPbobsCOM.Recordset orec;
                orec = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string str;
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    str = "Select \"U_AV_SMTPS\",\"U_AV_EMID\",\"U_AV_EPswd\",\"U_AV_ToEMID\",\"U_AV_Domain\",\"U_AV_UserName\",\"U_AV_Password\",\"U_AV_LPATH\",\"U_AV_Port\",\"U_AV_Essl\",\"U_AV_WorkStation\" from \"OADM\"";
                }
                else
                {
                    str = "Select U_AV_SMTPS,U_AV_EMID,U_AV_EPswd,U_AV_ToEMID,U_AV_Domain,U_AV_UserName,U_AV_Password,U_AV_LPATH,U_AV_Port,U_AV_Essl,U_AV_WorkStation from OADM";
                }
                orec.DoQuery(str);
                if (orec.RecordCount > 0)
                {
                    SAPbouiCOM.CheckBox oChk = objform.Items.Item("Item_65").Specific;
                    oChk2 = objform.Items.Item("Item_23").Specific;
                    if (orec.Fields.Item("U_AV_SMTPS").Value != "" & orec.Fields.Item("U_AV_EMID").Value != "")
                    {
                        objform.Freeze(true);
                        objform.Items.Item("Item_58").Visible = true;
                        objform.Items.Item("Item_59").Visible = true;
                        objform.Items.Item("Item_4").Specific.value = orec.Fields.Item("U_AV_SMTPS").Value;
                        objform.Items.Item("Item_6").Specific.value = orec.Fields.Item("U_AV_EMID").Value;
                        objform.Items.Item("Item_39").Specific.value = orec.Fields.Item("U_AV_Domain").Value;
                        objform.Items.Item("Item_53").Specific.value = orec.Fields.Item("U_AV_UserName").Value;
                        objform.Items.Item("Item_56").Specific.value = Cls_Encrypt_Helper.Decrypt(orec.Fields.Item("U_AV_Password").Value);
                        objform.Items.Item("Item_31").Specific.value = orec.Fields.Item("U_AV_LPATH").Value;
                        objform.Items.Item("Item_35").Specific.value = orec.Fields.Item("U_AV_Port").Value;
                        objform.Items.Item("Item_3").Specific.value = Cls_Encrypt_Helper.Decrypt(orec.Fields.Item("U_AV_EPswd").Value);
                        objform.Items.Item("Item_7").Specific.value = orec.Fields.Item("U_AV_ToEMID").Value;
                        if (orec.Fields.Item("U_AV_Essl").Value == "Y")
                        {
                            oChk.Checked = true;
                        }
                        else
                        {
                            oChk.Checked = false;
                        }

                        if (orec.Fields.Item("U_AV_WorkStation").Value == "Y")
                        {
                            oChk2.Checked = true;
                        }
                        else
                        {
                            oChk2.Checked = false;
                        }
                        objform.Items.Item("Item_4").Click(SAPbouiCOM.BoCellClickType.ct_Linked);
                        objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        objform.Freeze(false);
                    }
                    else
                    {
                        objform.Freeze(true);
                        objform.Items.Item("Item_58").Visible = false;
                        objform.Items.Item("Item_59").Visible = false;
                        objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                        objform.Items.Item("Item_4").Specific.value = "";
                        objform.Items.Item("Item_6").Specific.value = "";
                        objform.Items.Item("Item_39").Specific.value = "";
                        objform.Items.Item("Item_53").Specific.value = "";
                        objform.Items.Item("Item_56").Specific.value = "";
                        objform.Items.Item("Item_31").Specific.value = "";
                        objform.Items.Item("Item_35").Specific.value = "";
                        objform.Items.Item("Item_3").Specific.value = "";
                        objform.Items.Item("Item_7").Specific.value = "";
                        oChk.Checked = false;
                        objform.Items.Item("Item_4").Click(SAPbouiCOM.BoCellClickType.ct_Linked);
                        objform.Freeze(false);
                    }
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

        #region Data Adding
        public void load_data()
        {
            try
            {
                if (objform.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    objform.Freeze(true);
                    SAPbobsCOM.Recordset orec;
                    orec = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string EnableSsl = "";
                    string EnableWS = "";
                    SAPbouiCOM.CheckBox oChk = objform.Items.Item("Item_65").Specific;
                    oChk2 = objform.Items.Item("Item_23").Specific;
                    if (oChk.Checked == true)
                    {
                        EnableSsl = "Y";
                    }
                    else
                    {
                        EnableSsl = "N";
                    }

                    if (oChk2.Checked == true)
                    {
                        EnableWS = "Y";
                    }
                    else
                    {
                        EnableWS = "N";
                    }

                    string str = "";

                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        str = "Update \"OADM\" set \"U_AV_SMTPS\" = '" + objform.Items.Item("Item_4").Specific.value + "',\"U_AV_EMID\"='" + objform.Items.Item("Item_6").Specific.value + "',\"U_AV_Domain\"='" + objform.Items.Item("Item_39").Specific.value + "',\"U_AV_UserName\"='" + objform.Items.Item("Item_53").Specific.value + "',\"U_AV_Password\"='" + Cls_Encrypt_Helper.Encrypt(objform.Items.Item("Item_56").Specific.value) + "',\"U_AV_LPATH\"='" + objform.Items.Item("Item_31").Specific.value + "',\"U_AV_EPswd\"='" + Cls_Encrypt_Helper.Encrypt(objform.Items.Item("Item_3").Specific.value) + "',\"U_AV_ToEMID\"='" + objform.Items.Item("Item_7").Specific.value + "',\"U_AV_Port\"='" + objform.Items.Item("Item_35").Specific.value + "',\"U_AV_Essl\"='" + EnableSsl + "',\"U_AV_WorkStation\" = '" + EnableWS + "' where \"CompnyName\" = '" + objSBOAPI.oCompany.CompanyName + "'";
                    }
                    else
                    {
                        str = "Update OADM set U_AV_SMTPS = '" + objform.Items.Item("Item_4").Specific.value + "',U_AV_EMID='" + objform.Items.Item("Item_6").Specific.value + "'," + "U_AV_Domain='" + objform.Items.Item("Item_39").Specific.value + "',U_AV_UserName='" + objform.Items.Item("Item_53").Specific.value + "'," + "U_AV_Password='" + Cls_Encrypt_Helper.Encrypt(objform.Items.Item("Item_56").Specific.value) + "',U_AV_LPATH='" + objform.Items.Item("Item_31").Specific.value + "'," + "U_AV_EPswd='" + Cls_Encrypt_Helper.Encrypt(objform.Items.Item("Item_3").Specific.value) + "',U_AV_ToEMID='" + objform.Items.Item("Item_7").Specific.value + "',U_AV_Port='" + objform.Items.Item("Item_35").Specific.value + "',U_AV_Essl='" + EnableSsl + "',U_AV_WorkStation = '" + EnableWS + "'" + "where CompnyName = '" + objSBOAPI.oCompany.CompanyName + "'";
                    }
                    orec.DoQuery(str);
                    objform.Freeze(false);
                    if (objform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Datas are Added Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        objform.Close();
                    }
                    else
                    {
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Datas are Updated Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        //objform.Close()
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

        #region Validate Email Credentials
        public bool Validate_EmailCredentials()
        {
            try
            {
                if (objform.Items.Item("Item_4").Specific.value == "")
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("SMTP Server is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else if (objform.Items.Item("Item_35").Specific.value == "")
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Port is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else if (objform.Items.Item("Item_6").Specific.value == "")
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Email Id is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else if (objform.Items.Item("Item_3").Specific.value == "")
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Email Password is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else if (objform.Items.Item("Item_7").Specific.value == "")
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("ToMail is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
        }

        #endregion

        #region Validate Server Credentials
        public bool Validate_ServerCredentials()
        {
            try
            {
                oChk2 = objform.Items.Item("Item_23").Specific;
                if (objform.Items.Item("Item_39").Specific.value == "")
                {
                    if(oChk2.Checked == false)
                    {
                        objSBOAPI.SBO_Appln.StatusBar.SetText("Domain Name is Missing / Enable WorkStation", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                }
                else if (objform.Items.Item("Item_53").Specific.value == "")
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("UserName is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else if (objform.Items.Item("Item_56").Specific.value == "")
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Password is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else if (objform.Items.Item("Item_31").Specific.value == "")
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Layout Path is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
        }

        #endregion

        #region Test Mail Service
        public void Test_Mail_Service()
        {
            try
            {
                objform.Freeze(true);
                SAPbouiCOM.CheckBox oChk = objform.Items.Item("Item_65").Specific;
                string MailId = objform.Items.Item("Item_6").Specific.value;
                string ToMailId = objform.Items.Item("Item_7").Specific.value;
                string SMTPServer = objform.Items.Item("Item_4").Specific.value;
                string Domain = objform.Items.Item("Item_39").Specific.value;
                string Password = objform.Items.Item("Item_3").Specific.value;
                Int32 Port = Convert.ToInt32(objform.Items.Item("Item_35").Specific.value);
                MailMessage message = new MailMessage(MailId.Trim(), ToMailId.Trim(), "Add-on Test Email", "Test Email");
                SmtpClient emailClient = new SmtpClient(SMTPServer, Port);
                if (oChk.Checked == true)
                {
                    emailClient.EnableSsl = true;
                }
                else
                {
                    emailClient.EnableSsl = false;
                }

                NetworkCredential SMTPUserInfo = new NetworkCredential(MailId, Password);
                emailClient.UseDefaultCredentials = false;
                emailClient.Credentials = SMTPUserInfo;
                ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(customCertValidation);
                emailClient.Send(message);

                objSBOAPI.SBO_Appln.StatusBar.SetText("Email Sent Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private static bool customCertValidation(object sender, X509Certificate cert, X509Chain chain, SslPolicyErrors errors)
        {
            return true;
        }

        #endregion
    }
}
