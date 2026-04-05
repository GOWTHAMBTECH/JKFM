using Sap.Data.Hana;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JKFM_Source
{
    class Cls_LicenseUpload
    {
        #region Declaration
        SAPbouiCOM.Form objform;
        ClsSBO objSBOAPI;
        string v_sql = "";
        string UID = "";

        byte[] _certPubicKeyData;
        #endregion        

        #region Constructor
        public Cls_LicenseUpload(ClsSBO objSBO)
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
                            if (pval.ItemUID == "Btn_UPL")
                            {
                                Cls_LicenseVerify obj_ClsVerify = new Cls_LicenseVerify(objSBOAPI);
                                if (objform.Items.Item("Et_SName").Specific.Value == "")
                                {
                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Server Name is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    bubbleevent = false;
                                }
                                else if (objform.Items.Item("Et_User").Specific.Value == "")
                                {
                                    objSBOAPI.SBO_Appln.StatusBar.SetText("UserName is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    bubbleevent = false;
                                }
                                else if (objform.Items.Item("Et_Pswd").Specific.Value == "")
                                {
                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Password is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    bubbleevent = false;
                                }
                                else if (objform.Items.Item("Et_File").Specific.Value == "")
                                {
                                    objSBOAPI.SBO_Appln.StatusBar.SetText("License File Path is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    bubbleevent = false;
                                }
                                else if (objform.Items.Item("Et_Pswd").Specific.Value != "" && objform.Items.Item("Et_User").Specific.Value != "")
                                {
                                    if(objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                    {
                                        v_sql = string.Format("Server = {0}; UserID = {1}; Password = {2}", objform.Items.Item("Et_SName").Specific.Value, objform.Items.Item("Et_User").Specific.Value, objform.Items.Item("Et_Pswd").Specific.Value);
                                        HanaConnection myConnection = new HanaConnection(v_sql);
                                        try
                                        {
                                            myConnection.Open();
                                        }
                                        catch (HanaException ex)
                                        {
                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Server Credentials Failed", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                            bubbleevent = false;
                                            return;
                                        }
                                        finally
                                        {
                                            myConnection.Dispose();
                                        }
                                    }
                                    else
                                    {
                                        v_sql = string.Format("Data Source = {0}; Initial Catalog = {1}; User ID = {2}; Password = {3}", objform.Items.Item("Et_SName").Specific.Value, objSBOAPI.oCompany.CompanyDB, objform.Items.Item("Et_User").Specific.Value, objform.Items.Item("Et_Pswd").Specific.Value);
                                        SqlConnection myConnection = new SqlConnection(v_sql);
                                        try
                                        {
                                            myConnection.Open();
                                        }
                                        catch (SqlException)
                                        {
                                            objSBOAPI.SBO_Appln.StatusBar.SetText("Server Credentials Failed", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                            bubbleevent = false;
                                            return;
                                        }
                                        finally
                                        {
                                            myConnection.Dispose();
                                        }
                                    }
                                }
                                if (!Validate_License())
                                {
                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Not a valid License File", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    bubbleevent = false;
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
                            if (pval.ItemUID == "Btn_Cancel")
                            {
                                objform.Close();
                                System.Windows.Forms.Application.Exit();
                            }
                            else if (pval.ItemUID == "Btn_Browse")
                            {
                                objSBOAPI.Browse("Open", objform.UniqueID, "license", "Et_File");
                            }
                            else if (pval.ItemUID == "Btn_UPL")
                            {
                                Upload_License();
                            }
                            break;

                        case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                            if(pval.ItemUID== "Et_SName")
                            {
                                objform.Items.Item("Et_SName").Enabled = true;
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
                    if (pval.MenuUID == "")
                    {
                        
                    }
                    else if (pval.MenuUID == "1281")
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

        #region Upload License

        public void Upload_License()
        {
            try
            {
                string ServerName = Cls_Encrypt_Helper.Encrypt(objform.Items.Item("Et_SName").Specific.Value);
                string UserName = Cls_Encrypt_Helper.Encrypt(objform.Items.Item("Et_User").Specific.Value);
                string Password = Cls_Encrypt_Helper.Encrypt(objform.Items.Item("Et_Pswd").Specific.Value);
                string LicenseKey = File.ReadAllText(objform.Items.Item("Et_File").Specific.Value);
                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    int Table = Convert.ToInt32(objSBOAPI.Query_Execute("SELECT Count(*) FROM TABLES where SCHEMA_NAME = '" + objSBOAPI.oCompany.CompanyDB + "' and TABLE_NAME = 'AVA_AddonLicense'"));
                    if (Table == 0)
                    {
                        string Str = objSBOAPI.Reading_Addon_Query_NotePad("LicenseTable_HANA.txt");
                        Str = Str.Replace("[%1]","" + objSBOAPI.oCompany.CompanyDB +"");
                        using (HanaConnection myConnection = new HanaConnection(v_sql))
                        {
                            myConnection.Open();
                            using (HanaCommand cmd = new HanaCommand(Str, myConnection))
                            {
                                cmd.ExecuteNonQuery();
                            }
                            myConnection.Close();
                        }
                    }
                    objSBOAPI.Query_Execute("Delete from \"AVA_AddonLicense\" Where \"LicenseUID\"='" + UID + "'");
                    objSBOAPI.Query_Execute("Insert into \"AVA_AddonLicense\" (\"LicenseUID\",\"ServerName\",\"UserName\",\"Password\",\"LicenseKey\")Values('" + UID+"','"+ ServerName+"','"+UserName+"','"+ Password+"','"+ LicenseKey+"')");
                }
                else
                {
                    int Table = Convert.ToInt32(objSBOAPI.Query_Execute("SELECT Count(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_CATALOG = '" + objSBOAPI.oCompany.CompanyDB + "' and TABLE_NAME = 'AVA_AddonLicense'"));
                    if (Table == 0)
                    {
                        string Str = objSBOAPI.Reading_Addon_Query_NotePad("LicenseTable_SQL.txt");
                        using (SqlConnection myConnection = new SqlConnection(v_sql))
                        {
                            myConnection.Open();
                            using (SqlCommand cmd = new SqlCommand(Str, myConnection))
                            {
                                cmd.ExecuteNonQuery();
                            }
                            myConnection.Close();
                        }
                    }
                    objSBOAPI.Query_Execute("Delete from AVA_AddonLicense Where LicenseUID='" + UID + "'");
                    objSBOAPI.Query_Execute("Insert into AVA_AddonLicense (LicenseUID,ServerName,UserName,Password,LicenseKey)Values('"+ UID+ "','" + ServerName + "','" + UserName + "','" + Password + "','"+ LicenseKey+"')");
                }

                objSBOAPI.objMain.HANA_ServerName = objform.Items.Item("Et_SName").Specific.Value;
                objSBOAPI.objMain.HANA_UserID = objform.Items.Item("Et_User").Specific.Value;
                objSBOAPI.objMain.HANA_Pwd = objform.Items.Item("Et_Pswd").Specific.Value;

                objSBOAPI.SBO_Appln.MessageBox("License Upload Successfully...Restart Addon to Proceed");

                objform.Close();
                System.Windows.Forms.Application.Exit();
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Validate the License
        public bool Validate_License()
        {
            try
            {
                
                Cls_MyLicense _lic = null;
                string _msg = string.Empty;
                LicenseStatus _status = LicenseStatus.UNDEFINED;

                string CertKey = "QmFnIEF0dHJpYnV0ZXMNCiAgICBsb2NhbEtleUlEOiBFOSBEMiAyRSBERCA3QiAwNyBGOSBGQyAwMCA0QiA5NiA3MSBFMCBGOSBGOCA0QiA1NCA0NyAxQyBBOSANCnN1YmplY3Q9L0NOPUFuYW5kIEtyaXNoYW1vb3J0aHkvTz1BdmFuaWtvIFRlY2hub2xvZ2llcy9PVT1DRU8vZW1haWxBZGRyZXNzPWFuYW5kLmtyaXNobmFtb29ydGh5QGF2YW5pa28uY29tL0M9SU4NCmlzc3Vlcj0vQ049QW5hbmQgS3Jpc2hhbW9vcnRoeS9PPUF2YW5pa28gVGVjaG5vbG9naWVzL09VPUNFTy9lbWFpbEFkZHJlc3M9YW5hbmQua3Jpc2huYW1vb3J0aHlAYXZhbmlrby5jb20vQz1JTg0KLS0tLS1CRUdJTiBDRVJUSUZJQ0FURS0tLS0tDQpNSUlEdURDQ0FxQ2dBd0lCQWdJS3Z6bDVhZ0JkZVhUYThqQU5CZ2txaGtpRzl3MEJBUXNGQURDQmlURWNNQm9HDQpBMVVFQXhNVFFXNWhibVFnUzNKcGMyaGhiVzl2Y25Sb2VURWRNQnNHQTFVRUNoTVVRWFpoYm1scmJ5QlVaV05vDQpibTlzYjJkcFpYTXhEREFLQmdOVkJBc1RBME5GVHpFdk1DMEdDU3FHU0liM0RRRUpBUllnWVc1aGJtUXVhM0pwDQpjMmh1WVcxdmIzSjBhSGxBWVhaaGJtbHJieTVqYjIweEN6QUpCZ05WQkFZVEFrbE9NQjRYRFRJd01EUXhNakUxDQpNekl5TmxvWERUSTFNRFF4TWpFMU16SXlObG93Z1lreEhEQWFCZ05WQkFNVEUwRnVZVzVrSUV0eWFYTm9ZVzF2DQpiM0owYUhreEhUQWJCZ05WQkFvVEZFRjJZVzVwYTI4Z1ZHVmphRzV2Ykc5bmFXVnpNUXd3Q2dZRFZRUUxFd05EDQpSVTh4THpBdEJna3Foa2lHOXcwQkNRRVdJR0Z1WVc1a0xtdHlhWE5vYm1GdGIyOXlkR2g1UUdGMllXNXBhMjh1DQpZMjl0TVFzd0NRWURWUVFHRXdKSlRqQ0NBU0l3RFFZSktvWklodmNOQVFFQkJRQURnZ0VQQURDQ0FRb0NnZ0VCDQpBTWJvenFkU3FXb09HWWRYVWlsWDRuZFJZMmpQeUJ4eHZPNXVQT2tBdWJGQkVld2ZGUWJYb2U4NHRMN3FIOW82DQpEMVlIa3Y2WHFJMUtrZzZwRjJkaUk2VTRkeU9IWSs0N3g2VHBwcFNkYVpNVzhvTkZpU01pWVQ3MW5jZHF0NVEyDQoxQ09GSkhoQXoxSGErcnNtUERkQnU1V2VpMmgxZStDRFdBYWNCREF6WEp0OFIrZVhMMy8rdnFaVHRpRTdBUGIrDQpJQ040NzZQUXpoSnNTNDFvRElrczFIYmxsYzFHRkVzNUVxUlpXVUhlK1NFSTN5VlF0ZS8wMmVEcVlDWTdiMEFODQp2N2ZCWXpZajJiK0RDa1FZTnJJbEFjOGl3cXpsUTBYQXhyYmREd3VPb1BIdWNxWHF0TzJMWjNCdjZhVjhNMHlFDQovNDljbHpXMXVoRjFDWWZjbkF3aGlWY0NBd0VBQWFNZ01CNHdEd1lKS29aSWh2Y3ZBUUVLQkFJRkFEQUxCZ05WDQpIUThFQkFNQ0I0QXdEUVlKS29aSWh2Y05BUUVMQlFBRGdnRUJBRVA4VDdCYXhGeVVDYm50OFkvUGJLTUN2aXYyDQpGUG9OM0d3T2tFZ3A3Zy9FQVIxd2prbERXdGFxVnJwTzhnVFJ6UWk4ZXk0SGVGeXlqWG11L0N1WnpSMmN4RW5BDQpEQWg3azRRYkpvZkl4STlkOUdaY1FUR21FT25TNXJISnFNNGZ4dXJLemU3b1BJZTFIay9yRUZMRFpPb3JWaFczDQozWnBNOTBRSlAwelZRNXd3SWNCNUhYditYMENnQ0cvb3lUZWNmcG9LNzJaZHJJK053Rlc2NGFyYWtudUp5KzBSDQptS0ZVZ0ZEMjRXYkpnRklGV0RaQzNVWEI3RnYyWm82ZlNlaC9YNjJZcnlYeTEvclV0SDc4RkJ0dFVKUExRamUwDQpCWi9DSVNQdjhTWVdsRFlaaDJoV2h5L05TN0Z2MW9zWVk5MDUrYWFseC9tNm81R0xnYVZXRDF2MmhjUT0NCi0tLS0tRU5EIENFUlRJRklDQVRFLS0tLS0NCg==";
                byte[] data = System.Convert.FromBase64String(CertKey);

                _certPubicKeyData = data;

                UID = Cls_LicenseHandler.GenerateUID(objSBOAPI.AddonName, objSBOAPI.oCompany.Server, objSBOAPI.oCompany.CompanyDB);

                if (File.Exists(objform.Items.Item("Et_File").Specific.Value))
                {

                    string str = File.ReadAllText(objform.Items.Item("Et_File").Specific.Value);

                    _lic = (Cls_MyLicense)Cls_LicenseHandler.ParseLicenseFromBASE64String(
                        typeof(Cls_MyLicense),
                        str,
                        _certPubicKeyData,
                        out _status,
                        out _msg);
                }
                else
                {
                    _status = LicenseStatus.INVALID;
                }

                switch (_status)
                {
                    case LicenseStatus.VALID:
                        return true;

                    default:
                            return false;
                }
            }
            catch (Exception ex)
            {
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
        }
        #endregion

    }
}
