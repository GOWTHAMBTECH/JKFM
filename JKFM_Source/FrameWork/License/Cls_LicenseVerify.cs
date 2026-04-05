using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JKFM_Source
{    
    public class Cls_LicenseVerify
    {

        #region Declaration
        static ClsSBO objSBOAPI;

        byte[] _certPubicKeyData;
        string LicenseKey = "";

        string UserName = "";
        string Password = "";
        string Server = "";
        #endregion

        #region Constructor
        public Cls_LicenseVerify(ClsSBO objSBO)
        {
            objSBOAPI = objSBO;
        }
        #endregion

        #region Validate the License
        public bool Validate_License()
        {
            try
            {
                string UID = "";
                Cls_MyLicense _lic = null;
                string _msg = string.Empty;
                LicenseStatus _status = LicenseStatus.UNDEFINED;

                UID = Cls_LicenseHandler.GenerateUID(objSBOAPI.AddonName, objSBOAPI.oCompany.Server, objSBOAPI.oCompany.CompanyDB);

                string CertKey = "QmFnIEF0dHJpYnV0ZXMNCiAgICBsb2NhbEtleUlEOiBFOSBEMiAyRSBERCA3QiAwNyBGOSBGQyAwMCA0QiA5NiA3MSBFMCBGOSBGOCA0QiA1NCA0NyAxQyBBOSANCnN1YmplY3Q9L0NOPUFuYW5kIEtyaXNoYW1vb3J0aHkvTz1BdmFuaWtvIFRlY2hub2xvZ2llcy9PVT1DRU8vZW1haWxBZGRyZXNzPWFuYW5kLmtyaXNobmFtb29ydGh5QGF2YW5pa28uY29tL0M9SU4NCmlzc3Vlcj0vQ049QW5hbmQgS3Jpc2hhbW9vcnRoeS9PPUF2YW5pa28gVGVjaG5vbG9naWVzL09VPUNFTy9lbWFpbEFkZHJlc3M9YW5hbmQua3Jpc2huYW1vb3J0aHlAYXZhbmlrby5jb20vQz1JTg0KLS0tLS1CRUdJTiBDRVJUSUZJQ0FURS0tLS0tDQpNSUlEdURDQ0FxQ2dBd0lCQWdJS3Z6bDVhZ0JkZVhUYThqQU5CZ2txaGtpRzl3MEJBUXNGQURDQmlURWNNQm9HDQpBMVVFQXhNVFFXNWhibVFnUzNKcGMyaGhiVzl2Y25Sb2VURWRNQnNHQTFVRUNoTVVRWFpoYm1scmJ5QlVaV05vDQpibTlzYjJkcFpYTXhEREFLQmdOVkJBc1RBME5GVHpFdk1DMEdDU3FHU0liM0RRRUpBUllnWVc1aGJtUXVhM0pwDQpjMmh1WVcxdmIzSjBhSGxBWVhaaGJtbHJieTVqYjIweEN6QUpCZ05WQkFZVEFrbE9NQjRYRFRJd01EUXhNakUxDQpNekl5TmxvWERUSTFNRFF4TWpFMU16SXlObG93Z1lreEhEQWFCZ05WQkFNVEUwRnVZVzVrSUV0eWFYTm9ZVzF2DQpiM0owYUhreEhUQWJCZ05WQkFvVEZFRjJZVzVwYTI4Z1ZHVmphRzV2Ykc5bmFXVnpNUXd3Q2dZRFZRUUxFd05EDQpSVTh4THpBdEJna3Foa2lHOXcwQkNRRVdJR0Z1WVc1a0xtdHlhWE5vYm1GdGIyOXlkR2g1UUdGMllXNXBhMjh1DQpZMjl0TVFzd0NRWURWUVFHRXdKSlRqQ0NBU0l3RFFZSktvWklodmNOQVFFQkJRQURnZ0VQQURDQ0FRb0NnZ0VCDQpBTWJvenFkU3FXb09HWWRYVWlsWDRuZFJZMmpQeUJ4eHZPNXVQT2tBdWJGQkVld2ZGUWJYb2U4NHRMN3FIOW82DQpEMVlIa3Y2WHFJMUtrZzZwRjJkaUk2VTRkeU9IWSs0N3g2VHBwcFNkYVpNVzhvTkZpU01pWVQ3MW5jZHF0NVEyDQoxQ09GSkhoQXoxSGErcnNtUERkQnU1V2VpMmgxZStDRFdBYWNCREF6WEp0OFIrZVhMMy8rdnFaVHRpRTdBUGIrDQpJQ040NzZQUXpoSnNTNDFvRElrczFIYmxsYzFHRkVzNUVxUlpXVUhlK1NFSTN5VlF0ZS8wMmVEcVlDWTdiMEFODQp2N2ZCWXpZajJiK0RDa1FZTnJJbEFjOGl3cXpsUTBYQXhyYmREd3VPb1BIdWNxWHF0TzJMWjNCdjZhVjhNMHlFDQovNDljbHpXMXVoRjFDWWZjbkF3aGlWY0NBd0VBQWFNZ01CNHdEd1lKS29aSWh2Y3ZBUUVLQkFJRkFEQUxCZ05WDQpIUThFQkFNQ0I0QXdEUVlKS29aSWh2Y05BUUVMQlFBRGdnRUJBRVA4VDdCYXhGeVVDYm50OFkvUGJLTUN2aXYyDQpGUG9OM0d3T2tFZ3A3Zy9FQVIxd2prbERXdGFxVnJwTzhnVFJ6UWk4ZXk0SGVGeXlqWG11L0N1WnpSMmN4RW5BDQpEQWg3azRRYkpvZkl4STlkOUdaY1FUR21FT25TNXJISnFNNGZ4dXJLemU3b1BJZTFIay9yRUZMRFpPb3JWaFczDQozWnBNOTBRSlAwelZRNXd3SWNCNUhYditYMENnQ0cvb3lUZWNmcG9LNzJaZHJJK053Rlc2NGFyYWtudUp5KzBSDQptS0ZVZ0ZEMjRXYkpnRklGV0RaQzNVWEI3RnYyWm82ZlNlaC9YNjJZcnlYeTEvclV0SDc4RkJ0dFVKUExRamUwDQpCWi9DSVNQdjhTWVdsRFlaaDJoV2h5L05TN0Z2MW9zWVk5MDUrYWFseC9tNm81R0xnYVZXRDF2MmhjUT0NCi0tLS0tRU5EIENFUlRJRklDQVRFLS0tLS0NCg==";
                byte[] data = System.Convert.FromBase64String(CertKey);

                _certPubicKeyData = data;

                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    int Table = Convert.ToInt32(objSBOAPI.Query_Execute("SELECT Count(*) FROM TABLES where SCHEMA_NAME = '" + objSBOAPI.oCompany.CompanyDB + "' and TABLE_NAME = 'AVA_AddonLicense'"));
                    if (Table != 0)
                    {
                        UserName = objSBOAPI.Query_Execute("Select \"UserName\" from \"AVA_AddonLicense\" Where \"LicenseUID\"='" + UID + "'");
                        Password = objSBOAPI.Query_Execute("Select \"Password\" from \"AVA_AddonLicense\" Where \"LicenseUID\"='" + UID + "'");
                        Server = objSBOAPI.Query_Execute("Select \"ServerName\" from \"AVA_AddonLicense\" Where \"LicenseUID\"='" + UID + "'");

                        LicenseKey = objSBOAPI.Query_Execute("Select \"LicenseKey\" from \"AVA_AddonLicense\" Where \"LicenseUID\"='" + UID + "'");
                        if (UserName != string.Empty && Password != string.Empty && Server != string.Empty && LicenseKey != string.Empty )
                        {
                            UserName = Cls_Encrypt_Helper.Decrypt(UserName);
                            Password = Cls_Encrypt_Helper.Decrypt(Password);
                            Server = Cls_Encrypt_Helper.Decrypt(Server);

                            _lic = (Cls_MyLicense)Cls_LicenseHandler.ParseLicenseFromBASE64String(
                                                typeof(Cls_MyLicense),
                                                LicenseKey,
                                                _certPubicKeyData,
                                                out _status,
                                                out _msg);
                        }
                        else
                        {
                            _status = LicenseStatus.INVALID;
                            _msg = objSBOAPI.AddonName + " Add-on License is not activated";
                        }
                    }
                    else
                    {
                        _status = LicenseStatus.INVALID;
                        _msg = objSBOAPI.AddonName + " Add-on License is not activated";
                    }
                }
                else
                {
                    int Table = Convert.ToInt32(objSBOAPI.Query_Execute("SELECT Count(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_CATALOG = '" + objSBOAPI.oCompany.CompanyDB + "' and TABLE_NAME = 'AVA_AddonLicense'"));
                    if (Table != 0)
                    {
                        UserName = objSBOAPI.Query_Execute("Select UserName from AVA_AddonLicense Where LicenseUID='" + UID + "'");
                        Password = objSBOAPI.Query_Execute("Select Password from AVA_AddonLicense Where LicenseUID='" + UID + "'");
                        Server = objSBOAPI.Query_Execute("Select ServerName from AVA_AddonLicense Where LicenseUID='" + UID + "'");

                        LicenseKey = objSBOAPI.Query_Execute("Select LicenseKey from AVA_AddonLicense Where LicenseUID='" + UID + "'");
                        if (UserName != string.Empty && Password != string.Empty && Server != string.Empty && LicenseKey != string.Empty )
                        {
                            UserName = Cls_Encrypt_Helper.Decrypt(UserName);
                            Password = Cls_Encrypt_Helper.Decrypt(Password);
                            Server = Cls_Encrypt_Helper.Decrypt(Server);

                            _lic = (Cls_MyLicense)Cls_LicenseHandler.ParseLicenseFromBASE64String(
                                                typeof(Cls_MyLicense),
                                                LicenseKey,
                                                _certPubicKeyData,
                                                out _status,
                                                out _msg);
                        }
                        else
                        {
                            _status = LicenseStatus.INVALID;
                            _msg = objSBOAPI.AddonName + " Add-on License is not activated";
                        }
                    }
                    else
                    {
                        _status = LicenseStatus.INVALID;
                        _msg = objSBOAPI.AddonName + " Add-on License is not activated";
                    }

                }
                

                switch (_status)
                {
                    case LicenseStatus.VALID:

                        objSBOAPI.objMain.HANA_ServerName = Server;
                        objSBOAPI.objMain.HANA_UserID = UserName;
                        objSBOAPI.objMain.HANA_Pwd = Password;
                        return true;

                    default:
                        objSBOAPI.SBO_Appln.MessageBox(_msg);
                        SAPbouiCOM.Form oForm;
                        string FormName = "";
                        try
                        {
                            oForm = objSBOAPI.SBO_Appln.Forms.ActiveForm;
                            FormName = oForm.TypeEx;
                        }
                        catch(Exception)
                        {
                        }
                        if (FormName == "AV_ADLUF")
                        {
                            return false;
                        }
                        else
                        {
                            oForm = objSBOAPI.LoadForm("LicenseUpload.xml", "AV_ADLUF");

                            if(Server!=string.Empty)
                            {
                                oForm.Items.Item("Et_SName").Specific.Value = Server;
                            }
                            else
                            {
                                oForm.Items.Item("Et_SName").Specific.Value = objSBOAPI.oCompany.Server;
                            }
                            oForm.Items.Item("Et_Pswd").Specific.Value = Password;
                            oForm.Items.Item("Et_User").Specific.Value = UserName;
                            oForm.Items.Item("Et_LicUID").Specific.Value = UID;
                            return false;
                        }                        
                }
            }
            catch (Exception ex)
            {
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
        }
        #endregion

        public static LicenseStatus DoExtraValidation(string UID, DateTime ValidateDate, out string validationMsg)
        {
            LicenseStatus _licStatus = LicenseStatus.UNDEFINED;
            validationMsg = string.Empty;

            string AddonName = objSBOAPI.AddonName;
            string ServerName = objSBOAPI.oCompany.Server;
            string DatabaseName = objSBOAPI.oCompany.CompanyDB;

            DateTime currentDate = DateTime.Now;

            if (currentDate.Date <= ValidateDate.Date)
            {
                if (UID == Cls_LicenseHandler.GenerateUID(AddonName, ServerName, DatabaseName))
                {
                    _licStatus = LicenseStatus.VALID;
                }
                else
                {
                    validationMsg = "The license is not valid";
                    _licStatus = LicenseStatus.INVALID;
                }
            }
            else
            {
                validationMsg = "The license Expired";
                _licStatus = LicenseStatus.INVALID;
            }           

            return _licStatus;
        }

    }
}
