using System.ComponentModel;
using System.Xml.Serialization;
using System.Configuration;

namespace JKFM_Source
{
    public class Cls_MyLicense : Cls_LicenseEntity
    {

        #region Constructor
        public Cls_MyLicense()
        {
            this.AppName = ConfigurationManager.AppSettings["AddonName"];

            this.Type = LicenseTypes.Single;
        }
        #endregion

        public override LicenseStatus DoExtraValidation(out string validationMsg)
        {
            return Cls_LicenseVerify.DoExtraValidation(this.UID, this.ValidityDateTime, out validationMsg);
        }
    }
}
