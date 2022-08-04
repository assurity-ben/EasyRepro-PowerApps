using System;
using System.Security;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;

namespace Microsoft.Dynamics365.UIAutomation.Sample.UCI
{
    [TestClass]
    public class PPMTestSuite
    {
        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["PPMOnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["PPMOnlinePassword"].ToSecureString();
        private readonly SecureString _mfaSecretKey = System.Configuration.ConfigurationManager.AppSettings["PPMMfaSecretKey"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["PPMOnlineCrmUrl"].ToString());

        [TestMethod]
        public void PPMTest1()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);
                xrmApp.Navigation.OpenApp("Vacancy and Market Rent Approvals");
                xrmApp.Navigation.OpenSubArea("Main", "Providers");
                xrmApp.Grid.SwitchView("Active Accounts");
                xrmApp.Grid.Search("Aarangi Motel");
                xrmApp.Grid.OpenRecord(0);
                xrmApp.Entity.SelectTab("Related", "Contacts");
                xrmApp.Entity.SubGrid.ClickCommand("Contact Commands", "New Contact");
                xrmApp.QuickCreate.SetValue("firstname", "Bugs");
                xrmApp.QuickCreate.SetValue("lastname", "Bunny");
                xrmApp.QuickCreate.SetValue("jobtitle", "Grey Hare");
                xrmApp.QuickCreate.SetValue("emailaddress1", "bugs@email.com");
                xrmApp.QuickCreate.SetValue("mobilephone", "021123456");
                xrmApp.QuickCreate.SetValue("telephone1", "044780668");
                xrmApp.QuickCreate.SetValue("description", "Sample test account for Bugs");
                xrmApp.ThinkTime(5000);
            }
        }

    }
}