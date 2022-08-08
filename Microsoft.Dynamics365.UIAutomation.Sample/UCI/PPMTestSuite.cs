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
                xrmApp.Navigation.OpenApp(UCIAppName.Vacancy);
                //Open providers page
                xrmApp.Navigation.OpenSubArea("Main", "Providers");
                //Switch view and search for a provider
                xrmApp.Grid.SwitchView("Active Accounts");
                xrmApp.Grid.Search("Aarangi Motel");
                //Open the provider details
                xrmApp.Grid.OpenRecord(0);
                //QuickCreate a new contact
                Random rnd = new Random();
                String lastname = "Bunny" + rnd.Next(1000);
                xrmApp.Entity.SelectTab("Related", "Contacts");
                xrmApp.Entity.SubGrid.ClickCommand("Contact Commands", "New Contact");
                xrmApp.QuickCreate.SetValue("firstname", "Bugs");
                xrmApp.QuickCreate.SetValue("lastname", lastname);
                xrmApp.QuickCreate.SetValue("jobtitle", "Grey Hare");
                xrmApp.QuickCreate.SetValue("emailaddress1", "bugs." + lastname + "@email.com");
                xrmApp.QuickCreate.Save();
                //Serach for and open the new contact just created
                xrmApp.Grid.Search("Bugs " + lastname);
                xrmApp.RelatedGrid.OpenGridRow(0);
                //Create an invitation
                xrmApp.CommandBar.ClickCommand("Create Invitation");
                xrmApp.CommandBar.ClickCommand("Save");
                xrmApp.ThinkTime(5000);
            }
        }

    }
}