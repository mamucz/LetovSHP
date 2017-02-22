using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
using System.Net;
using System.Security;

namespace LetovSharepoint
{
    public class Sharepoint
    {
        private string _siteUrl;
        private ClientContext clientContext;
        private Web oWebsite;

        string SiteUrl
        {
            get { return _siteUrl; }
            
        }

        public Sharepoint(string siteUrl)
        {
            _siteUrl = siteUrl;
        }

        public string Authenicate(string UserName, string Password)
        {
            string siteUrl = "https://ermsro.sharepoint.com/Firma%20Letov/";

            clientContext = new ClientContext(siteUrl);
            oWebsite = clientContext.Web;
            ListCollection collList = oWebsite.Lists;

            SecureString passWord = new SecureString();
            foreach (char c in Password.ToCharArray()) passWord.AppendChar(c);
            clientContext.Credentials = new SharePointOnlineCredentials(UserName, passWord);

            clientContext.Load(oWebsite);
            try
            {
                clientContext.ExecuteQuery();
            }
            catch (Microsoft.SharePoint.Client.IdcrlException e)
            {
                return e.Message.ToString();
            }
            return "Connected sucesfully";
        }

        public List<SP.List> GetLists()
        {
            ListCollection collList = oWebsite.Lists;
            clientContext.Load(collList);

            clientContext.ExecuteQuery();
            return collList.ToList();

            //foreach (SP.List oList in collList)
            //{
            //    Console.WriteLine("Title: {0} Created: {1}", oList.Title, oList.Created.ToString());
            //}
        }

       
    }
}
