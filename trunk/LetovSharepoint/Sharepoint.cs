using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
using System.Net;
using System.Security;
using System.Data;

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
        }

        public DataTable LoadList(string Title)
        {
            DataTable dtGetReqForm = new DataTable(); ;
            try
            {
                Microsoft.SharePoint.Client.List spList = clientContext.Web.Lists.GetByTitle(Title);
                clientContext.Load(spList);
                clientContext.ExecuteQuery();

                if (spList != null && spList.ItemCount > 0)
                {
                    Microsoft.SharePoint.Client.CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ViewXml =
                    @"<View>" +
                    "<Query> " +
                        "<Where>" +
                            "<And>" +
                                    "<IsNotNull><FieldRef Name='ID' /></IsNotNull>" +
                                    "<Eq><FieldRef Name='ReqNo' /><Value Type='Text'>123</Value></Eq>" +
                            "</And>" +
                        "</Where>" +
                    "</Query> " +
                    "<ViewFields>" +
                        "<FieldRef Name='ID' />" +
                    "</ViewFields>" +      
                    "</View>";

                    SP.ListItemCollection listItems = spList.GetItems(camlQuery);
                    clientContext.Load(listItems);
                    clientContext.ExecuteQuery();

                    if (listItems != null && listItems.Count > 0)
                    {
                        foreach (var field in listItems[0].FieldValues.Keys)
                        {
                            dtGetReqForm.Columns.Add(field);
                        }

                        foreach (var item in listItems)
                        {
                            DataRow dr = dtGetReqForm.NewRow();

                            foreach (var obj in item.FieldValues)
                            {
                                if (obj.Value != null)
                                {
                                    string key = obj.Key;
                                    string type = obj.Value.GetType().FullName;

                                    if (type == "Microsoft.SharePoint.Client.FieldLookupValue")
                                    {
                                        dr[obj.Key] = ((FieldLookupValue)obj.Value).LookupValue;
                                    }
                                    else if (type == "Microsoft.SharePoint.Client.FieldUserValue")
                                    {
                                        dr[obj.Key] = ((FieldUserValue)obj.Value).LookupValue;
                                    }
                                    else if (type == "Microsoft.SharePoint.Client.FieldUserValue[]")
                                    {
                                        FieldUserValue[] multValue = (FieldUserValue[])obj.Value;
                                        foreach (FieldUserValue fieldUserValue in multValue)
                                        {
                                            dr[obj.Key] += (fieldUserValue).LookupValue;
                                        }
                                    }
                                    else if (type == "System.DateTime")
                                    {
                                        if (obj.Value.ToString().Length > 0)
                                        {
                                            var date = obj.Value.ToString().Split(' ');
                                            if (date[0].Length > 0)
                                            {
                                                dr[obj.Key] = date[0];
                                            }
                                        }
                                    }
                                    else
                                    {
                                        dr[obj.Key] = obj.Value;
                                    }
                                }
                                else
                                {
                                    dr[obj.Key] = null;
                                }
                            }
                            dtGetReqForm.Rows.Add(dr);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
            finally
            {
                if (clientContext != null)
                    clientContext.Dispose();
            }
            return dtGetReqForm;
        }


    }
}
