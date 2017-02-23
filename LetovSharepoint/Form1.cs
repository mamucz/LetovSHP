using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LetovSharepoint
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        Sharepoint shp;
        List<SP.List> collList;

        public Form1()
        {
            InitializeComponent();
        }

      
        private void btConnect_click(object sender, EventArgs e)
        {
            shp = new Sharepoint(tbSite.Text);
            lbConnectionStatus.Text = "Connecting...";
            lbConnectionStatus.Refresh();
            lbConnectionStatus.Text = shp.Authenicate(tbUserName.Text, tbPassword.Text);
            collList = shp.GetLists();
            foreach(var l in from list in collList select list.Title)
            {
                cbLists.Items.Add(l);
            }
        }

        private void cbLists_SelectedIndexChanged(object sender, EventArgs e)
        {
            //SP.List selectedList = collList.FirstOrDefault(x => x.Title == cbLists.Text);
            DataTable dt = shp.LoadList(cbLists.Text);
            dataGridView1.DataSource = dt;
            dataGridView1.Update();
           
        }
    }
}
