using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System.Web;
using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excalibur.Models;

namespace Excalibur.ExcelClient
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        public void onSubButton(Office.IRibbonControl control)
        {
            SubForm frm = new SubForm();
            frm.Show();

        }

        public void onPubButton(Office.IRibbonControl control)
        {
            PubForm frm = new PubForm();
            frm.Show();

        }

        public void onRegButton(Office.IRibbonControl control)
        {
            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel.Workbook wb = exApp.ActiveWorkbook as Excel.Workbook;
            string filename = wb.Name.ToString();

            Microsoft.Office.Core.DocumentProperties properties;
            properties = (Office.DocumentProperties)wb.CustomDocumentProperties;

            Channel ch = new Channel();
            if (ch.checkSpreadSheetID(wb) == "Nil")
            {
                string fileID = ch.getSpreadSheetID(filename);
                MessageBox.Show(fileID, "File ID");

                properties.Add("Excalibur ID", false,
                    Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, fileID);
            }
            else
            {
                MessageBox.Show("ID Exists - Excalibur ID: " + ch.checkSpreadSheetID(wb), "File Already Registered");
            }

        }

        public void onRefreshButton(Office.IRibbonControl control)
        {
            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel.Workbook wb = exApp.ActiveWorkbook as Excel.Workbook;
            Channel ch = new Channel();

            string txt = ch.channelsRefresh(wb);
            MessageBox.Show(txt, "Refresh");


        }

        public void onLoginButton(Office.IRibbonControl control)
        {
            LoginForm frm = new LoginForm();
            frm.Show();

        }
    }
}
