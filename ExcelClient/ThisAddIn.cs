using System;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Net;
using System.Web;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excalibur.Models;

namespace Excalibur.ExcelClient
{
    public partial class ThisAddIn
    {
        Office.CommandBarButton subButton;
        Office.CommandBarButton pubButton;
        Office.CommandBarButton refreshButton;
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.AddContextMenu();

        }

        public void AddContextMenu()
        {
            Office.CommandBar cellbar = this.Application.CommandBars["Cell"];

            //subscribe button
            subButton = (Office.CommandBarButton)cellbar.Controls.Add();
            subButton.Caption = "Subscribe";
            subButton.BeginGroup = true;
            subButton.Tag = "subButton";
            subButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(showSubForm);

            //publish button 
            pubButton = (Office.CommandBarButton)cellbar.Controls.Add();
            pubButton.Caption = "Publish";
            pubButton.Tag = "pubButton";
            pubButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(showPubForm);

            //refresh button 
            refreshButton = (Office.CommandBarButton)cellbar.Controls.Add();
            refreshButton.Caption = "Refresh";
            refreshButton.Tag = "refreshButton";
            refreshButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(refreshAll);
        }

        private void showSubForm(Office.CommandBarButton cmdBarbutton, ref bool cancel)
        {
            SubForm frm1 = new SubForm();
            frm1.Show();
        }

        private void showPubForm(Office.CommandBarButton cmdBarbutton, ref bool cancel)
        {
            PubForm frm2 = new PubForm();
            frm2.Show();       
        }

        private void refreshAll(Office.CommandBarButton cmdBarbutton, ref bool cancel)
        {
            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel.Workbook wb = exApp.ActiveWorkbook as Excel.Workbook;
            Channel ch = new Channel();
            string txt;

            txt = ch.channelsRefresh(wb);
            MessageBox.Show(txt, "Refresh");
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Application.CommandBars["Cell"].Reset();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new ExcaliburRibbon();
        }

    }
}
