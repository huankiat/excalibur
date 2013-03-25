using System;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Net;
using System.Web;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Excalibur.ExcelClient
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Office.CommandBar cellbar = this.Application.CommandBars["Cell"];
            Office.CommandBarButton button = (Office.CommandBarButton) cellbar.FindControl
                (Office.MsoControlType.msoControlButton, 0, 
                "MYRIGHTCLICKMENU", Missing.Value, Missing.Value);
            Office.CommandBarButton button2 = (Office.CommandBarButton)cellbar.FindControl
                (Office.MsoControlType.msoControlButton, 1,
                "MYRIGHTCLICKMENU", Missing.Value, Missing.Value);

            if (button == null & button2== null)
            {
                //subscribe button
                button = (Office.CommandBarButton)cellbar.Controls.
                    Add(Office.MsoControlType.msoControlButton,
                    Missing.Value, Missing.Value, 1, true);
                button.Caption = "Subscribe";
                button.BeginGroup = true;
                button.Tag = "MYRIGHTCLICKMENU";
                button.Click += new Office._CommandBarButtonEvents_ClickEventHandler(showSubForm);

                //publish button
                button2 = (Office.CommandBarButton)cellbar.Controls.
                    Add(Office.MsoControlType.msoControlButton,
                    Missing.Value, Missing.Value, 2, true);
                button2.Caption = "Publish";
                button2.BeginGroup = false;
                button2.Tag = "MYRIGHTCLICKMENU";
                button2.Click += new Office._CommandBarButtonEvents_ClickEventHandler(showPubForm);
            }

           

        }

        private void showSubForm(Office.CommandBarButton cmdBarbutton, ref bool cancel)
        {
            Form1 frm = new Form1();
            frm.Show();
        }

        private void showPubForm(Office.CommandBarButton cmdBarbutton, ref bool cancel)
        {
            Form2 frm = new Form2();
            frm.Show();
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
            return new Ribbon1();
        }

    }
}
