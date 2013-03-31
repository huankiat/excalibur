using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System.Web;
using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excalibur.Models;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace Excalibur.ExcelClient
{
    [ComVisible(true)]
    public class ExcaliburRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public ExcaliburRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Excalibur.ExcelClient.ExcaliburRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion

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
            Excel.Workbook wb = exApp.ActiveWorkbook as Excel.Workbook ;
            string filename = wb.Name.ToString();

            Microsoft.Office.Core.DocumentProperties properties;
            properties = (Office.DocumentProperties)wb.CustomDocumentProperties;

            Channel ch = new Channel();
            if (ch.checkFileID(wb) == "Nil")
            {
                string fileID = ch.getFileID(filename);
                MessageBox.Show(fileID, "File ID");

                properties.Add("Excalibur ID", false,
                    Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, fileID);
            }
            else
            {
                MessageBox.Show("ID Exists - Excalibur ID: " + ch.checkFileID(wb), "File Already Registered");
            }

        }

        public void onRefreshButton(Office.IRibbonControl control)
        {
            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel.Workbook wb = exApp.ActiveWorkbook as Excel.Workbook;
            Channel ch = new Channel();

            foreach (Excel.Name nRange in wb.Names)
            {
                string full_name = nRange.Name.ToString();
                nRange.Name = full_name;

                if (full_name.Substring(0, 3) == "SUB" | full_name.Substring(0, 3) == "PUB")
                {
                    string partial_name = full_name.Substring(4);
                    char[] delim = { '_' };
                    string[] splitTxt = partial_name.Split(delim);
                    string cellID = splitTxt[0];
                    int channelID = Convert.ToInt32(cellID);
                    string cellDesc = splitTxt[1];


                    if (full_name.Substring(0, 3) == "SUB")
                    {
                        nRange.RefersToRange.Value = ch.getChannelData(cellID);
                    }
                    else
                    {
                        if (ch.checkFileID(wb) == "Nil")
                        {
                            MessageBox.Show("You need to register this workbook", "Alert");
                        }
                        else
                        {
                            int r_value = (int)nRange.RefersToRange.Value;
                            int fileID = Convert.ToInt32(ch.checkFileID(wb));
                            string txt = ch.rePublishChannel(channelID, cellDesc, r_value, fileID, false);
                            MessageBox.Show(txt, "Response");
                        }
                    }
                }
            }

        }

        public void onLoginButton(Office.IRibbonControl control)
        {
            LoginForm frm = new LoginForm();
            frm.Show();

        }


    }
}
