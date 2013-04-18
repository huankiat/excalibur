using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Web;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Excalibur.Models;
using VBIDE = Microsoft.Vbe.Interop;
using Office = Microsoft.Office.Core;

namespace Excalibur.ExcelClient
{
    public partial class PubForm : Form
    {
        
        int spreadSheetID;
        
        public PubForm()
        {
            InitializeComponent();
            checkSSIDPassword();
        }

        private void checkSSIDPassword()
        {
            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel.Workbook wb = exApp.ActiveWorkbook as Excel.Workbook;
            Channel ch = new Channel();

            if (!TokenStore.checkTokenInStore())
            {
                LoginForm frm = new LoginForm();
                frm.Show();
            }
            else
            {
                ch.setAuthToken(TokenStore.getTokenFromStore());
            }

            if (ch.checkSpreadSheetID(wb) == "0")
            {
                string id = ch.getSpreadSheetID(wb.Name.ToString());
                spreadSheetID = Convert.ToInt32(id);

                //write to wb property when a new ID is obtained
                Microsoft.Office.Core.DocumentProperties properties;
                properties = (Office.DocumentProperties)wb.CustomDocumentProperties;

                properties.Add("Excalibur ID", false,
                   Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, id);
            }
            else
            {
                string id = ch.checkSpreadSheetID(wb);
                spreadSheetID = Convert.ToInt32(id);

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel.Workbook wb = exApp.ActiveWorkbook as Excel.Workbook;
            Excel.Worksheet ws = exApp.ActiveSheet as Excel.Worksheet;
            Excel.Range rng = (Excel.Range)exApp.ActiveCell;
           
            Channel ch = new Channel();
            string returnID;
            string token;
            
            //get token and send token together with publication request
            token = TokenStore.getTokenFromStore();
            ch.setAuthToken(token);
            returnID = ch.publishChannel(textBox1.Text.ToString(), rng.Value.ToString(), spreadSheetID);
            rng.Name = "PUB_" + returnID;
            MessageBox.Show("Published as Channel ID:" + returnID, "Response from Server");

            //Add indicator to show publication status
            Excel.Shape aShape;
            aShape = ws.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeCross, rng.Left,
                                        rng.Top, 3, 3);
            aShape.Name = "Pub";
            aShape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            aShape.Fill.Solid();
            aShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            aShape.Fill.ForeColor.RGB = Color.FromArgb(200, 200, 90).ToArgb();
            aShape.Placement = Excel.XlPlacement.xlMove;
              
            PubForm.ActiveForm.Close();
        }

        public void mssgCall()
        {
            MessageBox.Show("Publish");
        }


    }
    

}
