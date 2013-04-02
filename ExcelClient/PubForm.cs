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

namespace Excalibur.ExcelClient
{
    public partial class PubForm : Form
    {
        public PubForm()
        {
            InitializeComponent();
            //ThisAddIn.AddMacro();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel.Workbook wb = exApp.ActiveWorkbook as Excel.Workbook;
            Excel.Worksheet ws = exApp.ActiveSheet as Excel.Worksheet;
            Excel.Range rng = (Excel.Range)exApp.ActiveCell;
           
 
            Channel ch = new Channel();
            AuthToken at = new AuthToken();
            string returnID;
            string token;


            if (AuthToken.cContainer == null & ch.checkSpreadSheetID(wb) == "Nil") 
            {
                MessageBox.Show("Please login and register your workbook.");
            }
            else if (AuthToken.cContainer == null)
            {
                MessageBox.Show("Please login first.");
            }
            else if (ch.checkSpreadSheetID(wb) == "Nil")
            {
                MessageBox.Show("Please register your workbook.");
            }
            else     
            {
                Excel.Shape aShape;
                aShape = ws.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeCross, rng.Left,
                                            rng.Top, 3, 3);
                aShape.Name = "Pub";
                aShape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                aShape.Fill.Solid();
                aShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                aShape.Fill.ForeColor.RGB = Color.FromArgb(200, 200, 90).ToArgb();
                aShape.Placement = Excel.XlPlacement.xlMove;

                int spreadSheetID = Convert.ToInt32(ch.checkSpreadSheetID(wb));
                token = at.readTokenFromCookie();
                ch.setAuthToken(token);

                returnID = ch.publishChannel(textBox1.Text.ToString(), rng.Value.ToString(), spreadSheetID);
                rng.Name = "PUB_" + returnID;
                MessageBox.Show("Published as Channel ID:" + returnID, "Response from Server");
            }
              
            PubForm.ActiveForm.Close();
        }

        public void mssgCall()
        {
            MessageBox.Show("Publish");
        }


    }
    

}
