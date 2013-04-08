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

namespace Excalibur.ExcelClient
{
    public partial class SubForm : Form
    {

        string currentValue;
        string currentFormula;
        string channelID;
        string channelName;
        string channelValue;
        Channel ch;
        
        public SubForm()
        {
            ch = new Channel();
            currentValue = "";
            currentFormula = "";
            InitializeComponent();
            InitializeComboBox();

        }
      

        private void InitializeComboBox()
        {

            ch.setAuthToken(Properties.Settings.Default.Token.ToString());
            JArray d = ch.getAllChannels();
            JArray datafeed = ch.filterPermittedChannels(1, d);
            if (datafeed.ToString() != "[]")
            {
                foreach (dynamic data in datafeed)
                {
                    comboBox2.Items.Add(data.id.ToString() + "-" + data.description.ToString());
                }
            }
            else
            {
                comboBox2.Items.Add("No data in the channel");
                comboBox2.SelectedIndex = 0;
            }
        }

        private void comboxbox1_Select(object sender, EventArgs e)
        {
            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel._Worksheet ws = exApp.ActiveSheet as Excel.Worksheet;
            Excel.Range rng = (Excel.Range)exApp.ActiveCell;

            
            string[] nameArray = comboBox2.SelectedItem.ToString().Split(new string[] {"-"}, StringSplitOptions.None);
            
            channelID = nameArray[0];
            channelName = nameArray[1];
            channelValue = ch.getChannelData(channelID);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel._Worksheet ws = exApp.ActiveSheet as Excel.Worksheet;
            Excel._Workbook wb = exApp.ActiveWorkbook as Excel.Workbook;
            Excel.Range rng = (Excel.Range)exApp.ActiveCell;

            string[] nameArray = comboBox2.SelectedItem.ToString().Split(new string[] { "-" }, StringSplitOptions.None);

            channelID = nameArray[0];
            channelValue = ch.getChannelData(channelID);

            rng.Name = "SUB_" + channelID;
            rng.Value = channelValue;

            //Add indicator to show subscription status
            Excel.Shape aShape;
            aShape = ws.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeCross, rng.Left,
                                        rng.Top, 3, 3);
            aShape.Name = "Sub";
            aShape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            aShape.Fill.Solid();
            aShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            aShape.Fill.ForeColor.RGB = Color.FromArgb(90, 90, 200).ToArgb();
            aShape.Placement = Excel.XlPlacement.xlMove;
         

            SubForm.ActiveForm.Close();
       
        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel.Range rng = (Excel.Range)exApp.ActiveCell;       

            if (rng.Value == null)
            {
                currentValue = "";
            }
            else
            {
                currentValue = rng.Value.ToString();

                if (rng.Formula != null)
                {
                    currentFormula = rng.Formula;
                }

            }
            rng.Value = channelValue;

        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel.Range rng = (Excel.Range)exApp.ActiveCell;

            rng.Value = currentValue;
            rng.Formula = currentFormula;
            
            currentValue = "";

        }

        
    }
}
