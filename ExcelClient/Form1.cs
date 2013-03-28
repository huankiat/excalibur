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
    public partial class Form1 : Form
    {

        string currentValue;
        string dataID;
        string dataName;
        string dataValue;
        Channel ch;
        
        public Form1()
        {
            
            ch = new Channel();
            currentValue = "";
            InitializeComponent();
            InitializeComboBox();

        

        }

           

        private void InitializeComboBox()
        {

            JArray datafeed = ch.getAllChannels();
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
            
            this.currentValue = rng.Value.ToString();
            
            string[] channelID = comboBox2.SelectedItem.ToString().Split(new string[] {"-"}, StringSplitOptions.None);
            
            this.dataID = channelID[0];
            this.dataName = channelID[1];
            this.dataValue = ch.getChannelData(this.dataID);
            Clipboard.SetDataObject(this.dataValue);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel._Worksheet ws = exApp.ActiveSheet as Excel.Worksheet;
            Excel.Range rng = (Excel.Range)exApp.ActiveCell;

            rng.Name = "SUB_" + this.dataID + "_" + this.dataName;

            rng.PasteSpecial(Excel.XlPasteType.xlPasteAll, 
                Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                Type.Missing, 
                Type.Missing);

            Form1.ActiveForm.Close();
       
        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel.Range rng = (Excel.Range)exApp.ActiveCell;
            
            currentValue = rng.Value.ToString();
            rng.Value = this.dataValue;

        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel.Range rng = (Excel.Range)exApp.ActiveCell;

            rng.Value = this.currentValue;

        }

        
    }
}
