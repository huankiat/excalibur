using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Excalibur.Models;

namespace Excalibur.ExcelClient
{
    public partial class RePubForm : Form
    {
        int channelSelected;
        int spreadSheetID;

        public RePubForm()
        {           
            InitializeComponent();
            InitializePubComboBox();
            //checkSSID();
        }

        private void checkSSID()
        {
            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel.Workbook wb = exApp.ActiveWorkbook as Excel.Workbook;
            Channel ch = new Channel();

            if (ch.checkSpreadSheetID(wb) == "0")
            {
                string id = ch.getSpreadSheetID(wb.Name.ToString());
                spreadSheetID = Convert.ToInt32(id);
            }
            else
            {
                string id = ch.checkSpreadSheetID(wb);
                spreadSheetID = Convert.ToInt32(id);
            }
        }


        private void InitializePubComboBox()
        {
            Channel ch = new Channel();
            JArray d = ch.getAllChannels();

            //Need User ID to send into filterPermittedChannels
            JArray datafeed = ch.filterPermittedChannels(1, d);
            if (datafeed.ToString() != "[]")
            {
                foreach (dynamic data in datafeed)
                {
                    pubComboBox.Items.Add(data.id.ToString() + "-" + data.description.ToString());
                }
            }
            else
            {
                pubComboBox.Items.Add("No data in the channel");
                pubComboBox.SelectedIndex = 0;
            }
        }

        private void readSelectedChannel()
        {
            string[] nameArray = pubComboBox.SelectedItem.ToString().
                    Split(new string[] { "-" }, StringSplitOptions.None);
            channelSelected = Convert.ToInt32(nameArray[0]);
        }

        private void updateChannelDescription(object sender, EventArgs e)
        {
            Channel ch = new Channel();
            descriptionTextBox.Text = ch.getChannelDesc(channelSelected.ToString());
        }

        private void rePubButton_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel.Workbook wb = exApp.ActiveWorkbook as Excel.Workbook;
            Excel.Worksheet ws = exApp.ActiveSheet as Excel.Worksheet;
            Excel.Range rng = (Excel.Range)exApp.ActiveCell;
            Channel ch = new Channel();

            ch.rePublishChannel(channelSelected, descriptionTextBox.Text, rng.Value.ToString(),
                spreadSheetID, forceCheckBox.Checked);
            rng.Name = "PUB_" + channelSelected.ToString();
            RePubForm.ActiveForm.Close();
        }

        
    }



}
