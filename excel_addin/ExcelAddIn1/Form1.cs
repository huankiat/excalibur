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

namespace ExcelAddIn1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            Channel ch = new Channel();
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

           

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel.Worksheet ws = exApp.ActiveSheet as Excel.Worksheet;
            Excel.Range rng = (Excel.Range)exApp.ActiveCell;
            
            string[] channelID = comboBox2.SelectedItem.ToString().Split(new string[] {"-"}, StringSplitOptions.None);

            Channel ch = new Channel();
            rng.Value = ch.getChannelData(channelID[0]);

        }


        
    }
}
