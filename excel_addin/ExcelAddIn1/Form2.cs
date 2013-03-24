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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel.Worksheet ws = exApp.ActiveSheet as Excel.Worksheet;
            Excel.Range rng = (Excel.Range)exApp.ActiveCell;

            Channel ch = new Channel();
            string responseTxt = ch.publishChannel(textBox1.Text.ToString(), rng.Value.ToString());
            Form2.ActiveForm.Close();
            MessageBox.Show(responseTxt, "Response from Server");
        }
    }
}
