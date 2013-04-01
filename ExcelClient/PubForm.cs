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
    public partial class PubForm : Form
    {
        public PubForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel.Workbook wb = exApp.ActiveWorkbook as Excel.Workbook;
            Excel.Worksheet ws = exApp.ActiveSheet as Excel.Worksheet;
            Excel.Range rng = (Excel.Range)exApp.ActiveCell;
 
            Channel ch = new Channel();
            string returnID;
            Excel.IconSetCondition isc = (Excel.IconSetCondition)rng.FormatConditions.AddIconSetCondition();
            isc.ShowIconOnly = true;
            isc.IconSet = wb.IconSets[Excel.XlIconSet.xl3Triangles];
            isc.IconCriteria[2].Type = Excel.XlConditionValueTypes.xlConditionValueNumber;
            isc.IconCriteria[2].Value = 0;
            isc.IconCriteria[2].Operator = (int)(Excel.XlFormatConditionOperator.xlNotEqual);

            if (ch.checkFileID(wb) == "Nil")
            {
                MessageBox.Show("Need to register the workbook first", "Alert");
            }
            else
            {
                int fileID = Convert.ToInt32(ch.checkFileID(wb));
                returnID = ch.publishChannel(textBox1.Text.ToString(), rng.Value.ToString(), fileID);
                rng.Name = "PUB_" + returnID + "_" + textBox1.Text.ToString();
                MessageBox.Show("Published as " + returnID, "Response from Server");
            }
            PubForm.ActiveForm.Close();
        
        }


    }
}
