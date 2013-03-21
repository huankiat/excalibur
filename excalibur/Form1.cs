using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Net;
using System.Web;
using System.IO;

namespace Excalibur
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            try
            {
                //start Excel 
                oXL = new Excel.Application();
                oXL.Visible = true;

                //new workbook with activesheet selected
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create("http://panoply-staging.herokuapp.com/api/channels.json");
                request.Method = "GET";
                request.Accept = "application/json";
                request.ContentType = "application/json";

                WebResponse response = request.GetResponse();
                using (Stream stream = response.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(stream, Encoding.UTF8);
                    String responseString = reader.ReadToEnd();

                    MessageBox.Show(responseString, "Response");
                    reader.Close();
                }
                response.Close();


            }
            catch
            {
            
            }

        }
    }
}
