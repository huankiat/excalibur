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
using Office = Microsoft.Office.Core;

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


        private void InitializePubComboBox()
        {
            Channel ch = new Channel();      
            string token = TokenStore.getTokenFromStore();
            ch.setAuthToken(token);
            JArray d = ch.getAllBroadcastsChannels();

            //Need User ID to send into filterPermittedChannels
            JArray datafeed = ch.filterChannelsInBroadcast(d);
            if (datafeed.ToString() != "[]")
            {
                foreach (dynamic data in datafeed)
                {
                    foreach (dynamic c in data.channels)
                    {
                        pubComboBox.Items.Add(c.id.ToString() + "-" + c.description.ToString());
                    }
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
            this.readSelectedChannel();
            descriptionTextBox.Text = ch.getChannelDesc(channelSelected.ToString());
        }

        private void rePubButton_Click(object sender, EventArgs e)
        {
            string responseFromServer;

            Excel.Application exApp = Globals.ThisAddIn.Application as Excel.Application;
            Excel.Workbook wb = exApp.ActiveWorkbook as Excel.Workbook;
            Excel.Worksheet ws = exApp.ActiveSheet as Excel.Worksheet;
            Excel.Range rng = (Excel.Range)exApp.ActiveCell;
            Channel ch = new Channel();
            this.readSelectedChannel();
            ch.setAuthToken(TokenStore.getTokenFromStore());

            string to_replace = "";
            if (forceCheckBox.Checked)
            {
                to_replace = "true";
            }
            else
            {
                to_replace = "false";
            }

            responseFromServer = ch.rePublishChannel(channelSelected, descriptionTextBox.Text, rng.Value.ToString(),
                spreadSheetID, to_replace);
            MessageBox.Show(responseFromServer.ToString());
            if (responseFromServer == "409")
            {
                MessageBox.Show(@"This workbook is not the original publisher. 
                        Please check 'OverWrite' and retry if you want to overwrite data in the channel", "Cannot Overwrite");
            }
            else if (responseFromServer == "401")
            {
                MessageBox.Show(@"You are not authorized to publish into the channel", "Unauthorized");
            }
            else
            {
                rng.Name = "PUB_" + channelSelected.ToString();

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

                RePubForm.ActiveForm.Close();
            }
        }

        
    }



}
