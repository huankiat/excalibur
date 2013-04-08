using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.Windows.Forms;
using Excalibur.Models;

namespace Excalibur.ExcelClient
{
    public partial class LoginForm : Form
    {
        public AuthToken at;


        public LoginForm()
        {
            InitializeComponent();
            InitializePasswordBox();
            loginErrorLabel.Visible = false;
            at = new AuthToken();
        }

        public AuthToken getAuthToken()
        {
            return at;
        }


        private void loginButton_Click(object sender, EventArgs e)
        {
            Channel ch = new Channel();
            AuthToken authtoken = new AuthToken();
            
            string username = usernameBox.Text.ToString();
            string password = passwordBox.Text.ToString();
            string token = ch.getToken(username, password);

            if (token == "404")
            {
                loginErrorLabel.Visible = true;
            }
            else
            {
                loginErrorLabel.Visible = false;
                authtoken.setToken(token);
                authtoken.createCookieInContainer();
                at = authtoken;
                string readouttoken = authtoken.readTokenFromStore();
                MessageBox.Show(readouttoken, "Token Obtained and Stored");
                
                //write to registry
                Registry.CurrentUser.CreateSubKey("SOFTWARE\\Excalibur");
                RegistryKey myKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Excalibur", true);
                myKey.SetValue("Token", readouttoken, RegistryValueKind.String);

                Properties.Settings.Default.Token = readouttoken;
                Properties.Settings.Default.Save();
                LoginForm.ActiveForm.Close();
                MessageBox.Show(Properties.Settings.Default.Token, "Token Obtained and Stored");
             
                
            }
            
        }

        private void passwordOrUserName_KeyDown(object sender, EventArgs e)
        {
            loginErrorLabel.Visible = false;
        }

        private void InitializePasswordBox()
        {
            passwordBox.Text = "";
            passwordBox.PasswordChar = '*';

        }
    }
}
