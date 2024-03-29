﻿using System;
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


        public LoginForm()
        {
            InitializeComponent();
            InitializePasswordBox();
            loginErrorLabel.Visible = false;
        }



        private void loginButton_Click(object sender, EventArgs e)
        {
            Channel ch = new Channel();
            
            string username = usernameBox.Text.ToString();
            string password = passwordBox.Text.ToString();
            string token = ch.getToken(username, password);

            if (token == "404"|token == "401")
            {
                loginErrorLabel.Visible = true;
            }
            else
            {
                loginErrorLabel.Visible = false;
                TokenStore.addTokenToStore(token);
                MessageBox.Show(TokenStore.getTokenFromStore(), "Token Obtained and Stored");
                
                LoginForm.ActiveForm.Close();
             
                
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
