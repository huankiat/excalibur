using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
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
        }

        private void loginButton_Click(object sender, EventArgs e)
        {
            Channel ch = new Channel();
            
            string username = usernameBox.Text.ToString();
            string password = passwordBox.Text.ToString();
            string token = ch.getToken(username, password);
            LoginForm.ActiveForm.Close();
            MessageBox.Show(token);
        }

        private void InitializePasswordBox()
        {
            passwordBox.Text = "";
            passwordBox.PasswordChar = '*';

        }
    }
}
