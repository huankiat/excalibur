namespace Excalibur.ExcelClient
{
    partial class SubForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.broadcastComboBox = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.channelComboBox = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // broadcastComboBox
            // 
            this.broadcastComboBox.FormattingEnabled = true;
            this.broadcastComboBox.Location = new System.Drawing.Point(56, 27);
            this.broadcastComboBox.Name = "broadcastComboBox";
            this.broadcastComboBox.Size = new System.Drawing.Size(159, 21);
            this.broadcastComboBox.TabIndex = 0;
            this.broadcastComboBox.SelectedIndexChanged += new System.EventHandler(this.broadcastCombox_Select);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(88, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(88, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Select Broadcast";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(91, 158);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "Subscribe";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.RosyBrown;
            this.label1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.label1.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.label1.Location = new System.Drawing.Point(101, 123);
            this.label1.Name = "label1";
            this.label1.Padding = new System.Windows.Forms.Padding(5);
            this.label1.Size = new System.Drawing.Size(55, 23);
            this.label1.TabIndex = 3;
            this.label1.Text = "Preview";
            this.label1.MouseEnter += new System.EventHandler(this.button1_MouseEnter);
            this.label1.MouseLeave += new System.EventHandler(this.button1_MouseLeave);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(92, 62);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(79, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Select Channel";
            // 
            // channelComboBox
            // 
            this.channelComboBox.FormattingEnabled = true;
            this.channelComboBox.Location = new System.Drawing.Point(56, 80);
            this.channelComboBox.Name = "channelComboBox";
            this.channelComboBox.Size = new System.Drawing.Size(159, 21);
            this.channelComboBox.TabIndex = 5;
            this.channelComboBox.SelectedIndexChanged += new System.EventHandler(this.channelComboBox_Select);
            // 
            // SubForm
            // 
            this.ClientSize = new System.Drawing.Size(269, 198);
            this.Controls.Add(this.channelComboBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.broadcastComboBox);
            this.Name = "SubForm";
            this.Text = "Subscription";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox broadcastComboBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox channelComboBox;
    }
}