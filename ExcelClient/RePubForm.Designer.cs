namespace Excalibur.ExcelClient
{
    partial class RePubForm
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
            this.rePubButton = new System.Windows.Forms.Button();
            this.pubComboBox = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.forceCheckBox = new System.Windows.Forms.CheckBox();
            this.descriptionTextBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // rePubButton
            // 
            this.rePubButton.Location = new System.Drawing.Point(93, 181);
            this.rePubButton.Name = "rePubButton";
            this.rePubButton.Size = new System.Drawing.Size(75, 23);
            this.rePubButton.TabIndex = 0;
            this.rePubButton.Text = "RePublish";
            this.rePubButton.UseVisualStyleBackColor = true;
            this.rePubButton.Click += new System.EventHandler(this.rePubButton_Click);
            // 
            // pubComboBox
            // 
            this.pubComboBox.FormattingEnabled = true;
            this.pubComboBox.Location = new System.Drawing.Point(60, 47);
            this.pubComboBox.Name = "pubComboBox";
            this.pubComboBox.Size = new System.Drawing.Size(158, 21);
            this.pubComboBox.TabIndex = 1;
            this.pubComboBox.SelectedIndexChanged += new System.EventHandler(this.updateChannelDescription);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(76, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(119, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Available Pub Channels";
            // 
            // forceCheckBox
            // 
            this.forceCheckBox.AutoSize = true;
            this.forceCheckBox.Location = new System.Drawing.Point(93, 144);
            this.forceCheckBox.Name = "forceCheckBox";
            this.forceCheckBox.Size = new System.Drawing.Size(80, 17);
            this.forceCheckBox.TabIndex = 3;
            this.forceCheckBox.Text = "OverWrite?";
            this.forceCheckBox.UseVisualStyleBackColor = true;
            // 
            // descriptionTextBox
            // 
            this.descriptionTextBox.Location = new System.Drawing.Point(60, 105);
            this.descriptionTextBox.Name = "descriptionTextBox";
            this.descriptionTextBox.Size = new System.Drawing.Size(158, 20);
            this.descriptionTextBox.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(90, 89);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(102, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Channel Description";
            // 
            // RePubForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 227);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.descriptionTextBox);
            this.Controls.Add(this.forceCheckBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pubComboBox);
            this.Controls.Add(this.rePubButton);
            this.Name = "RePubForm";
            this.Text = "Republish";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button rePubButton;
        private System.Windows.Forms.ComboBox pubComboBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox forceCheckBox;
        private System.Windows.Forms.TextBox descriptionTextBox;
        private System.Windows.Forms.Label label2;
    }
}