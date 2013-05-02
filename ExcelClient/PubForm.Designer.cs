namespace Excalibur.ExcelClient
{
    partial class PubForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.feedNametextBox = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.broadcastComboBox = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(46, 85);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(98, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Name the datafeed";
            // 
            // feedNametextBox
            // 
            this.feedNametextBox.Location = new System.Drawing.Point(36, 103);
            this.feedNametextBox.Name = "feedNametextBox";
            this.feedNametextBox.Size = new System.Drawing.Size(122, 20);
            this.feedNametextBox.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(60, 133);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "Publish";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // broadcastComboBox
            // 
            this.broadcastComboBox.FormattingEnabled = true;
            this.broadcastComboBox.Location = new System.Drawing.Point(36, 46);
            this.broadcastComboBox.Name = "broadcastComboBox";
            this.broadcastComboBox.Size = new System.Drawing.Size(121, 21);
            this.broadcastComboBox.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(50, 26);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(94, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Choose Broadcast";
            // 
            // PubForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(204, 176);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.broadcastComboBox);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.feedNametextBox);
            this.Controls.Add(this.label1);
            this.Name = "PubForm";
            this.Text = "Publish";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox feedNametextBox;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox broadcastComboBox;
        private System.Windows.Forms.Label label2;
    }
}