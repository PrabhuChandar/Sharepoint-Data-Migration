namespace CopyInfopathFields
{
    partial class MoveInfopathFields
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
            this.btnCopy = new System.Windows.Forms.Button();
            this.txtlistName = new System.Windows.Forms.TextBox();
            this.txtLibraryName = new System.Windows.Forms.TextBox();
            this.txtDesturl = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtSourceURL = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.txtLogFileName = new System.Windows.Forms.TextBox();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtUserName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnCopy
            // 
            this.btnCopy.Location = new System.Drawing.Point(592, 299);
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Size = new System.Drawing.Size(137, 23);
            this.btnCopy.TabIndex = 10;
            this.btnCopy.Text = "Start";
            this.btnCopy.UseVisualStyleBackColor = true;
            this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
            // 
            // txtlistName
            // 
            this.txtlistName.Location = new System.Drawing.Point(551, 207);
            this.txtlistName.Name = "txtlistName";
            this.txtlistName.Size = new System.Drawing.Size(178, 20);
            this.txtlistName.TabIndex = 8;
            this.txtlistName.TextChanged += new System.EventHandler(this.txtlistName_TextChanged);
            // 
            // txtLibraryName
            // 
            this.txtLibraryName.Location = new System.Drawing.Point(214, 207);
            this.txtLibraryName.Name = "txtLibraryName";
            this.txtLibraryName.Size = new System.Drawing.Size(198, 20);
            this.txtLibraryName.TabIndex = 9;
            this.txtLibraryName.TextChanged += new System.EventHandler(this.txtLibraryName_TextChanged);
            // 
            // txtDesturl
            // 
            this.txtDesturl.Location = new System.Drawing.Point(214, 91);
            this.txtDesturl.Name = "txtDesturl";
            this.txtDesturl.Size = new System.Drawing.Size(515, 20);
            this.txtDesturl.TabIndex = 7;
            this.txtDesturl.TextChanged += new System.EventHandler(this.txtDesturl_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(472, 210);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(60, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "List Name :";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(56, 214);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(118, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "InfoPath Library Name :";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(81, 54);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(93, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Source Site URL :";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // txtSourceURL
            // 
            this.txtSourceURL.Location = new System.Drawing.Point(214, 54);
            this.txtSourceURL.Name = "txtSourceURL";
            this.txtSourceURL.Size = new System.Drawing.Size(515, 20);
            this.txtSourceURL.TabIndex = 7;
            this.txtSourceURL.TextChanged += new System.EventHandler(this.txtSourceURL_TextChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(35, 98);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(139, 13);
            this.label6.TabIndex = 6;
            this.label6.Text = "Destination o365 Site URL :";
            this.label6.Click += new System.EventHandler(this.label6_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(99, 250);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(81, 13);
            this.label7.TabIndex = 4;
            this.label7.Text = "Log File Name :";
            this.label7.Click += new System.EventHandler(this.label7_Click);
            // 
            // txtLogFileName
            // 
            this.txtLogFileName.Location = new System.Drawing.Point(214, 247);
            this.txtLogFileName.Name = "txtLogFileName";
            this.txtLogFileName.Size = new System.Drawing.Size(515, 20);
            this.txtLogFileName.TabIndex = 8;
            this.txtLogFileName.Text = "LogFile.txt";
            this.txtLogFileName.TextChanged += new System.EventHandler(this.txtLogFileName_TextChanged);
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(214, 175);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '*';
            this.txtPassword.Size = new System.Drawing.Size(515, 20);
            this.txtPassword.TabIndex = 13;
            this.txtPassword.TextChanged += new System.EventHandler(this.txtPassword_TextChanged);
            // 
            // txtUserName
            // 
            this.txtUserName.Location = new System.Drawing.Point(214, 133);
            this.txtUserName.Name = "txtUserName";
            this.txtUserName.Size = new System.Drawing.Size(515, 20);
            this.txtUserName.TabIndex = 14;
            this.txtUserName.TextChanged += new System.EventHandler(this.txtUserName_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(115, 182);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(59, 13);
            this.label5.TabIndex = 11;
            this.label5.Text = "Password :";
            this.label5.Click += new System.EventHandler(this.label5_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(108, 140);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(66, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "User Name :";
            this.label4.Click += new System.EventHandler(this.label4_Click);
            // 
            // MoveInfopathFields
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(782, 344);
            this.Controls.Add(this.txtPassword);
            this.Controls.Add(this.txtUserName);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btnCopy);
            this.Controls.Add(this.txtLogFileName);
            this.Controls.Add(this.txtlistName);
            this.Controls.Add(this.txtLibraryName);
            this.Controls.Add(this.txtSourceURL);
            this.Controls.Add(this.txtDesturl);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label1);
            this.Name = "MoveInfopathFields";
            this.Text = "Move InfoPath Fields";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCopy;
        private System.Windows.Forms.TextBox txtlistName;
        private System.Windows.Forms.TextBox txtLibraryName;
        private System.Windows.Forms.TextBox txtDesturl;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtSourceURL;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtLogFileName;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.TextBox txtUserName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
    }
}

