namespace LetovSharepoint
{
    partial class Form1
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
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.tbSite = new System.Windows.Forms.TextBox();
            this.tbPassword = new System.Windows.Forms.TextBox();
            this.tbUserName = new System.Windows.Forms.TextBox();
            this.btConnect = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.lbConnectionStatus = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.cbLists = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(28, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Site:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(396, 8);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(60, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "UserName:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(396, 38);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Password:";
            // 
            // tbSite
            // 
            this.tbSite.Location = new System.Drawing.Point(46, 5);
            this.tbSite.Name = "tbSite";
            this.tbSite.Size = new System.Drawing.Size(337, 20);
            this.tbSite.TabIndex = 3;
            this.tbSite.Text = "https://ermsro.sharepoint.com/Firma%20Letov/";
            // 
            // tbPassword
            // 
            this.tbPassword.Location = new System.Drawing.Point(462, 35);
            this.tbPassword.Name = "tbPassword";
            this.tbPassword.Size = new System.Drawing.Size(100, 20);
            this.tbPassword.TabIndex = 4;
            this.tbPassword.Text = "Pipe2495";
            this.tbPassword.UseSystemPasswordChar = true;
            // 
            // tbUserName
            // 
            this.tbUserName.Location = new System.Drawing.Point(462, 5);
            this.tbUserName.Name = "tbUserName";
            this.tbUserName.Size = new System.Drawing.Size(100, 20);
            this.tbUserName.TabIndex = 5;
            this.tbUserName.Text = "petr.ondra@erm.cz";
            // 
            // btConnect
            // 
            this.btConnect.Location = new System.Drawing.Point(46, 35);
            this.btConnect.Name = "btConnect";
            this.btConnect.Size = new System.Drawing.Size(75, 23);
            this.btConnect.TabIndex = 6;
            this.btConnect.Text = "Connect";
            this.btConnect.UseVisualStyleBackColor = true;
            this.btConnect.Click += new System.EventHandler(this.btConnect_click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(43, 61);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(95, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Connection status:";
            // 
            // lbConnectionStatus
            // 
            this.lbConnectionStatus.Location = new System.Drawing.Point(144, 61);
            this.lbConnectionStatus.Name = "lbConnectionStatus";
            this.lbConnectionStatus.Size = new System.Drawing.Size(505, 23);
            this.lbConnectionStatus.TabIndex = 8;
            this.lbConnectionStatus.Text = "Not connected";
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 116);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(550, 224);
            this.dataGridView1.TabIndex = 9;
            // 
            // cbLists
            // 
            this.cbLists.FormattingEnabled = true;
            this.cbLists.Location = new System.Drawing.Point(15, 87);
            this.cbLists.Name = "cbLists";
            this.cbLists.Size = new System.Drawing.Size(212, 21);
            this.cbLists.TabIndex = 10;
            this.cbLists.SelectedIndexChanged += new System.EventHandler(this.cbLists_SelectedIndexChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(577, 352);
            this.Controls.Add(this.cbLists);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.lbConnectionStatus);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btConnect);
            this.Controls.Add(this.tbUserName);
            this.Controls.Add(this.tbPassword);
            this.Controls.Add(this.tbSite);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tbSite;
        private System.Windows.Forms.TextBox tbPassword;
        private System.Windows.Forms.TextBox tbUserName;
        private System.Windows.Forms.Button btConnect;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lbConnectionStatus;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.ComboBox cbLists;
    }
}

