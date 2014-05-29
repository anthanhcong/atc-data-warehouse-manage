namespace WarehouseManager
{
    partial class User_Login
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
            this.Login_grp = new System.Windows.Forms.GroupBox();
            this.Login_BT = new System.Windows.Forms.Button();
            this.HidePass_check = new System.Windows.Forms.CheckBox();
            this.Password_txt = new System.Windows.Forms.TextBox();
            this.UserName_txt = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.Login_grp.SuspendLayout();
            this.SuspendLayout();
            // 
            // Login_grp
            // 
            this.Login_grp.Controls.Add(this.Login_BT);
            this.Login_grp.Controls.Add(this.HidePass_check);
            this.Login_grp.Controls.Add(this.Password_txt);
            this.Login_grp.Controls.Add(this.UserName_txt);
            this.Login_grp.Controls.Add(this.label2);
            this.Login_grp.Controls.Add(this.label1);
            this.Login_grp.Location = new System.Drawing.Point(12, 12);
            this.Login_grp.Name = "Login_grp";
            this.Login_grp.Size = new System.Drawing.Size(296, 148);
            this.Login_grp.TabIndex = 0;
            this.Login_grp.TabStop = false;
            this.Login_grp.Text = "Login";
            // 
            // Login_BT
            // 
            this.Login_BT.Location = new System.Drawing.Point(30, 109);
            this.Login_BT.Name = "Login_BT";
            this.Login_BT.Size = new System.Drawing.Size(75, 23);
            this.Login_BT.TabIndex = 4;
            this.Login_BT.Text = "Login";
            this.Login_BT.UseVisualStyleBackColor = true;
            this.Login_BT.Click += new System.EventHandler(this.Login_BT_Click);
            // 
            // HidePass_check
            // 
            this.HidePass_check.AutoSize = true;
            this.HidePass_check.Location = new System.Drawing.Point(119, 86);
            this.HidePass_check.Name = "HidePass_check";
            this.HidePass_check.Size = new System.Drawing.Size(97, 17);
            this.HidePass_check.TabIndex = 3;
            this.HidePass_check.Text = "Hide Password";
            this.HidePass_check.UseVisualStyleBackColor = true;
            this.HidePass_check.CheckedChanged += new System.EventHandler(this.HidePass_check_CheckedChanged);
            // 
            // Password_txt
            // 
            this.Password_txt.Location = new System.Drawing.Point(119, 60);
            this.Password_txt.Name = "Password_txt";
            this.Password_txt.Size = new System.Drawing.Size(130, 20);
            this.Password_txt.TabIndex = 2;
            this.Password_txt.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.Password_txt_KeyPress);
            // 
            // UserName_txt
            // 
            this.UserName_txt.Location = new System.Drawing.Point(119, 32);
            this.UserName_txt.Name = "UserName_txt";
            this.UserName_txt.Size = new System.Drawing.Size(130, 20);
            this.UserName_txt.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(27, 63);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Password";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "User Name";
            // 
            // TCA_Login
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(320, 172);
            this.Controls.Add(this.Login_grp);
            this.Name = "TCA_Login";
            this.Text = "Login";
            this.Login_grp.ResumeLayout(false);
            this.Login_grp.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox Login_grp;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button Login_BT;
        private System.Windows.Forms.CheckBox HidePass_check;
        private System.Windows.Forms.TextBox Password_txt;
        private System.Windows.Forms.TextBox UserName_txt;
    }
}

