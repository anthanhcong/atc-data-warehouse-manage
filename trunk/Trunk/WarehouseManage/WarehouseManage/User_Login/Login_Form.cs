using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Threading;

namespace WarehouseManager
{
    public partial class User_Login : Form
    {
        //Form1 f1 = new Form1();
        public User_Login()
        {
            InitializeComponent();
            Init_form_Configure();
            //f1.Hide();
        }
        DataTable user_info;
        DataSet user_info_dataset = new DataSet();
        SqlDataAdapter user_info_dataAdapter = new SqlDataAdapter();

        private void Login_BT_Click(object sender, EventArgs e)
        {
            string username = UserName_txt.Text.Trim();
            string password;

            username = Map_UserName(username);
            user_info = Get_UserInfo(username, ref user_info_dataAdapter, ref user_info_dataset);
            if (user_info.Rows.Count == 0)
            {
                MessageBox.Show("This Username did not register. \nPlease contact with Admin!");
                return;
            }
            password = ((string)user_info.Rows[0]["password"]).Trim();
            if (Password_txt.Text.Trim() == Decrypt_Pass(password))
            {
                //TCA_User.Bill_List.Bill_List bill_list = new TCA_User.Bill_List.Bill_List();
                //bill_list.Show();
                Thread Material_Manage_Thread = new Thread(run_WarehouseManage);
                Material_Manage_Thread.SetApartmentState(ApartmentState.STA);
                Material_Manage_Thread.Start();
                //this.Hide();
                //f1.Show();
                this.Close();
                // Application.Exit();

            }
            else {
                MessageBox.Show("Wrong Password");
            }
        }

        
        private void run_WarehouseManage()
        {
            Application.Run(new Form1());
        }

        private void HidePass_check_CheckedChanged(object sender, EventArgs e)
        {
            if (HidePass_check.Checked == true)
            {
                Password_txt.UseSystemPasswordChar = true;
            }
            else 
            {
                Password_txt.UseSystemPasswordChar = false;
            }
        }

        private void Password_txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(e.KeyChar==13)
            {
                Login_BT_Click(null, null);
            }
        }

        private string Decrypt_Pass(string encrypt_pass)
        {
            return Change_HexString2String(encrypt_pass.Trim());

        }

        private string Change_HexString2String(string indata)
        {
            int i, in_len;
            string char_str;
            Int32 value;

            // check correct data
            if (indata == "") return "";
            in_len = indata.Length;
            StringBuilder sb = new StringBuilder(in_len);
            for (i = 0; i < in_len - 2; i = i + 3)
            {
                char_str = indata.Substring(i, 2);
                value = Int32.Parse(char_str, System.Globalization.NumberStyles.HexNumber);
                if ((value < 127) || (value > 0))
                {
                    sb.Append(Convert.ToString(Convert.ToChar(Int32.Parse(char_str, System.Globalization.NumberStyles.HexNumber))));
                }
                else return "";
            }
            return sb.ToString();
        }
    }
}
