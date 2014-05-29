using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.IO.Ports;

namespace WarehouseManager
{
    public class UserInfo_struct
    {
        public string UserName;
        public string Password;
        public string Empl_ID;
        public string Department;
        public string MailAddr;
        public string Permission;
        public string Group;
    }

    class PosSize
    {
        public int pos_x, pos_y;
        public int width, height;
    }

    enum TextBox_Type
    {
        TEXT,
        NUMBER
    }

    enum AnchorType
    {
        LEFT,
        RIGHT,
        NONE
    }
    enum DateStringType
    {
        DDMMYY,
        MMDDYY,
        YYMMDD
    }

}
