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
using System.IO;
using System.IO.Ports;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace WarehouseManager
{
    public partial class Form1
    {
        public enum Excel_Col_Type
        {
            COL_STRING,
            COL_INT,
            COL_FLOAT,
            COL_DATE
        };

        public class ExcelImportStruct : Form
        {
            public string Name;
            public string Col_str;
            public string DB_str;
            public int Col;
            public Excel_Col_Type Col_type;
            public int Data_Max_len;

            public ExcelImportStruct(string name, string col_str, Excel_Col_Type type, int data_max_len)
            {
                Name = name;
                Col_str = col_str;
                DB_str = name;
                Col_type = type;
                Data_Max_len = data_max_len;
                Col = 0;
            }
        }

    }
}