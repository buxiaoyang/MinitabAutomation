using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MinitabAutomation
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelHelper excelHelper = new ExcelHelper(@"C:\Users\bux\Desktop\minitab\Measdata ROR 101 1010-CB P1X - first tc -  1.xls");
            Models.RowData rowData = excelHelper.getRowData("Raw data");

            MinitabHelper.GeneratePictures(rowData);

            return;
        }
    }
}
