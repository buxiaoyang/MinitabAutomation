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

            this.textBoxOutPut.Text = "Reading Excel File...";
            ExcelHelper excelHelper = new ExcelHelper(this.textBoxFile.Text.Trim());
            Models.RowData rowData = excelHelper.getRowData("Raw data");
            if (rowData == null)
            {
                this.textBoxOutPut.AppendText("Failed\r\n");
            }
            else
            {
                this.textBoxOutPut.AppendText("Succeed\r\n");           
                MinitabHelper minitab = new MinitabHelper();
                minitab.GeneratePictures(rowData, this.textBoxOutPut);

                
                this.textBoxOutPut.AppendText("Inserting picture to Excel...\r\n");
                excelHelper.insertPicture("Graphs Minitab", rowData, this.textBoxOutPut);
                MessageBox.Show("Operation succeed, Please open excel file and check \"Graphs Minitab\" worksheet.");
             
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter =
               "Excel file (*.xls)|*.xls";
            dialog.InitialDirectory = this.textBoxFile.Text;
            dialog.Title = "Select a project file";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                this.textBoxFile.Text = dialog.FileName;
            }
        }
    }
}
