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

            this.textBoxOutPut.Text = "正在读取Excel文件...";
            ExcelHelper excelHelper = new ExcelHelper(this.textBoxFile.Text.Trim());
            Models.RowData rowData = excelHelper.getRowData("Raw data");
            if (rowData == null)
            {
                this.textBoxOutPut.AppendText("失败\r\n");
            }
            else
            {
                this.textBoxOutPut.AppendText("成功\r\n");           
                MinitabHelper minitab = new MinitabHelper();
                minitab.GeneratePictures(rowData, this.textBoxOutPut);

                
                this.textBoxOutPut.AppendText("正在插入图片到Excel...\r\n");
                excelHelper.insertPicture("Graphs Minitab", rowData, this.textBoxOutPut);
                MessageBox.Show("操作完成，请打开Excel文件查看\"Graphs Minitab\"工作表。");
             
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
