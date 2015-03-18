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
            ExcelHelper excelHelper = new ExcelHelper(@"C:\Copy of Measdata ROR 101 1010-CB P1X - first tc -1 .xls");
            DataTable dt = excelHelper.ExcelToDataTable("Raw data", true);
            
            return;

            try
            {
                foreach (Process proc in Process.GetProcessesByName("Mtb"))
                {
                    proc.Kill();
                }
            }
            catch (Exception ex)
            {

            }

            Mtb.Application MtbApp = new Mtb.Application();
            MtbApp.UserInterface.Visible = true;
            Console.WriteLine("Status = " + MtbApp.Status);
            Console.WriteLine("LastError = " + MtbApp.LastError);
            Console.WriteLine("Application Path = " + MtbApp.AppPath);
            Console.WriteLine("Window Handle = " + MtbApp.Handle);

            Mtb.Project MtbProj = MtbApp.ActiveProject;
            /*
            Mtb.Columns MtbColumns = MtbProj.ActiveWorksheet.Columns;
            Mtb.Column MtbColumn1 = MtbColumns.Add(null,null,1);
            MtbColumn1.Name = "缺陷项";
            String[] data1 = {"虚焊","漏焊","强度不够","外观受损","其他"};
            MtbColumn1.SetData(data1);

            Mtb.Column MtbColumn2 = MtbColumns.Add(null, null, 1);
            MtbColumn2.Name = "数量";
            Double[] data2 = { 500, 300, 200, 150, 160};
            MtbColumn2.SetData(data2);
            */

            Mtb.Columns MtbColumns = MtbProj.ActiveWorksheet.Columns;
            Mtb.Column MtbColumn1 = MtbColumns.Add(null, null, 1);
            int[] data1 = { 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6193, 6192, 6193, 6193, 6192, 6193, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6193, 6192, 6193, 6193, 6193, 6192, 6193, 6193, 6193, 6193, 6192, 6193, 6193, 6192, 6193, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192, 6192 };
            MtbColumn1.SetData(data1);

            Mtb.Column MtbColumn2 = MtbColumns.Add(null, null, 1);
            DateTime[] data2 = { DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-08"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07"), DateTime.Parse("2013-10-07") };
            MtbColumn2.SetData(data2);

            Mtb.Column MtbColumn3 = MtbColumns.Add(null, null, 1);
            Double[] data3 = { 1808, 1796.3, 1799, 1800.8, 1800.5, 1803.6, 1794.5, 1796.9, 1802.1, 1796.5, 1795.5, 1812.8, 1795.8, 1805.6, 1798.6, 1799.6, 1803.7, 1800.9, 1799.4, 1808.6, 1800.3, 1804, 1804.4, 1798.8, 1802.1, 1810.1, 1797, 1809.4, 1793.3, 1803.6, 1801, 1797.5, 1797.8, 1796.4, 1796.1, 1793.9, 1809.7, 1795.3, 1798.9, 1796.6, 1793.6, 1794.4, 1798.5, 1806.2, 1807, 1805.4, 1806.9, 1800.7, 1799.5, 1802.3, 1791.8, 1795.5, 1798.5, 1801.1, 1794.2, 1799.4, 1792.5, 1803.7, 1805.1, 1808.9, 1795.6, 1806.3, 1799.5, 1798.9, 1799.6, 1795.8, 1799.9, 1795.1, 1806.4, 1809.6, 1808, 1796.1, 1797.1, 1793.1, 1802.6, 1800, 1796.5, 1799.3, 1798.6, 1808, 1803.2, 1810.2, 1793.9, 1798.6, 1808, 1797, 1793.2, 1801.6, 1803.4, 1801.5, 1794.5, 1801.2, 1807.3, 1796.3, 1801.7, 1797.5, 1794.5, 1797.7, 1798.5, 1794.2, 1797.2, 1800.4, 1794.9, 1797.5, 1797.1, 1796.3, 1796.8, 1805.7, 1798.7, 1796.7, 1799, 1793.3, 1798.4, 1809.6, 1793.2, 1791.8, 1796.8, 1797.1, 1800.7, 1797.3, 1799.5, 1807, 1802.7, 1800.9, 1790, 1805.9, 1802.6, 1800.8, 1795.9, 1801.9, 1807.8, 1798.4, 1795.6, 1789.7, 1801.9, 1797.9, 1799.5, 1805.8, 1804.8, 1795.7, 1797.4, 1801.8, 1788.5, 1802.3, 1800.3, 1798.3, 1802.6, 1794.8, 1799, 1794.5, 1800.2, 1799.8, 1798, 1797.7, 1804.7, 1800.9, 1799.5, 1797, 1796.4, 1801.5, 1801.4, 1801.8, 1813.7, 1796.4, 1805.4, 1797.6, 1805, 1792.6, 1809.8, 1799.8, 1804.8, 1798.9, 1801.4, 1798.1, 1802.4, 1803.7, 1796.9, 1808.8, 1798.9, 1795.8, 1803.8, 1796.2, 1797.2, 1797.5, 1803.2, 1803, 1805.4, 1796.7, 1795.1, 1796, 1794.5, 1801.7, 1803.5, 1806.5, 1799, 1793.6, 1812.2, 1809.6, 1801.2, 1802.2, 1802.4, 1796.5, 1812.8, 1795.2, 1793.1, 1804, 1799.9, 1797, 1799.3, 1801.6, 1798.1, 1794.8, 1801.5, 1810, 1799.2, 1798.5, 1795.7, 1792.9, 1801.8, 1803.4, 1798.9, 1801.5, 1804, 1802.3, 1797.1, 1795.7, 1797.8, 1801.3, 1796.6, 1800.4, 1798, 1803.4, 1805.7, 1800.6, 1801.4, 1801.8, 1798.7, 1805.9, 1794.8, 1800.6, 1795.7, 1811.2, 1798.3, 1792.4, 1801.8, 1805.6, 1806.9, 1798.4, 1797.9, 1795.4, 1801.2, 1800, 1801, 1793, 1797.4, 1805.6, 1793.9, 1794.9, 1797, 1795.3, 1801.7, 1799.7, 1805.1, 1799.5, 1796.9, 1797.9, 1795.7, 1794.4, 1796.4, 1797.1, 1798.4, 1804.3, 1800.2, 1799.1, 1792, 1811.3, 1802, 1798.8, 1804.9, 1792.8, 1792.4, 1791.4, 1794.7, 1796.6, 1793.8, 1799.2, 1794.7, 1799.9, 1790.1, 1795.1, 1799.6, 1795.7, 1796.5, 1800.1, 1801.6, 1797.2, 1790.5, 1799, 1794.3, 1803.5, 1794.1, 1794.5, 1800, 1801.8 };
            MtbColumn3.SetData(data3);

            MtbProj.ExecuteCommand(" Capa C3 304;   Lspec 1764;   Uspec 1836;   Pooled;   AMR;   UnBiased;   OBiased;   Toler 6;   Within;   Percent;   Title \"aaaaa\";   CStat.");
            Mtb.Graph MtbGraph = MtbProj.Commands.Item(1).Outputs.Item(1).Graph;
            MtbGraph.SaveAs("C:\\MyGraph" + DateTime.Now.ToString("yyyy-MM-dd HHmmssffff"), true, Mtb.MtbGraphFileTypes.GFPNGHighColor, 600, 400);

            MtbProj.ExecuteCommand("  Indplot ( C3 ) * C1;   Title \"bbbbb\";   Individual.");
            Mtb.Graph MtbGraph2 = MtbProj.Commands.Item(2).Outputs.Item(1).Graph;
            MtbGraph2.SaveAs("C:\\MyGraph" + DateTime.Now.ToString("yyyy-MM-dd HHmmssffff"), true, Mtb.MtbGraphFileTypes.GFPNGHighColor, 600, 400);

            MtbProj.ExecuteCommand("  Plot C3*C2;   Symbol C1;   Title \"ccccc\";   JITTER.");
            Mtb.Graph MtbGraph3 = MtbProj.Commands.Item(3).Outputs.Item(1).Graph;
            MtbGraph3.SaveAs("C:\\MyGraph" + DateTime.Now.ToString("yyyy-MM-dd HHmmssffff"), true, Mtb.MtbGraphFileTypes.GFPNGHighColor, 600, 400);

            MtbProj.ExecuteCommand(" PPlot C3;   Normal;   Symbol;   FitD;     NoCI;   Grid 2;   Grid 1;   MGrid 1;   Title \"dddddd\".");
            Mtb.Graph MtbGraph4 = MtbProj.Commands.Item(4).Outputs.Item(1).Graph;
            MtbGraph4.SaveAs("C:\\MyGraph" + DateTime.Now.ToString("yyyy-MM-dd HHmmssffff"), true, Mtb.MtbGraphFileTypes.GFPNGHighColor, 600, 400);

            MtbApp.Quit();
        }
    }
}
