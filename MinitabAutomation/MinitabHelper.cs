using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace MinitabAutomation
{
    public class MinitabHelper
    {
        public void GeneratePictures(Models.RowData modelRowData, System.Windows.Forms.TextBox textBox)
        {
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
            //创建图片文件夹
            textBox.AppendText("新建图片文件夹：" + modelRowData.filePath + "\r\n");
            if (!Directory.Exists(modelRowData.filePath))//判断文件夹是否已经存在
            {
                Directory.CreateDirectory(modelRowData.filePath);//创建文件夹
            }
            textBox.AppendText("正在生成图片...\r\n");
            foreach (Models.Instance modelInstance in modelRowData.instances)
            {
                try
                {
                    textBox.AppendText("    " + modelInstance.title + "    ");
                    GeneratePicturesInstance(MtbApp, modelInstance, modelRowData);
                    textBox.AppendText("成功\r\n");
                }
                catch
                {
                    Mtb.Project MtbProj = MtbApp.ActiveProject;
                    MtbProj.Delete();
                    MtbApp.New();
                    textBox.AppendText("失败\r\n");
                }

            }
            textBox.AppendText("图片生成完成\r\n");
            MtbApp.Quit();
        }

        private ArrayList column1 = new ArrayList();
        private ArrayList column2 = new ArrayList();
        private ArrayList column3 = new ArrayList();

        public void parseColumn(Models.Instance modelInstance, Models.RowData modelRowData)
        {
            column1 = new ArrayList();
            column2 = new ArrayList();
            column3 = new ArrayList();
            for (int i=0; i< modelInstance.data.Count; i++)
            {
                if (modelInstance.data[i].ToString().Trim() != "")
                {
                    try
                    {
                        Double data = Double.Parse(modelInstance.data[i].ToString().Trim());
                        column1.Add(modelRowData.node[i]);
                        column2.Add(modelRowData.dataTime[i]);
                        column3.Add(data);
                    }
                    catch
                    {
                        
                    }
                }
            }
        }

        public void GeneratePicturesInstance(Mtb.Application MtbApp, Models.Instance modelInstance, Models.RowData modelRowData)
        {
            Mtb.Project MtbProj = MtbApp.ActiveProject;

            parseColumn(modelInstance, modelRowData);

            Mtb.Columns MtbColumns = MtbProj.ActiveWorksheet.Columns;
            Mtb.Column MtbColumn1 = MtbColumns.Add(null, null, 1);
            MtbColumn1.SetData(column1.ToArray());

            Mtb.Column MtbColumn2 = MtbColumns.Add(null, null, 1);
            MtbColumn2.SetData(column2.ToArray());

            Mtb.Column MtbColumn3 = MtbColumns.Add(null, null, 1);
            MtbColumn3.SetData(column3.ToArray());

            try
            {
                string imgPath = Path.Combine(modelRowData.filePath, modelInstance.title + " Process Capability");
                MtbProj.ExecuteCommand(" Capa C3 " + column1.Count + ";   Lspec " + modelInstance.lowerLimit + ";   Uspec " + modelInstance.upLimit + ";   Pooled;   AMR;   UnBiased;   OBiased;   Toler 6;   Within;   Percent;   Title \"" + getPictureTitle(0, modelInstance) + "\";   CStat.");
                Mtb.Graph MtbGraph = MtbProj.Commands.Item(MtbProj.Commands.Count).Outputs.Item(1).Graph;
                MtbGraph.SaveAs(imgPath, true, Mtb.MtbGraphFileTypes.GFPNGHighColor, 512, 354);
                modelInstance.pictures.Add(imgPath + ".png");
            }
            catch
            {
                modelInstance.pictures.Add(null);
            }

            try
            {
                string imgPath = Path.Combine(modelRowData.filePath, modelInstance.title + " Individual Polt");
                MtbProj.ExecuteCommand("  Indplot ( C3 ) * C1;   Title \"" + getPictureTitle(1, modelInstance) + "\";   Individual.");
                Mtb.Graph MtbGraph2 = MtbProj.Commands.Item(MtbProj.Commands.Count).Outputs.Item(1).Graph;
                MtbGraph2.SaveAs(imgPath, true, Mtb.MtbGraphFileTypes.GFPNGHighColor, 512, 354);
                modelInstance.pictures.Add(imgPath + ".png");
            }
            catch
            {
                modelInstance.pictures.Add(null);
            }
            try
            {
                string imgPath = Path.Combine(modelRowData.filePath, modelInstance.title + " Scatter Plot");
                MtbProj.ExecuteCommand("  Plot C3*C2;   Symbol C1;   Title \"" + getPictureTitle(2, modelInstance) + "\";   JITTER.");
                Mtb.Graph MtbGraph3 = MtbProj.Commands.Item(MtbProj.Commands.Count).Outputs.Item(1).Graph;
                MtbGraph3.SaveAs(imgPath, true, Mtb.MtbGraphFileTypes.GFPNGHighColor, 512, 354);
                modelInstance.pictures.Add(imgPath + ".png");
            }
            catch
            {
                modelInstance.pictures.Add(null);
            }
            try
            {
                string imgPath = Path.Combine(modelRowData.filePath, modelInstance.title + " Probability Plot");
                MtbProj.ExecuteCommand(" PPlot C3;   Normal;   Symbol;   FitD;     NoCI;   Grid 2;   Grid 1;   MGrid 1;   Title \"" + getPictureTitle(3, modelInstance) + "\".");
                Mtb.Graph MtbGraph4 = MtbProj.Commands.Item(MtbProj.Commands.Count).Outputs.Item(1).Graph;
                MtbGraph4.SaveAs(imgPath, true, Mtb.MtbGraphFileTypes.GFPNGHighColor, 512, 354);
                modelInstance.pictures.Add(imgPath + ".png");
            }
            catch
            {
                modelInstance.pictures.Add(null);
            }
            MtbProj.Delete();
            MtbApp.New();
        }

        public string getPictureTitle(int Type, Models.Instance modelInstance)
        {
            string title = "";
            if (Type == 0)
            {
                title += modelInstance.name + " : " + modelInstance.title + " : L=" + modelInstance.lowerLimit + " H=" + modelInstance.upLimit + " (" + modelInstance.unit + ")";
            }
            else if (Type == 1)
            {
                title += "Individual Polt of " + modelInstance.title + " per Node w avg conf";
            }
            else if (Type == 2)
            {
                title += "Scatter Plot of " + modelInstance.title + " vs Date Time";
            }
            else if (Type == 3)
            {
                title += "Probability Plot of " + modelInstance.title + "";
            }
            return "用于演示 " + title.Replace("\"", "_");
        }

       
    }
}
