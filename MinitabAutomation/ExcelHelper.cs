using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace MinitabAutomation
{
    public class ExcelHelper : IDisposable
    {
        private string fileName = null; //文件名
        private IWorkbook workbook = null;
        private FileStream fs = null;
        private bool disposed;

        public ExcelHelper(string fileName)
        {
            this.fileName = fileName;
            disposed = false;
        }

        /// <summary>
        /// 将DataTable数据导入到excel中
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="isColumnWritten">DataTable的列名是否要导入</param>
        /// <param name="sheetName">要导入的excel的sheet的名称</param>
        /// <returns>导入数据行数(包含列名那一行)</returns>
        public int DataTableToExcel(DataTable data, string sheetName, bool isColumnWritten)
        {
            int i = 0;
            int j = 0;
            int count = 0;
            ISheet sheet = null;

            fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                workbook = new XSSFWorkbook();
            else if (fileName.IndexOf(".xls") > 0) // 2003版本
                workbook = new HSSFWorkbook();

            try
            {
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet(sheetName);
                }
                else
                {
                    return -1;
                }

                if (isColumnWritten == true) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                    }
                    count = 1;
                }
                else
                {
                    count = 0;
                }

                for (i = 0; i < data.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                    }
                    ++count;
                }
                workbook.Write(fs); //写入到excel
                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return -1;
            }
        }

        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public DataTable ExcelToDataTable(string sheetName, bool isFirstRowColumn)
        {
            ISheet sheet = null;
            DataTable data = new DataTable();
            int startRow = 0;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                    workbook = new XSSFWorkbook(fs);
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook(fs);

                if (sheetName != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                    if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    {
                        sheet = workbook.GetSheetAt(0);
                    }
                }
                else
                {
                    sheet = workbook.GetSheetAt(0);
                }
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

                    if (isFirstRowColumn)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                string cellValue = cell.StringCellValue;
                                if (cellValue != null)
                                {
                                    DataColumn column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                            }
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }

                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　　　　　　　

                        DataRow dataRow = data.NewRow();
                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {
                            if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                                dataRow[j] = row.GetCell(j).ToString();
                        }
                        data.Rows.Add(dataRow);
                    }
                }

                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return null;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    if (fs != null)
                        fs.Close();
                }

                fs = null;
                disposed = true;
            }
        }

        public Models.RowData getRowData(string sheetName)
        {
            ISheet sheet = null;
            Models.RowData modelRowData = new Models.RowData();
            try
            {
                modelRowData.filePath = fileName.Substring(0, fileName.LastIndexOf("."));
                fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite);
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                    workbook = new XSSFWorkbook(fs);
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook(fs);

                if (sheetName != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                    if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    {
                        sheet = workbook.GetSheetAt(0);
                    }
                }
                else
                {
                    sheet = workbook.GetSheetAt(0);
                }
                if (sheet != null)
                {
                    int startRow = 14;
                    int startColumn = 9;

                    //获取实例的个数
                    IRow titleRow = sheet.GetRow(startRow);
                    int columnCount = titleRow.LastCellNum;
                    //测试专用，用于减小列数
                    //columnCount = columnCount > 14 ? 14 : columnCount;
                    //获取数据行数
                    int rowCount = sheet.LastRowNum;
                    for (int i = startColumn; i < columnCount; i++)
                    {
                        modelRowData.instances.Add(new Models.Instance());
                        //添加title
                        if (titleRow.GetCell(i) != null)
                        {
                            ((Models.Instance)modelRowData.instances[i - startColumn]).title = titleRow.GetCell(i).ToString();
                        }
                    }
                    //更新实例信息

                    //添加限制类型
                    IRow limTypeRow = sheet.GetRow(7);
                    for (int i = startColumn; i < columnCount; i++)
                    {
                        if (limTypeRow.GetCell(i) != null)
                        {
                            ((Models.Instance)modelRowData.instances[i - startColumn]).limType = limTypeRow.GetCell(i).ToString();
                        }
                    }

                    //添加下限
                    IRow lowerLimitRow = sheet.GetRow(8);
                    for (int i = startColumn; i < columnCount; i++)
                    {
                        if (lowerLimitRow.GetCell(i) != null)
                        {
                            try
                            {
                                ((Models.Instance)modelRowData.instances[i - startColumn]).lowerLimit = Double.Parse(lowerLimitRow.GetCell(i).ToString());
                            }
                            catch (Exception exx)
                            { 
                            
                            }
                        }
                    }

                    //添加上限
                    IRow upLimitRow = sheet.GetRow(9);
                    for (int i = startColumn; i < columnCount; i++)
                    {
                        if (upLimitRow.GetCell(i) != null)
                        {
                            try
                            {
                            ((Models.Instance)modelRowData.instances[i - startColumn]).upLimit = Double.Parse(upLimitRow.GetCell(i).ToString());
                            }
                            catch (Exception exx)
                            {

                            }
                        }
                    }

                    //添加名称
                    IRow nameRow = sheet.GetRow(10);
                    for (int i = startColumn; i < columnCount; i++)
                    {
                        if (nameRow.GetCell(i) != null)
                        {
                            ((Models.Instance)modelRowData.instances[i - startColumn]).name = nameRow.GetCell(i).ToString();
                        }
                    }

                    //添加单位
                    IRow unitRow = sheet.GetRow(11);
                    for (int i = startColumn; i < columnCount; i++)
                    {
                        if (unitRow.GetCell(i) != null)
                        {
                            ((Models.Instance)modelRowData.instances[i - startColumn]).unit = unitRow.GetCell(i).ToString();
                        }
                    }

                    //更新实例数据信息
                    for (int i = startRow+1; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　　　　　　　
                        //Node
                        if (row.GetCell(3) != null)
                        {
                            try
                            {
                                modelRowData.node.Add(float.Parse(row.GetCell(3).ToString()));
                            }
                            catch {
                                modelRowData.node.Add(null);
                            }
                        }
                        else {
                            modelRowData.node.Add(null);
                        }

                        //Datetime
                        if (row.GetCell(8) != null)
                        {
                            try
                            {
                                modelRowData.dataTime.Add(DateTime.Parse(row.GetCell(8).ToString()));
                            }
                            catch {
                                modelRowData.dataTime.Add(null);
                            }
                        }
                        else
                        {
                            modelRowData.dataTime.Add(null);
                        }

                        //instances
                        for (int j = startColumn; j < columnCount; j++)
                        {
                            if (row.GetCell(j) != null)
                            {
                                try
                                {
                                    ((Models.Instance)modelRowData.instances[j - startColumn]).data.Add(row.GetCell(j).ToString());
                                }
                                catch
                                {
                                    ((Models.Instance)modelRowData.instances[j - startColumn]).data.Add(null);
                                }
                                
                            }
                            else
                            {
                                ((Models.Instance)modelRowData.instances[j - startColumn]).data.Add(null);
                            }
                        }
                    }
                }
                fs.Close();
                return modelRowData;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                fs.Close();
                return null;
            }
        }

        public void insertPicture(string sheetName, Models.RowData modelRowData, System.Windows.Forms.TextBox textBox)
        {
            //删除空白的实例
            for (int i = modelRowData.instances.Count - 1; i >= 0; i--)
            {
                bool isDelete = true;
                Models.Instance model_Instance = (Models.Instance)modelRowData.instances[i];
                for (int j = 0; j < model_Instance.pictures.Count; j++)
                {
                    if (model_Instance.pictures[j] != null)
                    {
                        isDelete = false;
                    }
                }
                if (isDelete)
                {
                    modelRowData.instances.RemoveAt(i);
                }
            }
            //delete sheet if exist
            ISheet sheet = null;
            sheet = workbook.GetSheet(sheetName);
            if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
            {
                sheet = workbook.CreateSheet(sheetName);
            }

            HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();

            for (int i = 0; i < modelRowData.instances.Count; i++)
            {
                Models.Instance model_Instance = (Models.Instance)modelRowData.instances[i];
                for (int j = 0; j < model_Instance.pictures.Count; j++)
                {
                    if(model_Instance.pictures[j] != null)
                    {
                        string picturePath = model_Instance.pictures[j].ToString();
                        //读取图片
                        byte[] bytes = System.IO.File.ReadAllBytes(picturePath);
                        int pictureIdx = workbook.AddPicture(bytes, PictureType.PNG);
                        //add a picture
                        HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 255, 255, 4 + (j * 8), 2 + (i * 21), 14 + (j * 8), 10 + (i * 21));
                        HSSFPicture pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);
                        pict.Resize();
                        textBox.AppendText("    " + picturePath.Substring(picturePath.LastIndexOf("\\")+1) + "\r\n");
                    }
                }
            }

            

            fs = new FileStream(fileName, FileMode.Open, FileAccess.Write);
            workbook.Write(fs);
            fs.Close();
            textBox.AppendText("插入图片完成\r\n");
        }

    }
}
