using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
//using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Drawing;
//using System.Text;
//using System.Threading.Tasks;

namespace SWS
{
    public class IOExcel
    {

        //public static 
        private static Microsoft.Office.Interop.Excel.Application Excel;//  = new Microsoft.Office.Interop.Excel.Application();
        private static Microsoft.Office.Interop.Excel.Workbook ExcelBook;
        private static Microsoft.Office.Interop.Excel.Worksheet ExcelSheet;


        public static void Open(string fileName)
        {
            Excel = new Microsoft.Office.Interop.Excel.Application();
            ExcelBook = Excel.Workbooks.Open(fileName, Type.Missing);
            Excel.Visible = false;
        }

        public static void Open(string fileName, string sheetName)
        {
            Excel = new Microsoft.Office.Interop.Excel.Application();
            ExcelBook = Excel.Workbooks.Open(fileName, Type.Missing);
            ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Sheets[sheetName];

            ExcelSheet.Select();           
            Excel.Visible = false;
            ExcelSheet.Activate();

        }

     


        public static void SaveClose()
        {
            Excel.DisplayAlerts = false;
            Excel.AlertBeforeOverwriting = false;

            object misValue = System.Reflection.Missing.Value;

            
            ExcelBook.Save();
            ExcelBook.Saved = true;


            ExcelBook.Close(true, misValue, misValue);

            Excel.Quit();

            PublicMethod.Kill(Excel);//调用kill当前Excel进程  

            releaseObject(ExcelSheet);

            releaseObject(ExcelBook);

            releaseObject(Excel);


            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel);
            Excel = null;
            GC.Collect();
        }


        public static void SaveClose(object objectLock)
        {
            Excel.DisplayAlerts = false;
            Excel.AlertBeforeOverwriting = false;

            object misValue = System.Reflection.Missing.Value;


            ExcelBook.Save();
            ExcelBook.Saved = true;


            ExcelBook.Close(true, misValue, misValue);

            Excel.Quit();

            PublicMethod.Kill(Excel);//调用kill当前Excel进程  

            releaseObject(ExcelSheet);

            releaseObject(ExcelBook);

            releaseObject(Excel);


            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel);
            Excel = null;
            GC.Collect();
        }


        public static void SetCellValue(string fileName, string sheetName,int cellX,int cellY,string value)
        {
            Open(fileName, sheetName);

            try
            {
                // DataTable dt = dataTable;

                ExcelSheet.Activate();
                ExcelSheet.Cells[cellX, cellY] = value;

             
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {

                SaveClose();

            }
        }

        public  void SetCellValue(string fileName, string sheetName, int cellX, int cellY, string value,object lockObject)
        {
            // Open(fileName, sheetName);

            

        Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbook excelBook = excel.Workbooks.Open(fileName, Type.Missing);
        Microsoft.Office.Interop.Excel.Worksheet excelSheet  = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets[sheetName];

            excelSheet.Select();
            excel.Visible = false;
            excelSheet.Activate();

            try
            {
                // DataTable dt = dataTable;

                excelSheet.Activate();
                excelSheet.Cells[cellX, cellY] = value;


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {

                // SaveClose();

                excel.DisplayAlerts = false;
                excel.AlertBeforeOverwriting = false;

                object misValue = System.Reflection.Missing.Value;


                excelBook.Save();
                excelBook.Saved = true;


                excelBook.Close(true, misValue, misValue);

                excel.Quit();

                PublicMethod.Kill(excel);//调用kill当前Excel进程  

                releaseObject(excelSheet);

                releaseObject(excelBook);

                releaseObject(excel);


                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                excel = null;
                GC.Collect();

            }
        }

        /// <summary>
        /// 给Excel插入一列
        /// </summary>
        /// <param name="fileName">文件路径</param>
        /// <param name="sheetName">Excel工作表</param>
        /// <param name="newColumnIndex">待插入的列序号</param>
        /// <returns></returns>
        public DataTable InserOneColumn(string fileName, string sheetName, out int newColumnIndex)
        {

            // 1.确定要插入列的位置         
            DataTable excelData = (this.GetFileDataSet(fileName)).Tables[sheetName];
            DataRow firstRow = excelData.Rows[0];
            newColumnIndex = 1;
            foreach (DataColumn item in excelData.Columns)
            {
                if (string.IsNullOrEmpty(firstRow[item.ColumnName].ToString()))
                    break;
                newColumnIndex++;
            }

            // 2.获取最后一列 Range          
            string str = (sheetName.Contains("$") && sheetName.Contains("'")) ? sheetName.Replace("$", "").Replace("'", "") : sheetName;
            Open(fileName, str);
            Microsoft.Office.Interop.Excel.Range xlsxColumns = (Microsoft.Office.Interop.Excel.Range)ExcelSheet.Columns[newColumnIndex, System.Type.Missing];

            // 3.执行插入一列的操作
            xlsxColumns.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftToRight, Type.Missing);

            SaveClose();
                           
            // 4.返回插入后的数据
            excelData = (this.GetFileDataSet(fileName)).Tables[sheetName];

            return excelData;
        }

        /// <summary>
        /// 读取Excel文件数据
        /// </summary>
        /// <param name="fileName">文件路径</param>
        /// <returns></returns>
        public DataSet GetFileDataSet(string fileName)
        {
            DataSet ds = new DataSet();
            if (!String.IsNullOrEmpty(fileName))
            {
                string connStr = "";
                string fileType = System.IO.Path.GetExtension(fileName);
                if (string.IsNullOrEmpty(fileType)) return null;
                string path = fileName;

                if (fileType == ".xls")
                {
                    connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + path + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
                }
                else
                {
                    connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + path + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
                }
                OleDbConnection conn = null;
                OleDbDataAdapter da = null;
                DataTable dtSheetName = null;
                string sql_F = "Select * FROM [{0}]";

                try
                {
                    // 初始化连接，并打开 
                    conn = new OleDbConnection(connStr);
                    conn.Open();
                    // 获取数据源的表定义元数据      
                    string SheetName = "";
                    dtSheetName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                    // 初始化适配器  
                    da = new OleDbDataAdapter();
                    for (int i = 0; i < dtSheetName.Rows.Count; i++)
                    {

                        SheetName = (string)dtSheetName.Rows[i]["TABLE_NAME"];
                        if (SheetName.Contains("$") && !SheetName.Replace("'", "").EndsWith("$"))
                        {
                            continue;
                        }

                        da.SelectCommand = new OleDbCommand(String.Format(sql_F, SheetName), conn);
                        DataSet dsItem = new DataSet();
                        da.Fill(dsItem, SheetName);
                        ds.Tables.Add(dsItem.Tables[0].Copy());
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }

                finally
                {
                    // 关闭连接  
                    if (conn.State == ConnectionState.Open)
                    {
                        conn.Close();
                        da.Dispose();
                        conn.Dispose();
                    }
                }
            }
            return ds;
        }



        /// <summary>
        /// 循环去除datatable中的空行(静态方法)
        /// </summary>
        /// <param name="dt"></param>
        public static void RemoveEmpty(DataTable dt)
        {
            List<DataRow> removelist = new List<DataRow>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                bool rowdataisnull = true;
                for (int j = 0; j < dt.Columns.Count; j++)
                {

                    if (!string.IsNullOrEmpty(dt.Rows[i][j].ToString().Trim()))
                    {

                        rowdataisnull = false;
                    }

                }
                if (rowdataisnull)
                {
                    removelist.Add(dt.Rows[i]);
                }

            }
            for (int i = 0; i < removelist.Count; i++)
            {
                dt.Rows.Remove(removelist[i]);
            }
        }

        /// <summary>
        /// 该方法需要引用Microsoft.Office.Interop.Excel.dll 和 Microsoft.CSharp.dll
        /// 将数据由DataTable导出到Excel
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="fileName"></param>
        /// <param name="filePath"></param>
       

        //将DataTable写入已存在Excel
        public void ExportDataTableToExcelByRange(DataTable dataTable, string fileName, string sheetName,int cellRowStartIndex=2)
        {


            Open(fileName, sheetName);

            try
            {
                DataTable dt = dataTable;
                DataTableToExcel(dt, ExcelSheet,cellRowStartIndex);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {

                SaveClose();

            }
        }
        public void DataTableToExcel(DataTable dt, Microsoft.Office.Interop.Excel.Worksheet excelSheet, int cellRowStartIndex = 2)
        {
            // excelSheet.Rows.Clear();         
            excelSheet.Activate();

            #region 导入数据方式(直接对单元格赋值)
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)  // j=2 跳过前 2 列
                {
                    // dataArray[i + 1, j] = dt.Rows[i][j].ToString();
                    if (dt.Rows[i][j] is "")
                        continue;
                    string value = dt.Rows[i][j].ToString();
                    excelSheet.Cells[i + cellRowStartIndex, j + 1] = value;
                }
            }
            #endregion

        }

        //将DataTable写入已存在Excel
        public void ExportDataTableToExcelByRange(DataTable dataTable, string fileName, string sheetName)
        {


            Open(fileName, sheetName);

            try
            {
                DataTable dt = dataTable;
                DataTableToExcel(dt, ExcelSheet);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {

                SaveClose();

            }
        }
        public void DataTableToExcel(DataTable dt, Microsoft.Office.Interop.Excel.Worksheet excelSheet)
        {
            #region 导入数据方式(通过数组)
            // 通过 Range 方法整体赋值

            int rowCount = dt.Rows.Count;
            int colCount = dt.Columns.Count;
            object[,] dataArray = new object[rowCount, colCount];

            // 赋值行数据
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < colCount; j++)
                {
                    dataArray[i, j] = dt.Rows[i][j].ToString();
                }
            }

            // 从 Excel 第3行 开始写入
            
            excelSheet.Range[excelSheet.Cells[3, 1], excelSheet.Cells[3 + rowCount - 1, colCount]].Value2 = dataArray;
            ////////////////////////////////////////////////////////////////////////////////////



            #endregion

        }


        public void ExportDataTableToExcelByRange(object[,] dataArray, string fileName, string sheetName)
        {


            Open(fileName, sheetName);

            try
            {
             
                DataTableToExcel(dataArray, ExcelSheet);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {

                SaveClose();

            }
        }
        public void DataTableToExcel(object[,] dataArray, Microsoft.Office.Interop.Excel.Worksheet excelSheet)
        {
            #region 导入数据方式(通过数组)
            // 通过 Range 方法整体赋值

            //int rowCount = dt.Rows.Count;
            //int colCount = dt.Columns.Count;
            //object[,] dataArray = new object[rowCount, colCount];

            // 赋值行数据
            //for (int i = 0; i < rowCount; i++)
            //{
            //    for (int j = 0; j < colCount; j++)
            //    {
            //        dataArray[i, j] = dt.Rows[i][j].ToString();
            //    }
            //}
        
            int rowCount = dataArray.GetLength(0);//获取维数，这里指行数
            int colCount = dataArray.GetLength(1); //获取指定维度中的元素个数，这里也就是列数了。（0是第一维，1表示的是第二维）

            // 从 Excel 第3行 开始写入
            excelSheet.Range[excelSheet.Cells[1, 11], excelSheet.Cells[1, 11+colCount-1]].Value2 = dataArray;
            ////////////////////////////////////////////////////////////////////////////////////



            #endregion

        }


        public void ExportDataTableToExcelByRange(DataTable dataTable,object[,] dataArray, string fileName, string sheetName)
        {


            Open(fileName, sheetName);

            try
            {

                DataTableToExcel(dataTable,dataArray, ExcelSheet);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {

                SaveClose();

            }
        }
        public void DataTableToExcel(DataTable dt,object[,] dataTitle, Microsoft.Office.Interop.Excel.Worksheet excelSheet)
        {
            #region 导入数据方式(通过数组)
            // 通过 Range 方法整体赋值

            int rowCount = dt.Rows.Count;
            int colCount = dt.Columns.Count;
            object[,] dataArray = new object[rowCount, colCount];

            // 赋值行数据
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < colCount; j++)
                {
                    dataArray[i, j] = dt.Rows[i][j].ToString();
                }
            }

            // 从 Excel 第3行 开始写入
            excelSheet.Range[excelSheet.Cells[3, 1], excelSheet.Cells[3+ rowCount-1 , colCount]].Value2 = dataArray;
            ////////////////////////////////////////////////////////////////////////////////////

            int rowTitle = dataTitle.GetLength(0);//获取维数，这里指行数
            int colTitle = dataTitle.GetLength(1); //获取指定维度中的元素个数，这里也就是列数了。（0是第一维，1表示的是第二维）

            // 从 Excel 第1行 开始写入
            excelSheet.Range[excelSheet.Cells[1, 11], excelSheet.Cells[1, 11 + colTitle - 1]].Value2 = dataTitle;

            #endregion

        }

        /// <summary>
        /// 获取Excel Sheet名称
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static List<string> GetAllSheetName(string fileName)
        {
            Open(fileName);

            List<string> list = new List<string>();
            try
            {
                foreach (Microsoft.Office.Interop.Excel.Worksheet item in ExcelBook.Worksheets)
                {
                    list.Add(item.Name);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                SaveClose();
            }

            return list;
        }


        /// <summary>
        /// 计算单元格的地址
        /// </summary>
        /// <param name="row">单元格行</param>
        /// <param name="col">单元格列</param>
        /// <returns>返回地址</returns>
        public static string GetExcelCellName(int col,int row=0 ) 
        {
            string CellName = GetExcelColumnName(col);
            if (CellName.Length > 2)
            {
                CellName = string.Format("{0}{1}", CellName[0], CellName[CellName.Length - 1]);
            }
            
            
            return row==0? CellName: CellName + row.ToString();
        }

        public static string GetExcelColumnName(int col)
        {
            string colNameString = "";

            int num = (col - 1) / 26;

            char Chr;
            if (num == 0)
            {
                Chr = Convert.ToChar(col - 1 + 65);  // ASCII A --- 65
                return Chr.ToString();
            }

            colNameString = Convert.ToChar(num - 1 + 65).ToString();
            colNameString += GetExcelColumnName(col - 26);

            return colNameString;

        }

        /// <summary>
        /// 释放COM组件对象
        /// </summary>
        /// <param name="obj"></param>
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        /// <summary>
        /// 关闭进程的内部类
        /// </summary>
        public class PublicMethod
        {
            [DllImport("User32.dll", CharSet = CharSet.Auto)]

            public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);

            public static void Kill(Microsoft.Office.Interop.Excel.Application excel)
            {
                //如果外层没有try catch方法这个地方需要抛异常。
                IntPtr t = new IntPtr(excel.Hwnd);

                int k = 0;

                GetWindowThreadProcessId(t, out k);

                System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);

                p.Kill();
            }
        }
    }
}