using System;

using System.IO;


namespace Test
{
    public  class Class1
    {
       
 
     
        public static void SplitWorksheetToExcel(string srcExcelPath,string destExcelFolder)
        {


            #region 尝试使用线程 处理 数据 (待测试)


            //  VendorFilePath = 

            // 将总的数据 拆分 分别写 到 若干个 Excel 文件

            // 开5条线程

            //int totalCount = dtNewPOChangeList.Rows.Count;

            //int taskNum = 4;

            //int perCount = Convert.ToInt32(Math.Ceiling(totalCount*1.0 / taskNum));

            //TaskFactory taskFactory = new TaskFactory();
            //List<Task> taskList = new List<Task>();

            //for (int i=0; i<taskNum; i++)
            //{
            //    int perIndex = i;
            //  taskList.Add(  taskFactory.StartNew(() =>
            //    {
            //        DataTable dtPer = dtNewPOChangeList.Clone();
            //        perCount = perIndex == 3 ? totalCount - (perCount * perIndex) : perCount;
            //        dtPer.Rows.Add(dtNewPOChangeList.Rows.OfType<DataRow>().Skip(perCount* perIndex).Take( perCount ) );

            //        DealWithVendorData(dtPer, sheetNames[0], stripVendorList, "前导");
            //    }));
            //}


            //Task.WaitAll(taskList.ToArray());

            //// 先分组
            //var vendors = dtNewPOChangeList.Rows.OfType<DataRow>().GroupBy(t => t.Field<string>("A"));
            //// 获得组 key
            //string option = "仅前导";
            //// 遍历 Key 

            //ManualResetEvent reSet = new ManualResetEvent(true);
            //vendors.ToList().ForEach(vendor =>
            //{
            //    // 通过Key 查询数据集合
            //    bool doWork = true;     // 干事不 ? 
            //    var stripVendors = stripVendorList.Where(t => t.Field<string>("B") == vendor.Key).ToList();
            //    if (stripVendors.Count() > 0)
            //    {
            //        string backNote = stripVendors[0].Field<string>("G");
            //        switch (backNote)
            //        {
            //            case "仅前导":
            //                if (option == "后到")
            //                    doWork = false;         // 不干事
            //                break;
            //            case "仅后到":
            //                if (option == "前导")
            //                    doWork = false;         // 不干事
            //                break;
            //            case "忽略":
            //                doWork = false;             // 不干事
            //                break;
            //        }
            //    }

            //    if (doWork)
            //    {

            //        // 加入线程池
            //        ThreadPool.QueueUserWorkItem(state =>
            //        {
            //            string buyer = "";
            //            // 获取 Vendor 数据
            //            DataTable dtVendor = dtNewPOChangeList.Clone();  // vendor 供应商表
            //            var vendorData = dtNewPOChangeList.Rows.OfType<DataRow>().Where(t => t.Field<string>("A") == vendor.Key);
            //            vendorData.ToList().ForEach(t =>
            //            {
            //                buyer = t["B"].ToString();
            //                dtVendor.Rows.Add(t.ItemArray);
            //            });
            //            // 写入 Vendor 数据 至 Excel 中
            //            //string VendorReportFile = string.Format(@"{0}\{1}_Vendor_{2}_{3}{4}", remoteDir, buyer, vendor.Key, DateTime.Now.ToString("yyyyMMdd"), ".xlsx");
            //            string VendorReportFile = string.Format(@"{0}\{1}_Vendor_{2}_{3}{4}", Workstaion, buyer, vendor.Key, DateTime.Now.ToString("yyyyMMdd"), ".xlsx");

            //            if (!File.Exists(VendorReportFile))
            //            {   // 文件不存在
            //                File.Copy(ReportFile, VendorReportFile, true);       // Copy 模板文件到指定路径
            //            }
            //            else
            //            {

            //            }


            //            //IOExcel ioExcel = new IOExcel();

            //          //  ioExcel.SetCellValue(VendorReportFile, sheetNames[0], 1, 14, DateTime.Now.ToString("dd/MM/yyyy"), "");// 设置供应商名称 

            //            // SetCellValue(VendorReportFile, 1, 2, DateTime.Now.ToString("yyyy/MM/dd"));
            //            reSet.WaitOne();

            //            reSet.Reset();

            //            IOExcel.SetCellValue(VendorReportFile, sheetNames[0], 1, 11, vendor.Key);                           // 设置供应商名称 
            //            iOExcel.ExportDataTableToExcelByRange(dtVendor, VendorReportFile, sheetNames[0]);

            //            reSet.Set();

            //        });

            //    }
            //});

            #endregion
            object missing = Type.Missing;
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Open(srcExcelPath, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            try
            {
                if (workBook != null)
                {
                    for (int i = 1; i < workBook.Sheets.Count; i++)
                    {
                        Microsoft.Office.Interop.Excel.Worksheet _wSheets = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets[i];
                        Microsoft.Office.Interop.Excel.Workbook newBook = app.Workbooks.Add(missing);
                        Microsoft.Office.Interop.Excel.Worksheet mySheet = newBook.Sheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
                        try
                        {


                            string filename = string.Format(@"{0}\{1}_{2}.xlsx", destExcelFolder, Path.GetFileNameWithoutExtension(srcExcelPath), _wSheets.Name); // @"D:\temp\" + _wSheets.Name + ".xls";
                            if (File.Exists(filename))
                            {
                                File.Delete(filename);
                            }
                            //mySheet.Name = _wSheets.Name;
                            _wSheets.Copy(mySheet, missing);
                            newBook.SaveAs(filename, missing
                                , missing, missing, missing, missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, missing
                                , missing, missing, missing, missing);
                        }
                        catch (Exception ex)
                        {

                        }
                        finally
                        {
                            _wSheets = null;
                            mySheet = null;
                            newBook.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                workBook.Close();
                app.Quit();
                app = null;
            }
        }

       
    }
}
