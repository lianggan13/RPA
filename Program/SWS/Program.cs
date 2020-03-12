using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace SWS
{
    class Program
    {
       
        ////////////////////////////////////
        ///
        private static string Workstaion;
        private static string POChangeListFiles;
        private static string ForecastChangeListFiles;
        private static string ReportTempFilePath;
        private static string ReportFile;
        private static string VendorFilePath;
        private static string LastWeekReportFile;
        private static string POCancelFile="";



        private static IOExcel iOExcel = new IOExcel();

        static void Main(string[] args)
        {

         

            if (args.Count() > 0)
            {
                string param = args[0].ToString().Replace("*", " ");
                Workstaion = param.Split('|')[0].ToString();
                POChangeListFiles = param.Split('|')[1].ToString().TrimStart(new char[] { ';' });
                ForecastChangeListFiles = param.Split('|')[2].ToString().TrimStart(new char[] { ';' });
                ReportTempFilePath = param.Split('|')[3].ToString();     

                VendorFilePath = param.Split('|')[4].ToString();    // RemoteFile
                LastWeekReportFile = param.Split('|')[5].ToString();    // RemoteFile 
                 if(param.Split('|').Count() == 7)
                    POCancelFile = param.Split('|')[6].ToString();

               // Console.WriteLine(POCancelFile, args.Count());
            }
            else
            {


                Workstaion = @"C:\Users\mac\Desktop\RPA\CodeTest";
                // POChangeListFiles = "J1_OrderHistorySearch20190426;A0_OrderHistorySearch20190426;S0_OrderHistorySearch20190426;H0_OrderHistorySearch20190426;C1_OrderHistorySearch20190426;C0_OrderHistorySearch20190426";
                POChangeListFiles = ";C1_OrderHistorySearch20200312.xls;A0_OrderHistorySearch20200312.xls;C0_OrderHistorySearch20200312.xls";
                POChangeListFiles = POChangeListFiles.Trim(new char[] { ';' });
                ForecastChangeListFiles = ";C1_Comg0010Download20200312.xls;A0_Comg0010Download20200312.xls;C0_Comg0010Download20200312.xls"; //A0_Comg0010Download20190426;S0_Comg0010Download20190426;H0_Comg0010Download20190426;C1_Comg0010Download20190426;C0_Comg0010Download20190426;";
                ForecastChangeListFiles = ForecastChangeListFiles.Trim(new char[] { ';' });
                ReportTempFilePath = @"C:\Users\mac\Desktop\RPA\file\Remote\自动生成模板.xls";             
                LastWeekReportFile = @"C:\Users\mac\Desktop\RPA\file\Remote\自动生成模板_20200228.xls";

                // POCancelFile
                POCancelFile = @"C:\Users\mac\Desktop\RPA\file\Remote\POCancle_20200312.xls";

            }



            // Step1: 复制Report模板文件到工作目录中
            Console.WriteLine("Step1.复制Report模板文件 [{0}] 到工作目录中 [{1}] ...",Path.GetFileName(ReportTempFilePath),Workstaion);            
            ReportFile =  string.Format(@"{0}\临时模板{1}", Workstaion, ".xls");
            File.Copy(ReportTempFilePath, ReportFile, true);       // Copy 模板文件到指定路径


            //Step2. 读取文件集合 --> dtPOChangeList,dtForecastChangeList
            Console.WriteLine("Step2.读取文件集合 [{0}] [{1}] 数据", POChangeListFiles, ForecastChangeListFiles);
            DataTable dtPOChangeList = ReadFilesToDataTable(POChangeListFiles);     //IOExcel RemoveEmpty
           DataTable dtForecastChangeList = ReadFilesToDataTable(ForecastChangeListFiles,true);

           
            //Step3. 整理文件集合数据
            Console.WriteLine("Step3.整理文件集合数据");
            List<string> sheetNames = IOExcel.GetAllSheetName(ReportFile);
            DataTable dtOldPOChangeList = (iOExcel.GetFileDataSet(ReportFile)).Tables[string.Format("{0}$", sheetNames[0])];
            DataTable dtOldForecastChangeList = (iOExcel.GetFileDataSet(ReportFile)).Tables[string.Format("{0}$", sheetNames[1])];      // Sheet1 (2)$
            if (dtOldForecastChangeList is null)
            {
                dtOldForecastChangeList = (iOExcel.GetFileDataSet(ReportFile)).Tables[string.Format("'{0}$'", sheetNames[1])]; // 'Sheet1 (2)$'
            }
            dtOldForecastChangeList = ChangeDtColumnTypeName(dtOldForecastChangeList);    // 改变 字段类型 和 名称

            // ******PO Change List******
            DataTable dtNewPOChangeList = MakeNewPOChangeList(dtOldPOChangeList, dtPOChangeList);
            dtNewPOChangeList = ChangeDtColumnTypeName(dtNewPOChangeList);
             
            // ******Forecast Change List******
             DataTable dtNewForecastChangeList = MakeNewForecastChangeList(dtOldForecastChangeList, dtForecastChangeList);
            dtNewForecastChangeList = ChangeDtColumnTypeName(dtNewForecastChangeList);


            // step4：读取上一周Rport文件   
            Console.WriteLine("step4.读取上一周Report文件 [{0}]", Path.GetFileName(LastWeekReportFile));
            sheetNames = IOExcel.GetAllSheetName(LastWeekReportFile);
            dtOldPOChangeList = (iOExcel.GetFileDataSet(LastWeekReportFile)).Tables[string.Format("{0}$", sheetNames[0])];
            dtOldPOChangeList = ChangeDtColumnTypeName(dtOldPOChangeList);

            
           


            // 匹配 
            Console.WriteLine("Step5-1.将文件集合数据 与 上一周Report文件数据 进行匹配 并写入到新的Report文件 [{0}] 中", Path.GetFileName(ReportFile));
            MatchPOChangeList(dtOldPOChangeList, dtNewPOChangeList);                 // 匹配 Status_1        //MatchPOChangeList(dtNewPOChangeList);  //  不是读取上一周文件 做匹配      
        
                // 读取 POCancel 文件
                POCancelFile = Path.Combine(Workstaion, POCancelFile);
                Console.WriteLine(@"step5-2.匹配“20259发注纳入列表检索”文件 至 [{0}]", Path.GetFileName(LastWeekReportFile));
                DataTable dtPOCancle = (iOExcel.GetFileDataSet(POCancelFile)).Tables[0];
                dtPOCancle = ChangeDtColumnTypeName(dtPOCancle);
                MatchPOCancel(dtNewPOChangeList, dtNewForecastChangeList, dtPOCancle);   // 匹配 Status_2        OLD_PO_QTY   PO_QTY 
            
         

            // 将总的数据 写到一个 Excel 文件
            string ResultReportFile = string.Format(@"{0}\{1}_{2}{3}", Workstaion, "自动生成模板", DateTime.Now.ToString("yyyyMMdd"), ".xls");
            File.Copy(ReportFile, ResultReportFile, true);       // Copy 模板文件到指定路径
            iOExcel.ExportDataTableToExcelByRange(dtNewPOChangeList, ResultReportFile, sheetNames[0]);        
            iOExcel.ExportDataTableToExcelByRange(dtNewForecastChangeList, ResultReportFile, sheetNames[1]);

          


            // 前导数据
            Console.WriteLine("Step6-1.分割 前导 数据...");
            DealWithVendorData(dtNewPOChangeList, sheetNames,sheetNames[0], dtPOChangeList, "前导");
            // 后到数据
            Console.WriteLine("Step6-2.分割 后到 数据...");
            DataTable dtNewForecastChangeListFilter = dtNewForecastChangeList.Clone();
            foreach (DataRow item in dtNewForecastChangeList.Rows.OfType<DataRow>().ToList().Where(t => t.Field<string>("K").ToString() == "1"))// 只要 ORDER_STATUS=1 的数据
            {
                dtNewForecastChangeListFilter.Rows.Add(item.ItemArray);
            }
              
            DealWithVendorData(dtNewForecastChangeListFilter, sheetNames,sheetNames[1], dtForecastChangeList, "后导");

           

            ///////////////////////////////////////////////////////////////////


           Environment.Exit(0);
        }

       
        /// <summary>
        /// 读取文件集合
        /// </summary>
        /// <param name="FileNames"></param>
        /// <returns></returns> 
        public static DataTable ReadFilesToDataTable(string FileNames,bool addBuyer=false)
        {     
            DataTable dtDest = new DataTable();
            FileNames.Split(new char[] { ';' }).ToList().ForEach(FileName =>
            {
                string filePath = Path.Combine(Workstaion, FileName);
                string sheetName = string.Format("{0}$",IOExcel.GetAllSheetName(filePath)[0]);
                DataTable dt = iOExcel.GetFileDataSet(filePath).Tables[sheetName];
                dt = ChangeDtColumnTypeName(dt);    // 改变 字段类型 和 名称
                if (dtDest.Columns.Count == 0)
                    dtDest = dt.Clone();                // 先 clone 一下表结构再说
                // 过滤首行
                var newRows = dt.Rows.OfType<DataRow>().Where(t => !(t.Field<string>("A") == "Vendor"|| t.Field<string>("A") == "PO Line" || t.Field<string>("A") == "JOB_ORDER_NO")).ToList();
                newRows.ForEach(t =>
                {
                    if (addBuyer)   // 增加一列
                    {
                        string Buyer = FileName.Split(new char[] { '_' })[0];
                        t["L"] = Buyer;
                    }
                    dtDest.Rows.Add(t.ItemArray);
                });
            });

            return dtDest;
        }

       
        public static DataTable MakeNewPOChangeList(DataTable SrcCollection, DataTable DataCollection)
        {
            IOExcel.RemoveEmpty(SrcCollection); // 清空 空行

            DataTable tempTable =   SrcCollection.Clone();

            // string[] matchColName = new string[] { "A", "P", "G", "B", "C", "F", "N", "O", "Y", "Z","I","J" };      // 匹配Excel对应的字段名称
             string[] matchColName = new string[] { "A", "P", "G", "B", "C", "F", "N", "O", "Y", "Z","I","J" };      // 匹配Excel对应的字段名称


            foreach (DataRow dataRow in DataCollection.Rows)
            {
                DataRow tempRow = tempTable.NewRow();
                int i = 0;
                
                matchColName.ToList().ForEach(t =>
                {
                    tempRow[i++] = dataRow[t];
                });

                tempTable.Rows.Add(tempRow);
            }  
      
            return tempTable;
        }


       

        public static DataTable MakeNewForecastChangeList(DataTable SrcCollection, DataTable DataCollection)
        {
            IOExcel.RemoveEmpty(SrcCollection); // 清空 空行

            DataTable tempTable = SrcCollection.Clone();

            
            string[] matchColName= new string[] { "G", "L", "*", "B", "C", "D", "E", "J", "*", "*","I","F" };      // 匹配Excel对应的字段名称


            foreach (DataRow dataRow in DataCollection.Rows)
            {
                DataRow tempRow = tempTable.NewRow();
                int i = 0;

                matchColName.ToList().ForEach(t =>
                {

                    // tempRow[i++] = dataRow[t];

                    switch (t)
                    {
                        //case "Buyer":
                        //    tempRow[i++] = "";
                        //    break;
                        case "*":
                            tempRow[i++] = "*";
                            break;
                        default:
                            tempRow[i++] = dataRow[t];
                            break;

                    }


                });

                tempTable.Rows.Add(tempRow);
            }


            return tempTable;

        }


      
        public static void MatchPOChangeList(DataTable dtOldPOChangeList, DataTable dtNewPOChangeList)
        {
           
            foreach (DataRow rowNew in dtNewPOChangeList.Rows)
            {
                string PoNo = rowNew["C"].ToString();

                //DataRow[] rows = dtOldPOChangeList.Select(string.Format("C='{0}'", PoNo));  // dtOldPOChangeList.Rows.OfType<DataRow>().FirstOrDefault(t => t.Field<string>("C") == PoNo)
                DataRow rowOld =  dtOldPOChangeList.Rows.OfType<DataRow>().FirstOrDefault(t => t.Field<string>("C") == PoNo);

                string NewDueText = rowNew["L"].ToString(); // 这周文件的 [Po Due Date] --> NewDueText
                rowNew["M"] = NewDueText;                   // NewDueText --> 移动到这周文件的 [New Due Date]

                // string NewDueText = rowNew["L"].ToString();
                //string NewDueText = DateTime.Now.ToString("yyyy/MM/dd"); // 为本周 新生成文件 指定[New Due Date]

                string DueText = "";
               

                if (rowOld != null)
                {
                    DueText = rowOld["M"].ToString();   //  上一周文件的  [New Due Date] 
                    rowNew["L"] = DueText;              //  上一周文件的  [New Due Date] --> 这一周文件的 [Po Due Date]

                    if (string.IsNullOrEmpty(DueText) ||  DueText == "")
                    {// 上一周文件的  [New Due Date]  == ""  ---> 直接新规
                        rowNew["N"] = "新規";
                        continue;
                    }

                    DateTime Due = DateTime.Parse(DueText);           //Due = 上一次的Due
                    DateTime NewDue = DateTime.Parse(NewDueText);     //New Due = 这一次的Due

                    if (Due > NewDue)
                        //Due > New Due，前导
                        rowNew["N"] = "前倒";
                    else if (Due < NewDue)
                        //Due < New Due, 后到
                        rowNew["N"] = "後倒";

                }
                else  // 本周有 上周没有  --> 新规
                {
                    rowNew["L"] = "";
                    rowNew["N"] = "新規";
                }
               
            }

            ////看看  上周有     本周没有
            //dtOldPOChangeList.Rows.OfType<DataRow>().Skip(1).ToList().ForEach(rowOld =>
            //{
            //    string PoNo = rowOld["C"].ToString();
           
            //    DataRow rowNew = dtNewPOChangeList.Rows.OfType<DataRow>().FirstOrDefault(t => t.Field<string>("C") == PoNo);
            //    if (rowNew == null)
            //    {
            //        // 追加至 本周
            //        // 并 更改 Status
            //        rowOld["N"] = "Cancel";
            //        rowOld["L"] = rowOld["M"];
            //        rowOld["M"] = "";
            //        dtNewPOChangeList.Rows.Add(rowOld.ItemArray);

            //    }

            //});

        }


        public static void MatchPOCancel(DataTable dtNewPOChangeList, DataTable dtNewForecastChangeList, DataTable dtPOCancel)
        {
            // 对 dtNewForecastChangeList dtNewPOChangeList 新增2两列  O P
            // dtNewForecastChangeList.Columns.Add("O", typeof(string));
            // dtNewForecastChangeList.Columns.Add("P", typeof(string));
            if(!dtNewPOChangeList.Columns.Contains("Q"))
            dtNewPOChangeList.Columns.Add("Q");
            if (!dtNewForecastChangeList.Columns.Contains("Q"))
                dtNewForecastChangeList.Columns.Add("Q");

            dtPOCancel.Rows.OfType<DataRow>().ToList().ForEach(rowCancel => {
                string PoNo = rowCancel["J"].ToString();

                // 注: 查询时候 要跳过 头几行
                var vendorMatch = dtNewPOChangeList.Rows.OfType<DataRow>().FirstOrDefault(t => t.Field<string>("C") == PoNo);
                if(vendorMatch != null && PoNo!="*")
                {
                    vendorMatch["O"] = rowCancel["B"];  // Status_2
                    vendorMatch["P"] = rowCancel["K"];  // OLD_PO_QTY
                    vendorMatch["Q"] = rowCancel["L"];  // PO_QTY

                }

                var vendorMatchFor = dtNewForecastChangeList.Rows.OfType<DataRow>().FirstOrDefault(t => t.Field<string>("C") == PoNo);
                if (vendorMatchFor != null && PoNo != "*")
                {
                    vendorMatchFor["O"] = rowCancel["B"];
                    vendorMatchFor["P"] = rowCancel["K"];
                    vendorMatchFor["Q"] = rowCancel["L"];

                }

            });
        }

        public static void MatchForecastChangeList(DataTable dtOldForecastChangeList, DataTable dtNewForecastChangeList)
        { // MatchForecastChangeList(dtOldForecastChangeList, dtNewForecastChangeList);
            foreach (DataRow rowNew in dtNewForecastChangeList.Rows)
            {
                string PoNo = rowNew["C"].ToString();        

                DataRow[] rows = dtOldForecastChangeList.Select(string.Format("C='{0}'", PoNo));

                string DueText="";
                string NewDueText="";
                if (rows.Count() > 0)
                {
                    DueText = rows[0]["L"].ToString();
                    NewDueText = rowNew["L"].ToString();

                    //Due有，没有New Due，Cancel
                    if (!(string.IsNullOrEmpty(DueText) || DueText.Equals("*"))
                         &&
                         ((string.IsNullOrEmpty(NewDueText) || NewDueText.Equals("*"))
                        ))
                    {
                        rowNew["N"] = "Cancle";
                    }


                    //Due没，有New Due，新规
                    if ((string.IsNullOrEmpty(DueText) || DueText.Equals("*"))
                         &&
                         (!(string.IsNullOrEmpty(NewDueText) || NewDueText.Equals("*"))
                        ))
                    {
                        rowNew["N"] = "新規";
                    }


                    //Due = 上一次的Due
                    //DateTime Due = DateTime.ParseExact(DueText, "yyyyMMdd",null);
                    DateTime Due = DateTime.Parse(DueText);
                    //New Due = 这一次的Due
                    //DateTime NewDue = DateTime.ParseExact(NewDueText, "yyyyMMdd", null);
                    DateTime NewDue = DateTime.Parse(NewDueText);

                    if (Due > NewDue)
                        //Due > New Due，前导
                        rowNew["N"] = "前倒";
                    else if (Due < NewDue)
                        //Due < New Due, 后到
                        rowNew["N"] = "後倒";

                    rowNew["M"] = NewDueText;     //更新  New Due Date
                }
                else if (!(string.IsNullOrEmpty(DueText) || DueText.Equals("*"))) // Due没
                {
                    rowNew["N"] = "新規";
                }

                //var rowsOld =   dtOldPOChangeList.Rows.OfType<DataRow>().Where(t => t.Field<string>("C") == PoNo);
                //if (rowsOld.Count() > 0)
                //{
                //    rowsOld.ToArray()[0]
                //}

            }
        }

        public static void DealWithVendorData(DataTable dtTotal,List<string>sheetNames,string sheetName, DataTable dtFiles, string option="前导")
        {
         
            string remoteDir = Path.GetDirectoryName(LastWeekReportFile);
            remoteDir = string.Format(@"{0}\{1}\", remoteDir, DateTime.Now.ToString("yyyy-MM-dd"));
            // List.GroupBy(x => new{x.x1,x.x2}).Select(g=>new {g.key.x1,g.Key.x2}); 



            // 分组表 按 Vendor  和  Buyer 进行分组
            DataTable filterDt = dtTotal.Clone();
            foreach (DataRow item in dtTotal.Rows)
            {
                string filterVendor = item["A"].ToString();
                string filterBuyer = item["B"].ToString();
                DataRow[]  filterRow =  filterDt.Select(string.Format("A='{0}' and B='{1}'", filterVendor, filterBuyer));
                if(filterRow.Count() <= 0)   
                {
                    // 如果 不存在 则追加
                    filterDt.Rows.Add(item.ItemArray);

                }
                

            }

            // 遍历分组表   查询总表
            foreach (DataRow filterRow in filterDt.Rows)
            {
                string filterVendor = filterRow["A"].ToString();    // Vendor Code
                string filterBuyer = filterRow["B"].ToString();     // Buyer Code

                DataRow[] vendorRows = dtTotal.Select(string.Format("A='{0}' and B='{1}'", filterVendor, filterBuyer));
                string buyer = "";

                if (vendorRows.Count() < 0)
                    continue;

                DataTable dtVendor = dtTotal.Clone();  // vendor 供应商表
                foreach (DataRow vendorRow in vendorRows)
                {
                    buyer = vendorRow["B"].ToString();
                    dtVendor.Rows.Add(vendorRow.ItemArray);
                }

                string vendorCodeIndex = "";
                string vendorNameIndex = "";
                switch (option)
                {
                    case "前导":
                        vendorCodeIndex = "A";
                        vendorNameIndex = "Q";
                        break;
                    case "后导":
                        vendorCodeIndex = "G";
                        vendorNameIndex = "H";
                        break;
                    default:
                        break;

                }

                
               var vendorFirst = dtFiles.Rows.OfType<DataRow>().FirstOrDefault(t => t.Field<string>(vendorCodeIndex) == filterVendor);       // 这里可能有 Bug 

                if (vendorFirst != null)
                {
                    string venderName = vendorFirst[vendorNameIndex].ToString();        //

                    // [Workstation] & "\"&"Backup"&"\"&FormatDate(Today(), "yyyy - MM - dd")&"\"&"Vendors"
                    string VendorReportFile = string.Format(@"{0}\{1}\{2}\{3}\{4}_{5}_{6}_{7}{8}", Workstaion, "Backup", DateTime.Now.ToString("yyyy-MM-dd"), "Vendors", buyer, filterVendor, venderName, DateTime.Now.ToString("yyyyMMdd"), ".xls");

                    if (!File.Exists(VendorReportFile))
                    {   // 文件不存在
                        File.Copy(ReportTempFilePath, VendorReportFile, true);       // Copy 模板文件到指定路径

                        // 判断一下是否 是后到
                        if (option == "后导")
                        {
                            // 若是 在Sheet1 写入  vendorTitle
                            object[,] vendorName = new object[,] {
                            {  filterVendor,venderName,"", DateTime.Now.ToString("yyyy/MM/dd")} };

                            iOExcel.ExportDataTableToExcelByRange(vendorName, VendorReportFile, sheetNames[0]);     // 后补一下 前到数据 【仅标题】
                        }

                    }
                    else
                    {

                    }

                    object[,] vendorTitle = new object[,] {
                            {  filterVendor,venderName,"", DateTime.Now.ToString("yyyy/MM/dd")} };
                    iOExcel.ExportDataTableToExcelByRange(dtVendor, vendorTitle, VendorReportFile, sheetName);      // 写入数据 【标题和内容】
                    Console.WriteLine(Path.GetFileName(VendorReportFile));
                }
                else
                {
                    //Console.WriteLine("注意了  居然没有找到");
                }


            }

        }

      

       

        /// <summary>
        /// 改变字段类型
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static DataTable ChangeDtColumnType(DataTable dt)
        {
            DataTable dtResult = new DataTable();
            //克隆表结构
            dtResult = dt.Clone();
            foreach (DataColumn col in dtResult.Columns)
            {
                //修改列类型
                col.DataType = typeof(String);
            }
            foreach (DataRow row in dt.Rows)
            {
                DataRow rowNew = dtResult.NewRow();
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    rowNew[i] = row[i].ToString();
                }
                dtResult.Rows.Add(rowNew);
            }

            return dtResult;

        }

        /// <summary>
        /// 修改字段名称和类型
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static DataTable ChangeDtColumnTypeName(DataTable dt)
        {
            DataTable dtResult = new DataTable();
            //克隆表结构
            dtResult = dt.Clone();

            int coliIndex = 0;
            foreach (DataColumn col in dtResult.Columns)
            {
                //修改列类型
                col.DataType = typeof(String);
                dtResult.Columns[coliIndex].ColumnName =  IOExcel.GetExcelCellName(++coliIndex);
                    
            }
            foreach (DataRow row in dt.Rows)
            {
                DataRow rowNew = dtResult.NewRow();
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    rowNew[i] = row[i].ToString();
                }
                dtResult.Rows.Add(rowNew);
            }

            return dtResult;

        }


    }
}
