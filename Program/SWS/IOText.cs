using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace SWS
{
    public class IOText
    {
        /// <summary>
        /// 读取Txt文件数据
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public DataTable GetTxtFileDataSet(string fileName)
        {
            DataTable dt = new DataTable();
            
            if (!String.IsNullOrEmpty(fileName))
            {
                string fileType = System.IO.Path.GetExtension(fileName);
                if (string.IsNullOrEmpty(fileType)) return null;

                if (fileType.ToLower() == ".txt")
                {
                    File.WriteAllText(fileName,File.ReadAllText(fileName, Encoding.GetEncoding("GB18030")).Replace("德国制造||", "德国制造|"), Encoding.GetEncoding("GB18030"));
                    //File.WriteAllText(fileName, File.ReadAllText(fileName, System.Text.Encoding.UTF8).Replace("德国制造\t\t", "德国制造\t"), Encoding.GetEncoding("UTF-8"));
                    string[] rows = File.ReadAllLines(fileName, Encoding.GetEncoding("GB18030"));
                    for (int i=0;i<rows.Count();i++)
                    {
                        if (i == 0)
                        {
                            //首行，添加表头
                            
                            string[] headers = rows[i].Split('|');
                            for (int j = 0; j < headers.Count(); j++)
                            {
                                if (dt.Columns.Contains(headers[j]))
                                {
                                    dt.Columns.Add(headers[j] + j);
                                }
                                else
                                {
                                    dt.Columns.Add(headers[j], typeof(string));
                                }
                            }
                        }
                        else
                        {
                            DataRow row = dt.NewRow();
                            string[] contents = rows[i].Split('|');
                            for (int k = 0; k < contents.Count(); k++)
                            {
                                row[k] = contents[k];
                            }
                            dt.Rows.Add(row);
                        }
                    }
                }
                
            }
            return dt;
        }
    }
}
