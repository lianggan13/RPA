using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWS
{
   public static class GlobalConfig
    {
        
        /// <summary>
        /// 匹配 下标
        /// 如:               PXW-X180	        84351371	          机种结束
        ///                 MatchIndex-2      MatchIndex-1           MatchIndex
        ///                     1                  2                       3
        /// </summary>
        public static int MatchIndex = 3;


        /// <summary>
        /// 匹配字段 consoli在库(周别).xlsx
        /// </summary>
        public static  string[] MatchFieldWeekly = new string[] { "SSGE Inventory", "Dealer Inventory", "PSC Inventory", "Total Inventory", "Sell In", "WOS", "" };


        /// <summary>
        /// 匹配字段 consoli在库(月别).xlsx
        /// </summary>
        public static string[] MatchFieldMonthly = new string[] { "SSGE", "Dealer", "PSC", "Total", "Sell", "WOS", "" };


        /// <summary>
        /// 枚举:看下 consoli在库(周别).xlsx 就知道了
        /// </summary>
        enum EnumMatchFieldWeekly
        {
            SSGE_Inventory=1,
            Dealer_Inventory = 2,
            PSC_Inventory = 3,
            Total_Inventory = 4,
            Sell_In = 5,
            WOS = 6, 
        }

        /// <summary>
        /// 枚举:看下 consoli在库(月别).xlsx 就知道了
        /// </summary>
        enum EnumMatchFieldMonthly
        {
            SSGE = 1,
            Dealer = 2,
            PSC = 3,
            Total = 4,
            Sell = 5,
            WOS = 6,
        }



    }
}
