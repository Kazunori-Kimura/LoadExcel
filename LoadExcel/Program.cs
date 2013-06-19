using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;

using OfficeOpenXml;
using LoadExcel.Models;

namespace LoadExcel
{
    class Program
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        static void Main(string[] args)
        {
            string path = @"D:\workspace\透明性ガイドライン\Category-C\謝礼開示情報サンプルDisclosureSummary_20130618(1).xlsx";
            string sheetName = @"DisclosureSummary_20130618(1)";
            int rowIndex = 2;

            DataManager dm = new DataManager();

            //Excelを読み込みDBに登録
            List<CpData> items = dm.LoadExcel(path, sheetName, rowIndex);

            //MegaCOARAの各マスタの情報を反映
            dm.UpdateMasterInfo(items);
        }

    }
}
