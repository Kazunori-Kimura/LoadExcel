using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;
using System.Configuration;
using System.Text.RegularExpressions;

using OfficeOpenXml;
using LoadExcel.Models;

namespace LoadExcel
{
    class Program
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        static void Main(string[] args)
        {
            //TODO 引数からファイルパスを取得する
            string path = @"D:\workspace\透明性ガイドライン\Category-C\謝礼開示情報サンプルDisclosureSummary_20130618(1).xlsx";

            log.InfoFormat("** Category-Cデータ作成 処理開始 **");
            log.DebugFormat("  -- input_file={0}", path);

            var job = new JobExecute();

            bool isDebug = false; //ファイル読み込みをスキップする(デバッグ時 時間短縮のため)

            if (!isDebug)
            {
                //(1) ファイル読み込み
                if (!job.LoadExcelFile(path))
                {
                    if (!IsIgnoreWarning()) //警告を無視するか？
                    {
                        //警告がある際はエラーメールを送信して終了
                        job.SendErrorMail();
                        return;
                    }
                }
            }

            //(2) Excelファイル出力
            job.OutputExcelFiles();
            
            //(3) 処理完了メール通知
            job.SendMail();

            log.InfoFormat("** Category-Cデータ作成 処理終了 **");
        }

        /// <summary>
        /// ignore_warn設定時、警告を無視する
        /// </summary>
        /// <returns></returns>
        private static bool IsIgnoreWarning()
        {
            return Utils.ParseBool(ConfigurationManager.AppSettings["ignore_warn"]);
        }

    } //end class Program

    /// <summary>
    /// データ作成処理本体
    /// </summary>
    class JobExecute
    {
        #region properties
        /// <summary>
        /// Logger
        /// </summary>
        private readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        /// <summary>
        /// DataManager
        /// </summary>
        DataManager dm;

        /// <summary>
        /// MailManager
        /// </summary>
        MailManager mm;
        #endregion properties

        #region コンストラクタ
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public JobExecute()
        {
            dm = new DataManager();
            mm = new MailManager();
            //開始時間設定
            mm.parameters.Add("$start_date", DateTime.Now.ToString());
        }
        #endregion コンストラクタ

        #region Excel読み込み / マスタ反映

        /// <summary>
        /// Excelファイル読み込み
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public bool LoadExcelFile(string path)
        {
            //取込み対象ファイル設定
            mm.parameters.Add("$import_file", path);

            string sheetName = @"DisclosureSummary_20130618(1)";
            int rowIndex = 2;

            //シート名を設定ファイルから取得
            string _sn = ConfigurationManager.AppSettings["sheetName"];
            if (!string.IsNullOrEmpty(_sn))
            {
                sheetName = _sn;
            }
            //開始行を設定ファイルから取得
            string _row = ConfigurationManager.AppSettings["start_row"];
            if (!string.IsNullOrEmpty(_row))
            {
                rowIndex = Utils.ParseInt(_row);
            }

            //Excelファイル読み込み
            List<XlsData> xd = dm.LoadExcel(path, sheetName, rowIndex);

            bool ret = false;
            if (xd.Count > 0)
            {
                //MegaCOARAの各マスタの情報を反映
                ret = dm.UpdateMasterInfo();
            }
            return ret;
        }

        #endregion Excel読み込み / マスタ反映

        #region 支払い回数集計 / Excel出力

        /// <summary>
        /// Excelファイル出力
        /// </summary>
        /// <returns></returns>
        public bool OutputExcelFiles()
        {
            //支払い回数集計
            dm.Summarize();

            //ファイル出力
            string message = dm.OutputExcelFiles();
            
            bool ret = false;
            if (!string.IsNullOrEmpty(message))
            {
                //正常終了
                mm.SetParam("$status", "正常終了");
                mm.SetParam("$message", message);
                ret = true;
            }
            else
            {
                //エラー
                mm.SetParam("$status", "エラー");
                mm.SetParam("$message", @"
Excelファイルの出力にてエラーが発生しました。
ログを確認してください。
");
            }

            return ret;
        }

        #endregion

        #region メール送信

        /// <summary>
        /// エラーメール送信
        /// </summary>
        public void SendErrorMail()
        {
            //警告発生時はメール送付して終了
            mm.parameters.Add("$status", "エラー");
            mm.parameters.Add("$end_date", DateTime.Now.ToString());

            //メッセージ
            string message = @"
Excelファイルに問題があります。
ログファイルを確認してください。

    ログファイル: D:\workspace\survey\logs
";
            mm.parameters.Add("$message", message);

            mm.SendMail();
        }

        /// <summary>
        /// メール送信
        /// </summary>
        public void SendMail()
        {
            //終了日時セット
            mm.SetParam("$end_date", DateTime.Now.ToString());
            mm.SendMail();
        }

        #endregion
    }
}
