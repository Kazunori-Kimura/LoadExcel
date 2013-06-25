using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Configuration;
using System.Text.RegularExpressions;
using System.IO;
using System.Net.Mail;

namespace LoadExcel
{
    class MailManager
    {
        #region properties
        /// <summary>
        /// logger
        /// </summary>
        readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        /// <summary>
        /// メールテンプレートに設定するパラメータ
        /// </summary>
        public Dictionary<string, string> parameters { get; set; }
        #endregion

        #region コンストラクタ
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public MailManager()
        {
            parameters = new Dictionary<string, string>();
        }

        #endregion

        #region パラメータ設定

        /// <summary>
        /// パラメータ設定
        /// </summary>
        /// <param name="key"></param>
        /// <param name="value"></param>
        public void SetParam(string key, string value)
        {
            if (parameters.ContainsKey(key))
            {
                parameters[key] = value; //上書き設定
            }
            else
            {
                parameters.Add(key, value);
            }
        }

        #endregion

        #region メール送信

        /// <summary>
        /// メール送信処理
        /// </summary>
        public void SendMail()
        {
            //送信日時設定
            if (!parameters.ContainsKey("$send_date"))
            {
                parameters.Add("$send_date", DateTime.Now.ToString());
            }

            string from_addr = ConfigurationManager.AppSettings["from_addr"];
            string to_addr = ConfigurationManager.AppSettings["to_addr"];
            string smtp_server = ConfigurationManager.AppSettings["smtp_server"];
            string subject = ReplaceParameters(ConfigurationManager.AppSettings["subject"]);
            bool mailEnable = Utils.ParseBool(ConfigurationManager.AppSettings["mail_enable"]);

            //MessageBody
            string body = GetMessageBody();

            if (mailEnable)
            {
                var smtp = new SmtpClient();
                smtp.Host = smtp_server;

                //メール送信
                smtp.Send(from_addr, to_addr, subject, body);
            }
            else
            {
                #region メール送信代替 ログ出力
                string log_message = @"smtp_server: $smtp_server
from:    $from
to:      $to
subject: $subject
-------------------
";
                log_message += body;
                
                var prm = new Dictionary<string, string>();
                prm.Add(@"\$smtp_server", smtp_server);
                prm.Add(@"\$from", from_addr);
                prm.Add(@"\$to", to_addr);
                prm.Add(@"\$subject", subject);

                foreach (string key in prm.Keys)
                {
                    log_message = Regex.Replace(log_message, key, prm[key], RegexOptions.IgnoreCase);
                }

                log.Debug(log_message);
                #endregion
            }

            log.InfoFormat("メールを送信しました。TO={0}", to_addr);
        }

        /// <summary>
        /// MessageBody取得
        /// </summary>
        /// <returns></returns>
        private string GetMessageBody()
        {
            //テンプレートファイル読み込み
            string template_path = ConfigurationManager.AppSettings["mail_template"];
            StreamReader sr = new StreamReader(template_path, System.Text.Encoding.UTF8);
            string body = sr.ReadToEnd();

            //変数を置換して返す
            return ReplaceParameters(body);
        }

        /// <summary>
        /// parametersの内容で変数($xxx)を置換する
        /// </summary>
        /// <param name="org_text">元文字列</param>
        /// <returns>置換された文字列</returns>
        private string ReplaceParameters(string org_text)
        {
            string ret_text = org_text;
            if (string.IsNullOrEmpty(org_text))
            {
                return string.Empty;
            }

            //parametersの内容を反映
            foreach (string key in parameters.Keys)
            {
                string var_name = key;
                string value = parameters[key];
                if (!Regex.IsMatch(key, @"^\$"))
                {
                    var_name = @"$" + var_name;
                }

                //                                 ↓先頭の $ をエスケープ
                ret_text = Regex.Replace(ret_text, @"\"+var_name, value, RegexOptions.ECMAScript | RegexOptions.IgnoreCase);
            }

            return ret_text;
        }

        #endregion
    }
}
