using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LoadExcel
{
    class Utils
    {
        /// <summary>
        /// logger
        /// </summary>
        readonly static log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        #region 共通メソッド
        /// <summary>
        /// 値をstring型に変換する
        /// </summary>
        /// <param name="value">Worksheet.Cells.Value</param>
        /// <returns>string</returns>
        public static string ParseString(Object value)
        {
            if (value == null)
            {
                return string.Empty;
            }
            else
            {
                return value.ToString();
            }
        }

        /// <summary>
        /// 値をint型に変換する
        /// </summary>
        /// <param name="value">Worksheet.Cells.Value</param>
        /// <returns>int</returns>
        public static int ParseInt(Object value)
        {
            if (value == null)
            {
                return 0;
            }
            else
            {
                try
                {
                    var ret = int.Parse(value.ToString());
                    return ret;
                }
                catch (Exception exp)
                {
                    log.Debug(string.Format("変換エラー: value= {0}", value), exp);
                    return 0;
                }
            }
        }

        /// <summary>
        /// 値をlong型に変換する
        /// </summary>
        /// <param name="value">Object</param>
        /// <returns>long</returns>
        public static long ParseLong(Object value)
        {
            if (value == null)
            {
                return 0;
            }
            else
            {
                try
                {
                    var ret = long.Parse(value.ToString());
                    return ret;
                }
                catch (Exception exp)
                {
                    log.Debug(string.Format("変換エラー: value= {0}", value), exp);
                    return 0;
                }
            }
        }

        /// <summary>
        /// 値をbool型に変換する
        /// </summary>
        /// <param name="value"></param>
        /// <returns>bool</returns>
        public static bool ParseBool(Object value)
        {
            if (value == null)
            {
                return false;
            }
            else
            {
                try
                {
                    var ret = bool.Parse(value.ToString());
                    return ret;
                }
                catch (Exception exp)
                {
                    log.Debug(string.Format("変換エラー: value= {0}", value), exp);
                    return false;
                }
            }
        }

        #endregion
    }
}
