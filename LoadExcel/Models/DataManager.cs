using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.IO;
using System.Text.RegularExpressions;

using Oracle.DataAccess.Client;
using OfficeOpenXml;

namespace LoadExcel.Models
{
    class DataManager
    {
        /// <summary>
        /// logger
        /// </summary>
        readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        /// <summary>
        /// Database Connection
        /// </summary>
        private PetaPoco.Database db = null;

        /// <summary>
        /// Excel開始位置 (ログ出力に使用)
        /// </summary>
        private int _start_row = 2;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public DataManager()
        {
            string cs = string.Format(@"User Id={0};Password={1};Data Source={2};",
                "TRA_USER", "1qaz2wsx", "HOGE");
            //string cs = "Driver={Oracle in OraClient11g_home2};DBQ=HOGE;UID=TRA_USER;PWD=1qaz2wsx;";

            db = new PetaPoco.Database(cs, Oracle.DataAccess.Client.OracleClientFactory.Instance);
        }

        #region Excel読み込み

        /// <summary>
        /// 指定されたExcelを読み込み、Databaseに登録する
        /// </summary>
        /// <param name="path"></param>
        /// <param name="sheetName"></param>
        /// <param name="start_row"></param>
        /// <returns>Excel読み込みデータ</returns>
        public List<XlsData> LoadExcel(string path, string sheetName, int start_row)
        {
            log.Debug(string.Format("target file= {0}; sheet= {1}", path, sheetName));

            _start_row = start_row; //開始位置 (ログ出力に使用)
            int rowIndex = start_row;
            List<XlsData> records = new List<XlsData>();

            try
            {
                using (ExcelPackage xls = new ExcelPackage())
                {
                    using (FileStream fs = new FileStream(path, FileMode.Open))
                    {
                        xls.Load(fs);
                        ExcelWorksheet sheet = xls.Workbook.Worksheets[sheetName];

                        while (true)
                        {
                            XlsData data = new XlsData();

                            data.id = rowIndex - start_row;
                            // read row data
                            data.drcode = ToString(sheet.Cells[rowIndex, 1].Value);
                            data.drname = ToString(sheet.Cells[rowIndex, 2].Value);
                            if (string.IsNullOrEmpty(data.drname))
                            {
                                break; //Dr名が無ければ終了
                            }
                            data.ncc_cd = ToString(sheet.Cells[rowIndex, 3].Value);
                            data.ncc_name = ToString(sheet.Cells[rowIndex, 4].Value);
                            data.ncc_dept = ToString(sheet.Cells[rowIndex, 5].Value);
                            data.title = ToString(sheet.Cells[rowIndex, 6].Value);
                            data.category = ToString(sheet.Cells[rowIndex, 7].Value);
                            data.kingaku = ToLong(sheet.Cells[rowIndex, 8].Value);
                            data.kaisu = ToInt(sheet.Cells[rowIndex, 9].Value);

                            // 頭 '0' 埋め
                            data.drcode = data.drcode.PadLeft(6, '0');
                            if (data.drcode == "000000")
                            {
                                data.drcode = string.Empty;
                            }

                            records.Add(data);
                            rowIndex++;
                        }

                    }
                }

                //読み込んだデータをDatabaseに登録
                int registed = BulkInsert(records);

                log.Info(string.Format("登録件数= {0}", registed));
            }
            catch (Exception exp)
            {
                log.Error(string.Format(@"Fatal Error.: {0}", rowIndex), exp);
            }

            return records;
        }

        /// <summary>
        /// xls_dataにデータを登録する
        /// </summary>
        /// <param name="items"></param>
        /// <returns></returns>
        public int BulkInsert(List<XlsData> items)
        {
            int registed = 0;

            //truncate table
            truncate();

            using (var trans = db.GetTransaction())
            {
                foreach (var item in items)
                {
                    db.Insert(item);
                    registed++;
                }
                trans.Complete();
            }
            
            return registed;
        }

        /// <summary>
        /// XLS_DATAテーブルを初期化
        /// </summary>
        private void truncate()
        {
            //一時テーブルを初期化
            string query = @"truncate table xls_data";
            db.Execute(query);
        }

        #endregion

        #region MegaCOARA 施設マスタ取得

        /// <summary>
        /// 施設名を元に施設マスタから施設情報を取得し、更新する
        /// </summary>
        /// <returns></returns>
        private int UpdateHspByName()
        {
            #region SQL

            //施設名の部分一致で施設情報を取得
            string sql = @"
select
    t1.id,
    t1.drcode,
    t1.drname,
    t1.ncc_cd,
    t1.ncc_name,
    t1.ncc_dept,
    t1.title,
    t1.category,
    t1.kingaku,
    t1.kaisu,
    t1.dm_cont_id,
    t1.dm_doccd,
    t1.dm_name,
    t1.dm_ktkn_nm,
    h1.act_uni_id as hm_act_uni_id,
    h1.ncc_cd as hm_ncc_cd,
    h1.acnt_nm as hm_acnt_nm,
    h1.acnt_jpln_nm as hm_acnt_jpln_nm,
    h1.acnt_ktkn_nm as hm_acnt_ktkn_nm
from
    (select *
        from tmp_data1
        where hm_acnt_nm is null) t1
    left join (
        /* 施設マスタ */
        select
            act_uni_id,
            hs01 || hs02 as ncc_cd,
            acnt_nm,
            acnt_jpln_nm,
            acnt_ktkn_nm
        from
            coa_dcfhsp
        where
            hpdymd = 0 or hpdymd is null
        ) h1
        on h1.acnt_nm like '%' || t1.ncc_name || '%'
order by
    t1.id, h1.ncc_cd
";

            #endregion SQL
            //更新件数
            int ret = 0;
            //更新用データ
            var data = new Dictionary<int, List<TmpData>>();

            //施設マスタからデータを取得
            var items = db.Fetch<TmpData>(sql);

            //取得データを走査し、更新用データを作成する
            foreach (var item in items)
            {
                if (!data.ContainsKey(item.id))
                {
                    data.Add(item.id, new List<TmpData>());
                }
                data[item.id].Add(item);
            }

            //更新処理
            using (var trans = db.GetTransaction())
            {
                foreach (var key in data.Keys) //key = TmsData.id
                {
                    if (data[key].Count == 1)
                    {
                        var item = data[key][0]; //TmpData
                        if (!string.IsNullOrEmpty(item.hm_ncc_cd))
                        {
                            //ユニークなデータを発見
                            ret += db.Update("tmp_data1", "id", item);
                            log.DebugFormat("  -- update hsp data: row={0},ncc_cd={1}", item.id, item.hm_ncc_cd);
                        }
                        else
                        {
                            //該当データなし
                            log.WarnFormat("施設マスタに該当のデータがありません。 Excel行={0},施設名={1}", item.id, item.ncc_name);
                        }
                    }
                    else
                    {
                        //複数の候補がある
                        log.WarnFormat("施設情報が特定できませんでした。 Excel行={0},施設名={1}", data[key][0].id, data[key][0].ncc_name);
                        log.Warn("  候補データ:");

                        foreach (var d in data[key])
                        {
                            log.WarnFormat("    施設コード={0},施設名={1}", d.hm_ncc_cd, d.hm_acnt_nm);
                        }
                    }
                } //end foreach

                trans.Complete();
            } //end using

            return ret;
        }

        #endregion

        #region MegaCOARA 医師マスタ取得

        /// <summary>
        /// 医師情報更新
        /// </summary>
        /// <returns></returns>
        private int UpdateDoctorByName()
        {
            #region SQL
            //氏名を元に医師マスタを検索
            var sql = @"
select
    t1.id,
    t1.drcode,
    t1.drname,
    t1.ncc_cd,
    t1.ncc_name,
    t1.ncc_dept,
    t1.title,
    t1.category,
    t1.kingaku,
    t1.kaisu,
    d1.cont_id as dm_cont_id,
    d1.doccd as dm_doccd,
    d1.name as dm_name,
    d1.ktkn_nm as dm_ktkn_nm
from
    /* 医師情報未取得データ */
    (select *
        from tmp_data1
        where dm_cont_id is null) t1
    left join (
        /* 医師マスタ */
        select
            cont_id, doccd, name, ktkn_nm
        from
            coa_doc1p
        where
            (drdymd = 0 or drdymd is null) and
            doccd is not null
        ) d1
    on replace(replace(t1.drname, '　', ''), ' ', '') 
        = replace(replace(d1.name, '　', ''), ' ', '')
order by
    t1.id, d1.cont_id
";
            #endregion SQL
            //更新件数
            int ret = 0;
            //更新用データ
            var data = new Dictionary<int, List<TmpData>>();


            //医師マスタからデータを取得
            var ds = db.Fetch<TmpData>(sql);

            //取得データを走査し、更新用データを作成する
            foreach (var item in ds)
            {
                if (!data.ContainsKey(item.id))
                {
                    data.Add(item.id, new List<TmpData>());
                }
                data[item.id].Add(item);
            }

            //更新処理
            using (var trans = db.GetTransaction())
            {
                foreach (var key in data.Keys) //key = TmsData.id
                {
                    if (data[key].Count == 1)
                    {
                        var item = data[key][0]; //TmpData
                        if (!string.IsNullOrEmpty(item.dm_doccd))
                        {
                            //ユニークなデータを発見
                            ret += db.Update("tmp_data1", "id", item);
                            log.DebugFormat("  -- update doctor data: row={0},doccd={1}", item.id, item.dm_doccd);
                        }
                        else
                        {
                            //該当データなし
                            log.WarnFormat("医師マスタに該当のデータがありません。 Excel行={0},医師氏名={1}", item.id, item.drname);
                        }
                    }
                    else
                    {
                        //複数の候補がある
                        log.WarnFormat("医師情報が特定できませんでした。 Excel行={0},医師氏名={1}", data[key][0].id, data[key][0].drname);
                        log.Warn("  候補データ:");

                        foreach (var d in data[key])
                        {
                            log.WarnFormat("    医師コード={0},医師氏名={1}",d.dm_doccd, d.dm_name);
                        }
                    }
                } //end foreach

                trans.Complete();
            } //end using

            return ret;
        }

        #endregion

        #region マスタデータ登録
        /// <summary>
        /// 医師コード、施設コードを元に各マスタから一時テーブルにデータを取得
        /// </summary>
        /// <returns></returns>
        private int InsertTempTable()
        {
            #region SQL

            #region - 一時テーブル初期化
            var sql_temp_truncate = @"truncate table tmp_data1";
            #endregion

            #region - 一時テーブル登録
            //一時テーブル登録
            var sql = PetaPoco.Sql.Builder.Append(@"
insert into tmp_data1
select
    c1.id,
    c1.drcode,
    c1.drname,
    c1.ncc_cd,
    c1.ncc_name,
    c1.ncc_dept,
    c1.title,
    c1.category,
    c1.kingaku,
    c1.kaisu,
    d.cont_id as dm_cont_id,
    d.doccd as dm_doccd,
    d.name as dm_name,
    d.ktkn_nm as dm_ktkn_nm,
    h.act_uni_id as hm_act_uni_id,
    h.hs01 || h.hs02 as hm_ncc_cd,
    h.acnt_nm as hm_acnt_nm,
    h.acnt_jpln_nm as hm_acnt_jpln_nm,
    h.acnt_ktkn_nm as hm_acnt_ktkn_nm
from
    xls_data c1
    left join (
        /* doctor */
        select
            c2.id, d1.cont_id, d1.doccd, d1.name, d1.ktkn_nm
        from
            xls_data c2
            left join (
                select
                    cont_id, doccd, name, ktkn_nm
                from
                    coa_doc1p
                where
                    (drdymd = 0 or drdymd is null) and
                    doccd is not null
                ) d1
            on c2.drcode = d1.doccd
    ) d on c1.id = d.id
    left join (
        /* hsp */
        select
            c3.id, h1.act_uni_id, h1.hs01, h1.hs02,
            h1.acnt_nm, h1.acnt_jpln_nm, h1.acnt_ktkn_nm
        from
            xls_data c3
            left join (
                select
                    act_uni_id, hs01, hs02, acnt_nm,
                    acnt_jpln_nm, acnt_ktkn_nm
                from
                    coa_dcfhsp
                where
                    (hpdymd = 0 or hpdymd is null) and
                    hs01 is not null and
                    hs02 is not null
            ) h1
            on c3.ncc_cd = h1.hs01 || h1.hs02
    ) h on c1.id = h.id
");
            #endregion - 一時テーブル登録

            #endregion SQL

            //一時テーブル初期化
            db.Execute(sql_temp_truncate);

            //一時テーブル登録
            int ret = 0;
            using (var trans = db.GetTransaction())
            {
                ret = db.Execute(sql);
                trans.Complete();
            }

            log.DebugFormat("  tmp_data1 登録件数={0}", ret);
            return ret;
        }

        #endregion

        #region 一括更新

        /// <summary>
        /// 一括更新
        /// </summary>
        public void UpdateMasterInfo()
        {
            
            //コードを元に一時テーブル登録
            int rows = InsertTempTable();
            log.DebugFormat("  -- update data by code: {0}", rows);

            //医師情報更新
            rows = UpdateDoctorByName();
            log.DebugFormat("  -- update data by doctor name: {0}", rows);

            //施設情報更新
            rows = UpdateHspByName();
            log.DebugFormat("  -- update data by hsp name: {0}", rows);

        }

        #endregion


        #region 共通メソッド
        /// <summary>
        /// セルの値をstring型に変換する
        /// </summary>
        /// <param name="value">Worksheet.Cells.Value</param>
        /// <returns>string</returns>
        public string ToString(Object value)
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
        /// セルの値をint型に変換する
        /// </summary>
        /// <param name="value">Worksheet.Cells.Value</param>
        /// <returns>int</returns>
        public int ToInt(Object value)
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
        /// セルの値をlong型に変換する
        /// </summary>
        /// <param name="value">Worksheet.Cells.Value</param>
        /// <returns>long</returns>
        public long ToLong(Object value)
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
        #endregion
    }
}
