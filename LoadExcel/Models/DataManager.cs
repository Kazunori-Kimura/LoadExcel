using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.IO;
using System.Configuration;

using Oracle.DataAccess.Client;
using OfficeOpenXml;

namespace LoadExcel.Models
{
    class DataManager
    {
        #region properties

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

        #endregion

        #region コンストラクタ

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public DataManager()
        {
            //TODO DB接続文字列をapp.configから取得する
            string cs = string.Format(@"User Id={0};Password={1};Data Source={2};",
                "TRA_USER", "1qaz2wsx", "HOGE");
            //string cs = "Driver={Oracle in OraClient11g_home2};DBQ=HOGE;UID=TRA_USER;PWD=1qaz2wsx;";

            db = new PetaPoco.Database(cs, Oracle.DataAccess.Client.OracleClientFactory.Instance);
        }

        #endregion

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
                            data.drcode = Utils.ParseString(sheet.Cells[rowIndex, 1].Value);
                            data.drname = Utils.ParseString(sheet.Cells[rowIndex, 2].Value);
                            if (string.IsNullOrEmpty(data.drname))
                            {
                                break; //Dr名が無ければ終了
                            }
                            data.ncc_cd = Utils.ParseString(sheet.Cells[rowIndex, 3].Value);
                            data.ncc_name = Utils.ParseString(sheet.Cells[rowIndex, 4].Value);
                            data.ncc_dept = Utils.ParseString(sheet.Cells[rowIndex, 5].Value);
                            data.title = Utils.ParseString(sheet.Cells[rowIndex, 6].Value);
                            data.category = Utils.ParseString(sheet.Cells[rowIndex, 7].Value);
                            data.kingaku = Utils.ParseLong(sheet.Cells[rowIndex, 8].Value);
                            data.kaisu = Utils.ParseInt(sheet.Cells[rowIndex, 9].Value);

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

        #region MegaCOARA マスタ反映

        #region MegaCOARA 施設マスタ取得

        /// <summary>
        /// 施設名を元に施設マスタから施設情報を取得し、更新する
        /// </summary>
        /// <returns>正常終了時: true</returns>
        private bool UpdateHspByName()
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
            //警告有無
            bool isWarn = false;
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
                            isWarn = true;
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
                        isWarn = true;
                    }
                } //end foreach

                trans.Complete();
            } //end using

            log.DebugFormat("  -- UpdateHspByName: 更新件数={0}", ret);
            return !isWarn;
        }

        #endregion

        #region MegaCOARA 医師マスタ取得

        /// <summary>
        /// 医師情報更新
        /// </summary>
        /// <returns>警告ありの場合、true</returns>
        private bool UpdateDoctorByName()
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
            //警告あり
            bool isWarn = false;
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
                            isWarn = true;
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
                        isWarn = true;
                    }
                } //end foreach

                trans.Complete();
            } //end using

            log.DebugFormat("  -- UpdateDoctorByName: 更新件数={0}", ret);
            return isWarn;
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
            //TODO 実績ゼロのデータを除く
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
        /// <returns>警告有無</returns>
        public bool UpdateMasterInfo()
        {
            
            //コードを元に一時テーブル登録
            int rows = InsertTempTable();
            log.DebugFormat("  -- update data by code: {0}", rows);

            //医師情報更新
            //  医師情報の取得警告はとりあえず無視する
            UpdateDoctorByName();
            
            //施設情報更新
            //  施設情報の取得結果(警告有無)を返す
            return UpdateHspByName();
        }

        #endregion

        #endregion MegaCOARA マスタ反映

        #region 支払い回数 / 医師:施設 集計

        /// <summary>
        /// 支払い回数 / 医師:施設 集計
        /// </summary>
        public void Summarize()
        {
            #region SQL
            //テーブル初期化
            string sql_trunc = @"truncate table c_result";

            //支払い回数 / 医師:施設 を集計する
            string sql_summarize = @"
insert into c_result
select distinct
    t1.dr_code,
    t1.dr_name,
    t1.dm_ktkn_nm,
    t1.hm_ncc_cd,
    t1.hm_acnt_nm,
    t1.hm_acnt_ktkn_nm,
    (select
        max(ncc_dept)
    from
        tmp_data1 t2
    where
        t2.dm_doccd = t1.dr_code and
        t2.hm_ncc_cd = t1.hm_ncc_cd
    ) as ncc_dept,
    t1.category,
    count(t1.id) over (partition by t1.category, t1.dr_name) as siharai,
    sum(t1.kingaku) over (partition by t1.category, t1.dr_name) as goukei,
    t2.ncc_count
from
(
    select
        id,
        category,
        kingaku,
        dm_doccd as dr_code,
        case when dm_name is not null
            then dm_name
            else drname
        end as dr_name,
        dm_ktkn_nm,
        hm_ncc_cd,
        hm_acnt_nm,
        hm_acnt_ktkn_nm
    from
        TMP_DATA1
    where
        hm_ncc_cd is not null
) t1
left join (
    select distinct
        t3.category,
        t3.dr_name,
        count(t3.hm_ncc_cd) over (partition by t3.category, t3.dr_name) as ncc_count
    from
        (
            select distinct
                category,
                case when dm_name is not null
                    then dm_name
                    else drname
                end as dr_name,
                hm_ncc_cd
            from tmp_data1
            where hm_ncc_cd is not null
        ) t3
    ) t2
on t1.category = t2.category and t1.dr_name = t2.dr_name
order by
    t1.category,
    t1.dr_code
";
            #endregion

            //c_result初期化
            db.Execute(sql_trunc);

            using (var trans = db.GetTransaction())
            {
                int ret = db.Execute(sql_summarize);
                trans.Complete();

                log.DebugFormat("  -- summarize: insert rows={0}", ret);
            }
        }

        #endregion

        #region DWH 施設-役割 紐付け取得

        /// <summary>
        /// YEARMON を取得する
        /// </summary>
        /// <returns>yyyyMM</returns>
        private string GetYearMon()
        {
            string yearmon = DateTime.Now.ToString("yyyyMM");

            while (!ExistsYearMon(yearmon))
            {
                yearmon = GetLastYearMon(yearmon);
            }

            return yearmon;
        }

        /// <summary>
        /// 指定された年月の 前月のYEARMON を取得する
        /// </summary>
        /// <param name="yearmon">yyyyMM</param>
        /// <returns>yyyyMM</returns>
        private string GetLastYearMon(string yearmon)
        {
            string lastyearmon = DateTime.Now.ToString("yyyyMM"); //とりあえず当月
            if (!string.IsNullOrEmpty(yearmon))
            {
                lastyearmon = yearmon;
            }
            //yyyyMM を DateTime に変換
            DateTime dt = DateTime.ParseExact(lastyearmon, "yyyyMM", System.Globalization.DateTimeFormatInfo.InvariantInfo);
            DateTime lastDt = dt.AddMonths(-1); //前月

            return lastDt.ToString("yyyyMM");
        }

        /// <summary>
        /// 指定された年月が G_PSTN_HSP_MST,G_ACT_EMP_PSTN_MST に存在するか
        /// </summary>
        /// <param name="yearmon">yyyyMM</param>
        /// <returns></returns>
        private bool ExistsYearMon(string yearmon)
        {
            #region SQL

            //G_PSTN_HSP_MST
            var sql1 = PetaPoco.Sql.Builder.Append(@"
select
    yearmon
from
    g_pstn_hsp_mst
where
    yearmon = @yearmon and
    rownum = 1
", new { yearmon = yearmon });

            //G_ACT_EMP_PSTN_MST
            var sql2 = PetaPoco.Sql.Builder.Append(@"
select
    yearmon
from
    g_act_emp_pstn_mst
where
    yearmon = @yearmon and
    rownum = 1
", new { yearmon = yearmon });

            #endregion SQL

            //G_PSTN_HSP_MST
            var o1 = db.Fetch<YearMonData>(sql1);

            //G_ACT_EMP_PSTN_MST
            var o2 = db.Fetch<YearMonData>(sql2);

            bool ret = o1.Count > 0 && o2.Count > 0;
            log.DebugFormat("  -- ExistsYearMon: YEARMON={0}, EXISTS={1}", yearmon, ret);

            return ret;
        }

        /// <summary>
        /// DSMに送付するデータを生成する
        /// (支払い実績データが 医師:施設=1:1 のもの)
        /// </summary>
        /// <returns>Dictionary[key=GK_KB+POSITION_CD,value=List]</returns>
        private Dictionary<string, List<OutputData>> GetSendDataForDsm(string yearmon)
        {
            #region SQL

            //DWHの施設-役割紐付けデータから役割コードを取得する
            var sql = PetaPoco.Sql.Builder.Append(@"
select
    T1.dr_name,
    T1.dr_ktkn_nm,
    T1.acnt_nm,
    T1.acnt_ktkn_nm,
    T1.ncc_dept,
    T1.category,
    T1.siharai,
    T1.goukei,
    T1.gk_kb,
    T1.position_cd,
    AE.EMP_NM,
    AE.GKKB_NM,
    AE.RS_NM,
    AE.DS_NM
from
    (
        select distinct
            CR.*,
            PH.gk_kb,
            PH.position_cd
        from
            c_result CR
            left join g_pstn_hsp_mst PH
                on CR.ncc_cd = PH.ncc_cd
        where
            PH.yearmon = @yearmon and
            CR.ncc_count = 1
    ) T1
    left join g_act_emp_pstn_mst AE
        on AE.gk_kb = T1.gk_kb and
            AE.local_position_cd = T1.position_cd
where
    AE.yearmon = @yearmon
order by
    T1.dr_code,
    T1.ncc_cd,
    T1.gk_kb,
    T1.position_cd
", new { yearmon = yearmon });

            #endregion SQL

            //key=GK_KB + POSITION_CD
            //value=List<OutputData>
            var rd = new Dictionary<string, List<OutputData>>();

            List<OutputData> os = db.Fetch<OutputData>(sql);
            foreach (OutputData od in os)
            {
                //<治療領域>_<支店名>_<課名>
                string key = string.Format("{0}_{1}_{2}", od.gk_kb, od.rs_nm, od.ds_nm);

                if (!rd.ContainsKey(key))
                {
                    rd.Add(key, new List<OutputData>());
                }
                rd[key].Add(od);
            }

            return rd;
        }

        /// <summary>
        /// MRに送付するデータを生成する
        /// (支払い実績データが 医師:施設=1:N のもの)
        /// </summary>
        /// <returns>Dictionary[key=GK_KB+POSITION_CD,value=List]</returns>
        private Dictionary<string, List<OutputData>> GetSendDataForMr(string yearmon)
        {
            #region SQL
            //DWHの施設-役割紐付けデータから役割コードを取得する
            var sql = PetaPoco.Sql.Builder.Append(@"
select
    T1.dr_name,
    T1.dr_ktkn_nm,
    T1.acnt_nm,
    T1.acnt_ktkn_nm,
    T1.ncc_dept,
    T1.category,
    T1.siharai,
    T1.goukei,
    T1.gk_kb,
    T1.position_cd,
    AE.EMP_NM,
    AE.GKKB_NM,
    AE.RS_NM,
    AE.DS_NM
from
    (
        select distinct
            CR.*,
            PH.gk_kb,
            PH.position_cd
        from
            c_result CR
            left join g_pstn_hsp_mst PH
                on CR.ncc_cd = PH.ncc_cd
        where
            PH.yearmon = @yearmon and
            CR.ncc_count > 1
    ) T1
    left join g_act_emp_pstn_mst AE
        on AE.gk_kb = T1.gk_kb and
            AE.local_position_cd = T1.position_cd
where
    AE.yearmon = @yearmon
order by
    T1.dr_code,
    T1.ncc_cd,
    T1.gk_kb,
    T1.position_cd
", new { yearmon = yearmon });

            #endregion SQL
            
            //key=GK_KB + POSITION_CD
            //value=List<OutputData>
            var rd = new Dictionary<string, List<OutputData>>();

            List<OutputData> os = db.Fetch<OutputData>(sql);
            foreach (OutputData od in os)
            {
                string key = od.gk_kb + "_" + od.position_cd;

                if (!rd.ContainsKey(key))
                {
                    rd.Add(key, new List<OutputData>());
                }
                rd[key].Add(od);
            }

            return rd;
        }

        #endregion DWH 施設-役割 紐付け反映

        #region Excel出力

        /// <summary>
        /// Excelファイル出力処理
        /// </summary>
        /// <returns></returns>
        public string OutputExcelFiles()
        {
            //年月
            string yearmon = GetYearMon();
            //templateファイル
            string template_file = GetTemplateFile();
            if (string.IsNullOrEmpty(template_file))
            {
                return string.Empty;
            }
            //出力フォルダ
            string output_folder = CreateOutputFolder();
            if (!Directory.Exists(output_folder))
            {
                return string.Empty;
            }

            log.DebugFormat("  -- OutputExcelFiles: yearmon={0}, template_file={1}, output_folder={2}",
                yearmon, template_file, output_folder);

            //DSM向けファイル出力
            int dsm_count = OutputFileForDsm(yearmon, template_file, output_folder);

            //MR向けファイル出力
            int mr_count = OutputFileForMr(yearmon, template_file, output_folder);

            //完了メッセージ
            //TODO 出力フォルダを共有フォルダパスに変換する
            string message = string.Format(@"
データ作成が完了しました。

出力フォルダ: {0}
    DSM向け:     {1} 件
    MR向け:      {2} 件
", output_folder, dsm_count, mr_count);

            return message;
        }

        /// <summary>
        /// DSM向けファイル出力
        /// </summary>
        /// <param name="yearmon"></param>
        /// <param name="template_file"></param>
        /// <param name="output_dir"></param>
        /// <returns></returns>
        private int OutputFileForDsm(string yearmon, string template_file, string output_dir)
        {
            //DSM送付データの取得
            var data = GetSendDataForDsm(yearmon);
            int file_count = 0;
            foreach (var code in data.Keys)
            {
                //ファイル名: <治療領域>-<支店>-<課名>.xlsx
                var head_data = data[code][0]; //先頭行取得
                string file_name = output_dir + @"\" + string.Format("{0}-{1}-{2}.xlsx",
                    head_data.gkkb_nm, head_data.rs_nm, head_data.ds_nm).Replace(" ", "_");

                log.DebugFormat("  -- OutputFileForDsm: file_name={0}", file_name);

                //ファイルコピー (上書き)
                File.Copy(template_file, file_name, true);

                //ファイル書き込み
                WriteExcel(file_name, data[code]);
                
                file_count++;
            }

            return file_count;
        }

        /// <summary>
        /// MR向けファイル出力
        /// </summary>
        /// <param name="yearmon"></param>
        /// <param name="template_file"></param>
        /// <param name="output_dir"></param>
        /// <returns></returns>
        private int OutputFileForMr(string yearmon, string template_file, string output_dir)
        {
            //MR送付データの取得
            var data = GetSendDataForMr(yearmon);
            int file_count = 0;
            foreach (var code in data.Keys)
            {
                //ファイル名: <治療領域>-<支店>-<課名>-<MR名>.xlsx
                var head_data = data[code][0]; //先頭行取得
                string file_name = output_dir + @"\" + string.Format("{0}-{1}-{2}-{3}.xlsx",
                    head_data.gkkb_nm, head_data.rs_nm, head_data.ds_nm, head_data.emp_nm).Replace(" ", "_");

                log.DebugFormat("  -- OutputFileForMr: file_name={0}", file_name);

                //ファイルコピー (上書き)
                File.Copy(template_file, file_name, true);

                //ファイル書き込み
                WriteExcel(file_name, data[code]);

                file_count++;
            }

            return file_count;
        }

        /// <summary>
        /// app.configからtemplateファイルのパスを取得する
        /// </summary>
        /// <returns></returns>
        private string GetTemplateFile()
        {
            //template file取得
            string template_file = ConfigurationManager.AppSettings["excel_template"];
            if (string.IsNullOrEmpty(template_file))
            {
                log.Error("Excel出力用テンプレートファイルを指定してください。");
                return string.Empty;
            }
            if (!File.Exists(template_file))
            {
                log.ErrorFormat("Excel出力用テンプレートファイルが存在しません。 template_file=\"{0}\"", template_file);
                return string.Empty;
            }
            return template_file;
        }

        /// <summary>
        /// 出力フォルダ作成
        /// </summary>
        /// <returns></returns>
        private string CreateOutputFolder()
        {
            string _of = ConfigurationManager.AppSettings["output_folder"];
            if (string.IsNullOrEmpty(_of))
            {
                //実行ファイルのあるフォルダ
                _of = Path.GetDirectoryName(
                    System.Reflection.Assembly.GetExecutingAssembly().Location);
            }

            string tm = DateTime.Now.ToString("yyyyMMddHHmmss");
            _of += @"\" + tm;

            //フォルダ作成
            if (!Directory.Exists(_of))
            {
                Directory.CreateDirectory(_of);
                log.DebugFormat("  -- create directory={0}", _of);
            }

            return _of;
        }

        /// <summary>
        /// Excelファイル書き込み処理
        /// </summary>
        /// <param name="path">出力ファイル</param>
        /// <param name="data">出力データ</param>
        private void WriteExcel(string path, List<OutputData> data)
        {
            //シート名
            string sheetName = ConfigurationManager.AppSettings["output_sheet"];
            if (string.IsNullOrEmpty(sheetName))
            {
                sheetName = @"Sheet1";
            }

            try
            {
                //FileStream fs = new FileStream(path, FileMode.Open)
                FileInfo fs = new FileInfo(path);
                
                using (ExcelPackage xls = new ExcelPackage(fs))
                {
                    //xls.Load(fs);
                    ExcelWorksheet sheet = xls.Workbook.Worksheets[sheetName];

                    //書き込み
                    int rowIndex = 2;
                    foreach (OutputData d in data)
                    {
                        sheet.Cells[rowIndex, 1].Value = d.dr_name;         //医師氏名
                        sheet.Cells[rowIndex, 2].Value = d.dr_ktkn_nm;      //医師氏名カナ
                        sheet.Cells[rowIndex, 3].Value = d.acnt_nm;         //施設名
                        sheet.Cells[rowIndex, 4].Value = d.acnt_ktkn_nm;    //施設名カナ
                        sheet.Cells[rowIndex, 5].Value = d.ncc_dept;        //科名
                        sheet.Cells[rowIndex, 6].Value = d.category;        //カテゴリー
                        sheet.Cells[rowIndex, 7].Value = d.siharai;         //支払い回数
                        sheet.Cells[rowIndex, 8].Value = d.goukei;          //合計金額

                        rowIndex++;
                    }
                    xls.Save();

                }
                
                log.DebugFormat("  -- WriteExcel={0}, rows={1}", path, data.Count);
            }
            catch (Exception exp)
            {
                log.Error("  -- WriteExcel", exp);
            }
        }

        #endregion Excel出力
    }
}
