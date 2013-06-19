using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.IO;

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
        /// コンストラクタ
        /// </summary>
        public DataManager()
        {
            string cs = string.Format(@"User Id={0};Password={1};Data Source={2};",
                "TRA_USER", "1qaz2wsx", "HOGE");

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
        public List<CpData> LoadExcel(string path, string sheetName, int start_row)
        {
            log.Debug(string.Format("target file= {0}; sheet= {1}", path, sheetName));

            int rowIndex = start_row;
            List<CpData> records = new List<CpData>();

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
                            CpData cd = new CpData();

                            cd.id = rowIndex;
                            // read row data
                            cd.drcode = ToString(sheet.Cells[rowIndex, 1].Value);
                            cd.drname = ToString(sheet.Cells[rowIndex, 2].Value);
                            if (string.IsNullOrEmpty(cd.drname))
                            {
                                break; //Dr名が無ければ終了
                            }
                            cd.ncc_cd = ToString(sheet.Cells[rowIndex, 3].Value);
                            cd.ncc_name = ToString(sheet.Cells[rowIndex, 4].Value);
                            cd.ncc_dept = ToString(sheet.Cells[rowIndex, 5].Value);
                            cd.title = ToString(sheet.Cells[rowIndex, 6].Value);
                            cd.category = ToString(sheet.Cells[rowIndex, 7].Value);
                            cd.kingaku = ToLong(sheet.Cells[rowIndex, 8].Value);
                            cd.kaisu = ToInt(sheet.Cells[rowIndex, 9].Value);

                            // 頭 '0' 埋め
                            cd.drcode = cd.drcode.PadLeft(6, '0');
                            if (cd.drcode == "000000")
                            {
                                cd.drcode = string.Empty;
                            }

                            records.Add(cd);
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
        /// cp_dataにデータを登録する
        /// </summary>
        /// <param name="items"></param>
        /// <returns></returns>
        public int BulkInsert(List<CpData> items)
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
        /// CP_DATAテーブルを初期化
        /// </summary>
        private void truncate()
        {
            string query = @"truncate table cp_data";
            db.Execute(query);
        }

        #endregion

        #region MegaCOARA 医師マスタ・施設マスタ取得

        /// <summary>
        /// 医師マスタから医師の情報を取得する
        /// </summary>
        /// <param name="cp_data"></param>
        public CpData UpdateDoctorInfo(CpData cp_data)
        {
            if (!string.IsNullOrEmpty(cp_data.drcode))
            {
                //医師コードを元に医師マスタを検索
                var ds = GetDoctorByCode(cp_data.drcode);

                if (ds.Count == 0)
                {
                    //氏名を元に医師マスタを検索
                    ds = GetDoctorByName(cp_data.drname);
                }

                if (ds.Count != 1)
                {
                    log.Warn(string.Format("医師を特定することができませんでした。: Excel行={0}, 医師コード={1}, 氏名={2}",
                        cp_data.id, cp_data.drcode, cp_data.drname));

                    if (ds.Count == 0)
                    {
                        log.Warn("    候補データ: なし");
                    }
                    else
                    {
                        foreach (var d in ds)
                        {
                            log.Warn(string.Format("    候補データ: DOCCD={0}, NAME={1}", d.doccd, d.name));
                        }
                    }
                    return cp_data;
                }

                //更新
                cp_data.dm_cont_id = ds[0].cont_id;
                cp_data.dm_doccd = ds[0].doccd;
                cp_data.dm_ktkn_nm = ds[0].ktkn_nm;
                cp_data.dm_name = ds[0].name;

                log.Debug(string.Format("医師データ取得: CONT_ID={0}, DOCCD={1}, NAME={2}", cp_data.dm_cont_id, cp_data.dm_doccd, cp_data.dm_name));
            }

            return cp_data;
        }

        /// <summary>
        /// doccdを元に医師マスタよりデータを取得
        /// </summary>
        /// <param name="doccd">医師コード</param>
        /// <returns>医師マスタのデータ</returns>
        public List<doctor> GetDoctorByCode(string doccd)
        {
            var sql = PetaPoco.Sql.Builder.Append("select cont_id, doccd, name, ktkn_nm")
                .Append("from coa_doc1p")
                .Append("where (drdymd = 0 or drdymd is null) and")
                .Append("doccd=@doccd", new { doccd = doccd });

            //医師マスタよりデータ取得
            var ds = db.Fetch<doctor>(sql);

            log.Debug(string.Format("doccd={0}, Count={1}", doccd, ds.Count));
            return ds;
        }

        /// <summary>
        /// 氏名を元に医師マスタよりデータを取得
        /// </summary>
        /// <param name="name">氏名</param>
        /// <returns>医師マスタのデータ</returns>
        public List<doctor> GetDoctorByName(string name)
        {
            //nameから半角スペース、全角スペースを除去
            string docname = name.Replace("　", "").Replace(" ", "");

            var sql = PetaPoco.Sql.Builder.Append("select cont_id, doccd, name, ktkn_nm")
                .Append("from coa_doc1p")
                .Append("where (drdymd = 0 or drdymd is null) and")
                .Append("replace(replace(name, '　', ''), ' ', '') = @docname", new { docname = docname });

            //医師マスタよりデータ取得
            var ds = db.Fetch<doctor>(sql);

            log.Debug(string.Format("docname={0}, Count={1}", docname, ds.Count));
            return ds;
        }


        /// <summary>
        /// 施設マスタから施設の情報を取得する
        /// </summary>
        /// <param name="cp_data"></param>
        public CpData UpdateHspInfo(CpData cp_data)
        {


            if (!string.IsNullOrEmpty(cp_data.ncc_cd))
            {
                //施設コードを元に施設マスタを検索
                var hs = GetHspByNccCode(cp_data.ncc_cd);
                
                if (hs.Count == 0)
                {
                    //施設名で検索
                    hs = GetHspByName(cp_data.ncc_name);
                }

                if (hs.Count != 1)
                {
                    log.Warn(string.Format("施設を特定することができませんでした。: Excel行={0}, 施設コード={1}, 施設名={2}",
                        cp_data.id, cp_data.ncc_cd, cp_data.ncc_name));

                    if (hs.Count == 0)
                    {
                        log.Warn("    候補データ: なし");
                    }
                    else
                    {
                        foreach (var h in hs)
                        {
                            log.Warn(string.Format("    候補データ: 施設コード={0}{1}, 施設名={2}", h.hs01, h.hs02, h.acnt_nm));
                        }
                    }
                    return cp_data;
                }

                //施設情報を更新
                cp_data.hm_act_uni_id = hs[0].act_uni_id;
                cp_data.hm_ncc_cd = hs[0].hs01 + hs[0].hs02;
                cp_data.hm_name = hs[0].acnt_nm;
                cp_data.hm_jpln_nm = hs[0].acnt_jpln_nm;
                cp_data.hm_ktkn_nm = hs[0].acnt_ktkn_nm;

                log.Debug(string.Format("施設データ取得: ACT_UNI_ID={0}, NCC_CD={1}, NAME={2}", cp_data.hm_act_uni_id, cp_data.hm_ncc_cd, cp_data.hm_name));
            }

            return cp_data;
        }

        /// <summary>
        /// ncc_cdを元に施設マスタよりデータを取得
        /// </summary>
        /// <param name="ncc_cd">施設コード</param>
        /// <returns>施設マスタのデータ</returns>
        public List<hsp> GetHspByNccCode(string ncc_cd)
        {
            var sql = PetaPoco.Sql.Builder
                .Append("select act_uni_id, hs01, hs02, acnt_nm, acnt_jpln_nm, acnt_ktkn_nm")
                .Append("from coa_dcfhsp")
                .Append("where (hpdymd = 0 or hpdymd is null) and")
                .Append("hs01 || hs02 = @ncc_cd", new { ncc_cd = ncc_cd })
                .Append("order by hs01, hs02");

            //施設マスタよりデータ取得
            var hs = db.Fetch<hsp>(sql);

            log.Debug(string.Format("ncc_cd={0}, Count={1}", ncc_cd, hs.Count));
            return hs;
        }

        /// <summary>
        /// 施設名を元に施設マスタよりデータを取得
        /// (施設名の部分一致で検索)
        /// </summary>
        /// <param name="name">施設名</param>
        /// <returns></returns>
        public List<hsp> GetHspByName(string name)
        {
            //部分一致でヒットするようにワイルドカードではさむ
            string ncc_name = "%" + name + "%";

            var sql = PetaPoco.Sql.Builder
                .Append("select act_uni_id, hs01, hs02, acnt_nm, acnt_jpln_nm, acnt_ktkn_nm")
                .Append("from coa_dcfhsp")
                .Append("where (hpdymd = 0 or hpdymd is null) and")
                .Append("acnt_nm like @ncc_name or", new { ncc_name = ncc_name })
                .Append("m.acnt_jpln_nm like @ncc_name", new { ncc_name = ncc_name })
                .Append("order by hs01, hs02");

            //施設マスタよりデータ取得
            var hs = db.Fetch<hsp>(sql);

            log.Debug(string.Format("name={0}, Count={1}", name, hs.Count));
            return hs;
        }

        #endregion

        #region 一括更新

        /// <summary>
        /// 一括更新
        /// </summary>
        /// <param name="items"></param>
        public void UpdateMasterInfo(List<CpData> items)
        {
            using (var trans = db.GetTransaction())
            {
                foreach (var item in items)
                {
                    //施設情報取得
                    var c = UpdateHspInfo(item);
                    //医師情報取得
                    c = UpdateDoctorInfo(c);
                    //Database更新処理
                    db.Update("cp_data", "id", c);
                    log.Debug(string.Format("MegaCOARAマスタ情報反映: 行No={0}", c.id));
                }
                trans.Complete();
            }
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
