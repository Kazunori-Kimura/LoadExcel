using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LoadExcel.Models
{
    /// <summary>
    /// conference pack data
    /// </summary>
    [PetaPoco.TableName("cp_data")]
    class CpData
    {
        public int id { get; set; }
        public string drcode { get; set; }
        public string drname { get; set; }
        public string ncc_cd { get; set; }
        public string ncc_name { get; set; }
        public string ncc_dept { get; set; }
        public string title { get; set; }
        public string category { get; set; }
        public long kingaku { get; set; }
        public int kaisu { get; set; }
        //医師マスタより取得
        public string dm_cont_id { get; set; }
        public string dm_doccd { get; set; }
        public string dm_name { get; set; }
        public string dm_ktkn_nm { get; set; }
        //施設マスタより取得
        public long hm_act_uni_id { get; set; }
        public string hm_ncc_cd { get; set; }
        public string hm_name { get; set; }
        public string hm_jpln_nm { get; set; }
        public string hm_ktkn_nm { get; set; }
    }
}
