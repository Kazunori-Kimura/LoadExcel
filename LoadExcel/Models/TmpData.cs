using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LoadExcel.Models
{
    [PetaPoco.TableName("tmp_data1")]
    class TmpData
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
        public string dm_cont_id { get; set; }
        public string dm_doccd { get; set; }
        public string dm_name { get; set; }
        public string dm_ktkn_nm { get; set; }
        public long hm_act_uni_id { get; set; }
        public string hm_ncc_cd { get; set; }
        public string hm_acnt_nm { get; set; }
        public string hm_acnt_jpln_nm { get; set; }
        public string hm_acnt_ktkn_nm { get; set; }
    }
}
