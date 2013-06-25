using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LoadExcel.Models
{
    /// <summary>
    /// Excel出力データ
    /// </summary>
    class OutputData
    {
        public string dr_name { get; set; }
        public string acnt_nm { get; set; }
        public string ncc_dept { get; set; }
        public string category { get; set; }
        public int siharai { get; set; }
        public long goukei { get; set; }
        public string gk_kb { get; set; }
        public string position_cd { get; set; }
        public string emp_nm { get; set; }
        public string gkkb_nm { get; set; }
        public string rs_nm { get; set; }
        public string ds_nm { get; set; }
    }
}
