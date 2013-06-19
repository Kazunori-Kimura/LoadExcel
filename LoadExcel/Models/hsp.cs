using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LoadExcel.Models
{
    [PetaPoco.TableName("coa_dcfhsp")]
    class hsp
    {
        public long act_uni_id { get; set; }
        public string hs01 { get; set; }
        public string hs02 { get; set; }
        public string acnt_nm { get; set; }
        public string acnt_jpln_nm { get; set; }
        public string acnt_ktkn_nm { get; set; }
    }
}
