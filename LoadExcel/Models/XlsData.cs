using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LoadExcel.Models
{
    [PetaPoco.TableName("xls_data")]
    class XlsData
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
    }
}
