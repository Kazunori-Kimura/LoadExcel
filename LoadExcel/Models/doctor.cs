using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LoadExcel.Models
{
    /// <summary>
    /// MegaCOARA 医師マスタ
    /// </summary>
    [PetaPoco.TableName("coa_doc1p")]
    class doctor
    {
        public string cont_id { get; set; }
        public string doccd { get; set; }
        public string name { get; set; }
        public string ktkn_nm { get; set; }
    }
}
