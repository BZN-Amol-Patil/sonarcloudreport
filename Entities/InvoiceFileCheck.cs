using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace _001TN0173.Entities
{
    //[Table("File_Desc")]
    public class InvoiceFileCheck
    {
        //[Key]
        public string FileType { get; set; }
        public string File_Date { get; set; }
    }
}
