using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace _001TN0173.Entities
{
    [Table("File_Mail")]
    class File_MailF5
    {
        [Key]
        public string File_Mail_Name { get; set; }
        public string File_Mail_Date { get; set; }
    }
}
