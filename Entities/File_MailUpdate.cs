using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace _001TN0173.Entities
{
    //[Table("File_Mail")]
    class File_MailUpdate
    {
        //[Key]
        public string File_Mail_Time { get; set; }
        public string File_Mail_Dates { get; set; }
        public string File_Mail_Names { get; set; }
       
    }
}
