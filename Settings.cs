using System;
using System.Collections.Generic;
using System.Text;

namespace _001TN0173
{
  public  class Settings
    {
        //MSILSettings
        public string InputFilePath { get; set; }

        public string OutputFilePath { get; set; }
        public string BackupFilePath { get; set; }
        public string NonConvertedFile { get; set; }
        public string ErrorLog { get; set; }
        public string Document_Release_Template { get; set; }
        public string Document_Release { get; set; }
        public string Payment_Rejection { get; set; }
        public string Payment_Rejection_Date { get; set; }
        public string AuditLog { get; set; }
        public string BlankExcel { get; set; }
        public string TRADE_EMAIL { get; set; }
        public string Frequency { get; set; }
        public string INV_CONF_EMAIL { get; set; }
        public string INV_CONF_FNAME { get; set; }
        public string ORD_MIS_EMAIL { get; set; }
        public string PAYREC_TRADE_EMAIL { get; set; }
        public string PHY_INV { get; set; }
        public string NO_INV { get; set; }
        public string EOD_MIS { get; set; }
        public string ORD_DEL { get; set; }
        public string INV_PHY_NTRC { get; set; }
        public string IntraDayPath { get; set; }
        public string Sleep_Time_in_Mint { get; set; }
        public string MISOutputFilePath { get; set; }
        public string EOD_File_EMail { get; set; }
        public string FCC_EMail { get; set; }
        public string DRC_EMail { get; set; }
        public string Payment_Rejection_EMail { get; set; }
        public string DRC_EMail_BNGR { get; set; }
        public string DRC_EMail_SLGR { get; set; }
        public string DO_Cancel_Email { get; set; }
        public string DO_Invoice_Cancel_Email { get; set; }
        public string Invoice_Cancel_Email { get; set; }
      


        //EmailSetting
        public string Confirmation_Mail { get; set; }
        public string SMTP_HOST { get; set; }
        public string Port { get; set; }
        public string Email_FromID { get; set; }
        public string UserID { get; set; }
        public string Password { get; set; }

        //SystemSetting
        public string SysEmail_FromID { get; set; }
        public string PWD { get; set; }

    }
}
