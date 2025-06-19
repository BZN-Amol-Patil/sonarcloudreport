using _001TN0173.Entities;
using _001TN0173.Shared;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _001TN0173
{
    class ReadWriteAppSettings
    {
        public Settings sttg = new Settings();
        ClsErrorLog clserr = new ClsErrorLog();

        public Settings ReadGetSectionAppSettings()
        {
            try
            {
                var builder = new ConfigurationBuilder()
                                .SetBasePath(Directory.GetCurrentDirectory() + "\\")
                                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

                IConfigurationRoot configuration = builder.Build();


                //sttg.InputFilePath = Environment.CurrentDirectory + "\\" + configuration.GetSection("MSILSettings:InputFilePath").Value;
                //sttg.OutputFilePath = Application.StartupPath + configuration.GetSection("MSILSettings:OutputFilePath").Value;
                //sttg.BackupFilePath = Application.StartupPath + configuration.GetSection("MSILSettings:BackupFilePath").Value;
                //sttg.NonConvertedFile = Application.StartupPath + configuration.GetSection("MSILSettings:NonConvertedFile").Value;

                sttg.InputFilePath =  configuration.GetSection("MSILSettings:InputFilePath").Value;
                sttg.OutputFilePath =  configuration.GetSection("MSILSettings:OutputFilePath").Value;
                sttg.BackupFilePath =  configuration.GetSection("MSILSettings:BackupFilePath").Value;
                sttg.NonConvertedFile =  configuration.GetSection("MSILSettings:NonConvertedFile").Value;
                sttg.ErrorLog = Application.StartupPath + configuration.GetSection("MSILSettings:ErrorLog").Value;
                sttg.AuditLog = Application.StartupPath + configuration.GetSection("MSILSettings:AuditLog").Value;
                sttg.BlankExcel = Application.StartupPath + configuration.GetSection("MSILSettings:BlankExcel").Value;
                sttg.Document_Release_Template = Application.StartupPath + configuration.GetSection("MSILSettings:Document_Release_Template").Value;
                sttg.Document_Release = Application.StartupPath + configuration.GetSection("MSILSettings:Document_Release").Value;
                sttg.Payment_Rejection = Application.StartupPath + configuration.GetSection("MSILSettings:Payment_Rejection").Value;
                sttg.Payment_Rejection_Date = Application.StartupPath + configuration.GetSection("MSILSettings:Payment_Rejection_Date").Value;
                sttg.TRADE_EMAIL = configuration.GetSection("MSILSettings:TRADE_EMAIL").Value;
                sttg.Frequency = configuration.GetSection("MSILSettings:Frequency").Value;
                sttg.INV_CONF_EMAIL = configuration.GetSection("MSILSettings:INV_CONF_EMAIL").Value;
                sttg.INV_CONF_FNAME = configuration.GetSection("MSILSettings:INV_CONF_FNAME").Value;
                sttg.ORD_MIS_EMAIL = configuration.GetSection("MSILSettings:ORD_MIS_EMAIL").Value;
                sttg.PAYREC_TRADE_EMAIL = configuration.GetSection("MSILSettings:PAYREC_TRADE_EMAIL").Value;
                sttg.PHY_INV = configuration.GetSection("MSILSettings:PHY_INV").Value;
                sttg.NO_INV = configuration.GetSection("MSILSettings:NO_INV").Value;
                sttg.EOD_MIS = configuration.GetSection("MSILSettings:EOD_MIS").Value;
                sttg.ORD_DEL = configuration.GetSection("MSILSettings:ORD_DEL").Value;
                sttg.INV_PHY_NTRC = configuration.GetSection("MSILSettings:INV_PHY_NTRC").Value;
                //sttg.IntraDayPath = Application.StartupPath + configuration.GetSection("MSILSettings:IntraDayPath").Value;
                sttg.IntraDayPath =  configuration.GetSection("MSILSettings:IntraDayPath").Value;
                sttg.MISOutputFilePath = sttg.IntraDayPath;
                sttg.EOD_File_EMail = configuration.GetSection("MSILSettings:EOD_File_EMail").Value;
                sttg.FCC_EMail = configuration.GetSection("MSILSettings:FCC_EMail").Value;
                sttg.DRC_EMail = configuration.GetSection("MSILSettings:DRC_EMail").Value;
                sttg.Payment_Rejection_EMail = configuration.GetSection("MSILSettings:Payment_Rejection_EMail").Value;
                sttg.DRC_EMail_BNGR = configuration.GetSection("MSILSettings:DRC_EMail_BNGR").Value;
                sttg.DRC_EMail_SLGR = configuration.GetSection("MSILSettings:DRC_EMail_SLGR").Value;
                sttg.DO_Cancel_Email = configuration.GetSection("MSILSettings:DO_Cancel_Email").Value;
                sttg.DO_Invoice_Cancel_Email = configuration.GetSection("MSILSettings:DO_Invoice_Cancel_Email").Value;
                sttg.Invoice_Cancel_Email = configuration.GetSection("MSILSettings:Invoice_Cancel_Email").Value;
                sttg.Sleep_Time_in_Mint = configuration.GetSection("MSILSettings:Sleep_Time_in_Mint").Value;

                sttg.Confirmation_Mail = configuration.GetSection("EmailSetting:Confirmation_Mail").Value;
                sttg.SMTP_HOST = configuration.GetSection("EmailSetting:SMTP_HOST").Value;
                sttg.Port = configuration.GetSection("EmailSetting:Port").Value;
                sttg.Email_FromID = configuration.GetSection("EmailSetting:Email_FromID").Value;
                sttg.UserID = configuration.GetSection("EmailSetting:UserID").Value;
                sttg.Password = configuration.GetSection("EmailSetting:Password").Value;

                sttg.SysEmail_FromID = configuration.GetSection("SystemSetting:Pwd").Value;
                sttg.PWD = configuration.GetSection("SystemSetting:Pwd").Value;
                sttg.PWD = sttg.PWD.ToString();
            }
            catch(Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message + "ReadGetSectionAppSettings", "ReadGetSectionAppSettings", "");
            }
            return sttg;
        }

       
        public void BindAppSettings()
        {
            try
            {
                Settings appSettings = new Settings();
                var builder = new ConfigurationBuilder()
                                .SetBasePath(Directory.GetCurrentDirectory())
                                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

                var configuration = builder.Build();
                var configurationResult = configuration.GetSection("MSILSettings");
                ConfigurationBinder.Bind(configurationResult, appSettings);
                string result_KeyA_Sec1 = appSettings.InputFilePath;
                string result_KeyB_Sec1 = appSettings.OutputFilePath;
                string result_KeyC_Sec1 = appSettings.BackupFilePath;
                string result_KeyD_Sec1 = appSettings.NonConvertedFile;
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message + "ReadWriteAppSettings", "BindAppSettings", "");
            }


        }


    }
}
