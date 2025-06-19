using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using _001TN0173.Entities;

namespace _001TN0173.Shared
{
    class Clsbase
    {
        Settings sttg = new Settings();
        ClsErrorLog clserr = new ClsErrorLog();
        //public string[] Create_NEFT_FileSheet     = { "Transaction_Ref_No", "Amount", "Value_Date", "Branch_Code", "Sender_Account_Type", "Remitter_Account_No", "Remitter_Name",
        //                                              "IFSC_Code","Debit_Account","Beneficiary_Account_type","Bank_Account_Number","Beneficiary_Name","Remittance_Details","Debit_Account_System","Originator_Of_Remmittance"};
        ////public string[] Create_FT_FileSheet       = { "TDealer code,Name" , "Dealer account number" , "Amount" , "Ref no" , "DO Number" , "Status" }; ////Change in field on discuss with hdfc on 14-12-2023
        //public string[] Create_FT_FileSheet       = { "UTR_No", "Dealer account number", "Amount", "DO Number", "Rejection Status" };
        //public string[] Create_RTGS_FileSheet     = { "Related Reference no" , "IFSC Code" , "Value_Date" , "Branch_Code" , "Branch Code Ordering A/c No" , "Amount" , "Sender to Receiver-INFORMATION" };
        public string[] Create_NEFT_FileSheet = { "Product", "Remitter Name", "Remitter A/ C No", "Remitter Bank Name" ,"UTR No", "Amount", "IFSC Code", "VirtualAccountNumber", "RejectedReason" };
        //public string[] Create_FT_FileSheet       = { "TDealer code,Name" , "Dealer account number" , "Amount" , "Ref no" , "DO Number" , "Status" }; ////Change in field on discuss with hdfc on 14-12-2023
        public string[] Create_FT_FileSheet = { "Product", "Remitter Name", "Remitter A/ C No", "Remitter Bank Name", "UTR No", "Amount", "IFSC Code", "VirtualAccountNumber", "RejectedReason" };
        public string[] Create_RTGS_FileSheet = { "Product", "Remitter Name", "Remitter A/ C No", "Remitter Bank Name", "UTR No", "Amount", "IFSC Code", "VirtualAccountNumber", "RejectedReason" };

        public string[] Create_Reg_FileSheet = { "SR.No", "INVOICE NUMBER" , "Inv value" , "Currency" , "Description of goods", "DUE DATE", "BUYER  Name" , "Address" , "City",
                                                      "Transporter name","Transport number(L/R or D/C or GRN or MRIR)","Transport date","Dealer Code","Transporter Code","Dealer Address line 2",
                                                      "Dealer Address Line 3","Dealer Address Line 4","Trade refrence No","Physical Received date", "Remarks"};
        public string[] Create_Reg_File_DSheet = { "SR.No", "INVOICE NUMBER" , "Inv value" , "Currency" , "Description of goods", "DUE DATE", "BUYER  Name" , "Address" , "City",
                                                      "Transporter name","Transport number(L/R or D/C or GRN or MRIR)","Transport date","Dealer Code","Transporter Code","Dealer Address line 2",
                                                      "Dealer Address Line 3","Dealer Address Line 4", "Remarks"};
        public string[] Reg_Invoice_Cancel_Sheet = { "SR.No", "INVOICE NUMBER", "Amount" };
        public string[] FCC_FileSheet = { "SR NO.", "Dealer Code", "Dealer Name", "Trade Ref. NO.", "No of Invoices Physically Received", "Sum of Invoices Physcially Received", "Date of Receipt Physical" };

        public bool isCompleteFileAvailable(string szFilePath)
        {
            bool functionReturnValue = false;

            FileStream fsObj = default(FileStream);
            StreamWriter obOpenFile = default(StreamWriter);
            try
            {
                while (true)
                {
                    try
                    {
                        if (File.Exists(szFilePath))
                        {
                            fsObj = new FileStream(szFilePath, FileMode.Append, FileAccess.Write, FileShare.None);
                            obOpenFile = new StreamWriter(fsObj);
                            functionReturnValue = true;
                        }
                        else
                        {
                            functionReturnValue = false;
                            break; // TODO: might not be correct. Was : Exit While
                        }

                    }
                    catch (Exception ex)
                    {
                        clserr.Handle_Error(ex, "Form1", "isCompleteFileAvailable");
                        functionReturnValue = false;
                        System.Threading.Thread.Sleep(1000);

                    }
                    finally
                    {
                        if ((fsObj != null))
                            fsObj.Flush();
                        if ((obOpenFile != null))
                            obOpenFile.Dispose();
                        fsObj = null;
                        obOpenFile = null;

                    }

                    if (functionReturnValue == true)
                        break; // TODO: might not be correct. Was : Exit While
                }

            }
            catch (Exception ex)
            {
                clserr.Handle_Error(ex, "Form1", "isCompleteFileAvailable");

            }
            return functionReturnValue;

        }

        private int DetermineNumberofDays(string dtStartDate)
        {
            TimeSpan tsTimeSpan;
            int iNumberOfDays;
            DateTime DtSystem;
            DateTime _dtStartDate;
            bool isValid1 = DateTime.TryParseExact(DateTime.Now.ToString("dd/MM/yyyy"), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DtSystem);
            bool isValid2 = DateTime.TryParseExact(dtStartDate, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out _dtStartDate);
            tsTimeSpan = DtSystem.Subtract(_dtStartDate);
            iNumberOfDays = tsTimeSpan.Days;
            return iNumberOfDays;
        }

    }

    class ClsErrorLog : IDisposable
    {
        Settings sttg = new Settings();

        private bool disposed = false;
        #region IDisposable Members
        public void Dispose()
        {
            // Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {

                }
            }
            disposed = true;
        }
        #endregion

        #region Handle_Error Region
        public void Handle_Error(Exception oErr, string strFormName, string strFunctionName)
        {
            try
            {
                WriteErrorToTxtFile(oErr.Message, strFormName, strFunctionName);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region Write Error to Text File
        public void WriteErrorToTxtFile(string ErrorDesc, string ModuleName, string ProcName)
        {
            string strfilename = string.Empty;
            string strErrorString = string.Empty;
            try
            {
                string strExeFilePath = System.Reflection.Assembly.GetExecutingAssembly().Location;

                strErrorString = "[" + DateTime.Now.ToString("dd MM yyyy") + "] [ " + ErrorDesc + "] [ " + ModuleName + "] [ " + ProcName + "]";
                strfilename = sttg.ErrorLog + "\\" + ModuleName + ".log";

                FileStream fsObj;
                StreamWriter SwOpenFile;

                if (File.Exists(strfilename))
                {
                    fsObj = new FileStream(strfilename, FileMode.Append, FileAccess.Write, FileShare.Read);
                }
                else
                {
                    fsObj = new FileStream(strfilename, FileMode.Create, FileAccess.Write, FileShare.Read);
                }

                SwOpenFile = new StreamWriter(fsObj);
                SwOpenFile.WriteLine(strErrorString);

                fsObj.Flush();
                SwOpenFile.Dispose();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region  Build String as per given string and padding postion
        public string GetBuildString(string strString, int iTotPadd)
        {
            string strBuld = "";
            try
            {
                if (strString.Length >= iTotPadd)
                {
                    strBuld = strString.PadRight(iTotPadd, ' ').Substring(0, iTotPadd);
                }
                else
                {
                    strBuld = strString.PadRight(iTotPadd, ' ');
                }

                return strBuld;
            }
            catch (Exception ex)
            { }
            return strBuld;
        }
        #endregion

        public void LogEntry(string StrMessage, bool IsError = false)
        {
            try
            {
                string LogPath = string.Empty;
                string LogFileName = string.Empty;
                StrMessage = "[" + DateTime.Now.Day + " " + DateTime.Now.Month + " " + DateTime.Now.Year + " " + DateTime.Now.Hour + " " + DateTime.Now.Minute + " " + DateTime.Now.Second + " ]" + StrMessage;
                ReadWriteAppSettings readWriteAppSettings = new ReadWriteAppSettings();
                sttg = readWriteAppSettings.ReadGetSectionAppSettings();
                if (IsError == true)
                {
                    LogPath = sttg.ErrorLog + "\\";
                    LogFileName = LogPath + "Error_" + DateTime.Now.ToString("ddMMyyyy") + ".log";
                }
                else
                {
                    LogPath = sttg.AuditLog + "\\";
                    LogFileName = LogPath + "Log_" + DateTime.Now.ToString("ddMMyyyy") + ".log";
                }

                if (!Directory.Exists(LogPath))
                {
                    Directory.CreateDirectory(LogPath);
                }
                FileStream fsObj;
                StreamWriter SwOpenFile;

                if (File.Exists(LogFileName))
                {
                    fsObj = new FileStream(LogFileName, FileMode.Append, FileAccess.Write, FileShare.Read);
                }
                else
                {
                    fsObj = new FileStream(LogFileName, FileMode.Create, FileAccess.Write, FileShare.Read);
                }

                SwOpenFile = new StreamWriter(fsObj);
                SwOpenFile.WriteLine(StrMessage);

                fsObj.Flush();
                SwOpenFile.Dispose();
            }
            catch (Exception ex)
            {
                WriteErrorToTxtFile(ex.Message, "LogEntry" + "   LogEntry", StrMessage);
            }
        }

    }
}
