using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using _001TN0173.Entities;
using _001TN0173.Shared;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using CsvHelper;
using CsvHelper.Configuration;
using ClosedXML.Excel;
using Microsoft.VisualBasic.FileIO;
//using System.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Data.SqlClient;
namespace _001TN0173
{
    public partial class Form1 : Form
    {
        Clsbase objBaseClass = new Clsbase();
        ClsErrorLog clserr = new ClsErrorLog();
        public Settings sttg = new Settings();
        FunctionsCl fncl = new FunctionsCl();
        private DataService _dataService;
        public Form1()
        {
            InitializeComponent();
        }

        public void ReadGetSectionAppSettings()
        {
            //string encr = "plnkH9:;";

            //string decname = fncl.Decrypt(encr, encr.ToString().Length);
            //string dec = "hdfc@123";
            //string encname = fncl.Encrypt(dec, dec.ToString().Length);
            ReadWriteAppSettings readWriteAppSettings = new ReadWriteAppSettings();
            sttg = readWriteAppSettings.ReadGetSectionAppSettings();
        }

        static public void BindAppSettings()
        {
            ReadWriteAppSettings readWriteAppSettings = new ReadWriteAppSettings();
            readWriteAppSettings.BindAppSettings();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                timer1.Enabled = true;
                timer1.Interval = 10000;

                ReadGetSectionAppSettings();

                FolderCreate();


                Main();
                //AddOrUpdateAppSetting("MSILSettings:InputFilePath", "MSILSettings:OutputFilePath", "MSILSettings:BackupFilePath", "MSILSettings:NonConvertedFile", sttg.InputFilePath, sttg.OutputFilePath, sttg.BackupFilePath, sttg.NonConvertedFile);

            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message + "Form1", "Form1_Load", "");
            }

        }
        public void WritefAILiNVOICES(string invoices, string ModuleName, string ProcName)
        {
            string strfilename = string.Empty;
            string strErrorString = string.Empty;
            try
            {
                string strExeFilePath = Application.StartupPath + "\\Fail_MSIL";
               
                if (!Directory.Exists(strExeFilePath))
                {
                    Directory.CreateDirectory(strExeFilePath);
                }
                strErrorString = invoices;
                strfilename = strExeFilePath +"\\"+ "FailInvoices"+System.DateTime.Now.ToString("dd/MM/yyyy") + ".log";

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

        private void Main()
        {

        StartEXE:
            try
            {
                DeleteFilesAndFolders();// Delete all files from folders if exists
                DateTime dtStartDate = DateTime.ParseExact(DateTime.Now.ToString("dd/MM/yyyy").ToString().Replace("-", "/").Replace(".", "/"), "dd/MM/yyyy", CultureInfo.InvariantCulture); //Current Date

                // dttime = dtStartDate.ToString().Replace("-", "/"); ////Added by yogesh
                string dttime = dtStartDate.ToString("dd/MM/yyyy").Replace("-", "/"); ////Added by yogesh
                //DateTime dtStartDate;
                //dtStartDate = DateTime.TryParseExact(DateTime.Now.ToString("dd/MM/yyyy"), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dtStartDate);

                //if (service.InvoiceFileCheck("I", Convert.ToString(dtStartDate)))
                if (service.InvoiceFileCheck("I", Convert.ToString(dttime)))
                {
                }
                else
                {
                    ////Commented As per mail confirmation on 23/04/2024
                    //if (service.File_MailF5("F5", Convert.ToString(dttime))) ////Parameter change by yogesh 
                    //{
                    //}
                    //else
                    //{
                    //    fncl.SendEmail(sttg.NO_INV, "", "", "", "No Invoice Received for the day", "No Invoice Received for the day.");
                    //    clserr.LogEntry("No Invoice Received for the day" + "   MAIN", false);
                    //    //DateTime dtStartDate1 = DateTime.ParseExact(DateTime.Now.ToString("dd/MM/yyyy HH:mm").ToString().Replace("-", "/").Replace(".", "/"), "dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture); //Current Date
                    //    string File_Mail_Time = DateTime.Now.ToString("HH:mm"); //Current Date
                    //    DataTable dt = service.UpdateDetails("Update_FilemailDetails", "F5", "", "", "", "", "", "");
                    //    //if (service.File_MailUpdate(Convert.ToString(File_Mail_Time), Convert.ToString(DateTime.Now.ToString("dd/MM/yyyy")), "F5"))
                    //    //{ }
                    //}
                    ////Commented As per mail confirmation on 23/04/2024
                }

                fncl.RectifyUpdateError();

                DirectoryInfo di = new DirectoryInfo(sttg.InputFilePath);
                //clserr.LogEntry(sttg.InputFilePath, false);
                FileInfo[] TXTFiles = di.GetFiles("*.txt");
                //clserr.LogEntry("Test6", false);
                if (TXTFiles.Length == 0)
                {
                    if (Directory.Exists(sttg.BlankExcel) == false)
                    {
                        Directory.CreateDirectory(sttg.BlankExcel);
                    }
                    //clserr.LogEntry("Test7", false);
                    fncl.MISReportMails(sttg.BlankExcel);
                    goto StartEXE;
                }
                for (int f = 0; f < TXTFiles.Length; f++)
                {
                    //clserr.LogEntry("Test10", false);
                    DataTable dt = service.getDetails("Get_DupplicateFile", TXTFiles[f].Name.ToUpper(), "", "", "", "", "", "");
                    //clserr.LogEntry("Test11", false);
                    if (dt.Rows.Count > 0)
                    {
                        fncl.SendEmail(sttg.INV_CONF_EMAIL, "", "", TXTFiles[f].FullName, "Duplicate File Name", "Duplicate File Name : " + TXTFiles[f].Name);
                        clserr.LogEntry("Mail sent of Duplicate File Name : " + TXTFiles[f].Name + "   MAIN", false);
                        //File.Copy(TXTFiles[f].FullName, sttg.NonConvertedFile + "\\" + TXTFiles[f].Name,true);
                        //File.Delete(TXTFiles[f].FullName);

                        File.Move(TXTFiles[f].FullName, sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                        //clserr.LogEntry("Test12", false);
                    }
                    else
                    {
                        //clserr.LogEntry("Test13", false);
                        Boolean IsCompleteFile = isCompleteFileAvailable(TXTFiles[f].FullName);
                        if (IsCompleteFile == false)
                        {
                            clserr.LogEntry("Files Is not Completely available for Reading File : " + TXTFiles[f].Name + "   MAIN", false);
                            goto FileRecheck;
                        }
                        if (TXTFiles[f].Name.ToUpper().Contains("DOINVCAN") || TXTFiles[f].Name.ToUpper().Contains("INVCAN") ||
                            TXTFiles[f].Name.ToUpper().Contains("DOCAN") || TXTFiles[f].Name.ToUpper().Contains("BNGR") || TXTFiles[f].Name.ToUpper().Contains("ITK") ||
                            TXTFiles[f].Name.ToUpper().Contains("SLGR") || TXTFiles[f].Name.ToUpper().Contains("KHAR"))
                        {
                            if (TXTFiles[f].Name.ToUpper().Contains("INVDATA"))
                            {
                                if (Converter(TXTFiles[f].Name, TXTFiles[f].FullName.Trim(), "Invoice") == "ERROR")
                                {
                                    //File.Copy("Input\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name,true);
                                    //File.Delete("Input\\" + TXTFiles[f].Name.Trim());

                                    //File.Move("Input\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                                    File.Move(sttg.InputFilePath + "\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                                }
                                else
                                {
                                    //File.Copy("Input\\" + TXTFiles[f].Name.Trim(), sttg.BackupFilePath + "\\" + TXTFiles[f].Name, true);
                                    //File.Delete("Input\\" + TXTFiles[f].Name.Trim());

                                    //File.Move("Input\\" + TXTFiles[f].Name.Trim(), sttg.BackupFilePath + "\\" + TXTFiles[f].Name, true);
                                    File.Move(sttg.InputFilePath + "\\" + TXTFiles[f].Name.Trim(), sttg.BackupFilePath + "\\" + TXTFiles[f].Name, true);
                                }

                            }
                            else if (TXTFiles[f].Name.ToUpper().Contains("ORDDATA"))
                            {
                                if (Converter(TXTFiles[f].Name, TXTFiles[f].FullName.Trim(), "Order") == "ERROR")
                                {
                                    //File.Copy("Input\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                                    //File.Delete("Input\\" + TXTFiles[f].Name.Trim());

                                    //File.Move("Input\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                                    File.Move(sttg.InputFilePath + "\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                                }
                                else
                                {
                                    string backfilepath = sttg.BackupFilePath;
                                    //File.Copy("Input\\" + TXTFiles[f].Name.Trim(), sttg.BackupFilePath + "\\" + TXTFiles[f].Name, true);
                                    //File.Delete("Input\\" + TXTFiles[f].Name.Trim());

                                    //File.Move("Input\\" + TXTFiles[f].Name.Trim(), sttg.BackupFilePath + "\\" + TXTFiles[f].Name, true);
                                    File.Move(sttg.InputFilePath + "\\" + TXTFiles[f].Name.Trim(), sttg.BackupFilePath + "\\" + TXTFiles[f].Name, true);
                                }
                            }

                            else if (TXTFiles[f].Name.ToUpper().Contains("ORDDELETE"))
                            {
                                if (Converter(TXTFiles[f].Name, TXTFiles[f].FullName.Trim(), "Order_Delete") == "ERROR")
                                {
                                    //File.Copy("Input\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                                    //File.Delete("Input\\" + TXTFiles[f].Name.Trim());

                                    //File.Move("Input\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                                    File.Move(sttg.InputFilePath + "\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                                }
                                else
                                {
                                    //File.Copy("Input\\" + TXTFiles[f].Name.Trim(), sttg.BackupFilePath + "\\" + TXTFiles[f].Name, true);
                                    //File.Delete("Input\\" + TXTFiles[f].Name.Trim());

                                    //File.Move("Input\\" + TXTFiles[f].Name.Trim(), sttg.BackupFilePath + "\\" + TXTFiles[f].Name, true);
                                    File.Move(sttg.InputFilePath + "\\" + TXTFiles[f].Name.Trim(), sttg.BackupFilePath + "\\" + TXTFiles[f].Name, true);
                                }
                            }
                            else if (TXTFiles[f].Name.ToUpper().Contains("DOINVCAN"))
                            {
                                if (Converter(TXTFiles[f].Name, TXTFiles[f].FullName.Trim(), "DOINVCAN") == "ERROR")
                                {
                                    //File.Copy("Input\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                                    //File.Delete("Input\\" + TXTFiles[f].Name.Trim());

                                    //File.Move("Input\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                                    File.Move(sttg.InputFilePath + "\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                                }
                                else
                                {
                                    //File.Copy("Input\\" + TXTFiles[f].Name.Trim(), sttg.BackupFilePath + "\\" + TXTFiles[f].Name, true);
                                    //File.Delete("Input\\" + TXTFiles[f].Name.Trim());

                                    //File.Move("Input\\" + TXTFiles[f].Name.Trim(), sttg.BackupFilePath + "\\" + TXTFiles[f].Name, true);
                                    File.Move(sttg.InputFilePath + "\\" + TXTFiles[f].Name.Trim(), sttg.BackupFilePath + "\\" + TXTFiles[f].Name, true);
                                }
                            }
                            else if (TXTFiles[f].Name.ToUpper().Contains("INVCAN"))
                            {
                                if (Converter(TXTFiles[f].Name, TXTFiles[f].FullName.Trim(), "INVCAN") == "ERROR")
                                {
                                    //File.Copy("Input\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                                    //File.Delete("Input\\" + TXTFiles[f].Name.Trim());

                                    //File.Move("Input\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                                    File.Move(sttg.InputFilePath + "\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                                }
                                else
                                {
                                    //File.Copy("Input\\" + TXTFiles[f].Name.Trim(), sttg.BackupFilePath + "\\" + TXTFiles[f].Name, true);
                                    //File.Delete("Input\\" + TXTFiles[f].Name.Trim());

                                    //File.Move("Input\\" + TXTFiles[f].Name.Trim(), sttg.BackupFilePath + "\\" + TXTFiles[f].Name, true);
                                    File.Move(sttg.InputFilePath + "\\" + TXTFiles[f].Name.Trim(), sttg.BackupFilePath + "\\" + TXTFiles[f].Name, true);
                                }
                            }
                            else if (TXTFiles[f].Name.ToUpper().Contains("DOCAN"))
                            {
                                if (Converter(TXTFiles[f].Name, TXTFiles[f].FullName.Trim(), "DOCAN") == "ERROR")
                                {
                                    //File.Copy("Input\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                                    //File.Delete("Input\\" + TXTFiles[f].Name.Trim());

                                    //File.Move("Input\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                                    File.Move(sttg.InputFilePath + "\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                                }
                                else
                                {
                                    //File.Copy("Input\\" + TXTFiles[f].Name.Trim(), sttg.BackupFilePath + "\\" + TXTFiles[f].Name, true);
                                    //File.Delete("Input\\" + TXTFiles[f].Name.Trim());

                                    //File.Move("Input\\" + TXTFiles[f].Name.Trim(), sttg.BackupFilePath + "\\" + TXTFiles[f].Name, true);
                                    File.Move(sttg.InputFilePath + "\\" + TXTFiles[f].Name.Trim(), sttg.BackupFilePath + "\\" + TXTFiles[f].Name, true);
                                }
                            }
                            else
                            {
                                clserr.LogEntry("Wrong Naming Convention.Invalid File Name. : " + TXTFiles[f].Name + "   MAIN", false);
                                fncl.SendEmail(sttg.INV_CONF_EMAIL, "", "", TXTFiles[f].FullName, "Wrong Naming Convention", "Wrong Naming Convention.Invalid File Name");

                                //File.Copy("Input\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                                //File.Delete("Input\\" + TXTFiles[f].Name.Trim());

                                //File.Move("Input\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                                File.Move(sttg.InputFilePath + "\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                            }
                        }
                        else
                        {
                            clserr.LogEntry("Wrong Naming Convention.Invalid File Name. : " + TXTFiles[f].Name + "   MAIN", false);
                            fncl.SendEmail(sttg.INV_CONF_EMAIL, "", "", TXTFiles[f].FullName, "Wrong Naming Convention", "Wrong Naming Convention.Invalid File Name");

                            //File.Copy(TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                            //File.Delete(TXTFiles[f].Name.Trim());

                            File.Move(sttg.InputFilePath + "\\" + TXTFiles[f].Name.Trim(), sttg.NonConvertedFile + "\\" + TXTFiles[f].Name, true);
                        }

                    }

                FileRecheck: string str = "";
                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message + "Form1", "Main", "");
            }
        }
        private string Converter(string strInputFile, string InputFile, string File_Type_Desc)
        {

            string converte = "";
            try
            {
                clserr.LogEntry("Entered Converter Function" + "   Converter", false);

                if (File_Type_Desc == "Invoice")
                {
                    if (Create_Reg_File(InputFile, objBaseClass.Create_Reg_FileSheet, objBaseClass.Create_Reg_File_DSheet) == "ERROR") throw new Exception();
                }
                else if (File_Type_Desc == "Order")
                {
                     if (Create_Reg_File_Order_New(InputFile) == "ERROR") throw new Exception();

                    //if (Create_Reg_File_Order(InputFile) == "ERROR") throw new Exception();
                }
                else if (File_Type_Desc == "Order_Delete")
                {
                    if (Create_Reg_File_Order_Delete(InputFile) == "ERROR") throw new Exception();
                }
                else if (File_Type_Desc == "DOCAN")
                {
                    if (Create_Reg_File_DO_Delete(InputFile) == "ERROR") throw new Exception();
                }
                else if (File_Type_Desc == "DOINVCAN")
                {
                    if (Create_Reg_File_DOInvoiceCancel(InputFile) == "ERROR") throw new Exception();
                }
                else if (File_Type_Desc == "INVCAN")
                {
                    if (Create_Reg_Invoice_Cancel(InputFile, objBaseClass.Reg_Invoice_Cancel_Sheet) == "ERROR") throw new Exception();
                }

                if (strInputFile == "") strInputFile = "Admin.txt";
            }
            catch (Exception ex)
            {
                converte = "ERROR";
                if (strInputFile == "") strInputFile = "Admin.txt";
                clserr.WriteErrorToTxtFile(ex.Message + "Form1", "Converter", InputFile);
                clserr.LogEntry("Exiting function Converter" + "   Converter", false);

            }
            return converte;
        }
        private string Create_Reg_File(string InputFile, string[] Create_Reg_FileSheet, string[] Create_Reg_File_DSheet)
        {
            string OFileName = "", OutDuplicateFileName;
            try
            {
                if (Directory.Exists(sttg.OutputFilePath) == false)
                {
                    Directory.CreateDirectory(sttg.OutputFilePath);
                }
                OFileName = sttg.OutputFilePath + "\\" + Path.GetFileName(InputFile).ToUpper().Replace("CSV", "") + "_" + System.DateTime.Now.ToString("ddMMyyyyHHmmss") + ".xlsx";
                OutDuplicateFileName = sttg.OutputFilePath + "\\InvalidInvoice_" + Path.GetFileName(InputFile).ToUpper().Replace("CSV", "") + "_" + System.DateTime.Now.ToString("ddMMyyyyHHmmss") + ".xlsx";
                int TotalRows = 0;

                clserr.LogEntry("Entered Create_Reg_File Function" + "   Create_Reg_File", false);
                clserr.LogEntry("Process File name  " + Path.GetFileName(InputFile) + "Create_Reg_File", false);
                /////Commented and added by yogesh on 21062024
                //string OUTFileName = ChagneFromCSVToExcel_INV(InputFile);

                DataTable dtSucc = new DataTable();
                DataTable dtUnSucc = new DataTable();
                // Define your custom headers
                string[] customHeaders = new string[] { "Invoice_Number", "Invoice_Amount", "Currency", "Vehical_ID", "DueDate", "Dealer_Name", "Dealer_Address1", "Dealer_City", "Transporter_Name", "Transport_Number", "Transport_Date", "Dealer_Code", "Transporter_Code", "Dealer_Address2", "Dealer_Address3", "Dealer_Address4", "REASON" }; // Adjust as needed
                // Add custom headers to the DataTable as columns
                foreach (var header in customHeaders)
                {
                    dtSucc.Columns.Add(header);
                }
                // Add custom headers to the DataTable as columns
                foreach (var header in customHeaders)
                {
                    dtUnSucc.Columns.Add(header);
                }

                DataTable dtOrgInput = GetDatatable_Text(InputFile);
                //string bool1 = ChagneFromCSVToExcel_INV(InputFile);
                /////Commented and added by yogesh on 21062024
                //clserr.LogEntry("Create_Reg_File 1", false);

                //XSSFWorkbook workbook;
                //ISheet excelSheet;

                //Commented by yogesh 20062024
                //using (var fs = new FileStream(OUTFileName, FileMode.Open, FileAccess.Read))
                //{
                //    workbook = new XSSFWorkbook(fs);
                //    excelSheet = workbook.GetSheet("Sheet1");
                //    TotalRows = excelSheet.LastRowNum + 1;
                //}
                //Commented by yogesh 20062024

                TotalRows = dtOrgInput.Rows.Count;
                string ErrReason = ""; Boolean FoundFlag = false;

                // creating excel file and columns for excel file
                fncl.CreateExcel(Create_Reg_FileSheet, OFileName);
                fncl.CreateExcel(Create_Reg_File_DSheet, OutDuplicateFileName);
                //clserr.LogEntry("Create_Reg_File 4", false);
                if (TotalRows == 0)
                    clserr.LogEntry("Blank Invoice File. File Name : " + InputFile + "  Create_Reg_File  ", false);
                else
                {
                    int OpRow = 0, CntDuplicateFIle = 0;
                    string Email_To = sttg.TRADE_EMAIL;
                    string Email_CC = "", Email_BCC = "", Create_Reg_File = "";

                    string[] StrDataRow = new string[17];
                    object ArrLineData = new object();
                    int fldcounter = 0;
                    foreach (DataRow dtRow in dtOrgInput.Rows)
                    {
                        ClearArray(StrDataRow);
                        //ArrLineData = dtRow.ItemArray;
                        ErrReason = "";
                        FoundFlag = false;
                        //'*************************************Validation Start********************************************'
                        //1.Invoice Number
                        if (dtRow.ItemArray[0].ToString() == "")
                        {
                            FoundFlag = true;
                            clserr.LogEntry("Invoice Number is Blank." + InputFile + "  Create_Reg_File  ", false);
                            ErrReason = "Invoice Number is Blank.";
                            StrDataRow[16] = "Invoice Number is Blank." + dtRow.ItemArray[0].ToString();
                            goto UpdateRec;
                        }
                        if (dtRow.ItemArray[0].ToString().Length < 7)
                        {
                            FoundFlag = true;
                            clserr.LogEntry("Invoice Number length is less than 7. Invoice No. " + InputFile + "  Create_Reg_File  ", false);
                            ErrReason = "Invoice Number length is less than 7. Invoice No. ";
                            StrDataRow[16] = "Invoice Number length is less than 7. Invoice No. " + dtRow.ItemArray[0].ToString();
                            goto UpdateRec;
                        }
                        if (dtRow.ItemArray[0].ToString().Length > 16)
                        {
                            FoundFlag = true;
                            clserr.LogEntry("Invoice Number length is greater than 16. Invoice No. " + InputFile + "  Create_Reg_File  ", false);
                            ErrReason = "Invoice Number length is greater than 16. Invoice No. ";
                            StrDataRow[16] = "Invoice Number length is greater than 16. Invoice No. " + dtRow.ItemArray[0].ToString();
                            goto UpdateRec;
                        }
                        if (dtRow.ItemArray[0].ToString().Substring(0, 1).ToUpper() == "C" || dtRow.ItemArray[0].ToString().Substring(0, 1).ToUpper() == "D")
                        {
                        }
                        else
                        {
                            FoundFlag = true;
                            clserr.LogEntry("Invoice Number does not start with C or D.  Invoice No. " + InputFile + "  Create_Reg_File  ", false);
                            ErrReason = "Invoice Number does not start with C or D.  Invoice No. ";
                            StrDataRow[16] = "Invoice Number does not start with C or D.  Invoice No. " + dtRow.ItemArray[0].ToString();
                            goto UpdateRec;
                        }
                        DataTable DT_FT = service.getDetails("Get_InvoiceAsper", dtRow.ItemArray[0].ToString(), "", "", "", "", "", "");
                        if (DT_FT.Rows.Count > 0)
                        {
                            FoundFlag = true;
                            clserr.LogEntry("Duplicate Invoice No. " + dtRow.ItemArray[0].ToString() + "  " + InputFile + "  Create_Reg_File  ", false);
                            ErrReason = "Duplicate Invoice No. ";
                            StrDataRow[16] = "Duplicate Invoice No. " + dtRow.ItemArray[0].ToString();
                            goto UpdateRec;

                        }
                        if (dtRow.ItemArray[2].ToString() == "")
                        {
                            FoundFlag = true;
                            clserr.LogEntry("Currency Field Blank. Invoice No. " + InputFile + "  Create_Reg_File  ", false);
                            ErrReason = "Currency Field Blank. Invoice No. ";
                            StrDataRow[16] = "Currency Field Blank. Invoice No. " + dtRow.ItemArray[0].ToString();
                            goto UpdateRec;
                        }
                        if (dtRow.ItemArray[2].ToString().ToUpper() != "INR")
                        {
                            FoundFlag = true;
                            clserr.LogEntry("Currency Field contain other than INR.  Invoice No. " + dtRow.ItemArray[2].ToString() + "  Create_Reg_File  ", false);
                            ErrReason = "Currency Field contain other than INR.  Invoice No.";
                            StrDataRow[16] = "Currency Field contain other than INR.  Invoice No." + dtRow.ItemArray[0].ToString();
                            goto UpdateRec;
                        }
                        if (dtRow.ItemArray[3].ToString().Trim() == "")
                        {
                            FoundFlag = true;
                            clserr.LogEntry("Vehicle Indentiication Number is Null. Invoice No. " + dtRow.ItemArray[3].ToString() + "  Create_Reg_File  ", false);
                            ErrReason = "Vehicle Indentiication Number is Null. Invoice No. ";
                            StrDataRow[16] = "Vehicle Indentiication Number is Null. Invoice No. " + dtRow.ItemArray[0].ToString();
                            goto UpdateRec;
                        }
                        if (dtRow.ItemArray[3].ToString().Trim().Length < 17)
                        {
                            FoundFlag = true;
                            clserr.LogEntry("Vehicle indentificaiton Number length  not correct. Invoice No. " + dtRow.ItemArray[3].ToString().Trim() + "  Create_Reg_File  ", false);
                            ErrReason = "Vehicle indentificaiton Number length  not correct. Invoice No.  ";
                            StrDataRow[16] = "Vehicle indentificaiton Number length  not correct. Invoice No.  " + dtRow.ItemArray[0].ToString();
                            goto UpdateRec;
                        }
                        if (dtRow.ItemArray[11].ToString().Trim() == "")
                        {
                            FoundFlag = true;
                            clserr.LogEntry("Dealer Code Blank . Invoice No. " + dtRow.ItemArray[11].ToString().Trim() + "  Create_Reg_File  ", false);
                            ErrReason = "Dealer Code Blank . Invoice No. ";
                            StrDataRow[16] = "Dealer Code Blank . Invoice No. " + dtRow.ItemArray[0].ToString();
                            goto UpdateRec;
                        }
                        if (dtRow.ItemArray[11].ToString().Trim().Length < 4)
                        {
                            FoundFlag = true;
                            clserr.LogEntry("Dealer Code Length not Correct. Invoice No. " + dtRow.ItemArray[11].ToString().Trim() + "  Create_Reg_File  ", false);
                            ErrReason = "Dealer Code Length not Correct. Invoice No. ";
                            StrDataRow[16] = "Dealer Code Length not Correct. Invoice No. " + dtRow.ItemArray[0].ToString();
                            goto UpdateRec;
                        }
                        if (dtRow.ItemArray[12].ToString() == "")
                        {
                            FoundFlag = true;
                            clserr.LogEntry("Transporter Code Blank . Invoice No. " + dtRow.ItemArray[12].ToString() + "  Create_Reg_File  ", false);
                            ErrReason = "Transporter Code Blank . Invoice No. ";
                            StrDataRow[16] = "Transporter Code Blank . Invoice No. " + dtRow.ItemArray[0].ToString();
                            goto UpdateRec;
                        }
                        if (dtRow.ItemArray[12].ToString().Length < 4)
                        {
                            FoundFlag = true;
                            clserr.LogEntry("Transporter Code Length not Correct. Invoice No. " + dtRow.ItemArray[12].ToString() + "  Create_Reg_File  ", false);
                            ErrReason = "Transporter Code Length not Correct. Invoice No. ";
                            StrDataRow[16] = "Transporter Code Length not Correct. Invoice No. " + dtRow.ItemArray[0].ToString();
                            goto UpdateRec;
                        }
                        if (dtRow.ItemArray[1].ToString() == "")
                        {
                            FoundFlag = true;
                            clserr.LogEntry("Invoice amount is Blank . Invoice No. " + dtRow.ItemArray[1].ToString() + "  Create_Reg_File  ", false);
                            ErrReason = "Invoice amount is Blank . Invoice No. ";
                            StrDataRow[16] = "Invoice amount is Blank . Invoice No. " + dtRow.ItemArray[0].ToString();
                            goto UpdateRec;
                        }
                        Boolean ismatch = Regex.IsMatch(dtRow.ItemArray[1].ToString(), @"^[\+\-]?\d*\.?[Ee]?[\+\-]?\d*$");
                        if (ismatch == false)
                        {
                            FoundFlag = true;
                            clserr.LogEntry("Invoice amount is not numeric. Invoice No. " + dtRow.ItemArray[1].ToString() + "  Create_Reg_File  ", false);
                            ErrReason = "Invoice amount is not numeric. Invoice No. ";
                            StrDataRow[16] = "Invoice amount is not numeric. Invoice No. " + dtRow.ItemArray[0].ToString();
                            goto UpdateRec;
                        }
                        if (dtRow.ItemArray[5].ToString() == "")
                        {
                            FoundFlag = true;
                            clserr.LogEntry("Dealer Name is Blank . Invoice No. " + dtRow.ItemArray[5].ToString() + "  Create_Reg_File  ", false);
                            ErrReason = "Dealer Name is Blank . Invoice No. ";
                            StrDataRow[16] = "Dealer Name is Blank . Invoice No. " + dtRow.ItemArray[0].ToString();
                            goto UpdateRec;
                        }
                        if (dtRow.ItemArray[8].ToString() == "")
                        {
                            FoundFlag = true;
                            clserr.LogEntry("Transporter Name is Blank . Invoice No. " + dtRow.ItemArray[8].ToString() + "  Create_Reg_File  ", false);
                            ErrReason = "Transporter Name is Blank . Invoice No. ";
                            StrDataRow[16] = "Transporter Name is Blank . Invoice No. " + dtRow.ItemArray[0].ToString();
                            goto UpdateRec;
                        }
                        if (dtRow.ItemArray[9].ToString() == "")
                        {
                            FoundFlag = true;
                            clserr.LogEntry("Transport Number is Blank . Invoice No. " + dtRow.ItemArray[9].ToString() + "  Create_Reg_File  ", false);
                            ErrReason = "Transport Number is Blank . Invoice No. ";
                            StrDataRow[16] = "Transport Number is Blank . Invoice No. " + dtRow.ItemArray[0].ToString();

                            goto UpdateRec;
                        }
                    //'*************************************Validation End********************************************'
                    UpdateRec:
                        StrDataRow[0] = dtRow["Invoice_Number"].ToString();
                        StrDataRow[1] = dtRow["Invoice_Amount"].ToString();
                        StrDataRow[2] = dtRow["Currency"].ToString();
                        StrDataRow[3] = dtRow["Vehical_ID"].ToString();
                        StrDataRow[4] = dtRow["DueDate"].ToString();
                        StrDataRow[5] = dtRow["Dealer_Name"].ToString();
                        StrDataRow[6] = dtRow["Dealer_Address1"].ToString();
                        StrDataRow[7] = dtRow["Dealer_City"].ToString();
                        StrDataRow[8] = dtRow["Transporter_Name"].ToString();
                        StrDataRow[9] = dtRow["Transport_Number"].ToString();
                        StrDataRow[10] = dtRow["Transport_Date"].ToString();
                        StrDataRow[11] = dtRow["Dealer_Code"].ToString();
                        StrDataRow[12] = dtRow["Transporter_Code"].ToString();
                        StrDataRow[13] = dtRow["Dealer_Address2"].ToString();
                        StrDataRow[14] = dtRow["Dealer_Address3"].ToString();
                        StrDataRow[15] = dtRow["Dealer_Address4"].ToString();

                        //StrDataRow[fldcounter] = dtRow.ItemArray[fldcounter].ToString();

                        if (FoundFlag == false)
                        {

                            dtSucc.Rows.Add(StrDataRow);

                        }
                        else
                        {

                            dtUnSucc.Rows.Add(StrDataRow);

                        }
                        fldcounter = fldcounter + 1;
                    }

                    //AddColumnWithDefaultValues(dtSucc, "NewColumn1", "");
                    //AddColumnWithDefaultValues(dtSucc, "NewColumn2", System.DateTime.Now.ToString("dd/MM/yyyy"));
                    //AddColumnWithDefaultValues(dtSucc, "NewColumn3", "");
                    for (int jj = 0; jj <= dtSucc.Rows.Count - 1; jj++)
                    {
                        OpRow = OpRow + 1;
                        UpdateInvoiceRecordsNew(dtSucc.Rows[jj], dtUnSucc, jj, OpRow);
                        clserr.LogEntry("Success Invoice Records " + dtSucc.Rows[jj].ItemArray[0].ToString() + OpRow, false);

                    }
                    for (int ij = 0; ij <= dtUnSucc.Rows.Count - 1; ij++)
                    {
                        CntDuplicateFIle = CntDuplicateFIle + 1;

                        clserr.LogEntry("Unsuccess Invoice Records " + dtUnSucc.Rows[ij].ItemArray[0].ToString(), false);
                    }

                    ExportDataTableToExcel(dtSucc, OFileName);
                    ExportDataTableToExcel(dtUnSucc, OutDuplicateFileName);





                    /////Commented by yogesh on 21062024
                    //for (int jj = 0; jj < TotalRows; jj++)
                    //{
                    //    ErrReason = "";

                    //ICell cell = excelSheet.GetRow(jj).GetCell(0);
                    //if ("" + cell.StringCellValue + "" != "")

                    //    //'*************************************Validation Start********************************************'
                    //    //1. INVOICE NUMBER
                    //    if ("" + excelSheet.GetRow(jj).GetCell(0).StringCellValue == "")
                    //    {
                    //        FoundFlag = true;
                    //        clserr.LogEntry("Invoice Number is Blank." + InputFile + "  Create_Reg_File  ", false);
                    //        ErrReason = "Invoice Number is Blank.";
                    //        goto UpdateRec;
                    //    }
                    //if (("" + excelSheet.GetRow(jj).GetCell(0).StringCellValue).Length < 7)
                    //{
                    //    FoundFlag = true;
                    //    clserr.LogEntry("Invoice Number length is less than 7. Invoice No. " + InputFile + "  Create_Reg_File  ", false);
                    //    ErrReason = "Invoice Number length is less than 7. Invoice No. ";
                    //    goto UpdateRec;
                    //}

                    //if (("" + excelSheet.GetRow(jj).GetCell(0).StringCellValue).Length > 16)
                    //{
                    //    FoundFlag = true;
                    //    clserr.LogEntry("Invoice Number length is greater than 16. Invoice No. " + InputFile + "  Create_Reg_File  ", false);
                    //    ErrReason = "Invoice Number length is greater than 16. Invoice No. ";
                    //    goto UpdateRec;
                    //}

                    //if (excelSheet.GetRow(jj).GetCell(0).StringCellValue.Substring(0, 1).ToUpper() == "C" || excelSheet.GetRow(jj).GetCell(0).StringCellValue.Substring(0, 1).ToUpper() == "D")
                    //{
                    //}
                    //else
                    //{
                    //    FoundFlag = true;
                    //    clserr.LogEntry("Invoice Number does not start with C or D.  Invoice No. " + InputFile + "  Create_Reg_File  ", false);
                    //    ErrReason = "Invoice Number does not start with C or D.  Invoice No. ";
                    //    goto UpdateRec;
                    //}
                    //DataTable DT_FT = service.getDetails("Get_InvoiceAsper", excelSheet.GetRow(jj).GetCell(0).StringCellValue, "", "", "", "", "", "");
                    //if (DT_FT.Rows.Count > 0)
                    //{
                    //    FoundFlag = true;
                    //    clserr.LogEntry("Duplicate Invoice No." + InputFile + "  Create_Reg_File  ", false);
                    //    ErrReason = "Duplicate Invoice No. ";
                    //    goto UpdateRec;

                    //}
                    //if (excelSheet.GetRow(jj).GetCell(2).StringCellValue == "")
                    //{
                    //    FoundFlag = true;
                    //    clserr.LogEntry("Currency Field Blank. Invoice No. " + InputFile + "  Create_Reg_File  ", false);
                    //    ErrReason = "Currency Field Blank. Invoice No. ";
                    //    goto UpdateRec;
                    //}
                    //if (excelSheet.GetRow(jj).GetCell(2).StringCellValue.ToUpper() != "INR")
                    //{
                    //    FoundFlag = true;
                    //    clserr.LogEntry("Currency Field contain other than INR.  Invoice No. " + excelSheet.GetRow(jj).GetCell(2).StringCellValue + "  Create_Reg_File  ", false);
                    //    ErrReason = "Currency Field contain other than INR.  Invoice No.";
                    //    goto UpdateRec;
                    //}
                    //if (excelSheet.GetRow(jj).GetCell(3).StringCellValue.Trim() == "")
                    //{
                    //    FoundFlag = true;
                    //    clserr.LogEntry("Vehicle Indentiication Number is Null. Invoice No. " + excelSheet.GetRow(jj).GetCell(2).StringCellValue + "  Create_Reg_File  ", false);
                    //    ErrReason = "Vehicle Indentiication Number is Null. Invoice No. ";
                    //    goto UpdateRec;
                    //}
                    //if (excelSheet.GetRow(jj).GetCell(3).StringCellValue.Length < 17)
                    //{
                    //    FoundFlag = true;
                    //    clserr.LogEntry("Vehicle indentificaiton Number length  not correct. Invoice No. " + excelSheet.GetRow(jj).GetCell(3).StringCellValue + "  Create_Reg_File  ", false);
                    //    ErrReason = "Vehicle indentificaiton Number length  not correct. Invoice No.  ";
                    //    goto UpdateRec;
                    //}
                    //if (excelSheet.GetRow(jj).GetCell(11).StringCellValue == "")
                    //{
                    //    FoundFlag = true;
                    //    clserr.LogEntry("Dealer Code Blank . Invoice No. " + excelSheet.GetRow(jj).GetCell(11).StringCellValue + "  Create_Reg_File  ", false);
                    //    ErrReason = "Dealer Code Blank . Invoice No. ";
                    //    goto UpdateRec;
                    //}
                    //if (excelSheet.GetRow(jj).GetCell(11).StringCellValue.Length < 4)
                    //{
                    //    FoundFlag = true;
                    //    clserr.LogEntry("Dealer Code Length not Correct. Invoice No. " + excelSheet.GetRow(jj).GetCell(11).StringCellValue + "  Create_Reg_File  ", false);
                    //    ErrReason = "Dealer Code Length not Correct. Invoice No. ";
                    //    goto UpdateRec;
                    //}
                    //if (excelSheet.GetRow(jj).GetCell(12).StringCellValue == "")
                    //{
                    //    FoundFlag = true;
                    //    clserr.LogEntry("Transporter Code Blank . Invoice No. " + excelSheet.GetRow(jj).GetCell(11).StringCellValue + "  Create_Reg_File  ", false);
                    //    ErrReason = "Transporter Code Blank . Invoice No. ";
                    //    goto UpdateRec;
                    //}
                    //if (excelSheet.GetRow(jj).GetCell(12).StringCellValue.Length < 4)
                    //{
                    //    FoundFlag = true;
                    //    clserr.LogEntry("Transporter Code Length not Correct. Invoice No. " + excelSheet.GetRow(jj).GetCell(11).StringCellValue + "  Create_Reg_File  ", false);
                    //    ErrReason = "Transporter Code Length not Correct. Invoice No. ";
                    //    goto UpdateRec;
                    //}
                    //if (excelSheet.GetRow(jj).GetCell(1).StringCellValue == "")
                    //{
                    //    FoundFlag = true;
                    //    clserr.LogEntry("Invoice amount is Blank . Invoice No. " + excelSheet.GetRow(jj).GetCell(1).StringCellValue + "  Create_Reg_File  ", false);
                    //    ErrReason = "Invoice amount is Blank . Invoice No. ";
                    //    goto UpdateRec;
                    //}
                    //Boolean ismatch = Regex.IsMatch(excelSheet.GetRow(jj).GetCell(1).StringCellValue, @"^[\+\-]?\d*\.?[Ee]?[\+\-]?\d*$");
                    //if (ismatch == false)
                    //{
                    //    FoundFlag = true;
                    //    clserr.LogEntry("Invoice amount is not numeric. Invoice No. " + excelSheet.GetRow(jj).GetCell(1).StringCellValue + "  Create_Reg_File  ", false);
                    //    ErrReason = "Invoice amount is not numeric. Invoice No. ";
                    //    goto UpdateRec;
                    //}
                    //if (excelSheet.GetRow(jj).GetCell(5).StringCellValue == "")
                    //{
                    //    FoundFlag = true;
                    //    clserr.LogEntry("Dealer Name is Blank . Invoice No. " + excelSheet.GetRow(jj).GetCell(1).StringCellValue + "  Create_Reg_File  ", false);
                    //    ErrReason = "Dealer Name is Blank . Invoice No. ";
                    //    goto UpdateRec;
                    //}
                    //if (excelSheet.GetRow(jj).GetCell(8).StringCellValue == "")
                    //{
                    //    FoundFlag = true;
                    //    clserr.LogEntry("Transporter Name is Blank . Invoice No. " + excelSheet.GetRow(jj).GetCell(1).StringCellValue + "  Create_Reg_File  ", false);
                    //    ErrReason = "Transporter Name is Blank . Invoice No. ";
                    //    goto UpdateRec;
                    //}
                    //if (excelSheet.GetRow(jj).GetCell(9).StringCellValue == "")
                    //{
                    //    FoundFlag = true;
                    //    clserr.LogEntry("Transport Number is Blank . Invoice No. " + excelSheet.GetRow(jj).GetCell(1).StringCellValue + "  Create_Reg_File  ", false);
                    //    ErrReason = "Transport Number is Blank . Invoice No. ";
                    //    goto UpdateRec;
                    //}
                    //UpdateRec:


                    //    if (FoundFlag == false)
                    //    {
                    //        //Update Invoice Value in Database
                    //        OpRow = OpRow + 1;
                    //        clserr.LogEntry("Create_Reg_File 5" + OpRow, false);
                    //        UpdateInvoiceRecords(excelSheet, jj, OpRow);

                    //        using (var fs = new FileStream(OFileName, FileMode.Open, FileAccess.ReadWrite))
                    //        {
                    //            XSSFWorkbook wb = new XSSFWorkbook(fs);
                    //            ISheet excelSheetOut = wb.GetSheet("Sheet1");
                    //            IRow rowHeader = excelSheetOut.CreateRow(0);
                    //            rowHeader.CreateCell(0).SetCellValue(OpRow);
                    //            rowHeader.CreateCell(1).SetCellValue(excelSheet.GetRow(jj).GetCell(0).StringCellValue);
                    //            rowHeader.CreateCell(2).SetCellValue(excelSheet.GetRow(jj).GetCell(1).StringCellValue);
                    //            rowHeader.CreateCell(3).SetCellValue(excelSheet.GetRow(jj).GetCell(2).StringCellValue);
                    //            rowHeader.CreateCell(4).SetCellValue(excelSheet.GetRow(jj).GetCell(3).StringCellValue);
                    //            rowHeader.CreateCell(5).SetCellValue(excelSheet.GetRow(jj).GetCell(4).StringCellValue);
                    //            rowHeader.CreateCell(6).SetCellValue(excelSheet.GetRow(jj).GetCell(5).StringCellValue);
                    //            rowHeader.CreateCell(7).SetCellValue(excelSheet.GetRow(jj).GetCell(6).StringCellValue);
                    //            rowHeader.CreateCell(8).SetCellValue(excelSheet.GetRow(jj).GetCell(7).StringCellValue);
                    //            rowHeader.CreateCell(9).SetCellValue(excelSheet.GetRow(jj).GetCell(8).StringCellValue);
                    //            rowHeader.CreateCell(10).SetCellValue(excelSheet.GetRow(jj).GetCell(9).StringCellValue);
                    //            rowHeader.CreateCell(11).SetCellValue(excelSheet.GetRow(jj).GetCell(10).StringCellValue);
                    //            rowHeader.CreateCell(12).SetCellValue(excelSheet.GetRow(jj).GetCell(11).StringCellValue);
                    //            rowHeader.CreateCell(13).SetCellValue( excelSheet.GetRow(jj).GetCell(12).StringCellValue);
                    //            rowHeader.CreateCell(14).SetCellValue( excelSheet.GetRow(jj).GetCell(13).StringCellValue);
                    //            rowHeader.CreateCell(15).SetCellValue(excelSheet.GetRow(jj).GetCell(14).StringCellValue);
                    //            rowHeader.CreateCell(16).SetCellValue(excelSheet.GetRow(jj).GetCell(15).StringCellValue);
                    //            rowHeader.CreateCell(17).SetCellValue("");
                    //            rowHeader.CreateCell(18).SetCellValue( System.DateTime.Now.ToString("dd/MM/yyyy"));
                    //            rowHeader.CreateCell(19).SetCellValue("");
                    //            var fsNew = new FileStream(OFileName, FileMode.Create, FileAccess.Write);
                    //            //workbook.Write(fsNew, true);
                    //            wb.Write(fsNew, true);
                    //            fsNew.Close();
                    //            fs.Close();
                    //            wb.Dispose();
                    //            fs.Dispose();
                    //            clserr.LogEntry("Create_Reg_File 6" + OpRow, false);
                    //        }
                    //    }
                    //    else
                    //    {
                    //        CntDuplicateFIle = CntDuplicateFIle + 1;
                    //        using (var fs = new FileStream(OutDuplicateFileName, FileMode.Open, FileAccess.ReadWrite))
                    //        {
                    //            XSSFWorkbook wbD = new XSSFWorkbook(fs);
                    //            ISheet excelSheetOut = wbD.GetSheet("Sheet1");
                    //            IRow rowHeader = excelSheetOut.CreateRow(0);
                    //            rowHeader.CreateCell(0).SetCellValue(jj + 1);
                    //            string str = excelSheet.GetRow(jj).GetCell(0).StringCellValue;
                    //            rowHeader.CreateCell(1).SetCellValue(excelSheet.GetRow(jj).GetCell(0).StringCellValue);
                    //            rowHeader.CreateCell(2).SetCellValue(excelSheet.GetRow(jj).GetCell(1).StringCellValue);
                    //            rowHeader.CreateCell(3).SetCellValue(excelSheet.GetRow(jj).GetCell(2).StringCellValue);
                    //            rowHeader.CreateCell(4).SetCellValue(excelSheet.GetRow(jj).GetCell(3).StringCellValue);
                    //            rowHeader.CreateCell(5).SetCellValue(excelSheet.GetRow(jj).GetCell(4).StringCellValue);
                    //            rowHeader.CreateCell(6).SetCellValue(excelSheet.GetRow(jj).GetCell(5).StringCellValue);
                    //            rowHeader.CreateCell(7).SetCellValue(excelSheet.GetRow(jj).GetCell(6).StringCellValue);
                    //            rowHeader.CreateCell(8).SetCellValue(excelSheet.GetRow(jj).GetCell(7).StringCellValue);
                    //            rowHeader.CreateCell(9).SetCellValue(excelSheet.GetRow(jj).GetCell(8).StringCellValue);
                    //            rowHeader.CreateCell(10).SetCellValue(excelSheet.GetRow(jj).GetCell(9).StringCellValue);
                    //            rowHeader.CreateCell(11).SetCellValue(excelSheet.GetRow(jj).GetCell(10).StringCellValue);
                    //            rowHeader.CreateCell(12).SetCellValue(excelSheet.GetRow(jj).GetCell(11).StringCellValue);
                    //            rowHeader.CreateCell(13).SetCellValue(excelSheet.GetRow(jj).GetCell(12).StringCellValue);
                    //            rowHeader.CreateCell(14).SetCellValue(excelSheet.GetRow(jj).GetCell(13).StringCellValue);
                    //            rowHeader.CreateCell(15).SetCellValue(excelSheet.GetRow(jj).GetCell(14).StringCellValue);
                    //            rowHeader.CreateCell(16).SetCellValue(excelSheet.GetRow(jj).GetCell(15).StringCellValue);
                    //            rowHeader.CreateCell(19).SetCellValue(ErrReason);
                    //        }
                    //    }
                    //}
                    /////Commented by yogesh on 21062024
                    DataTable dtFile = service.InsertDetails("Insert_FileDesc", Path.GetFileName(InputFile), "I", OpRow.ToString(), "", "", "", "");
                    if (OpRow > 0)
                    {
                        fncl.SendEmail(Email_To, Email_CC, Email_BCC, OFileName, "MSIL INVOICE DATA", "");
                        clserr.LogEntry("Invoice file uploaded and mail sent to -" + sttg.INV_CONF_EMAIL + "  Create_Reg_File  ", false);
                    }
                    else if (OpRow == 0 && FoundFlag == false) //'For Blank File
                    {
                        fncl.SendEmail(Email_To, Email_CC, "", InputFile, "Blank File", "");
                        clserr.LogEntry("Invoice file uploaded and mail sent to - " + sttg.INV_CONF_EMAIL + "  Create_Reg_File  ", false);
                    }
                    if (CntDuplicateFIle >= 1)
                    {
                        fncl.SendEmail(Email_To, Email_CC, Email_BCC, OutDuplicateFileName, "MSIl INVALID DATA", "");
                        clserr.LogEntry("Invalid Invoice file uploaded and mail sent to - " + sttg.INV_CONF_EMAIL + "  Create_Reg_File  ", false);
                    }

                }

            }
            catch (Exception ex)
            {
                OFileName = "ERROR";
                //File.Move(sttg.InputFilePath + InputFile + ".CSV", sttg.NonConvertedFile);
                clserr.WriteErrorToTxtFile(ex.Message, "Create_Reg_File", InputFile);
                clserr.WriteErrorToTxtFile("Exiting function Create_Reg_File", "Create_Reg_File", InputFile);
            }
            return OFileName;
        }

        private string Create_Reg_Invoice_Cancel(string InputFile, string[] Create_Reg_FileSheet)
        {
            string OFileName = "", OutDuplicateFileName;
            try
            {
                if (Directory.Exists(sttg.OutputFilePath) == false)
                {
                    Directory.CreateDirectory(sttg.OutputFilePath);
                }
                OFileName = sttg.OutputFilePath + "\\" + Path.GetFileName(InputFile).ToUpper().Replace("CSV", "") + "_" + System.DateTime.Now.ToString("ddMMyyyyHHmmss") + ".xlsx";
                OutDuplicateFileName = sttg.OutputFilePath + "\\InvalidInvoice_" + InputFile.ToUpper().Replace("CSV", "") + "_" + System.DateTime.Now.ToString("ddMMyyyyHHmmss") + ".xlsx";
                int TotalRows = 0;
                clserr.LogEntry("Entered Create_Reg_Invoice_Cancel Function" + "   Create_Reg_Invoice_Cancel", false);
                string OUTFileName = ChagneFromCSVToExcel_INV(InputFile);
                XSSFWorkbook workbook;
                ISheet excelSheet;

                var fs1 = new FileStream(OUTFileName, FileMode.Open, FileAccess.Read);
                workbook = new XSSFWorkbook(fs1);
                try
                {
                    using (fs1)
                    {

                        excelSheet = workbook.GetSheet("Sheet1");
                        TotalRows = excelSheet.LastRowNum + 1;
                    }
                    string ErrReason = ""; Boolean FoundFlag = false;
                    Boolean PymtRcdFlag = false;
                    // creating excel file and columns for excel file
                    fncl.CreateExcel(Create_Reg_FileSheet, OFileName);
                    //CreateExcel(Create_Reg_File_DSheet, OutDuplicateFileName);
                    if (TotalRows == 0)
                        clserr.LogEntry("Blank Invoice File. File Name : " + InputFile + "  Create_Reg_Invoice_Cancel  ", false);
                    else
                    {
                        int OpRow = 0;

                        string Email_To = sttg.TRADE_EMAIL;
                        for (int kk = 0; kk < TotalRows; kk++)
                        {
                            DataTable DT_FT = service.getDetails("Get_InvoiceCancelAsper", excelSheet.GetRow(kk).GetCell(1).StringCellValue, "", "", "", "", "", "");
                            if (DT_FT.Rows.Count > 0)
                            {
                                PymtRcdFlag = true;
                                clserr.LogEntry("Payment Received for Invoice No :  " + excelSheet.GetRow(kk).GetCell(1).StringCellValue + " " + InputFile + "  Create_Reg_Invoice_Cancel  ", false);
                                goto Move1;
                            }

                            clserr.LogEntry("Records with payment received status not found.Processing records" + InputFile + "  Create_Reg_Invoice_Cancel  ", false);

                        }
                        for (int jj = 0; jj < TotalRows; jj++)
                        {
                            ErrReason = "";
                            ICell cell = excelSheet.GetRow(jj).GetCell(0);
                            if (cell.StringCellValue + "" != "")

                                //'*************************************Validation Start********************************************'
                                //1. INVOICE NUMBER
                                if (excelSheet.GetRow(jj).GetCell(1).StringCellValue == "")
                                {
                                    FoundFlag = true;
                                    clserr.LogEntry("Invoice Number is Blank." + InputFile + "  Create_Reg_Invoice_Cancel  ", false);
                                    ErrReason = "Invoice Number is Blank.";
                                    goto UpdateRec;
                                }
                            if ((excelSheet.GetRow(jj).GetCell(1).StringCellValue).Length < 7)
                            {
                                FoundFlag = true;
                                clserr.LogEntry("Invoice Number length is less than 7. Invoice No. " + excelSheet.GetRow(jj).GetCell(1).StringCellValue + InputFile + "  Create_Reg_Invoice_Cancel  ", false);
                                ErrReason = "Invoice Number length is less than 7. Invoice No. ";
                                goto UpdateRec;
                            }

                            if ((excelSheet.GetRow(jj).GetCell(1).StringCellValue).Length > 16)
                            {
                                FoundFlag = true;
                                clserr.LogEntry("Invoice Number length is greater than 16. Invoice No. " + excelSheet.GetRow(jj).GetCell(1).StringCellValue + InputFile + "  Create_Reg_Invoice_Cancel  ", false);
                                ErrReason = "Invoice Number length is greater than 16. Invoice No. ";
                                goto UpdateRec;
                            }

                            if (excelSheet.GetRow(jj).GetCell(1).StringCellValue.Substring(0, 1).ToUpper() == "C" || excelSheet.GetRow(jj).GetCell(1).StringCellValue.Substring(0, 1).ToUpper() == "D")
                            {
                            }
                            else
                            {
                                FoundFlag = true;
                                clserr.LogEntry("Invoice Number does not start with C or D.  Invoice No. " + excelSheet.GetRow(jj).GetCell(1).StringCellValue + InputFile + "  Create_Reg_File  ", false);
                                ErrReason = "Invoice Number does not start with C or D.  Invoice No. ";
                                goto UpdateRec;
                            }
                            string stramt = excelSheet.GetRow(jj).GetCell(2).StringCellValue;

                            Boolean ismatch = Regex.IsMatch(stramt, @"^[\+\-]?\d*\.?[Ee]?[\+\-]?\d*$");
                            if (ismatch == false)
                            {
                                FoundFlag = true;
                                clserr.LogEntry("Invoice amount is not numeric. Invoice No. " + excelSheet.GetRow(jj).GetCell(1).StringCellValue + "  Create_Reg_File  ", false);
                                ErrReason = "Invoice amount is not numeric. Invoice No. ";
                                goto UpdateRec;
                            }
                            if (excelSheet.GetRow(jj).GetCell(2).StringCellValue == "")
                            {
                                FoundFlag = true;
                                clserr.LogEntry("Invoice amount is Blank. Invoice No. " + excelSheet.GetRow(jj).GetCell(1).StringCellValue + InputFile + "  Create_Reg_Invoice_Cancel  ", false);
                                ErrReason = "Invoice amount is Blank.";
                                goto UpdateRec;
                            }
                        UpdateRec:
                            if (FoundFlag == false)
                            {
                                //Update Invoice Value in Database
                                OpRow = OpRow + 1;
                                UpdateInvoiceCancelRecords(excelSheet, jj, OpRow, ErrReason);

                                using (var fs = new FileStream(OFileName, FileMode.Open, FileAccess.ReadWrite))
                                {

                                    XSSFWorkbook wb = new XSSFWorkbook(fs);
                                    try
                                    {
                                        ISheet excelSheetOut = wb.GetSheet("Sheet1");
                                        IRow rowHeader = excelSheetOut.CreateRow(OpRow);
                                        rowHeader.CreateCell(0).SetCellValue(OpRow);
                                        rowHeader.CreateCell(1).SetCellValue(excelSheet.GetRow(jj).GetCell(1).StringCellValue);
                                        rowHeader.CreateCell(2).SetCellValue(excelSheet.GetRow(jj).GetCell(2).StringCellValue);
                                        var fsNew = new FileStream(OFileName, FileMode.Create, FileAccess.Write);
                                        workbook.Write(fsNew, true);
                                        fsNew.Close();
                                        fs.Close();
                                        wb.Dispose();
                                        fs.Dispose();
                                    }
                                    catch (Exception ex)
                                    {
                                        clserr.WriteErrorToTxtFile(ex.Message, "Create_Reg_File", InputFile);
                                        fs.Close();
                                        wb.Dispose();
                                        fs.Dispose();
                                    }
                                }
                            }

                        }
                        DataTable dtFile = service.InsertDetails("Insert_FileDesc", Path.GetFileName(InputFile), "InvC", OpRow.ToString(), "", "", "", "");

                    Move1:
                        if (PymtRcdFlag == false)
                        {
                            fncl.SendEmail(sttg.Invoice_Cancel_Email, "", "", OFileName, "MSIL INVOICE CANCELLATION DATA", "");
                            clserr.LogEntry("Invoice cancellation file uploaded and mail sent to -" + sttg.Invoice_Cancel_Email + "  Create_Reg_Invoice_Cancel  ", false);
                        }
                        else
                        {
                            fncl.SendEmail(sttg.Invoice_Cancel_Email, "", "", "", "MSIL INVOICE CANCELLATION DATA", "Payment Received for Invoices in File : " + InputFile);
                            clserr.LogEntry("Details of Payment Received Invoices sent to -" + sttg.Invoice_Cancel_Email + "  Create_Reg_Invoice_Cancel  ", false);
                            OFileName = "ERROR";
                        }
                    }
                }

                catch (Exception ex)
                {
                    //workbook.Close();
                    workbook.Dispose();
                    fs1.Close();
                    fs1.Dispose();
                    OFileName = "ERROR";
                    File.Move(sttg.InputFilePath + InputFile + ".CSV", sttg.BackupFilePath);
                    clserr.WriteErrorToTxtFile(ex.Message, "Create_Reg_File", InputFile);
                    clserr.WriteErrorToTxtFile("Exiting function Create_Reg_File", "Create_Reg_File", InputFile);
                }
            }
            catch (Exception ex)
            {
                OFileName = "ERROR";
                File.Move(sttg.InputFilePath + InputFile + ".CSV", sttg.BackupFilePath);
                clserr.WriteErrorToTxtFile(ex.Message, "Create_Reg_File", InputFile);
                clserr.WriteErrorToTxtFile("Exiting function Create_Reg_File", "Create_Reg_File", InputFile);
            }
            return OFileName;
        }
        private string Create_Reg_File_Order_Delete(string InputFile)
        {

            string OFileName = "", Create_Reg_File_Order_Delete = "";
            try
            {
                if (Directory.Exists(Application.StartupPath + "\\Order_Delete") == false)
                    Directory.CreateDirectory(Application.StartupPath + "\\Order_Delete");

                OFileName = Application.StartupPath + "\\Order_Delete" + "\\MSILIN01.ORDDELFAIL." + System.DateTime.Now.ToString("ddMMyyyyHHmmss") + ".CSV";
                int TotalRows = 0;
                clserr.LogEntry("Entered Create_Reg_Invoice_Cancel Function" + "   Create_Reg_Invoice_Cancel", false);
                ChagneFromCSVToExcel_OrderDelete(InputFile);
                XSSFWorkbook workbook;
                ISheet excelSheet;

                using (var fs = new FileStream(InputFile, FileMode.Open, FileAccess.Read))
                {
                    workbook = new XSSFWorkbook(fs);
                    excelSheet = workbook.GetSheet("Sheet1");
                    TotalRows = excelSheet.LastRowNum + 1;
                }
                DataTable dtFile = service.InsertDetails("Insert_FileDesc", InputFile, "OD", TotalRows.ToString(), "", "", "", "");
                StreamWriter ts_in = new StreamWriter(OFileName);
                try
                {
                    for (int kk = 0; kk < TotalRows; kk++)
                    {
                        string[] detailsCash = new string[25];

                        detailsCash[0] = "0"; //Order_Del_ID
                        detailsCash[1] = excelSheet.GetRow(kk).GetCell(0).StringCellValue; //Order_No
                        detailsCash[2] = System.DateTime.Now.Date.ToString("dd/MMM/yyyy"); //Order_Delete_Date
                        detailsCash[3] = System.DateTime.Now.Date.ToString("HH:MM:ss AMPM"); //Order_Delete_Time
                        detailsCash[4] = "";   //Order_Del_Status
                        detailsCash[5] = dtFile.Rows[0][0].ToString();   //FileID
                        detailsCash[6] = "NULL";   //Email_Flag                   

                        DataTable dtInsert = service.Insert_Order_DeleteDetails("Save", detailsCash);
                        DataTable DT_FT = service.getDetails("Get_Order_Desc_Del", excelSheet.GetRow(kk).GetCell(0).StringCellValue, "", "", "", "", "", "");
                        if (DT_FT.Rows.Count > 0)

                            DT_FT = service.getDetails("Get_Order_Desc", excelSheet.GetRow(kk).GetCell(0).StringCellValue, "", "", "", "", "", "");
                        if (DT_FT.Rows.Count > 0)
                        {
                            if (DT_FT.Rows[0]["Order_Status"].ToString() == "PENDING")
                            {

                                detailsCash[0] = dtInsert.Rows[0][0].ToString(); //Order_ID_Del
                                detailsCash[1] = DT_FT.Rows[0]["Order_ID"].ToString(); //Order_ID
                                detailsCash[2] = System.DateTime.Now.Date.ToString("ddmmyyyHHMMSS"); //Order_Date
                                detailsCash[3] = "NULL"; //Record_Identifier
                                detailsCash[4] = DT_FT.Rows[0]["DO_number"].ToString();   //DO_Number
                                detailsCash[5] = "NULL";   //Do_Date
                                detailsCash[6] = "NULL";   //   Dealer_Code     
                                detailsCash[7] = "0"; //Dealer_Destination_Code
                                detailsCash[8] = "NULL"; //Dealer_Outlet_Code
                                detailsCash[9] = "NULL"; //Financier_Code
                                detailsCash[10] = "NULL"; //Financier_Name
                                detailsCash[11] = "NULL";   //Email_IDs
                                detailsCash[12] = "NULL";   //Order_Amount
                                detailsCash[13] = "NULL";   //   Order_Status     
                                detailsCash[14] = "NULL"; //Order_Data_Received_On
                                detailsCash[15] = "NULL"; //Cash_Ops_ID 
                                DataTable dtInsrt = service.Insert_Order_DeleteDetails("SaveAllDetails", detailsCash);
                                clserr.LogEntry(excelSheet.GetRow(kk).GetCell(0).StringCellValue + " DELETED " + "Order Delete " + InputFile + "   Create_Reg_Invoice_Cancel", false);

                            }
                            else
                            {
                                DataTable dtUpdate = service.UpdateDetails("Update_Order_Delete", "ORDER ALREADY PAID", dtInsert.Rows[0][0].ToString(), "", "", "", "", "");
                                if (Convert.ToInt32(dtUpdate.Rows[0][0]) > 0)
                                {
                                    clserr.LogEntry(excelSheet.GetRow(kk).GetCell(0).StringCellValue + "  FAILED - ORDER ALREADY PAID " + "Order Delete " + InputFile + "   Create_Reg_Invoice_Cancel", false);
                                }
                                else
                                    clserr.LogEntry( " Record not Updated :"+ excelSheet.GetRow(kk).GetCell(0).StringCellValue + "  FAILED - ORDER ALREADY PAID " + InputFile , false);

                                ts_in.WriteLine(excelSheet.GetRow(kk).GetCell(0).StringCellValue + ",");
                            }
                        }
                        else
                        {
                            DataTable dtUpdate = service.UpdateDetails("Update_Order_Delete", "Invalid Order Number", dtInsert.Rows[0][0].ToString(), "", "", "", "", "");
                            if (Convert.ToInt32(dtUpdate.Rows[0][0]) > 0)
                            {
                                clserr.LogEntry(excelSheet.GetRow(kk).GetCell(0).StringCellValue + "  FAILED - Invalid Order Number " + "Order Delete " + InputFile + "   Create_Reg_Invoice_Cancel", false);
                            }
                            else
                                clserr.LogEntry(" Record not Updated :" + excelSheet.GetRow(kk).GetCell(0).StringCellValue + "  FAILED - ORDER ALREADY PAID " + InputFile, false);

                            //clserr.LogEntry(excelSheet.GetRow(kk).GetCell(0).StringCellValue + "  FAILED - Invalid Order Number " + "Order Delete " + InputFile + "   Create_Reg_Invoice_Cancel", false);

                            ts_in.WriteLine(excelSheet.GetRow(kk).GetCell(0).StringCellValue + ",");
                        }

                    }
                    ts_in.Close();
                }
                catch (Exception ex)
                {
                    ts_in.Close();
                    Create_Reg_File_Order_Delete = "ERROR";
                    File.Move(sttg.InputFilePath + InputFile + ".CSV", sttg.BackupFilePath);
                    clserr.WriteErrorToTxtFile(ex.Message, "Create_Reg_File_Order_Delete", InputFile);
                    clserr.WriteErrorToTxtFile("Exiting function Create_Reg_File_Order_Delete", "Create_Reg_File_Order_Delete", InputFile);
                }
                Create_Reg_File_Order_Delete = "";

            }
            catch (Exception ex)
            {
                Create_Reg_File_Order_Delete = "ERROR";
                File.Move(sttg.InputFilePath + InputFile + ".CSV", sttg.BackupFilePath);
                clserr.WriteErrorToTxtFile(ex.Message, "Create_Reg_File_Order_Delete", InputFile);
                clserr.WriteErrorToTxtFile("Exiting function Create_Reg_File_Order_Delete", "Create_Reg_File_Order_Delete", InputFile);
            }
            return Create_Reg_File_Order_Delete;
        }
        private string Create_Reg_File_DO_Delete(string InputFile)
        {

            string OFileName = "", Create_Reg_File_DOInvoiceCancel = "";
            try
            {
                if (Directory.Exists(Application.StartupPath + "\\Order_Delete") == false)
                    Directory.CreateDirectory(Application.StartupPath + "\\Order_Delete");
                Boolean Rej_Flag = false;
                string Rej_Reason = "", ORD_Rej_Reason = "", Mail_DO = "", Order_ID = "", Mail_Body = "", Mail_OrderNO = "", Mail_InvoiceNO = "", Order_Inv_ID = "";
                OFileName = Application.StartupPath + "\\Order_Delete" + "\\MSILIN01.ORDDELFAIL." + System.DateTime.Now.ToString("ddMMyyyyHHmmss") + ".CSV";
                int TotalRows = 0;
                clserr.LogEntry("Entered Create_Reg_File_DO_Delete Function" + "   Create_Reg_File_DO_Delete", false);
                string OUTFileName = ChagneFromCSVToExcel_DO_Delete(InputFile);
                XSSFWorkbook workbook;
                ISheet excelSheet;
                Boolean flgPymt = false;
                using (var fs = new FileStream(OUTFileName, FileMode.Open, FileAccess.Read))
                {
                    workbook = new XSSFWorkbook(fs);
                    excelSheet = workbook.GetSheet("Sheet1");
                    TotalRows = excelSheet.LastRowNum + 1;
                }
                DataTable dtFile = service.InsertDetails("Insert_FileDesc", Path.GetFileName(InputFile), "OD", TotalRows.ToString(), "", "", "", "");

                if (TotalRows == 0)
                {
                    fncl.SendEmail(sttg.DO_Cancel_Email, "", "", InputFile, "DO DATA CANCELLATION", "BLANK DO FILE.");
                    clserr.LogEntry("Blank DO Cancellation File mail sent to - " + "   Create_Reg_File_DO_Delete", false);

                }
                clserr.LogEntry("Checking for records with payment received status" + "   Create_Reg_File_DO_Delete", false);
                StreamWriter ts_in = new StreamWriter(OFileName);
                for (int kk = 0; kk < TotalRows; kk++)
                {
                    if (excelSheet.GetRow(kk).GetCell(0).StringCellValue == "ORD")
                    {
                        if (excelSheet.GetRow(kk).GetCell(1).StringCellValue != "")
                        {
                            DataTable DT_FT = service.getDetails("Get_Order_DescAsper", excelSheet.GetRow(kk).GetCell(1).StringCellValue, "Payment Received", "", "", "", "", "");
                            if (DT_FT.Rows.Count > 0)  ////16-12-2023
                            {
                                flgPymt = true;
                                clserr.LogEntry("Payment Received for Order No : " + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DO_Delete", false);
                                goto Move1;
                            }
                        }
                    }
                }
                clserr.LogEntry("Records with payment received status not found.Processing records" + "   Create_Reg_File_DO_Delete", false);

                for (int kk = 0; kk < TotalRows; kk++)
                {
                    Rej_Flag = false;
                    Rej_Reason = "";
                    ORD_Rej_Reason = "";

                    if (excelSheet.GetRow(kk).GetCell(0).StringCellValue == "ORD")
                    {
                        if (excelSheet.GetRow(kk).GetCell(1).StringCellValue == "")
                        {
                            clserr.LogEntry("DO Number is Blank" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DO_Delete", false);

                            Rej_Flag = true;
                            Rej_Reason = "DO Number is Blank.";
                            ORD_Rej_Reason = "DO Number is Blank.";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(2).StringCellValue == "")
                        {
                            clserr.LogEntry("DO Date is Blank." + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DO_Delete", false);
                            Rej_Flag = true;
                            Rej_Reason = "DO Date is Blank.";
                            ORD_Rej_Reason = "DO Date is Blank.";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(3).StringCellValue == "")
                        {
                            clserr.LogEntry("Dealer Code is Blank" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DO_Delete", false);
                            Rej_Flag = true;
                            Rej_Reason = "Dealer Code is Blank";
                            ORD_Rej_Reason = "Dealer Code is Blank";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(4).StringCellValue == "")
                        {
                            clserr.LogEntry("Dealer Destination Code is Blank" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DO_Delete", false);
                            Rej_Flag = true;
                            Rej_Reason = "Dealer Destination Code is Blank";
                            ORD_Rej_Reason = "Dealer Destination Code is Blank";
                            goto Starting;
                        }
                        Boolean ismatch = Regex.IsMatch(excelSheet.GetRow(kk).GetCell(6).StringCellValue, @"^[\+\-]?\d*\.?[Ee]?[\+\-]?\d*$");
                        if (ismatch == false)
                        {
                            clserr.LogEntry("Order Amount is Not Numeric" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DO_Delete", false);
                            Rej_Flag = true;
                            Rej_Reason = "Order Amount is Not Numeric";
                            ORD_Rej_Reason = "Order Amount is Not Numeric";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(1).StringCellValue.Length != 22)
                        {
                            clserr.LogEntry("Wrong Length. Order Number : " + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DO_Delete", false);
                            Rej_Flag = true;
                            Rej_Reason = "Lenght of Order Number is Invalid.";
                            ORD_Rej_Reason = "Lenght of Order Number is Invalid.";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(1).StringCellValue.Length > 22)
                        {
                            clserr.LogEntry("Length of DO number is greater than 22 " + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DO_Delete", false);
                            Rej_Flag = true;
                            Rej_Reason = "Length of DO number is greater than 22";
                            ORD_Rej_Reason = "Length of DO number is greater than 22";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(1).StringCellValue.Length < 22)
                        {
                            clserr.LogEntry("Length of DO number is less than 22" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DO_Delete", false);
                            Rej_Flag = true;
                            Rej_Reason = "Length of DO number is less than 22";
                            ORD_Rej_Reason = "Length of DO number is less than 22";
                            goto Starting;
                        }
                        if ((excelSheet.GetRow(kk).GetCell(1).StringCellValue.ToUpper().Substring(0, 6) != "HDFCIN") && (excelSheet.GetRow(kk).GetCell(1).StringCellValue.ToUpper().Substring(0, 6) != "MARBNG") &&
                            (excelSheet.GetRow(kk).GetCell(1).StringCellValue.ToUpper().Substring(0, 5) != "MARSL") && (excelSheet.GetRow(kk).GetCell(1).StringCellValue.ToUpper().Substring(0, 6) != "MRKKH3"))
                        {
                            clserr.LogEntry("DO No. start with HDFCIN or MARBNG or MARSL or MRKKH3. Wrong Order Number : " + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DO_Delete", false);
                            Rej_Flag = true;
                            Rej_Reason = "DO No. start with HDFCIN or MARBNG or MARSL or MRKKH3. Wrong Order Number : ";
                            ORD_Rej_Reason = "DO No. start with HDFCIN or MARBNG or MARSL or MRKKH3. Wrong Order Number : ";
                            goto Starting;
                        }
                    }

                Starting:

                    if (excelSheet.GetRow(kk).GetCell(0).StringCellValue == "ORD")
                    {

                        //Mail_DO = Mail_DO + ";" + excelSheet.GetRow(kk).GetCell(1).StringCellValue;
                        Mail_OrderNO = excelSheet.GetRow(kk).GetCell(1).StringCellValue;
                        Order_ID = "";
                        DataTable DT_FT1 = service.getDetails("Get_MaxDO_desc", "", "", "", "", "", "", "");
                        clserr.LogEntry("Updating DO records in DO Desc Table" + "Order Delete " + InputFile + "   Create_Reg_Invoice_Cancel", false);
                        InsertDOData("0", "DO_Desc", excelSheet, Rej_Reason, Rej_Flag, kk);
                        if (Mail_Body == "")
                            Mail_Body = Mail_OrderNO;
                        else
                            Mail_Body = Mail_Body + ";" + Mail_OrderNO;
                    }
                    //else if (excelSheet.GetRow(kk).GetCell(0).StringCellValue == "INV" && Order_ID != "")
                    //{
                    //    Mail_InvoiceNO = excelSheet.GetRow(kk).GetCell(1).StringCellValue;
                    //    DataTable DT = service.getDetails("Get_InvoiceDetailsAsper", excelSheet.GetRow(kk).GetCell(1).StringCellValue, "", "", "", "", "", "");
                    //    if (DT.Rows.Count > 0)
                    //    {
                    //        clserr.LogEntry("Payment Received for Invoice Number : " + InputFile + "   Create_Reg_Invoice_Cancel", false);
                    //        Rej_Reason = "Payment Received";
                    //    }
                    //    Order_Inv_ID = "";
                    //    //Srno =  1;
                    //    clserr.LogEntry("Updating DO records in Invoice Table" + InputFile + "   Create_Reg_Invoice_Cancel", false);
                    //    InsertDOData("0", "invoice_cancel_desc", excelSheet, Rej_Reason, Rej_Flag, kk);
                    //    Mail_Body = Mail_Body + ", " + Mail_InvoiceNO;
                    //}
                }

            Move1:
                if (flgPymt == false)
                {
                    //fncl.SendEmail(sttg.DO_Invoice_Cancel_Email, "", "", "", "DO and Invoice cancellation Data Uploaded for Follwing Order Nos. : ", "ORDER DATA : " + "\n" + Mail_Body);
                    fncl.SendEmail(sttg.DO_Invoice_Cancel_Email, "", "", "", "DO cancellation Data Uploaded for Following Order Nos. : ", "ORDER DATA : " + "\n" + Mail_Body);
                    clserr.LogEntry("Mail sent of Do cancellation to- " + sttg.DO_Invoice_Cancel_Email + "   Create_Reg_Invoice_Cancel", false);
                }
                else
                {
                    //fncl.SendEmail(sttg.DO_Invoice_Cancel_Email, "", "", "", "DO and Invoice Cancellation Data ", "Payment Received records found in file : " + InputFile.Trim());
                    fncl.SendEmail(sttg.DO_Invoice_Cancel_Email, "", "", "", "DO Cancellation Data ", "Payment Received records found in file : " + InputFile.Trim());
                    clserr.LogEntry("Payment Received data Mail sent- " + sttg.DO_Invoice_Cancel_Email + "   Create_Reg_Invoice_Cancel", false);
                    Create_Reg_File_DOInvoiceCancel = "ERROR";
                }
                Mail_OrderNO = "";
            }
            catch (Exception ex)
            {
                Create_Reg_File_DOInvoiceCancel = "ERROR";
                File.Move(sttg.InputFilePath + Path.GetFileName(InputFile) + ".CSV", sttg.BackupFilePath);
                clserr.WriteErrorToTxtFile(ex.Message, "Create_Reg_File_Order_Delete", InputFile);
                clserr.WriteErrorToTxtFile("Exiting function Create_Reg_File_DO_Delete", "Create_Reg_File_DO_Delete", InputFile);
            }
            return Create_Reg_File_DOInvoiceCancel;
        }
        private string Create_Reg_File_DOInvoiceCancel(string InputFile)
        {
            string OFileName = "", Create_Reg_File_DOInvoiceCancel = "";
            try
            {
                if (Directory.Exists(Application.StartupPath + "\\Order_Delete") == false)
                    Directory.CreateDirectory(Application.StartupPath + "\\Order_Delete");
                Boolean Rej_Flag = false, Bool_in = true, flgPymt = false, flgDOFound = false;
                string strDO = "", Rej_Reason = "", ORD_Rej_Reason = "", Mail_DO = "", Order_ID = "", Mail_Body = "", Mail_OrderNO = "", Mail_InvoiceNO = "", Order_Inv_ID = "";
                OFileName = Application.StartupPath + "\\Order_Delete" + "\\MSILIN01.ORDDELFAIL." + System.DateTime.Now.ToString("ddMMyyyyHHmmss") + ".CSV";
                int TotalRows = 0, intInvCnt = 0, intFlInvCnt = 0;
                double Ord_Inv_Amt = 0, Order_Amt = 0;
                clserr.LogEntry("Entered Create_Reg_File_DOInvoiceCancel Function" + "   Create_Reg_File_DOInvoiceCancel", false);
                string OpFileName = ChagneFromCSVToExcel(InputFile);
                XSSFWorkbook workbook;
                ISheet excelSheet;

                using (var fs = new FileStream(OpFileName, FileMode.Open, FileAccess.Read))
                {
                    workbook = new XSSFWorkbook(fs);
                    excelSheet = workbook.GetSheet("Sheet1");
                    TotalRows = excelSheet.LastRowNum + 1;
                }
                DataTable dtFile = service.InsertDetails("Insert_FileDesc", Path.GetFileName(InputFile), "DOInv", TotalRows.ToString(), "", "", "", "");

                if (TotalRows == 0)
                {
                    fncl.SendEmail(sttg.DO_Invoice_Cancel_Email, "", "", InputFile, "REJECTED DO file DATA", "BLANK DO FILE.");
                    clserr.LogEntry("No Records Found in DO and Invoice cancellation File - " + "   Create_Reg_File_DOInvoiceCancel", false);
                }
                clserr.LogEntry("Checking for records with payment received status" + "   Create_Reg_File_DOInvoiceCancel", false);
                StreamWriter ts_in = new StreamWriter(OFileName);
                for (int kk = 0; kk < TotalRows; kk++)
                {
                    if (excelSheet.GetRow(kk).GetCell(0).StringCellValue == "ORD")
                    {
                        if (intInvCnt == intFlInvCnt)
                        {
                            if (excelSheet.GetRow(kk).GetCell(1).StringCellValue.Trim().ToUpper() != "")
                            {
                                strDO = excelSheet.GetRow(kk).GetCell(1).StringCellValue.Trim().ToUpper();
                                intFlInvCnt = 0;
                                intInvCnt = 0;

                                DataTable DT_FT = service.getDetails("Get_Order_DescAsper", excelSheet.GetRow(kk).GetCell(1).StringCellValue, "Payment Received", "", "", "", "", "");
                                if (DT_FT.Rows.Count > 0)
                                {
                                    flgDOFound = true;
                                    goto Move2;
                                }
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                    else if (excelSheet.GetRow(kk).GetCell(0).StringCellValue == "INV")
                    {
                        if (flgDOFound == true)
                        {
                            if (excelSheet.GetRow(kk).GetCell(1).StringCellValue.Trim().ToUpper() != "")
                            {
                                if (intFlInvCnt == 0)
                                {
                                    strDO = excelSheet.GetRow(kk).GetCell(1).StringCellValue.Trim().ToUpper();

                                    DataTable DT_FT = service.getDetails("Get_InvoiceCancelAsper", strDO.Trim(), "Payment Received", "", "", "", "", "");
                                    if (DT_FT.Rows.Count > 0)
                                    {
                                        intFlInvCnt = Convert.ToInt32(dtFile.Rows[0][0].ToString());
                                        intInvCnt = 1;
                                        flgPymt = true;
                                    }
                                }
                                else
                                {
                                    intFlInvCnt = intFlInvCnt + 1;
                                }
                            }
                        }
                    }
                Move2: string rrr = "";
                }

                if (flgPymt == true)
                {
                    clserr.LogEntry("Payment Received for Order No : " + strDO + "   Create_Reg_File_DOInvoiceCancel", false);
                    goto Move1;

                }
                clserr.LogEntry("Records with payment received status not found.Processing records " + "   Create_Reg_File_DOInvoiceCancel", false);

                for (int kk = 0; kk < TotalRows; kk++)
                {
                    Rej_Flag = false;
                    Rej_Reason = "";
                    ORD_Rej_Reason = "";

                    if (excelSheet.GetRow(kk).GetCell(0).StringCellValue == "ORD")
                    {
                        if (excelSheet.GetRow(kk).GetCell(1).StringCellValue == "")
                        {
                            clserr.LogEntry("DO Number is Blank" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);

                            Rej_Flag = true;
                            Rej_Reason = "DO Number is Blank.";
                            ORD_Rej_Reason = "DO Number is Blank.";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(2).StringCellValue == "")
                        {
                            clserr.LogEntry("DO Date is Blank." + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "DO Date is Blank.";
                            ORD_Rej_Reason = "DO Date is Blank.";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(3).StringCellValue == "")
                        {
                            clserr.LogEntry("Dealer Code is Blank" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "Dealer Code is Blank";
                            ORD_Rej_Reason = "Dealer Code is Blank";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(4).StringCellValue == "")
                        {
                            clserr.LogEntry("Dealer Destination Code is Blank" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "Dealer Destination Code is Blank";
                            ORD_Rej_Reason = "Dealer Destination Code is Blank";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(6).StringCellValue == "")
                        {
                            clserr.LogEntry("Order Amount is Blank" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "Order Amount is Blank";
                            ORD_Rej_Reason = "Order Amount is Blank";
                            goto Starting;
                        }
                        Boolean ismatch = Regex.IsMatch(excelSheet.GetRow(kk).GetCell(6).StringCellValue, @"^[\+\-]?\d*\.?[Ee]?[\+\-]?\d*$");
                        if (ismatch == false)
                        {
                            clserr.LogEntry("Order Amount is Not Numeric" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "Order Amount is Not Numeric";
                            ORD_Rej_Reason = "Order Amount is Not Numeric";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk + 1).GetCell(0).StringCellValue != "INV" || excelSheet.GetRow(kk + 1).GetCell(0).StringCellValue.ToUpper() == "")
                        {
                            clserr.LogEntry("Order data Does not Contain Any Invoice Records. For Order Number :" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "Order data Does not Contain Any Invoice Records.";
                            ORD_Rej_Reason = "Order data Does not Contain Any Invoice Records.";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(1).StringCellValue.Length != 22)
                        {
                            clserr.LogEntry("Wrong Length. Order Number : " + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "Lenght of Order Number is Invalid";
                            ORD_Rej_Reason = "Lenght of Order Number is Invalid";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(1).StringCellValue.Length > 22)
                        {
                            clserr.LogEntry("Length of DO number is greater than 22 " + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "Length of DO number is greater than 22";
                            ORD_Rej_Reason = "Length of DO number is greater than 22";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(1).StringCellValue.Length < 22)
                        {
                            clserr.LogEntry("Length of DO number is less than 22" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "Length of DO number is less than 22";
                            ORD_Rej_Reason = "Length of DO number is less than 22";
                            goto Starting;
                        }
                        if ((excelSheet.GetRow(kk).GetCell(1).StringCellValue.ToUpper().Substring(0, 6) != "HDFCIN") && (excelSheet.GetRow(kk).GetCell(1).StringCellValue.ToUpper().Substring(0, 6) != "MARBNG") &&
                            (excelSheet.GetRow(kk).GetCell(1).StringCellValue.ToUpper().Substring(0, 5) != "MARSL") && (excelSheet.GetRow(kk).GetCell(1).StringCellValue.ToUpper().Substring(0, 6) != "MRKKH3"))
                        {
                            clserr.LogEntry("DO No. start with HDFCIN or MARBNG or MARSL or MRKKH3. Wrong Order Number : " + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "DO No. start with HDFCIN or MARBNG or MARSL or MRKKH3. Wrong Order Number : ";
                            ORD_Rej_Reason = "DO No. start with HDFCIN or MARBNG or MARSL or MRKKH3. Wrong Order Number : ";
                            goto Starting;
                        }
                        Order_Amt = Convert.ToDouble(excelSheet.GetRow(kk).GetCell(6).StringCellValue);
                        Ord_Inv_Amt = 0;

                        for (int NRow = kk; NRow < TotalRows; NRow++)
                        {
                            //if ((excelSheet.GetRow(NRow + 1).GetCell(1).StringCellValue.ToUpper() == "INV"))
                            //    Ord_Inv_Amt = Ord_Inv_Amt + Convert.ToDouble(excelSheet.GetRow(NRow + 1).GetCell(1).StringCellValue);
                            //else
                            //    break;


                            if (NRow == 0 && (excelSheet.GetRow(NRow).GetCell(0).StringCellValue.ToUpper() == "ORD"))
                            {
                                if ((excelSheet.GetRow(NRow + 1).GetCell(0).StringCellValue.ToUpper() == "INV"))
                                    Ord_Inv_Amt = Ord_Inv_Amt + Convert.ToDouble(excelSheet.GetRow(NRow + 1).GetCell(2).StringCellValue);
                                else
                                    ;
                            }
                            else if (NRow != TotalRows - 1)
                            {
                                if ((excelSheet.GetRow(NRow + 1).GetCell(0).StringCellValue.ToUpper() == "INV"))
                                    Ord_Inv_Amt = Ord_Inv_Amt + Convert.ToDouble(excelSheet.GetRow(NRow + 1).GetCell(2).StringCellValue);
                                else
                                    break;
                            }
                        }


                        if (Convert.ToDouble(Order_Amt).ToString("0.00") != Convert.ToDouble(Ord_Inv_Amt).ToString("0.00"))
                        {
                            clserr.LogEntry("Wrong Amount For Order No : : " + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);

                            Rej_Flag = true;
                            Rej_Reason = "DO amount does not match with Invoice amount.";
                            goto Starting;
                        }
                        if (Rej_Flag == false)
                        {

                            for (int Row_New = kk + 1; Row_New < TotalRows; Row_New++)
                            {
                                if (Row_New != TotalRows - 1) ////Correction on 17-12-2023
                                {
                                    //if ((excelSheet.GetRow(Row_New+1).GetCell(1).StringCellValue.ToUpper() == "ORD") || (excelSheet.GetRow(Row_New+1).GetCell(1).StringCellValue.ToUpper() == ""))
                                    if ((excelSheet.GetRow(Row_New).GetCell(1).StringCellValue.ToUpper() == "ORD") || (excelSheet.GetRow(Row_New).GetCell(1).StringCellValue.ToUpper() == ""))
                                        break;
                                }
                            }
                        }
                    }
                Starting:
                    if (excelSheet.GetRow(kk).GetCell(0).StringCellValue == "ORD")
                    {
                        Mail_OrderNO = excelSheet.GetRow(kk).GetCell(1).StringCellValue;
                        Order_ID = "";
                        DataTable DT_FT1 = service.getDetails("Get_MaxDO_desc", "", "", "", "", "", "", "");
                        Order_ID = DT_FT1.Rows[0][0].ToString();
                        clserr.LogEntry("Updating DO records in DO Desc Table" + "Order Delete " + InputFile + "   Create_Reg_Invoice_Cancel", false);
                        //InsertDOData("0", "DO_Desc", excelSheet, Rej_Reason, Rej_Flag, kk);
                        InsertDOData(Order_ID, "DO_Desc", excelSheet, Rej_Reason, Rej_Flag, kk); ////Correction on 17-12-2023
                        if (Mail_Body == "")
                            Mail_Body = Mail_OrderNO;
                        else
                            Mail_Body = Mail_Body + ";" + Mail_OrderNO;
                    }
                    else if (excelSheet.GetRow(kk).GetCell(0).StringCellValue == "INV" && Order_ID != "")
                    {
                        Mail_InvoiceNO = excelSheet.GetRow(kk).GetCell(1).StringCellValue;
                        DataTable DT = service.getDetails("Get_InvoiceDetailsAsper", excelSheet.GetRow(kk).GetCell(1).StringCellValue, "", "", "", "", "", "");
                        if (DT.Rows.Count > 0)
                        {
                            clserr.LogEntry("Payment Received for Invoice Number : " + InputFile + "   Create_Reg_Invoice_Cancel", false);
                            Rej_Reason = "Payment Received";
                        }
                        Order_Inv_ID = "";
                        //Srno =  1;

                        clserr.LogEntry("Updating DO records in Invoice Table" + InputFile + "   Create_Reg_Invoice_Cancel", false);
                        //InsertDOData("0", "invoice_cancel_desc", excelSheet, Rej_Reason, Rej_Flag, kk);
                        InsertDOData(Order_ID, "invoice_cancel_desc", excelSheet, Rej_Reason, Rej_Flag, kk); ////Correction on 17-12-2023
                        Mail_Body = Mail_Body + ", " + Mail_InvoiceNO;
                    }
                }
            Move1:
                if (flgPymt == false)
                {
                    fncl.SendEmail(sttg.DO_Invoice_Cancel_Email, "", "", "", "DO and Invoice cancellation Data Uploaded for Follwing Order Nos. : ", "ORDER DATA : " + "\n" + Mail_Body);
                    clserr.LogEntry("DO and Invoice File Uploaded and Mail sent - " + sttg.DO_Invoice_Cancel_Email + "   Create_Reg_Invoice_Cancel", false);
                }
                else
                {
                    fncl.SendEmail(sttg.DO_Invoice_Cancel_Email, "", "", "", "DO and Invoice Cancellation Data ", "Payment Received records found in file : " + InputFile.Trim());
                    clserr.LogEntry("DO and Invoice File for Payment Received data Mail sent- " + sttg.DO_Invoice_Cancel_Email + "   Create_Reg_Invoice_Cancel", false);
                    Create_Reg_File_DOInvoiceCancel = "ERROR";
                }
                Mail_OrderNO = "";
            }
            catch (Exception ex)
            {
                Create_Reg_File_DOInvoiceCancel = "ERROR";
                File.Move(sttg.InputFilePath + InputFile + ".CSV", sttg.BackupFilePath);
                clserr.WriteErrorToTxtFile(ex.Message, "Create_Reg_File_Order_Delete", InputFile);
                clserr.WriteErrorToTxtFile("Exiting function Create_Reg_File_DO_Delete", "Create_Reg_File_DO_Delete", InputFile);
            }
            return Create_Reg_File_DOInvoiceCancel;
        }
        public DataTable GetDatatable_Text_Order(string StrFilePath)
        {
            // Path to the comma-delimited text file
            string filePath = StrFilePath;
            // Create a new DataTable
            DataTable dataTable = new DataTable();
            try
            {
                // Define your custom headers
                string[] customHeaders = new string[] { "RecordType", "DO_Number", "DO_Date", "Dealer Code", "Dealer Destination Code", "Dealer_Outlet_Code", "Financier_Name", "Email_IDs", "Order_amount", "Order_Status", "Order_Data_Received_On", "ORD_Rej_Reason", "EmailStatus" }; // Adjust as needed
                // Add custom headers to the DataTable as columns
                foreach (var header in customHeaders)
                {
                    dataTable.Columns.Add(header);
                }
                // Initialize the CsvReader
                using (var reader = new StreamReader(filePath))

                using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
                {

                    Delimiter = ",",
                    Quote = '"',
                    BadDataFound = null, // Ignore bad data
                    MissingFieldFound = null, // Ignore missing fields
                    HasHeaderRecord = false // Indicate that the CSV file does not have a header row
                }))

                {
                    //// Read the header row
                    //if (csv.Read())
                    //{
                    //    csv.ReadHeader();
                    //    foreach (var header in csv.HeaderRecord)
                    //    {
                    //        dataTable.Columns.Add(header);
                    //    }
                    //}

                    // Read the data rows
                    while (csv.Read())
                    {
                        var row = dataTable.NewRow();
                        for (int i = 0; i < customHeaders.Length; i++)
                        {
                            row[i] = csv.GetField(i);
                        }
                        dataTable.Rows.Add(row);
                    }
                }
            }
            catch (Exception ex)
            {
                clserr.Handle_Error(ex, "Form", "GetDatatable_Text");
            }
            return dataTable.Copy();
        }
        private async Task SQLBulkInsert(DataTable dataTable, string fileType, string fileName, string tableName)
        {
            const int maxRetries = 3;
            const int delay = 2000; // 2 seconds
            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                using SqlConnection connection = new SqlConnection(DataService._connectionString);
                await connection.OpenAsync();
                using var transaction = connection.BeginTransaction();
                try
                {
                    using (var bulkCopy = new SqlBulkCopy(connection, SqlBulkCopyOptions.Default, transaction))
                    {
                        bulkCopy.DestinationTableName = tableName;
                        // Define column mappings based on tableName
                        DefineColumnMappings(bulkCopy, tableName);
                        await bulkCopy.WriteToServerAsync(dataTable);
                    }
                    
                    // Commit the transaction
                    await transaction.CommitAsync();
                    await connection.CloseAsync();
                    clserr.LogEntry("Bulk Insertion in " + tableName + " Table Successfully. bulk Count :" + dataTable.Rows.Count.ToString() + "   SQLBulkInsert", false);
                    return; // Exit method if successful
                }
                catch (Exception ex)
                {
                    // Rollback the transaction on error
                    await transaction.RollbackAsync();                  
                    if (ex is Microsoft.Data.SqlClient.SqlException sqlEx && sqlEx.Number == 1205) // SQL Server error code for deadlock
                    {
                        // Log the error (consider using a logging framework)
                        Console.WriteLine($"Attempt {attempt} - Transaction was deadlocked. Retrying in {delay / 1000} seconds...");
                        //await Task.Delay(delay); // Wait before retrying
                    }
                    else { clserr.WriteErrorToTxtFile(ex.Message, "SQLBulkInsert", fileType + ", " + fileName); }
                }
                await connection.CloseAsync();
            }
        }
        private async void DefineColumnMappings(SqlBulkCopy bulkCopy, string tableName)
        {
            try
            {
                var mappings = new Dictionary<string, string> { };

                if (tableName == "Order_Desc")
                {
                    mappings.Add("Order_ID", "Order_ID");
                    mappings.Add("Order_Date", "Order_Date");
                    mappings.Add("Record_Identifier", "Record_Identifier");
                    mappings.Add("DO_Number", "DO_Number");
                    mappings.Add("Do_Date", "Do_Date");
                    mappings.Add("Dealer_Code", "Dealer_Code");
                    mappings.Add("Dealer_Destination_Code", "Dealer_Destination_Code");
                    mappings.Add("Dealer_Outlet_Code", "Dealer_Outlet_Code");
                    mappings.Add("Financier_Code", "Financier_Code");
                    mappings.Add("Financier_Name", "Financier_Name");
                    mappings.Add("Order_Amount", "Order_Amount");
                    mappings.Add("Order_Status", "Order_Status");
                    mappings.Add("Order_Data_Received_On", "Order_Data_Received_On");
                    mappings.Add("Ord_Rej_Flag", "Ord_Rej_Flag");
                    mappings.Add("Ord_Rej_Reason", "Ord_Rej_Reason");
                    mappings.Add("Email_IDs", "Email_IDs");
                }
                else if (tableName == "order_invoice")
                {
                    mappings.Add("Order_Amount", "Ord_Inv_ID");
                    mappings.Add("Order_ID", "Order_ID");
                    mappings.Add("Record_Identifier", "Record_Identifier");
                    mappings.Add("DO_Number", "Order_Inv_Number");
                    mappings.Add("Do_Date", "Order_Inv_Amount");
                    mappings.Add("Order_Status", "Order_Inv_Status");

                }
                if (tableName == "Order_Rejected")
                {
                    mappings.Add("Order_ID_Rej", "Order_ID_Rej");
                    mappings.Add("Order_ID", "Order_ID");
                    mappings.Add("Order_Date", "Order_Date");
                    mappings.Add("Record_Identifier", "Record_Identifier");
                    mappings.Add("DO_Number", "DO_Number");
                    mappings.Add("Do_Date", "Do_Date");
                    mappings.Add("Dealer_Code", "Dealer_Code");
                    mappings.Add("Dealer_Destination_Code", "Dealer_Destination_Code");
                    mappings.Add("Dealer_Outlet_Code", "Dealer_Outlet_Code");
                    mappings.Add("Financier_Code", "Financier_Code");
                    mappings.Add("Financier_Name", "Financier_Name");
                    mappings.Add("Email_IDs", "Email_IDs");
                    mappings.Add("Order_Amount", "Order_Amount");
                    mappings.Add("Order_Status", "Order_Status");
                    mappings.Add("Order_Data_Received_On", "Order_Data_Received_On");
                    mappings.Add("Ord_Rej_Reason", "Rejected_Reson");
                }
                if (tableName == "Order_Invoice_Rejected")
                {
                    mappings.Add("Order_INV_ID_Rej", "Ord_Inv_ID_Rej");
                    mappings.Add("Order_INV_ID", "Ord_Inv_ID");
                    mappings.Add("Order_ID", "Order_ID");
                    mappings.Add("Record_Identifier", "Record_Identifier");
                    mappings.Add("DO_Number", "Order_Inv_Number");
                    mappings.Add("Do_Date", "Order_Inv_Amount");
                    mappings.Add("Order_Status", "Order_Inv_Status");
                    mappings.Add("Order_ID_Rej", "Order_ID_Rej");
                }
                foreach (var mapping in mappings)
                {
                    bulkCopy.ColumnMappings.Add(mapping.Key, mapping.Value);
                }
            }
            catch (Exception ex) { clserr.WriteErrorToTxtFile(ex.Message, "DefineColumnMappings", tableName); }
        }
        public void LogAndReject(string message, int row, ref bool isRejected, ref string rejectReason, ref string orderRejectReason)
        {
            clserr.LogEntry($"{message} - Create_Reg_File_OrderNew", false);
            isRejected = true;
            if (rejectReason == "")
                rejectReason = message;
        }

        private string Create_Reg_File_Order_New(string InputFile)
        {
            string Create_Reg_File_Order = "";
            try
            {
                DataTable dt = ReadCsvFileToDataTable(InputFile);
                string OPFileName = ChagneFromCSVToExcel(InputFile);   // use existing function because client has to sent file in csv format
                DataTable dtFile = service.InsertDetails("Insert_FileDesc", Path.GetFileName(InputFile).ToUpper(), "O", dt.Rows.Count.ToString(), "", "", "", "");
                if (dt.Rows.Count == 0)
                {
                    fncl.SendEmail(sttg.ORD_MIS_EMAIL, "", "", InputFile, "REJECTED ORDER DATA", "BLANK DO FILE.");
                    clserr.LogEntry("No Records Found in Order File  - " + "   Create_Reg_File_OrderNew", false);
                }
                if (dt.Rows.Count > 0)
                {
                    //Processing for Success Records
                    DataRow[] dataRows = dt.Select("Record_Identifier = 'ORD' AND (Ord_Rej_Flag = '0' OR Ord_Rej_Flag = '')");
                    if (dataRows.Length > 0)
                    {
                        DataTable dt_Order = dataRows.CopyToDataTable();
                        Task.Run(() => SQLBulkInsert(dt_Order, " Order ", InputFile, "Order_Desc")).GetAwaiter().GetResult();
                        DataRow[] dataRows_INV = dt.Select("Record_Identifier = 'INV' AND (Ord_Rej_Flag = '0' OR Ord_Rej_Flag = '')");
                        DataTable dt_Invoice = dataRows_INV.CopyToDataTable();
                        Task.Run(() => SQLBulkInsert(dt_Invoice, " Invoice ", InputFile, "order_invoice")).GetAwaiter().GetResult();
                        DataTable dt_bulkupdate = service.UpdateDetails("Update_Bulk_OrderToInvoice", "", "", "", "", "", "", "");
                        if (Convert.ToInt32(dt_bulkupdate.Rows[0][0]) > 0)
                        {
                            clserr.LogEntry("Bulk Update in Invoice Table Successfully. Update Count :" + dt_bulkupdate.Rows[0][0] + "   Create_Reg_File_Order_New", false);
                        }
                        else
                        { clserr.LogEntry("No Records found to Bulk Update." + "   Create_Reg_File_Order_New", false); }

                        //fncl.SendEmail(sttg.ORD_MIS_EMAIL, "", "", "", "Order Data Successful Uploaded for Follwing Order Nos. : ", "Successful ORDER DATA : " + Mail_OrderNO);
                        fncl.SendEmail(sttg.ORD_MIS_EMAIL, "", "", OPFileName, "MSIL ORDER DATA", ""); ////Added on 14-12-2023 which is discuss on call
                        clserr.LogEntry(" " + "   Create_Reg_File_Order", false);
                    }
                    //Processing for Rejected Records
                    DataRow[] dataRowsReject = dt.Select("Record_Identifier = 'ORD' AND (Ord_Rej_Flag = '1' )");
                    if (dataRowsReject.Length > 0)
                    {
                        DataTable dt_Order_Reject = dataRowsReject.CopyToDataTable();
                        Task.Run(() => SQLBulkInsert(dt_Order_Reject, " Order ", InputFile, "Order_Rejected")).GetAwaiter().GetResult();
                        DataRow[] dataRows_INV_rej = dt.Select("Record_Identifier = 'INV' AND (Ord_Rej_Flag = '1')");
                        DataTable dt_Invoice_Rej = dataRows_INV_rej.CopyToDataTable();
                        Task.Run(() => SQLBulkInsert(dt_Invoice_Rej, " Invoice ", InputFile, "Order_Invoice_Rejected")).GetAwaiter().GetResult();
                        DataTable dt_bulkupdate_Rej = service.UpdateDetails("Update_Bulk_OrderToInvoice", "", "", "", "", "", "", "");
                        if (Convert.ToInt32(dt_bulkupdate_Rej.Rows[0][0]) > 0)
                        {
                            clserr.LogEntry("Bulk Update in Invoice Table Successfully. Update Count :" + dt_bulkupdate_Rej.Rows[0][0] + "   Create_Reg_File_Order_New", false);
                        }
                        else
                        { clserr.LogEntry("No Records found to Bulk Update." + "   Create_Reg_File_Order_New", false); }
                    }
                }
                SendOrderRejectedEmails_WOAtt();
                SendOrderRejectedEmails(InputFile);

                Create_Reg_File_Order = "";
            }
            catch (Exception ex)
            {
                Create_Reg_File_Order = "ERROR";
                File.Move(InputFile + ".CSV", sttg.BackupFilePath);
                clserr.WriteErrorToTxtFile(ex.Message, "Create_Reg_File_Order_Delete", InputFile);
                clserr.WriteErrorToTxtFile("Exiting function Create_Reg_File_DO_Delete", "Create_Reg_File_DO_Delete", InputFile);
            }
            return Create_Reg_File_Order;
        }
        public DataTable ReadCsvFileToDataTable(string filePath)
        {
            DataTable dataTable = new DataTable();
            try
            {
                using (TextFieldParser parser = new TextFieldParser(filePath))
                {
                    string[] customHeaders = new string[] { "Record_Identifier", "DO_Number", "Do_Date", "Dealer_Code", "Dealer_Destination_Code", "Dealer_Outlet_Code", "Financier_Code", "Financier_Name",
                                                       "Email_IDs", "Order_Amount", "Order_Status","Order_ID","Order_Date","Ord_Rej_Flag","Ord_Rej_Reason","Order_Data_Received_On",
                     "Order_INV_ID","Order_ID_Rej","Order_INV_ID_Rej"}; // Adjust as needed
                                                                        // Add custom headers to the DataTable as columns
                    foreach (var header in customHeaders)
                    {
                        dataTable.Columns.Add(header);
                    }
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");
                    DataTable dt = service.getDetails("GetMaxOrderID", "", "", "", "", "", "", "");
                    DataTable dtInv = service.getDetails("Get_MaxOrdInvoice", "", "", "", "", "", "", "");
                    DataTable dt_RejectID = service.getDetails("GetMaxOrderID_Rej", "", "", "", "", "", "", "");
                    DataTable dtInv_RejectID = service.getDetails("GetMaxINVID_Rej", "", "", "", "", "", "", "");
                    int Max_OrderID = Convert.ToInt32(dt.Rows[0][0].ToString());
                    int Max_ORDinvoiceID = Convert.ToInt32(dtInv.Rows[0][0].ToString());
                    int Max_OrderIDRej = Convert.ToInt32(dt_RejectID.Rows[0][0].ToString());
                    int Max_invoiceIDRej = Convert.ToInt32(dtInv_RejectID.Rows[0][0].ToString());
                    // Assuming the first row contains the column names
                    //bool isFirstRow = true;
                    string ORDRNO = ""; double INV_Sum = 0; int ord_recordNo = -1, dt_rowNO = 0;
                    Boolean isRejected = false;
                    string rejectReason = "", orderRejectReason = "";
                    while (!parser.EndOfData)
                    {
                        string[] fields = new string[11];
                        fields = parser.ReadFields();
                        Array.Resize(ref fields, dataTable.Columns.Count);
                        dt_rowNO = dt_rowNO + 1;
                        if (fields[0] == "ORD")
                        {
                            // Add data rows
                            if (ORDRNO != fields[1] && ORDRNO != "")
                            {
                                double order_amount = Convert.ToDouble(dataTable.Rows[ord_recordNo]["Order_Amount"].ToString());
                                if (INV_Sum.ToString("0.00").ToString() == "0.00")
                                {
                                    LogAndReject(" Order data Does not Contain Any Invoice Records. ", dt_rowNO, ref isRejected, ref rejectReason, ref orderRejectReason);
                                }

                                if (order_amount.ToString("0.00") != INV_Sum.ToString("0.00"))
                                {
                                    LogAndReject(" DO amount does not match with Invoice amount. ", dt_rowNO, ref isRejected, ref rejectReason, ref orderRejectReason);
                                }
                                if (isRejected == true)
                                {
                                    DataRow[] dataRows = dataTable.Select($"DO_Number = '{ORDRNO}' or Dealer_Destination_Code ='{ORDRNO}'");
                                    if (dataRows.Length > 0)
                                    {
                                        for (int i = 0; i < dataRows.Length; i++)
                                        {
                                            dataTable.Rows[dataTable.Rows.IndexOf(dataRows[i])]["Ord_Rej_Flag"] = "1";
                                            if (dataTable.Rows[dataTable.Rows.IndexOf(dataRows[i])]["Ord_Rej_Reason"].ToString() == "")
                                                dataTable.Rows[dataTable.Rows.IndexOf(dataRows[i])]["Ord_Rej_Reason"] = rejectReason;
                                        }
                                        isRejected = false;
                                        rejectReason = "";
                                    }
                                }
                                //if (rejectReason.Length > 0)
                                //  dataTable.Rows[ord_recordNo]["Order_Status"] = rejectReason.Substring(1, rejectReason.Length - 1);
                            }
                            // Rejected Reason
                            INV_Sum = 0;
                            rejectReason = "";
                            isRejected = false;
                            if (string.IsNullOrEmpty(fields[1]))
                            {
                                LogAndReject("DO Number is Blank", 1, ref isRejected, ref rejectReason, ref orderRejectReason);
                            }
                            if (string.IsNullOrEmpty(fields[2]))
                            {
                                LogAndReject("DO Date is Blank", dt_rowNO, ref isRejected, ref rejectReason, ref orderRejectReason);
                            }
                            if (string.IsNullOrEmpty(fields[3]))
                            {
                                LogAndReject("Dealer Code is Blank", dt_rowNO, ref isRejected, ref rejectReason, ref orderRejectReason);
                            }
                            if (string.IsNullOrEmpty(fields[4]))
                            {
                                LogAndReject("Dealer Destination Code is Blank", dt_rowNO, ref isRejected, ref rejectReason, ref orderRejectReason);
                            }
                            if (string.IsNullOrEmpty(fields[9]))
                            {
                                LogAndReject("Order Amount is Blank", dt_rowNO, ref isRejected, ref rejectReason, ref orderRejectReason);
                            }
                            if (!Regex.IsMatch(fields[9], @"^[\+\-]?\d*\.?[Ee]?[\+\-]?\d*$"))
                            {
                                LogAndReject("Order Amount is Not Numeric", dt_rowNO, ref isRejected, ref rejectReason, ref orderRejectReason);
                            }
                            if (fields[1].Length != 22)
                            {
                                LogAndReject("Length of Order Number is invalid", dt_rowNO, ref isRejected, ref rejectReason, ref orderRejectReason);
                            }
                            if (fields[1].Length > 22)
                            {
                                LogAndReject("Length of DO number is greater than 22", dt_rowNO, ref isRejected, ref rejectReason, ref orderRejectReason);
                            }
                            if (fields[1].Length < 22)
                            {
                                LogAndReject("Length of DO number is less than 22", dt_rowNO, ref isRejected, ref rejectReason, ref orderRejectReason);
                            }
                            if (!new[] { "HDFCIN", "MARBNG", "MARSL", "MRKKH3" }.Any(prefix => fields[1].ToUpper().StartsWith(prefix)))
                            {
                                LogAndReject("Invalid Order Number prefix", dt_rowNO, ref isRejected, ref rejectReason, ref orderRejectReason);
                            }

                            // Check for duplicate order number
                            DataTable DT_FT = service.getDetails("Get_Order_Desc", fields[1], "", "", "", "", "", "");
                            if (DT_FT.Rows.Count > 0)
                            {
                                LogAndReject("Duplicate Order Number", dt_rowNO, ref isRejected, ref rejectReason, ref orderRejectReason);
                            }
                            // Check for missing email ID
                            if (string.IsNullOrEmpty(fields[8]))
                            {
                                LogAndReject("No Email ID for Order Number", dt_rowNO, ref isRejected, ref rejectReason, ref orderRejectReason);
                            }
                            if (isRejected == true)
                            {
                                fields[13] = "1";
                            }
                            else
                                fields[13] = "0";
                            fields[14] = rejectReason.ToString();

                            ORDRNO = fields[1];   // DO no
                            Max_OrderID = Max_OrderID + 1;   // auto order id
                            Max_OrderIDRej = Max_OrderIDRej + 1;   // auto order id
                            fields[11] = (Max_OrderID).ToString();
                            fields[17] = Max_OrderIDRej.ToString();

                            fields[10] = "PENDING";
                            fields[12] = System.DateTime.Now.ToString("dd/MMM/yyyyy");   //
                            fields[15] = System.DateTime.Now.ToString("dd/MM/yyyy HH:MM:ss tt");
                            dataTable.Rows.Add(fields);
                            ord_recordNo = dataTable.Rows.Count - 1;
                        }
                        else
                        {
                            if (fields[0].ToUpper() == "INV")
                            {
                                DataTable DT_k = service.getDetails("Get_Order_InvoiceAsPer", fields[1], "", "", "", "", "", "");
                                if (DT_k.Rows.Count > 0)
                                {
                                    LogAndReject("Duplicate Invoice Number :" + fields[1], dt_rowNO, ref isRejected, ref rejectReason, ref orderRejectReason);
                                    rejectReason = "Duplicate Invoice Number";
                                    goto checking;
                                }
                                DataTable DT_T = service.getDetails("Get_Order_InvoiceAsPerInvoice", fields[1], "", "", "", "", "", "");
                                if (DT_T.Rows.Count == 0)
                                {
                                    LogAndReject("Invoice Number Not Found. :" + fields[1], dt_rowNO, ref isRejected, ref rejectReason, ref orderRejectReason);
                                    rejectReason = "Invoice Bill Number not generated";
                                    goto checking;
                                }
                                DataTable DT_ = service.getDetails("Get_Order_InvoiceAsPerAmount", fields[1], Convert.ToDouble(fields[2]).ToString(), "", "", "", "", "");
                                if (DT_.Rows.Count == 0)
                                {
                                    LogAndReject("Invoice Amount not match. Invoice Number " + fields[1], dt_rowNO, ref isRejected, ref rejectReason, ref orderRejectReason);
                                    rejectReason = "Invoice Amount not match.";
                                    goto checking;
                                }
                            }
                            else if (fields[1].ToUpper() == "INV" || fields[0].Trim() == "")
                                break;
                            checking:
                            INV_Sum = INV_Sum + Convert.ToDouble(fields[2]);
                            fields[4] = ORDRNO;
                            fields[11] = Max_OrderID.ToString();
                            fields[17] = Max_OrderIDRej.ToString();
                            //Max_OrderIDRej = Max_OrderIDRej + 1;   // auto order id
                            fields[12] = System.DateTime.Now.ToString("dd/MMM/yyyyy");
                            if (isRejected == true)
                            {
                                fields[13] = "1";
                                fields[10] = "Not Found";  // set to insert into order_invoice table
                                fields[18] = Max_invoiceIDRej.ToString();
                                Max_invoiceIDRej = Max_invoiceIDRej + 1;
                            }
                            else
                            {
                                fields[13] = "0";

                                fields[10] = "Found";  // set to insert into order_invoice table

                            }
                            fields[9] = Max_ORDinvoiceID.ToString(); // set to insert into order_invoice table   map column order amount
                            fields[16] = Max_ORDinvoiceID.ToString(); // set to insert into order_invoice table
                            Max_ORDinvoiceID = Max_ORDinvoiceID + 1;

                            fields[14] = rejectReason.ToString();
                            fields[15] = System.DateTime.Now.ToString("dd/MM/yyyy HH:MM:ss tt");
                            dataTable.Rows.Add(fields);
                            if (dataTable.Rows[ord_recordNo][14].ToString() == "")  // add only if no any reason to avoid old reason removal
                                dataTable.Rows[ord_recordNo][14] = rejectReason;  // set rejected reason to respected order no
                            if (isRejected == true)
                            {
                                dataTable.Rows[ord_recordNo][13] = "1";
                                //Max_OrderIDRej = Max_OrderIDRej + 1;   // auto order id
                                dataTable.Rows[ord_recordNo][16] = Max_OrderIDRej.ToString();
                            }
                            else
                                dataTable.Rows[ord_recordNo][13] = "0";
                        }
                    }
                }
            }
            catch (Exception ex) { clserr.WriteErrorToTxtFile(ex.Message, "ReadCsvFileToDataTable", filePath); }
            return dataTable;
        }
        private string Create_Reg_File_Order(string InputFile)
        {
            string OFileName = "", Create_Reg_File_Order = "";
            try
            {
                if (Directory.Exists(Application.StartupPath + "\\Order_Delete") == false)
                    Directory.CreateDirectory(Application.StartupPath + "\\Order_Delete");
                Boolean Rej_Flag = false, Bool_in = true, flgPymt = false, flgDOFound = false;
                string strDO = "", Rej_Reason = "", ORD_Rej_Reason = "", Mail_DO = "", Order_ID = "", Mail_Body = "", Mail_OrderNO = "", Mail_InvoiceNO = "", Order_Inv_ID = "";
                OFileName = Application.StartupPath + "\\Order_Delete" + "\\MSILIN01.ORDDELFAIL." + System.DateTime.Now.ToString("ddMMyyyyHHmmss") + ".CSV";
                int TotalRows = 0, intInvCnt = 0, intFlInvCnt = 0;
                double Ord_Inv_Amt = 0, Order_Amt = 0;
                clserr.LogEntry("Entered Create_Reg_File_DOInvoiceCancel Create_Reg_File_Order" + "   Create_Reg_File_Order", false);
                clserr.LogEntry("Process File name  " + Path.GetFileName(InputFile) + "Create_Reg_File", false);
                string OPFileName = ChagneFromCSVToExcel(InputFile);
                XSSFWorkbook workbook;
                ISheet excelSheet;
                using (var fs = new FileStream(OPFileName, FileMode.Open, FileAccess.Read))
                {
                    workbook = new XSSFWorkbook(fs);
                    excelSheet = workbook.GetSheet("Sheet1");
                    TotalRows = excelSheet.LastRowNum + 1;
                    fs.Close();
                    fs.Dispose();
                }
                DataTable dtFile = service.InsertDetails("Insert_FileDesc", Path.GetFileName(InputFile).ToUpper(), "O", TotalRows.ToString(), "", "", "", "");

                if (TotalRows == 0)
                {
                    fncl.SendEmail(sttg.ORD_MIS_EMAIL, "", "", InputFile, "REJECTED ORDER DATA", "BLANK DO FILE.");
                    clserr.LogEntry("No Records Found in Order File  - " + "   Create_Reg_File_Order", false);
                }

                StreamWriter ts_in = new StreamWriter(OFileName);
                for (int kk = 0; kk < TotalRows; kk++)
                {
                    Rej_Flag = false;
                    Rej_Reason = "";
                    ORD_Rej_Reason = "";
                    if (excelSheet.GetRow(kk).GetCell(0).StringCellValue == "ORD")
                    {
                        if (excelSheet.GetRow(kk).GetCell(1).StringCellValue == "")
                        {
                            clserr.LogEntry("DO Number is Blank" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);

                            Rej_Flag = true;
                            Rej_Reason = "DO Number is Blank.";
                            ORD_Rej_Reason = "DO Number is Blank.";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(2).StringCellValue == "")
                        {
                            clserr.LogEntry("DO Date is Blank." + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "DO Date is Blank.";
                            ORD_Rej_Reason = "DO Date is Blank.";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(3).StringCellValue == "")
                        {
                            clserr.LogEntry("Dealer Code is Blank" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "Dealer Code is Blank";
                            ORD_Rej_Reason = "Dealer Code is Blank";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(4).StringCellValue == "")
                        {
                            clserr.LogEntry("Dealer Destination Code is Blank" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "Dealer Destination Code is Blank";
                            ORD_Rej_Reason = "Dealer Destination Code is Blank";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(9).StringCellValue == "")
                        {
                            clserr.LogEntry("Order Amount is Blank" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "Order Amount is Blank";
                            ORD_Rej_Reason = "Order Amount is Blank";
                            goto Starting;
                        }
                        Boolean ismatch = Regex.IsMatch(excelSheet.GetRow(kk).GetCell(9).StringCellValue, @"^[\+\-]?\d*\.?[Ee]?[\+\-]?\d*$");
                        if (ismatch == false)
                        {
                            clserr.LogEntry("Order Amount is Not Numeric" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "Order Amount is Not Numeric";
                            ORD_Rej_Reason = "Order Amount is Not Numeric";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk + 1).GetCell(0).StringCellValue != "INV" || excelSheet.GetRow(kk + 1).GetCell(0).StringCellValue.ToUpper() == "")
                        {
                            clserr.LogEntry("Order data Does not Contain Any Invoice Records. For Order Number :" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "Order data Does not Contain Any Invoice Records.";
                            ORD_Rej_Reason = "Order data Does not Contain Any Invoice Records.";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(1).StringCellValue.Length != 22)
                        {
                            clserr.LogEntry("Wrong Length. Order Number : " + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "Length of Order Number is invalid";
                            ORD_Rej_Reason = "Length of Order Number is invalid";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(1).StringCellValue.Length > 22)
                        {
                            clserr.LogEntry("Length of DO number is greater than 22 " + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "Length of DO number is greater than 22";
                            ORD_Rej_Reason = "Length of DO number is greater than 22";
                            goto Starting;
                        }
                        if (excelSheet.GetRow(kk).GetCell(1).StringCellValue.Length < 22)
                        {
                            clserr.LogEntry("Length of DO number is less than 22" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "Length of DO number is less than 22";
                            ORD_Rej_Reason = "Length of DO number is less than 22";
                            goto Starting;
                        }
                        if ((excelSheet.GetRow(kk).GetCell(1).StringCellValue.ToUpper().Substring(0, 6) != "HDFCIN") && (excelSheet.GetRow(kk).GetCell(1).StringCellValue.ToUpper().Substring(0, 6) != "MARBNG") &&
                            (excelSheet.GetRow(kk).GetCell(1).StringCellValue.ToUpper().Substring(0, 5) != "MARSL") && (excelSheet.GetRow(kk).GetCell(1).StringCellValue.ToUpper().Substring(0, 6) != "MRKKH3"))
                        {
                            clserr.LogEntry("DO No. start with HDFCIN or MARBNG or MARSL or MRKKH3. Wrong Order Number : " + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "DO No. start with HDFCIN or MARBNG or MARSL or MRKKH3. Wrong Order Number : ";
                            ORD_Rej_Reason = "DO No. start with HDFCIN or MARBNG or MARSL or MRKKH3. Wrong Order Number : ";
                            goto Starting;
                        }
                        DataTable DT_FT = service.getDetails("Get_Order_Desc", excelSheet.GetRow(kk).GetCell(1).StringCellValue, "", "", "", "", "", "");

                        if (DT_FT.Rows.Count > 0)
                        {
                            clserr.LogEntry("Duplicate Order Number :" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "Duplicate Order Number :";
                            ORD_Rej_Reason = "Duplicate Order Number :";
                            goto Starting;
                        }
                        if (Rej_Flag == false)
                        {
                            if (excelSheet.GetRow(kk).GetCell(8).StringCellValue == "")
                            {
                                clserr.LogEntry("No Email ID for Order Number : " + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                                Rej_Flag = true;
                                Rej_Reason = "No Email ID for Order Number : ";
                                ORD_Rej_Reason = "REJECTED NO EMAIL ID";
                                goto Starting;
                            }
                        }
                        Order_Amt = Convert.ToDouble(excelSheet.GetRow(kk).GetCell(9).StringCellValue);
                        Ord_Inv_Amt = 0;
                        for (int NRow = kk; NRow <= TotalRows - 1; NRow++)
                        {
                            //if ((excelSheet.GetRow(NRow + 1).GetCell(0).StringCellValue.ToUpper() == "INV"))
                            //    Ord_Inv_Amt = Ord_Inv_Amt + Convert.ToDouble(excelSheet.GetRow(NRow + 1).GetCell(2).StringCellValue);
                            //else
                            //    break;

                            if (NRow == 0 && (excelSheet.GetRow(NRow).GetCell(0).StringCellValue.ToUpper() == "ORD"))
                            {
                                if ((excelSheet.GetRow(NRow + 1).GetCell(0).StringCellValue.ToUpper() == "INV"))
                                    Ord_Inv_Amt = Ord_Inv_Amt + Convert.ToDouble(excelSheet.GetRow(NRow + 1).GetCell(2).StringCellValue);
                                else
                                    break;
                            }
                            else if (NRow != TotalRows - 1)
                            {
                                if ((excelSheet.GetRow(NRow + 1).GetCell(0).StringCellValue.ToUpper() == "INV"))
                                    Ord_Inv_Amt = Ord_Inv_Amt + Convert.ToDouble(excelSheet.GetRow(NRow + 1).GetCell(2).StringCellValue);
                                else
                                    break;
                            }
                        }

                        if (Order_Amt.ToString("0.00") != Ord_Inv_Amt.ToString("0.00"))
                        {
                            clserr.LogEntry("Wrong Amount For Order No :" + excelSheet.GetRow(kk).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                            Rej_Flag = true;
                            Rej_Reason = "DO amount does not match with Invoice amount. ";
                            ORD_Rej_Reason = "DO amount does not match with Invoice amount.";
                            goto Starting;
                        }
                        if (Rej_Flag == false)
                        {
                            for (int Row_New = kk; Row_New < TotalRows; Row_New++)
                            {
                                if (excelSheet.GetRow(Row_New).GetCell(0).StringCellValue.ToUpper() == "INV")
                                {
                                    DataTable DT_k = service.getDetails("Get_Order_InvoiceAsPer", excelSheet.GetRow(Row_New).GetCell(1).StringCellValue, "", "", "", "", "", "");
                                    if (DT_k.Rows.Count > 0)
                                    {
                                        clserr.LogEntry("Duplicate Invoice Number :" + excelSheet.GetRow(Row_New).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                                        Rej_Flag = true;
                                        Rej_Reason = "Duplicate Invoice Number:";
                                        ORD_Rej_Reason = "Duplicate Invoice Number :";
                                        goto Starting;
                                    }
                                    DataTable DT_T = service.getDetails("Get_Order_InvoiceAsPerInvoice", excelSheet.GetRow(Row_New).GetCell(1).StringCellValue, "", "", "", "", "", "");
                                    if (DT_T.Rows.Count == 0)
                                    {
                                        clserr.LogEntry("Invoice Number Not Found. :" + excelSheet.GetRow(Row_New).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                                        Rej_Flag = true;
                                        //Rej_Reason = "Invoice Number Not Found.";
                                        Rej_Reason = "Invoice Bill Number not generated"; ////Changeas discuss on call on 21-12-2023
                                        ORD_Rej_Reason = "INVOICE NOT FOUND";
                                        goto Starting;
                                    }
                                    DataTable DT_ = service.getDetails("Get_Order_InvoiceAsPerAmount", excelSheet.GetRow(Row_New).GetCell(1).StringCellValue, Convert.ToDouble(excelSheet.GetRow(Row_New).GetCell(2).StringCellValue).ToString(), "", "", "", "", "");
                                    if (DT_.Rows.Count == 0)
                                    {
                                        clserr.LogEntry("Invoice Amount not match. Invoice Number " + excelSheet.GetRow(Row_New).GetCell(1).StringCellValue + "   Create_Reg_File_DOInvoiceCancel", false);
                                        Rej_Flag = true;
                                        Rej_Reason = "Invoice Amount not match.";
                                        ORD_Rej_Reason = "INVOICE AMOUNT NOT MATCH.";
                                        fncl.SendEmail(sttg.ORD_MIS_EMAIL, "", "", InputFile, "Amount Mismatched in OrderData File", "INVOICE AMOUNT NOT MATCH.");
                                        goto Starting;
                                    }
                                }
                                else if (excelSheet.GetRow(Row_New).GetCell(0).StringCellValue.ToUpper() == "INV" || excelSheet.GetRow(Row_New).GetCell(0).StringCellValue.Trim() == "")
                                    break;
                            }
                        }
                        Order_Amt = Convert.ToDouble(excelSheet.GetRow(kk).GetCell(9).StringCellValue);
                    }
                Starting:
                    if (excelSheet.GetRow(kk).GetCell(0).StringCellValue == "ORD")
                    {
                        if (Rej_Flag == false)
                            //Mail_OrderNO = Mail_OrderNO + "," + excelSheet.GetRow(kk).GetCell(1).StringCellValue;
                            Mail_OrderNO = excelSheet.GetRow(kk).GetCell(1).StringCellValue;
                        Order_ID = "";
                        DataTable DT_FT1 = service.getDetails("Get_MaxOrder_desc", "", "", "", "", "", "", "");
                        Order_ID = DT_FT1.Rows[0][0].ToString();
                        clserr.LogEntry("Updating DO records in Order Desc Table" + "Order Desc " + InputFile + "   Create_Reg_File_Order", false);
                        InsertOrderData("0", "Order", excelSheet, Rej_Reason, Rej_Flag, kk);
                    }
                    else if (excelSheet.GetRow(kk).GetCell(0).StringCellValue == "INV" && Order_ID != "")
                    {
                        Order_Inv_ID = "0";
                        DataTable DT_FT1 = service.getDetails("Get_MaxOrder_Invoice", "", "", "", "", "", "", "");
                        Order_Inv_ID = DT_FT1.Rows[0][0].ToString();
                        clserr.LogEntry("Updating order records in Invoice Table" + InputFile + "   Create_Reg_Invoice_Cancel", false);
                        InsertOrderData(Order_ID, "Invoice", excelSheet, Rej_Reason, Rej_Flag, kk);
                    }
                }
                DataTable DT_R = service.getDetails("Get_order_DescRejected", "", "", "", "", "", "", "");
                for (int hh = 0; hh < DT_R.Rows.Count; hh++)
                {
                    string[] details = new string[25];
                    details[0] = "0"; //Order_ID_Rej
                    details[1] = DT_R.Rows[hh]["Order_ID"].ToString(); //Order_ID
                    details[2] = Convert.ToDateTime(DT_R.Rows[hh]["Order_Date"].ToString()).ToString("yyyy-MM-dd hh:MM:ss"); //Order_Date
                    details[3] = DT_R.Rows[hh]["Record_Identifier"].ToString(); //Record_Identifier
                    details[4] = DT_R.Rows[hh]["DO_number"].ToString();   //DO_Number
                    details[5] = DT_R.Rows[hh]["Do_Date"].ToString();   //Do_Date
                    details[6] = DT_R.Rows[hh]["Dealer_Code"].ToString();   //Dealer_Code
                    details[7] = DT_R.Rows[hh]["Dealer_Destination_code"].ToString(); //Dealer_Destination_Code
                    details[8] = DT_R.Rows[hh]["Dealer_Outlet_Code"].ToString(); //Dealer_Outlet_Co;de
                    details[9] = DT_R.Rows[hh]["Financier_Code"].ToString(); //Financier_Code
                    details[10] = DT_R.Rows[hh]["Financier_Name"].ToString(); //Financier_Name
                    details[11] = DT_R.Rows[hh]["Email_IDs"].ToString();   //Email_IDs
                    details[12] = DT_R.Rows[hh]["Order_amount"].ToString();   //Order_Amount                   
                    details[13] = DT_R.Rows[hh]["Order_Status"].ToString();   //Order_Status
                    details[14] = DT_R.Rows[hh]["Order_Data_Received_On"].ToString(); //Order_Data_Received_On
                    details[15] = DT_R.Rows[hh]["ORD_Rej_Reason"].ToString(); //Rejected_Reson
                    details[16] = "NULL"; //EmailStatus

                    DataTable dtInsert = service.Insert_Order_RejectedDetails("Save", details);
                    DataTable DT_REj = service.getDetails("Get_order_DescRejectedAll", DT_R.Rows[hh]["Order_ID"].ToString(), "", "", "", "", "", "");
                    for (int pp = 0; pp < DT_REj.Rows.Count; pp++)
                    {
                        string[] detailsCash = new string[25];
                        detailsCash[0] = "0"; //Ord_Inv_ID_Rej
                        detailsCash[1] = DT_REj.Rows[pp]["Ord_Inv_ID"].ToString(); //Ord_Inv_ID
                        detailsCash[2] = DT_REj.Rows[pp]["Order_ID"].ToString(); //Order_ID
                        detailsCash[3] = DT_REj.Rows[pp]["Record_Identifier"].ToString(); //Record_Identifier
                        detailsCash[4] = DT_REj.Rows[pp]["Order_Inv_Number"].ToString();   //Order_Inv_Number
                        detailsCash[5] = DT_REj.Rows[pp]["Order_Inv_Amount"].ToString();   //Order_Inv_Amount
                        detailsCash[6] = DT_REj.Rows[pp]["Order_Inv_Status"].ToString();   //Order_Inv_Status
                        detailsCash[7] = dtInsert.Rows[0][0].ToString(); //Order_ID_Rej                 
                        DataTable dtInsertR = service.Insert_Order_Invoice_RejectedDetails("Save", detailsCash);
                    }

                    DataTable DT_REject = service.UpdateDetails("Delete_OrderINVRejectedDetails", DT_R.Rows[hh]["Order_ID"].ToString(), "", "", "", "", "", "");
                }
                DataTable DT_check = service.getDetails("Get_CheckingInvoice", "", "", "", "", "", "", "");
                for (int i = 0; i < DT_check.Rows.Count; i++)
                {
                    DataTable DT_Get = service.getDetails("Get_InvoiceDetailCheck", DT_check.Rows[i]["Order_Inv_Number"].ToString(), "", "", "", "", "", "");
                    if (DT_Get.Rows.Count > 0)
                    {
                        DataTable DT_Update = service.UpdateDetails("Update_OrderINVDesc_Details", DT_check.Rows[i]["Order_Inv_Number"].ToString(), DT_check.Rows[i]["DO_number"].ToString(),
                             DT_check.Rows[i]["Ord_Inv_ID"].ToString(), System.DateTime.Now.ToString("dd/MMM/yyyy"), DT_check.Rows[i]["Order_ID"].ToString(), "", "");
                    }
                }

                SendOrderRejectedEmails_WOAtt();
                SendOrderRejectedEmails(InputFile);
                if (Mail_OrderNO.Trim() != "")
                {
                    //fncl.SendEmail(sttg.ORD_MIS_EMAIL, "", "", "", "Order Data Successful Uploaded for Follwing Order Nos. : ", "Successful ORDER DATA : " + Mail_OrderNO);
                    fncl.SendEmail(sttg.ORD_MIS_EMAIL, "", "", OPFileName, "MSIL ORDER DATA", ""); ////Added on 14-12-2023 which is discuss on call
                    clserr.LogEntry(" " + "   Create_Reg_File_Order", false);
                }
                Mail_OrderNO = "";
                Create_Reg_File_Order = "";
            }
            catch (Exception ex)
            {
                Create_Reg_File_Order = "ERROR";
                File.Move(InputFile + ".CSV", sttg.BackupFilePath);
                clserr.WriteErrorToTxtFile(ex.Message, "Create_Reg_File_Order_Delete", InputFile);
                clserr.WriteErrorToTxtFile("Exiting function Create_Reg_File_DO_Delete", "Create_Reg_File_DO_Delete", InputFile);
            }
            return Create_Reg_File_Order;
        }
        private void SendOrderRejectedEmails(string strFileName)
        {

            string OpLine = "", Rej_Order_FileName; Boolean ChkFlag = false;
            int cnt = 0;
            Rej_Order_FileName = Path.GetFileName(strFileName).ToUpper().Replace(".TXT", "") + "_" + System.DateTime.Now.ToString("ddMMyyyyHHmmss") + ".CSV";
            ChkFlag = false;
            if (!Directory.Exists(Application.StartupPath + "\\Order")) Directory.CreateDirectory(Application.StartupPath + "\\Order");

            StreamWriter ts_Order = new StreamWriter(File.Create(Application.StartupPath + "\\Order\\" + Rej_Order_FileName));
            try
            {
                DataTable DT_check = service.getDetails("Get_Order_Rejected", "", "", "", "", "", "", "");

                for (int k = 0; k < DT_check.Rows.Count; k++)
                {

                    ChkFlag = true;

                    OpLine = "";
                    OpLine = DT_check.Rows[k][3].ToString() + ",";
                    for (cnt = 3; cnt < 12; cnt++)
                        OpLine = OpLine + DT_check.Rows[k][cnt].ToString() + ",";

                    OpLine = OpLine + DT_check.Rows[k]["Rejected_Reson"].ToString().Replace(",", "") + ",";
                    ts_Order.WriteLine(OpLine, true);
                    OpLine = "";

                    DataTable DT_INV = service.getDetails("Get_Order_Invoice_Rejected", DT_check.Rows[k]["Order_ID_Rej"].ToString(), "", "", "", "", "", "");
                    for (int j = 0; j < DT_INV.Rows.Count; j++)
                    {
                        OpLine = "";
                        OpLine = DT_INV.Rows[j][2].ToString() + ",";
                        for (cnt = 3; cnt < 5; cnt++)
                            OpLine = OpLine + DT_INV.Rows[j][cnt].ToString().Replace(",", "") + ",";
                        ts_Order.WriteLine(OpLine);

                    }
                    DataTable DT_ch = service.UpdateDetails("Update_RejectedDetails", DT_check.Rows[k]["Order_ID_Rej"].ToString(), "", "", "", "", "", "");
                }

                ts_Order.Close();
                ts_Order = null;
                if (ChkFlag == true)
                {
                    fncl.SendEmail(sttg.ORD_MIS_EMAIL, "", "", Application.StartupPath + "\\Order\\" + Rej_Order_FileName, "REJECTED ORDER DATA", "FYI");
                    clserr.LogEntry("Order File Rejected due to " + DT_check.Rows[0]["Order_ID_Rej"].ToString() + "   Create_Reg_File_Order", false);
                }
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();
                if (File.Exists(Application.StartupPath + "\\Order\\" + Rej_Order_FileName))
                    File.Delete(Application.StartupPath + "\\Order\\" + Rej_Order_FileName);
                if (Directory.Exists(Application.StartupPath + "\\Order"))
                    Directory.Delete(Application.StartupPath + "\\Order");
            }
            catch (Exception ex)
            {
                ts_Order.Close();
                ts_Order.Dispose();
                clserr.WriteErrorToTxtFile(ex.Message + "Form1", "SendOrderRejectedEmails", strFileName);
            }
        }
        private void SendOrderRejectedEmails_WOAtt() //'Sending Order Rejected Emails Without attachment
        {
            try
            {
                int cnt = 0;
                Boolean ChkFlag = false;
                FileStream ts_Order;
                string Rej_Order_FileName = "", OpLine = "";


                DataTable DT_check = service.getDetails("Get_Order_Rejected", "", "", "", "", "", "", "");
                for (int i = 0; i < DT_check.Rows.Count; i++)
                {
                    ChkFlag = true;
                    string ToEmail = "", Emnail_Matter;

                    ToEmail = DT_check.Rows[i]["Email_IDs"].ToString().Replace(((Char)34).ToString(), "");
                    ToEmail = ToEmail.Replace(",", ";");
                    Emnail_Matter = "REJECTED ORDER DATA" + System.Environment.NewLine + "Order No : " + DT_check.Rows[i]["DO_number"].ToString() + "\t" + "\t" + " Rejected Reason : " + DT_check.Rows[i]["Rejected_Reson"].ToString();
                    fncl.SendEmail(sttg.ORD_MIS_EMAIL, "", "", "", "REJECTED ORDER DATA", Emnail_Matter);
                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message + "Form1", "SendOrderRejectedEmails", "");
                clserr.LogEntry(ex.Message + " Send Order Reject MAIL2" + "   SendOrderRejectedEmails_WOAtt", false);
            }
        }
        private void InsertOrderData(string Order_ID, string TableName, ISheet excelSheet, string Rej_Reason, Boolean Rej_Flag, int kk)
        {
            try
            {
                DataTable DT_FT = service.getDetails("Get_DO_DescAsOrd_ID", Order_ID, "", "", "", "", "", "");
                if (TableName == "Order")
                {
                    string[] detailsCash = new string[25];
                    detailsCash[0] = Order_ID; //Order_ID

                    detailsCash[1] = System.DateTime.Now.Date.ToString("dd/MMM/yyyy"); //Order_Date
                    detailsCash[2] = excelSheet.GetRow(kk).GetCell(0).StringCellValue; //Record_Identifier
                    detailsCash[3] = excelSheet.GetRow(kk).GetCell(1).StringCellValue; //DO_Number
                    detailsCash[4] = excelSheet.GetRow(kk).GetCell(2).StringCellValue;   //Do_Date
                    detailsCash[5] = excelSheet.GetRow(kk).GetCell(3).StringCellValue;   //Dealer_Code
                    detailsCash[6] = excelSheet.GetRow(kk).GetCell(4).StringCellValue;   //Dealer_Destination_Code
                    detailsCash[7] = excelSheet.GetRow(kk).GetCell(5).StringCellValue; //Dealer_Outlet_Code
                    detailsCash[8] = excelSheet.GetRow(kk).GetCell(6).StringCellValue; //Financier_Code
                    detailsCash[9] = excelSheet.GetRow(kk).GetCell(7).StringCellValue; //Financier_Name
                    detailsCash[10] = excelSheet.GetRow(kk).GetCell(8).StringCellValue; //Email_IDs
                    detailsCash[11] = excelSheet.GetRow(kk).GetCell(9).StringCellValue;   //Order_amount
                    detailsCash[12] = "";   //Order_Status
                    detailsCash[13] = System.DateTime.Now.ToString("dd/MM/yyyy HH:MM:ss tt");   //Order_Data_Received_On
                    detailsCash[14] = "";   //Cash_Ops_ID
                    if (Rej_Flag == false)
                        detailsCash[15] = "0";   //DO_Rej_Flag
                    else
                        detailsCash[15] = "1";   //DO_Rej_Flag
                    detailsCash[16] = Rej_Reason; //ORD_Rej_Reason
                    detailsCash[17] = "NULL"; //
                    detailsCash[18] = "NULL"; //
                    detailsCash[19] = "NULL";//
                    detailsCash[20] = "NULL";  //

                    DataTable dtInsert = service.Insert_Order_Desc_Details("Save", detailsCash);
                }
                if (TableName == "Invoice")
                {
                    string[] detailsCash = new string[15];
                    detailsCash[0] = "0"; //Ord_Inv_ID
                    detailsCash[1] = Order_ID; //Order_ID
                    detailsCash[2] = excelSheet.GetRow(kk).GetCell(0).StringCellValue; //Record_Identifier
                    detailsCash[3] = excelSheet.GetRow(kk).GetCell(1).StringCellValue; //Order_Inv_Number
                    detailsCash[4] = excelSheet.GetRow(kk).GetCell(2).StringCellValue;   //Order_Inv_Amount
                    detailsCash[5] = "Not Found";   //Order_Inv_Status   

                    DataTable dtInsert = service.Insert_Order_Invoice_Details("Save", detailsCash);
                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message + "Form1", "InsertOrderData", TableName);
            }
        }
        private void InsertDOData(string Order_ID, string TableName, ISheet excelSheet, string Rej_Reason, Boolean Rej_Flag, int kk)
        {
            try
            {
                DataTable DT_FT = service.getDetails("Get_DO_DescAsOrd_ID", Order_ID, "", "", "", "", "", "");
                if (TableName == "DO_Desc")
                {
                    string[] detailsCash = new string[25];
                    detailsCash[0] = Order_ID; //Order_ID
                    detailsCash[1] = System.DateTime.Now.Date.ToString("dd/MMM/yyyy"); //Order_Date
                    detailsCash[2] = excelSheet.GetRow(kk).GetCell(0).StringCellValue; //Record_Identifier
                    detailsCash[3] = excelSheet.GetRow(kk).GetCell(1).StringCellValue; //DO_Number
                    detailsCash[4] = excelSheet.GetRow(kk).GetCell(2).StringCellValue;   //Do_Date
                    detailsCash[5] = excelSheet.GetRow(kk).GetCell(3).StringCellValue;   //Dealer_Code
                    detailsCash[6] = excelSheet.GetRow(kk).GetCell(4).StringCellValue;   //Dealer_Destination_Code
                    detailsCash[7] = excelSheet.GetRow(kk).GetCell(5).StringCellValue; //Dealer_Outlet_Code
                    detailsCash[8] = excelSheet.GetRow(kk).GetCell(6).StringCellValue; //Order_Amount
                    detailsCash[9] = System.DateTime.Now.Date.ToString("dd/MMM/yyyy"); //DO_Requested_date
                    detailsCash[10] = "NULL"; //DO_Updated_date
                    detailsCash[11] = "NULL";   //Delete_Flag
                    detailsCash[12] = Rej_Reason;   //Reason
                    if (Rej_Flag == false)
                        detailsCash[13] = "0";   //DO_Rej_Flag
                    else
                        detailsCash[13] = "1";   //DO_Rej_Flag
                    detailsCash[14] = "0"; //F2_MIS
                    detailsCash[15] = "0"; //DO_IN_Flag
                    detailsCash[16] = "NULL"; //Authorize_Flag
                    detailsCash[17] = "NULL";//Deleted_By
                    detailsCash[18] = "NULL";  //Deleted_On
                    detailsCash[19] = "NULL"; ;   //Authorized_By
                    detailsCash[20] = "NULL"; //Authorized_On

                    DataTable dtInsert = service.Insert_DO_DescDetails("Save", detailsCash);
                }
                else if (TableName == "invoice_cancel_desc")
                {
                    string[] detailsCash = new string[25];
                    detailsCash[0] = "0";   //Invoice_ID
                    detailsCash[1] = "0"; //Sr_No
                    detailsCash[2] = excelSheet.GetRow(kk).GetCell(1).StringCellValue; //Invoice_Number
                    detailsCash[3] = excelSheet.GetRow(kk).GetCell(2).StringCellValue; //Invoice_Amount
                    detailsCash[4] = System.DateTime.Now.Date.ToString("dd/MMM/yyyy"); //Requested_Date
                    detailsCash[5] = "NULL";   //Updated_Date
                    detailsCash[6] = DT_FT.Rows[0].ItemArray[3].ToString(); //"1";   //DO_number ////Correction on 17-12-2023
                    detailsCash[7] = "NULL";   //Order_ID
                    detailsCash[8] = "0";   //Cancelled_Flag
                    detailsCash[9] = Rej_Reason;   //Reason
                    detailsCash[10] = "0";   //F2_MIS
                    detailsCash[11] = "1";   //DO_IN_Flag
                    detailsCash[12] = "NULL";   //Authorize_Flag
                    detailsCash[13] = "NULL";   //Deleted_By
                    detailsCash[14] = "NULL";   //Deleted_On
                    detailsCash[15] = "NULL";   //Authorized_By
                    detailsCash[16] = "NULL";   //Authorized_On

                    DataTable dtInsert = service.Insert_InvoiceCancel_DescDetails("Save", detailsCash);
                }
            }
            catch (Exception ex) { clserr.WriteErrorToTxtFile(ex.Message + "Form1", "InsertDOData", TableName); }
        }
        private string ChagneFromCSVToExcel_INV(string InputFile)
        {
            string OpFileName = "";
            try
            {
                clserr.LogEntry("Entering function ChangeFromCSVToExcel_INV" + "   ChagneFromCSVToExcel_INV", false);
                StreamReader objStrmReader = new StreamReader(InputFile);
                if (!Directory.Exists(Application.StartupPath + "\\Temp"))
                    Directory.CreateDirectory(Application.StartupPath + "\\Temp");
                OpFileName = (Application.StartupPath + "\\Temp\\Temp" + System.DateTime.Now.ToString("ddMMyyyyHHmmss")) + ".xlsx";
                string ReadOP_Line = "";
                string[] ExcelColumns = { "" };
                fncl.CreateExcel(ExcelColumns, OpFileName);
                int C = 0, i = 0, k = 0;
                using (var fs = new FileStream(OpFileName, FileMode.Open, FileAccess.ReadWrite))
                {
                    XSSFWorkbook wb = new XSSFWorkbook(fs);
                    while (!objStrmReader.EndOfStream)
                    {
                        ReadOP_Line = objStrmReader.ReadLine();
                        string[] RL_split;
                        C = 0;
                        RL_split = ReadOP_Line.Split(",");
                        ISheet excelSheetOut = wb.GetSheet("Sheet1");
                        IRow rowHeader = excelSheetOut.CreateRow(i);
                        for (int j = 0; j < RL_split.Length; j++)
                        {
                            if (RL_split[j].ToString() == "")
                                rowHeader.CreateCell(C).SetCellValue(RL_split[j]);
                            else if (RL_split[j].Substring(0, 1) == "0")
                            {
                                rowHeader.CreateCell(C).SetCellValue(RL_split[j]);
                            }
                            else
                            {
                                if (RL_split[j].Substring(0, 1) == "\"")
                                {
                                    if (RL_split[j].Substring(RL_split[j].Length - 1, 1) == "\"")
                                        rowHeader.CreateCell(C).SetCellValue(RL_split[j].Replace("\"", ""));
                                    else
                                    {
                                        ICell cell = rowHeader.CreateCell(C);
                                        cell.SetCellValue(RL_split[j].Replace("\"", ""));
                                        j = j + 1;
                                        for (k = j; k < RL_split.Length; k++)
                                        {
                                            string ss = RL_split[k].Substring(RL_split[k].Length - 1, 1);
                                            if (RL_split[k].Substring(RL_split[k].Length - 1, 1) == "\"")
                                            {
                                                cell.SetCellValue(rowHeader.GetCell(C).StringCellValue + "," + RL_split[k].Replace("\"", ""));
                                                j = k;
                                                break;
                                                //return "";
                                            }
                                            else
                                            {
                                                cell.SetCellValue(rowHeader.GetCell(C).StringCellValue + "," + RL_split[k].Replace("\"", ""));
                                            }
                                        }
                                    }
                                }
                                else
                                    rowHeader.CreateCell(C).SetCellValue(RL_split[j]);

                            }
                            C = C + 1;
                        }
                        i = i + 1;
                    }
                    var fsNew = new FileStream(OpFileName, FileMode.Create, FileAccess.Write);
                    wb.Write(fsNew);
                    fs.Close();
                    fs.Dispose();
                    fsNew.Close();
                    wb.Dispose();
                    fsNew.Dispose();
                }
                clserr.LogEntry("Exiting function ChagneFromCSVToExcel_INV" + "   ChagneFromCSVToExcel_INV", false);
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message, "ChagneFromCSVToExcel_INV", InputFile);
            }
            return OpFileName;
        }
        private string ChagneFromCSVToExcel(string InputFile)
        {
            string OpFileName = "";
            try
            {
                clserr.LogEntry("Entering function ChagneFromCSVToExcel" + "   ChagneFromCSVToExcel", false);
                File.Move(InputFile, InputFile + ".CSV");
                StreamReader objStrmReader = new StreamReader(InputFile + ".CSV");
                if (!Directory.Exists(Application.StartupPath + "\\Temp"))
                    Directory.CreateDirectory(Application.StartupPath + "\\Temp");

                //OpFileName = (Application.StartupPath + "\\Temp\\Temp" + System.DateTime.Now.ToString("ddMMyyyyHHmmss") + ".xlsx");
                OpFileName = (Application.StartupPath + "\\Temp\\" + Path.GetFileName(InputFile).ToUpper().Replace("CSV", "") + "_" + System.DateTime.Now.ToString("ddMMyyyyHHmmss") + ".xlsx");
                string ReadOP_Line = "", Email_IDs = "";
                string[] ExcelColumns = { "" };
                fncl.CreateExcel(ExcelColumns, OpFileName);
                int i = 0, j = 0, C = 0;
                using (var fs = new FileStream(OpFileName, FileMode.Open, FileAccess.ReadWrite))
                {
                    XSSFWorkbook wb = new XSSFWorkbook(fs);
                    while (!objStrmReader.EndOfStream)
                    {
                        Email_IDs = "";
                        ReadOP_Line = objStrmReader.ReadLine();
                        if (ReadOP_Line.Trim() != "" || ReadOP_Line.Trim() != null)
                        {

                            string[] RL_split;

                            RL_split = ReadOP_Line.Split(",");
                            ISheet excelSheetOut = wb.GetSheet("Sheet1");
                            IRow rowHeader = excelSheetOut.CreateRow(i);
                            for (j = 0; j < RL_split.Length; j++)
                            {

                                if (j >= 8)
                                {
                                    Email_IDs = Email_IDs + "," + RL_split[j];
                                    if (j == RL_split.Length - 1)
                                        rowHeader.CreateCell(9).SetCellValue(RL_split[j].Replace(((Char)32).ToString(), ""));
                                    else
                                        rowHeader.CreateCell(8).SetCellValue(Email_IDs.Substring(1, Email_IDs.Length - 1).Replace(((Char)32).ToString(), ""));
                                }
                                else
                                {
                                    if (j == 0)
                                        rowHeader.CreateCell(j).SetCellValue(RL_split[j].Replace(((Char)32).ToString(), ""));
                                    else
                                    {
                                        if (RL_split[j].ToString() == "")
                                            rowHeader.CreateCell(j).SetCellValue(RL_split[j].Replace(((Char)32).ToString(), ""));
                                        else if (RL_split[j].Substring(0, 1) == "0")
                                            rowHeader.CreateCell(j).SetCellValue(RL_split[j].Replace(((Char)32).ToString(), ""));
                                        else
                                            rowHeader.CreateCell(j).SetCellValue(RL_split[j].Replace(((Char)32).ToString(), ""));
                                    }
                                }
                            }
                            i = i + 1;
                        }
                    }
                    var fsNew = new FileStream(OpFileName, FileMode.Create, FileAccess.Write);
                    wb.Write(fsNew);
                    fs.Close();
                    fs.Dispose();
                    fsNew.Close();
                    wb.Dispose();
                    fsNew.Dispose();
                    objStrmReader.Close();
                    File.Move(InputFile + ".CSV", InputFile);
                }

                clserr.LogEntry("Exiting function ChagneFromCSVToExcel_OrderDelete" + "   ChagneFromCSVToExcel_OrderDelete", false);
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message, "ChagneFromCSVToExcel_OrderDelete", InputFile);
            }
            return OpFileName;
        }
        private void ChagneFromCSVToExcel_OrderDelete(string InputFile)
        {
            try
            {
                clserr.LogEntry("Entering function ChagneFromCSVToExcel_OrderDelete" + "   ChagneFromCSVToExcel_OrderDelete", false);
                File.Move(InputFile, InputFile + ".CSV");
                StreamReader objStrmReader = new StreamReader(InputFile + ".CSV");
                if (!Directory.Exists(Application.StartupPath + "\\Temp"))
                    Directory.CreateDirectory(Application.StartupPath + "\\Temp");

                string OpFileName = (Application.StartupPath + "\\Temp\\Temp" + System.DateTime.Now.ToString("ddMMyyyyHHmmss")) + ".xlsx";
                string ReadOP_Line = "";
                string[] ExcelColumns = { "" };
                fncl.CreateExcel(ExcelColumns, OpFileName);
                int i = 1, j = 0;
                using (var fs = new FileStream(OpFileName, FileMode.Open, FileAccess.ReadWrite))
                {
                    XSSFWorkbook wb = new XSSFWorkbook(fs);
                    while (objStrmReader.EndOfStream)
                    {
                        ReadOP_Line = objStrmReader.ReadLine();
                        ISheet excelSheetOut = wb.GetSheet("Sheet1");
                        IRow rowHeader = excelSheetOut.CreateRow(j);
                        rowHeader.CreateCell(i).SetCellValue(ReadOP_Line.Replace(((Char)34).ToString(), ""));
                        i = i + 1;
                    }
                    var fsNew = new FileStream(OpFileName, FileMode.Create, FileAccess.Write);
                    wb.Write(fsNew);
                    fs.Close();
                    fs.Dispose();
                    fsNew.Close();
                    wb.Dispose();
                    fsNew.Dispose();
                    objStrmReader.Close();
                }

                clserr.LogEntry("Exiting function ChagneFromCSVToExcel_OrderDelete" + "   ChagneFromCSVToExcel_OrderDelete", false);
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message, "ChagneFromCSVToExcel_OrderDelete", InputFile);
            }
        }
        private string ChagneFromCSVToExcel_DO_Delete(string InputFile)
        {
            string OpFileName = "";
            try
            {
                clserr.LogEntry("Entering function ChagneFromCSVToExcel_DO_Delete" + "   ChagneFromCSVToExcel_DO_Delete", false);
                File.Move(InputFile, InputFile + ".CSV");
                StreamReader objStrmReader = new StreamReader(InputFile + ".CSV");
                if (!Directory.Exists(Application.StartupPath + "\\Temp"))
                    Directory.CreateDirectory(Application.StartupPath + "\\Temp");

                OpFileName = (Application.StartupPath + "\\Temp\\Temp" + System.DateTime.Now.ToString("ddMMyyyyHHmmss")) + ".xlsx";
                string ReadOP_Line = "", Email_IDs = "";
                string[] ExcelColumns = { "" };
                fncl.CreateExcel(ExcelColumns, OpFileName);
                int i = 0, j = 0, C = 0;
                using (var fs = new FileStream(OpFileName, FileMode.Open, FileAccess.ReadWrite))
                {
                    XSSFWorkbook wb = new XSSFWorkbook(fs);
                    while (!objStrmReader.EndOfStream)
                    {
                        ReadOP_Line = objStrmReader.ReadLine();
                        string[] RL_split;
                        C = 0;
                        RL_split = ReadOP_Line.Split(",");
                        ISheet excelSheetOut = wb.GetSheet("Sheet1");
                        IRow rowHeader = excelSheetOut.CreateRow(i);
                        for (j = 0; j < RL_split.Length; j++)
                        {
                            C = C + 1;

                            if (j >= 8)
                            {
                                Email_IDs = Email_IDs + "," + RL_split[j];
                                if (j == RL_split.Length)
                                    rowHeader.CreateCell(9).SetCellValue(RL_split[j].Replace(((Char)32).ToString(), ""));
                                else
                                    rowHeader.CreateCell(8).SetCellValue(Email_IDs.Substring(2, Email_IDs.Length - 1).Replace(((Char)32).ToString(), ""));
                            }
                            else
                            {
                                if (j == 1)
                                    rowHeader.CreateCell(j).SetCellValue(RL_split[j].Replace(((Char)32).ToString(), ""));
                                else
                                {
                                    if (RL_split[j].Substring(0, 1) == "0")
                                        rowHeader.CreateCell(j).SetCellValue(RL_split[j].Replace(((Char)32).ToString(), ""));
                                    else
                                        rowHeader.CreateCell(j).SetCellValue(RL_split[j].Replace(((Char)32).ToString(), ""));
                                }
                            }
                        }
                        i = i + 1;
                    }
                    var fsNew = new FileStream(OpFileName, FileMode.Create, FileAccess.Write);
                    wb.Write(fsNew);
                    fs.Close();
                    fs.Dispose();
                    fsNew.Close();
                    wb.Dispose();
                    fsNew.Dispose();
                    objStrmReader.Close();
                    File.Move(InputFile + ".CSV", InputFile);
                }
                clserr.LogEntry("Exiting function ChagneFromCSVToExcel_DO_Delete" + "   ChagneFromCSVToExcel_DO_Delete", false);
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message, "ChagneFromCSVToExcel_DO_Delete", InputFile);
            }
            return OpFileName;
        }

        static void ExportDataTableToExcel(DataTable dt, string filePath)

        {

            using (var workbook = new XLWorkbook())

            {

                var worksheet = workbook.Worksheets.Add(dt, "Sheet1");



                // Optional: Set the table style

                worksheet.Tables.FirstOrDefault().Theme = XLTableTheme.TableStyleMedium9;



                // Save the workbook to the specified file path

                workbook.SaveAs(filePath);

            }



            Console.WriteLine($"DataTable exported to {filePath}");

        }

        private void UpdateInvoiceRecordsNew(DataRow dtSuccRow, DataTable dtunsucc, int jj, int SRNo)
        {
            try
            {
                //clserr.LogEntry("Entering UpdateInvoiceRecords Function", false);

                DataTable dt = service.getDetails("Get_MaxFileDesc", "", "", "", "", "", "", "");
                string[] detailsCash = new string[50];
                detailsCash[0] = "0";   //Invoice_ID
                detailsCash[1] = SRNo.ToString(); //Sr_No
                detailsCash[2] = dtSuccRow.ItemArray[0].ToString(); //Invoice_Number
                detailsCash[3] = dtSuccRow.ItemArray[1].ToString(); //Invoice_Amount
                detailsCash[4] = dtSuccRow.ItemArray[2].ToString(); //Currency
                detailsCash[5] = dtSuccRow.ItemArray[3].ToString();   //Vehical_ID
                detailsCash[6] = dtSuccRow.ItemArray[4].ToString(); //DueDate
                detailsCash[7] = dtSuccRow.ItemArray[5].ToString(); //Dealer_Name
                detailsCash[8] = dtSuccRow.ItemArray[6].ToString(); //Dealer_Address1
                detailsCash[9] = dtSuccRow.ItemArray[7].ToString(); //Dealer_City
                detailsCash[10] = dtSuccRow.ItemArray[8].ToString(); //Transporter_Name
                detailsCash[11] = dtSuccRow.ItemArray[9].ToString(); //Transport_Number
                detailsCash[12] = dtSuccRow.ItemArray[10].ToString(); //Transport_Date
                detailsCash[13] = dtSuccRow.ItemArray[11].ToString(); //Dealer_Code
                detailsCash[14] = dtSuccRow.ItemArray[12].ToString(); //Transporter_Code
                detailsCash[15] = dtSuccRow.ItemArray[13].ToString(); //Dealer_Address2
                detailsCash[16] = dtSuccRow.ItemArray[14].ToString(); //Dealer_Address3
                detailsCash[17] = dtSuccRow.ItemArray[15].ToString(); //Dealer_Address4
                detailsCash[18] = DBNull.Value.ToString(); //Invoice_data_Received_Date
                detailsCash[19] = DBNull.Value.ToString(); //Physical_Invoice_Received_Date
                detailsCash[20] = "NULL"; //IMEX_DEAL_NUMBER
                detailsCash[21] = "NULL"; //StepDate
                detailsCash[22] = DBNull.Value.ToString(); //Order_Data_Received_Date
                detailsCash[23] = "NULL"; //Order_Number
                detailsCash[24] = DBNull.Value.ToString(); //Payment_Received_Date
                detailsCash[25] = "NULL"; //Utr_Number
                detailsCash[26] = "NULL"; //Invoice_Status
                detailsCash[27] = "1"; //TradeOp_Selected_Invoice_Flag
                detailsCash[28] = DBNull.Value.ToString(); //TradeOp_Selected_Invoice_Date
                detailsCash[29] = dt.Rows[0][0].ToString(); //FileID
                detailsCash[30] = "NULL"; //Status
                detailsCash[31] = "NULL"; //Ord_Inv_ID
                detailsCash[32] = "NULL"; //Cash_Ops_ID
                detailsCash[33] = "NULL"; //F1_MIS
                detailsCash[34] = "NULL"; //F2_MIS
                detailsCash[35] = "NULL"; //F3_MIS
                detailsCash[36] = "NULL"; //F4_MIS
                detailsCash[37] = "NULL"; //Trade_OPs_Remarks
                detailsCash[38] = "NULL"; //LoginID_TradeOps               
                detailsCash[39] = "NULL"; //LoginID_CashOps
                detailsCash[40] = "NULL"; //F7_MIS
                detailsCash[41] = "NULL"; //TradeopsFileID


                DataTable dtInsert = service.Insert_InvoiceDetails("Save", detailsCash);
                // clserr.LogEntry("Exiting UpdateInvoiceRecords Function", false);



            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message + "Form1", "UpdateInvoiceRecordsNew", "");
                dtunsucc.Rows.Add(dtSuccRow);

            }
        }

        private void UpdateInvoiceCancelRecords(ISheet excelSheet, int jj, int SRNo, string Reason)
        {
            try
            {
                //clserr.LogEntry("Entering UpdateInvoiceCancelRecords Function", false);
                DataTable dt = service.getDetails("Get_MaxFileDesc", "", "", "", "", "", "", "");
                string[] detailsCash = new string[25];


                detailsCash[0] = "0";   //Invoice_ID
                detailsCash[1] = SRNo.ToString(); //Sr_No
                detailsCash[2] = excelSheet.GetRow(jj).GetCell(1).StringCellValue; //Invoice_Number
                detailsCash[3] = excelSheet.GetRow(jj).GetCell(2).StringCellValue; //Invoice_Amount
                detailsCash[4] = System.DateTime.Now.Date.ToString("dd/MMM/yyyy"); //Requested_Date
                detailsCash[5] = "NULL";   //Updated_Date
                detailsCash[6] = "NULL";   //DO_number
                detailsCash[7] = "NULL";   //Order_ID
                detailsCash[8] = "0";   //Cancelled_Flag
                detailsCash[9] = Reason;   //Reason
                detailsCash[10] = "0";   //F2_MIS
                detailsCash[11] = "0";   //DO_IN_Flag
                detailsCash[12] = "NULL";   //Authorize_Flag
                detailsCash[13] = "NULL";   //Deleted_By
                detailsCash[14] = "NULL";   //Deleted_On
                detailsCash[15] = "NULL";   //Authorized_By
                detailsCash[16] = "NULL";   //Authorized_On

                DataTable dtInsert = service.Insert_InvoiceCancel_DescDetails("Save", detailsCash);
                //clserr.LogEntry("Exiting UpdateInvoiceCancelRecords Function", false);



            }
            catch (Exception ex) { clserr.WriteErrorToTxtFile(ex.Message + "Form1", "UpdateInvoiceCancelRecords", ""); }
        }
        public Boolean isCompleteFileAvailable(string szFilePath)
        {
            Boolean isCompleteFileAvailableR;
            try
            {
                //clserr.LogEntry("Entered isCompleteFileAvailable Function" + " isCompleteFileAvailable" + "   MAIN", false);
                string strtmp;
                // File fsObj = new File;
                object obOpenFile;
                strtmp = szFilePath.Trim();
                isCompleteFileAvailableR = false;
                if (File.Exists(strtmp) == true)
                {
                    obOpenFile = File.ReadAllText(szFilePath, Encoding.Default);
                    isCompleteFileAvailableR = true;
                }

            }
            catch (Exception ex)
            {
                //' if file is still locked by another application then its still in process of copying               
                isCompleteFileAvailableR = false;
                clserr.WriteErrorToTxtFile(ex.Message + "Form1", "isCompleteFileAvailable", "");
                clserr.LogEntry("Exiting function isCompleteFileAvailable" + " isCompleteFileAvailable" + "   MAIN", false);
            }
            return isCompleteFileAvailableR;
        }
        public void FolderCreate()
        {
            //sttg.InputFilePath = Application.StartupPath + "Input";
            //sttg.OutputFilePath = Application.StartupPath + "Output";
            //sttg.BackupFilePath = Application.StartupPath + "BackupFile";
            //sttg.NonConvertedFile = Application.StartupPath + "NonConvertedFile";
            //sttg.ErrorLog = Application.StartupPath + "Error";
            //sttg.AuditLog = Application.StartupPath + "Audit";

            try
            {

                if (sttg.InputFilePath.ToString().Trim().Length > 0)
                {
                    if (Directory.Exists(sttg.InputFilePath) == false)
                    {
                        Directory.CreateDirectory(sttg.InputFilePath);
                    }
                }
                else
                {
                    clserr.LogEntry("InputFile path is blank", false);
                }
            }
            catch (Exception ex) { clserr.Handle_Error(ex, "Form", "InputFile path"); }

            try
            {
                if (sttg.OutputFilePath.ToString().Trim().Length > 0)
                {
                    if (Directory.Exists(sttg.OutputFilePath) == false)
                    {
                        Directory.CreateDirectory(sttg.OutputFilePath);
                    }
                }
                else
                {
                    clserr.LogEntry("OutputFile path is blank", false);
                }
            }
            catch (Exception ex) { clserr.Handle_Error(ex, "Form", "OutputFile path"); }

            try
            {
                if (sttg.BackupFilePath.ToString().Trim().Length > 0)
                {
                    if (Directory.Exists(sttg.BackupFilePath) == false)
                    {
                        Directory.CreateDirectory(sttg.BackupFilePath);
                    }
                }
                else
                {
                    clserr.LogEntry("BackupFile path is blank", false);
                }
            }
            catch (Exception ex) { clserr.Handle_Error(ex, "Form", "BackupFile path"); }

            try
            {
                if (sttg.NonConvertedFile.ToString().Trim().Length > 0)
                {
                    if (Directory.Exists(sttg.NonConvertedFile) == false)
                    {
                        Directory.CreateDirectory(sttg.NonConvertedFile);
                    }
                }
                else
                {
                    clserr.LogEntry("NonConvertedFile path is blank", false);
                }
            }
            catch (Exception ex) { clserr.Handle_Error(ex, "Form", "NonConvertedFile path"); }

            try
            {
                if (sttg.ErrorLog.ToString().Trim().Length > 0)
                {
                    if (Directory.Exists(sttg.ErrorLog) == false)
                    {
                        Directory.CreateDirectory(sttg.ErrorLog);
                    }
                }
                else
                {
                    clserr.LogEntry("ErrorLog path is blank", false);
                }
            }
            catch (Exception ex) { clserr.Handle_Error(ex, "Form", "ErrorLog path"); }

            try
            {
                if (sttg.AuditLog.ToString().Trim().Length > 0)
                {
                    if (Directory.Exists(sttg.AuditLog) == false)
                    {
                        Directory.CreateDirectory(sttg.AuditLog);
                    }
                }
                else
                {
                    clserr.LogEntry("AuditLog path is blank", false);
                }
            }
            catch (Exception ex) { clserr.Handle_Error(ex, "Form", "AuditLog path"); }

            try
            {
                if (sttg.BlankExcel.ToString().Trim().Length > 0)
                {
                    if (Directory.Exists(sttg.BlankExcel) == false)
                    {
                        Directory.CreateDirectory(sttg.BlankExcel);
                    }
                }
                else
                {
                    clserr.LogEntry("Blank Excel path is blank", false);
                }
            }
            catch (Exception ex) { clserr.Handle_Error(ex, "Form", "Blank Excel path"); }
            try
            {
                if (sttg.Document_Release.ToString().Trim().Length > 0)
                {
                    if (Directory.Exists(sttg.Document_Release) == false)
                    {
                        Directory.CreateDirectory(sttg.Document_Release);
                    }
                }
                else
                {
                    clserr.LogEntry("Document Release path is blank", false);
                }
            }
            catch (Exception ex) { clserr.Handle_Error(ex, "Form", "Document Release path"); }
            try
            {
                if (sttg.Payment_Rejection.ToString().Trim().Length > 0)
                {
                    if (Directory.Exists(sttg.Payment_Rejection) == false)
                    {
                        Directory.CreateDirectory(sttg.Payment_Rejection);
                    }
                }
                else
                {
                    clserr.LogEntry("Payment Rejection path is blank", false);
                }
            }
            catch (Exception ex) { clserr.Handle_Error(ex, "Form", "Payment Rejection path"); }
            try
            {
                //create  payment rejection folder as per date
                if (sttg.Payment_Rejection_Date.ToString().Trim().Length > 0)
                {
                    if (Directory.Exists((sttg.Payment_Rejection_Date + (System.DateTime.Now.Date.ToString("ddMMyyyy")))) == false)
                    {
                        Directory.CreateDirectory((sttg.Payment_Rejection_Date + (System.DateTime.Now.Date.ToString("ddMMyyyy"))));
                    }
                }
                else
                {
                    clserr.LogEntry("Payment Rejection path is blank", false);
                }
            }
            catch (Exception ex) { clserr.Handle_Error(ex, "Form", "Payment Rejection path"); }

        }

        public void DeleteFilesAndFolders()
        {
            try
            {
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();
                //File delete  from temp folder
                if (Directory.Exists(Application.StartupPath + "Temp"))
                {
                    System.IO.DirectoryInfo di = new DirectoryInfo(Application.StartupPath + "Temp");
                    foreach (FileInfo file in di.GetFiles())
                    {
                        file.Delete();
                    }
                }
                //File delete  from MIS_Temp folder
                if (Directory.Exists(Application.StartupPath + "MIS_Temp"))
                {
                    System.IO.DirectoryInfo diMIS = new DirectoryInfo(Application.StartupPath + "MIS_Temp");
                    foreach (FileInfo file in diMIS.GetFiles())
                    {
                        file.Delete();
                    }
                }
                if (Directory.Exists(Application.StartupPath + "Order"))
                {
                    //File delete  from Order folder
                    System.IO.DirectoryInfo diOrder = new DirectoryInfo(Application.StartupPath + "Order");
                    foreach (FileInfo file in diOrder.GetFiles())
                    {
                        file.Delete();
                    }
                }
                //File delete  from OLD_INV folder
                if (Directory.Exists(Application.StartupPath + "OLD_INV"))
                {
                    System.IO.DirectoryInfo diOLD_INV = new DirectoryInfo(Application.StartupPath + "OLD_INV");
                    foreach (FileInfo file in diOLD_INV.GetFiles())
                    {
                        file.Delete();
                    }
                }
                if (Directory.Exists(Application.StartupPath + "ORD_DEL"))
                {
                    //File delete  from ORD_DEL folder
                    System.IO.DirectoryInfo diORD_DEL = new DirectoryInfo(Application.StartupPath + "ORD_DEL");
                    foreach (FileInfo file in diORD_DEL.GetFiles())
                    {
                        file.Delete();
                    }
                }
                //File delete  from Invoice_Conf folder
                if (Directory.Exists(Application.StartupPath + "Invoice_Conf"))
                {
                    System.IO.DirectoryInfo diInvoice_Conf = new DirectoryInfo(Application.StartupPath + "Invoice_Conf");
                    foreach (FileInfo file in diInvoice_Conf.GetFiles())
                    {
                        file.Delete();
                    }
                }
            }
            catch (Exception ex) { clserr.Handle_Error(ex, "Form1", "DeleteFilesAndFolders"); }
        }

        private void AddColumnToTable(DataTable pDt, int pCols)
        {
            if (pDt is null)
            {
                pDt = new DataTable("Input");
            }
            if (pDt.Columns.Count < pCols)
            {
                pDt.Columns.Add(new DataColumn("Column_" + pDt.Columns.Count));
                AddColumnToTable(pDt, pCols);
            }
        }


        private void ClearArray(string[] ArrRow)
        {
            try
            {
                for (int i = 0; i < ArrRow.Length; i++)
                {
                    ArrRow[i] = "";
                }
            }
            catch (Exception ex) { clserr.Handle_Error(ex, "Form", "ClearArray"); }

        }



        public DataTable GetDatatable_Text(string StrFilePath)

        {
            // Path to the comma-delimited text file
            string filePath = StrFilePath;
            // Create a new DataTable
            DataTable dataTable = new DataTable();
            try
            {
                // Define your custom headers
                string[] customHeaders = new string[] { "Invoice_Number", "Invoice_Amount", "Currency", "Vehical_ID", "DueDate", "Dealer_Name", "Dealer_Address1", "Dealer_City", "Transporter_Name", "Transport_Number", "Transport_Date", "Dealer_Code", "Transporter_Code", "Dealer_Address2", "Dealer_Address3", "Dealer_Address4" }; // Adjust as needed
                                                                                                                                                                                                                                                                                                                                          // Add custom headers to the DataTable as columns
                foreach (var header in customHeaders)
                {
                    dataTable.Columns.Add(header);
                }
                // Initialize the CsvReader
                using (var reader = new StreamReader(filePath))

                using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
                {

                    Delimiter = ",",
                    //Quote = '',
                    BadDataFound = null, // Ignore bad data
                    MissingFieldFound = null, // Ignore missing fields
                    HasHeaderRecord = false // Indicate that the CSV file does not have a header row
                }))

                {
                    //// Read the header row
                    //if (csv.Read())
                    //{
                    //    csv.ReadHeader();
                    //    foreach (var header in csv.HeaderRecord)
                    //    {
                    //        dataTable.Columns.Add(header);
                    //    }
                    //}
                    string failinvoices = "";
                    // Read the data rows
                    while (csv.Read())
                    {
                        var row = dataTable.NewRow();
                        string str = csv.Parser.RawRecord;
                        var strarray = str.Split(",");
                        int rowadd = 0;
                        try
                        {
                            for (int i = 0; i < strarray.Length; i++)
                            {
                                if (strarray[i].Length == 0)
                                    row[rowadd] = strarray[i];
                                else if (strarray[i].Substring(0, 1) == "0")
                                    row[rowadd] = strarray[i];
                                else if (strarray[i].Substring(0, 1) == "\"")
                                {
                                    strarray[i] = strarray[i].Replace("\r\n", "");

                                    if (strarray[i].Substring(strarray[i].Length - 1, 1) == "\"")
                                        row[rowadd] = strarray[i].Replace("\"", "");
                                    else
                                    {
                                        //ICell cell = rowHeader.CreateCell(C);
                                        row[rowadd] = strarray[i].Replace("\"", "");
                                        i = i + 1;
                                        for (int k = i; k < strarray.Length; k++)
                                        {
                                            string ss = strarray[k].Substring(strarray[k].Length - 1, 1);
                                            strarray[k] = strarray[k].Replace("\n", "\"");
                                            if (strarray[k].Substring(strarray[k].Length - 1, 1) == "\"")
                                            {
                                                row[rowadd] = (row[rowadd] + "," + strarray[k].Replace("\"", "").Replace("\n", ""));
                                                i = k;
                                                break;
                                                //return "";
                                            }
                                            else
                                            {
                                                row[rowadd] = (row[rowadd] + "," + strarray[k].Replace("\"", "").Replace("\n", ""));
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    row[rowadd] = strarray[i];
                                }
                                rowadd = rowadd + 1;
                            }
                            row[row.ItemArray.Length - 1] = row[row.ItemArray.Length - 1].ToString().Replace("\r\n", "").Replace("\n", "");
                            dataTable.Rows.Add(row);
                        }
                        catch(Exception ex)
                        {
                            failinvoices = failinvoices + "\n" + strarray[0].ToString();
                        }
                    }
                    if (failinvoices != "")
                    { WritefAILiNVOICES(failinvoices, "", ""); }
                }               
            }

            catch (Exception ex)
            {
                clserr.Handle_Error(ex, "Form", "GetDatatable_Text");
            }
            //string bool1 = ChagneFromCSVToExcel_INV(filePath);
            return dataTable.Copy();
            
        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            Main();
        }
    }
}
public class DataService
{
    public static string _connectionString;

    public DataService(string connectionString)
    {
        _connectionString = connectionString;
    }

    // Methods for data access
}