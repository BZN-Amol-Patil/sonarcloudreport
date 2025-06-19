using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using _001TN0173.Entities;
using Microsoft.EntityFrameworkCore;
using System.Linq;
using Microsoft.Extensions.Configuration;
using System.Data.OleDb;
using System.IO;
using System.Data;
using _001TN0173.Shared;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Net.Mail;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
namespace _001TN0173
{
    class FunctionsCl
    {
        Clsbase objBaseClass = new Clsbase();
        private DatabaseContext db = new DatabaseContext();
        ClsErrorLog clserr = new ClsErrorLog();
        Settings sttg = new Settings();
        ReadWriteAppSettings readWriteAppSettings = new ReadWriteAppSettings();
        StreamWriter ts_inGlobal;
        public int Rows_NEFT = 0, Rows_FT = 0, Rows_RTGS = 0;
        public void RectifyUpdateError()
        {
            try
            {
                //clserr.LogEntry("Entered RectifyUpdateError Function", false);
                DataTable dt = service.getDetails("GetOrder_CashOpsDetails", "", "", "", "", "", "", "");
                //clserr.LogEntry("Test1", false);
                if (dt.Rows.Count > 0)
                {
                    //clserr.LogEntry("Test2", false);
                    for (int irow = 0; irow < dt.Rows.Count; irow++)
                    {
                        //clserr.LogEntry("Test3", false);
                        DataTable dtCash = service.UpdateDetails("Update_OrderINV_Details", dt.Rows[irow]["Cash_Ops_ID"].ToString(), dt.Rows[irow]["DO_Number"].ToString(),
                                           dt.Rows[irow]["UTR_No"].ToString(), dt.Rows[irow]["Order_ID"].ToString(), "", "", "");
                    }
                    clserr.LogEntry("Payment File uploaded and updated with status -Payment Received in Order & Invoice Table", false);
                    DataTable dtOrdInv = service.getDetails("GetCashOps_INVDetails", "", "", "", "", "", "", "");   //Method to update details 
                    //clserr.LogEntry("Test5", false);
                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message + "", "RectifyUpdateError", "");
                clserr.LogEntry("Exiting RectifyUpdateError Function", false);
            }
        }
        public void MISReportMails(string ExcelMISReportMails)
        {
            //clserr.LogEntry("Test8", false);
            Boolean ChkRecord_Flag, Eod_flag = false, Phy_flag = false, CancelDoInFlag = false;
            ChkRecord_Flag = true;
            int i, i_time, test;
            sttg = readWriteAppSettings.ReadGetSectionAppSettings();
            i = Convert.ToInt32(sttg.Sleep_Time_in_Mint.Substring(sttg.Sleep_Time_in_Mint.Length - 1, 1));
            i_time = Convert.ToInt32(i);

            test = i_time * 60000;
            //clserr.LogEntry("Entered MISReportMails Function", false);
            DataTable dtUpdate = new DataTable();
            try
            {
                DataTable dtR = service.getDetails("Get_LoginRC", "", "", "", "", "", "", "");
                //var inv = db.Set<RectifyUpdateEr>().FromSqlRaw("Select * from Login where Record_Chk_Flag=1 ").ToList();
                if (dtR.Rows.Count > 0)
                {
                    ChkRecord_Flag = true;
                }
                else
                {
                    ChkRecord_Flag = false;
                }

                FCC_User(ExcelMISReportMails);
                MatchPaymentFiles(); // avoid cmduploadclick method from vb6. it combined into matchpayment file

                if (ChkRecord_Flag == false)
                {
                    DataTable dtE = service.getDetails("Get_EODReport", "", "", "", "", "", "", "");
                    if (dtE.Rows.Count > 0)
                        Eod_flag = true;
                    else
                        Eod_flag = false;

                    DataTable dtP = service.getDetails("Get_PhyReports", "", "", "", "", "", "", "");
                    if (dtP.Rows.Count > 0)
                        Phy_flag = true;
                    else
                        Phy_flag = false;
                    DataTable dtC = service.getDetails("Get_cancelDOReports", "", "", "", "", "", "", "");
                    if (dtC.Rows.Count > 0)
                        CancelDoInFlag = true;
                    else
                        CancelDoInFlag = false;
                    dtUpdate = service.UpdateDetails("Update_ReportFlagDetails", "", "", "", "", "", "", "");

                    DataTable dtF = service.getDetails("Get_FileMail", "", "", "", "", "", "", "");
                    if (dtF.Rows.Count > 0)
                    {
                        for (int k = 0; k < dtF.Rows.Count; k++)
                        {
                            //1.First Check if "File_Mail_Time" isNull or not if it's null then Call FillF1 / F2F3 / F4 Function update date and Time In DBase
                            //2.If it's not null then check date if "Database date" and "Current Date" if both are "Not equal" then Call FillF1 / F2F3 / F4 Function update date and Time In DBase
                            //3.if it is equal then check Time Diffrence then Call FillF1/F2F3/F4 Function update date and Time In DBase

                            if (dtF.Rows[k]["File_Mail_Name"].ToString().ToUpper() == "F1" && Eod_flag == true)//   'The utility should give out the file in every 2 hours and send to Trade OPS
                            {
                                FillF1();   //'payment file uploaded
                                FillEODACC();// 'Payment file is uploaded it is update the invoice table
                                dtUpdate = service.UpdateDetails("Update_FilemailDetails", "F1", "", "", "", "", "", "");
                            }
                            else if (dtF.Rows[k]["File_Mail_Name"].ToString().ToUpper() == "F2") //End of day MIS should be sent to MSIL containing the list of Invoices for which status is “PHYSICAL INV REC” for all the data that has been updated on the same day.
                            {
                                if (Phy_flag == true)
                                {
                                    FillF2(); //Physical inv mail is send and updated.
                                    dtUpdate = service.UpdateDetails("Update_PhyReport", "", "", "", "", "", "", "");
                                }

                                if (CancelDoInFlag == true)
                                {
                                    CancelReport();
                                    UnsucessCancelReport();
                                    Cancel_DO_Report();
                                    UnsucCancel_DO_Report();
                                    Cancel_DO_Invoice_Report();
                                    UnCancel_DO_Invoice_Report();
                                }

                                if (Eod_flag == true)
                                {
                                    OtherReports(); //check if payment/order is not recieved but invoice is physc is done.
                                    dtUpdate = service.UpdateDetails("Update_FilemailDetails", "F2", "", "", "", "", "", "");
                                }
                            }
                            else if (dtF.Rows[k]["File_Mail_Name"].ToString().ToUpper() == "F3") //'MIS for Payment Credited to MSIL - HOURLY basis.
                            {
                                if (dtF.Rows[k]["File_Mail_Time"].ToString() == null || dtF.Rows[k]["File_Mail_date"].ToString() == null)
                                {
                                    DataTable dts = service.getDetails("Get_StockyardDetails", "", "", "", "", "", "", "");
                                    for (int j = 0; j < dts.Rows.Count; j++)
                                    {
                                        clserr.LogEntry("MIS For " + dts.Rows[0]["Name"].ToString() + " Location", false);
                                        FillF3(dts.Rows[j]["Code"].ToString(), dts.Rows[j]["Name"].ToString(), dts.Rows[j]["DO_Number"].ToString(), dts.Rows[j]["Email"].ToString());
                                        clserr.LogEntry("Sleep Time (in mintues) = " + sttg.Sleep_Time_in_Mint, false);
                                        Thread.Sleep(test);
                                    }

                                    dtUpdate = service.UpdateDetails("Update_FilemailDetails", "F3", "", "", "", "", "", "");
                                }
                                else
                                {
                                    if (dtF.Rows[k]["File_Mail_date"].ToString().Replace("-", "/").Replace(".", "/") == System.DateTime.Now.ToString("dd/MM/yyyy").Replace("-", "/").Replace(".", "/"))
                                    {
                                        DateTime time_File = DateTime.ParseExact(dtF.Rows[k]["File_Mail_Time"].ToString(), "HH:mm", CultureInfo.CurrentCulture);
                                        DateTime time_current = DateTime.ParseExact(System.DateTime.Now.ToString("HH:mm"), "HH:mm", CultureInfo.CurrentCulture);

                                        TimeSpan timeResult = (time_current - time_File).Duration();
                                        if (timeResult.TotalMinutes >= 30)
                                        {
                                            DataTable dts = service.getDetails("Get_StockyardDetails", "", "", "", "", "", "", ""); // get all stock yards
                                            for (int j = 0; j < dts.Rows.Count; j++)
                                            {
                                                clserr.LogEntry("MSIL - HALF HOURLY basis MIS For " + dts.Rows[0]["Name"].ToString() + " Location", false);
                                                FillF3(dts.Rows[j]["Code"].ToString(), dts.Rows[j]["Name"].ToString(), dts.Rows[j]["DO_Number"].ToString(), dts.Rows[j]["Email"].ToString());

                                                clserr.LogEntry("Sleep time for next Paycon to generate(in mintues) = " + sttg.Sleep_Time_in_Mint, false);
                                                Thread.Sleep(test);
                                                dtUpdate = service.UpdateDetails("Update_FilemailDetails", "F3", "", "", "", "", "", "");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        DataTable dts = service.getDetails("Get_StockyardDetails", "", "", "", "", "", "", "");
                                        for (int j = 0; j < dts.Rows.Count; j++)
                                        {
                                            clserr.LogEntry("MSIL - HALF HOURLY basis MIS For " + dts.Rows[0]["Name"].ToString() + " Location", false);
                                            FillF3(dts.Rows[j]["Code"].ToString(), dts.Rows[j]["Name"].ToString(), dts.Rows[j]["DO_Number"].ToString(), dts.Rows[j]["Email"].ToString());

                                            clserr.LogEntry("Sleep time for next Paycon to generate(in mintues) = " + sttg.Sleep_Time_in_Mint, false);
                                            Thread.Sleep(test);
                                        }
                                        dtUpdate = service.UpdateDetails("Update_FilemailDetails", "F3", "", "", "", "", "", "");
                                    }
                                }
                            }
                            else if (dtF.Rows[k]["File_Mail_Name"].ToString().ToUpper() == "F4") //End of day consolidated MIS Format of consolidated MIS will be same as that of Intra-day report format.

                            {
                                if (Eod_flag == true)
                                {
                                    FillF4();// 'consolidate paycon file is send at EOD
                                    dtUpdate = service.UpdateDetails("Update_FilemailDetails", "F4", "", "", "", "", "", "");
                                }
                            }
                            else if (dtF.Rows[k]["File_Mail_Name"].ToString().ToUpper() == "F6")// 'End of day consolidated MIS Format of consolidated MIS will be same as that of Intra-day report format.

                            {
                                //If Not fso.FolderExists(App.Path & "\Invoice_Conf") Then fso.CreateFolder(App.Path & "\Invoice_Conf")
                                if (Eod_flag == true)
                                {
                                    SendInvoiceConfirmation();// 'invoice recived confirmation mail
                                    dtUpdate = service.UpdateDetails("Update_FilemailDetails", "F6", "", "", "", "", "", "");
                                }
                                if (Directory.Exists("\\Invoice_Conf"))
                                    Directory.Delete("\\Invoice_Conf");
                            }
                        }
                    }
                    dtUpdate = service.UpdateDetails("Update_AllReportFlagDetails", "", "", "", "", "", "", "");
                    //clserr.LogEntry("Exiting function MISReportMails", false);

                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message + "", "MISReportMails", "");
            }
            //clserr.LogEntry("Test9", false);
        }
        private void FCC_User(string ExcelMISReportMails)
        {
            try
            {
                DataTable dtR = service.getDetails("Get_IRFileDetails", "", "", "", "", "", "", "");

                if (dtR.Rows.Count > 0)
                {
                    for (int k = 0; k < dtR.Rows.Count; k++)
                    {
                        FCC(dtR.Rows[k]["FileID"].ToString(), ExcelMISReportMails);
                        DataTable dtF = service.UpdateDetails("Update_FileDesc", dtR.Rows[k]["FileID"].ToString(), "1", "", "", "", "", "");
                    }
                    clserr.LogEntry("FCC File Is updated in DB", false);
                }

                //clserr.LogEntry("Exiting function FCC_User", false);
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message + "", "FCC_User", "");
            }

        }
        public void FCC(string IRFile_ID, string ExcelMISReportMails)
        {
            try
            {
                DataTable dtR = service.getDetails("Get_RectifyUpdateEr", IRFile_ID, "", "", "", "", "", "");
                // var inv = db.Set<RectifyUpdateEr>().FromSqlRaw("(SELECT COUNT(*) AS Inv_Count, Dealer_Name, Dealer_Code, IMEX_DEAL_NUMBER, StepDate, SUM(CAST(Invoice_Amount AS FLOAT)) AS Inv_Amt From Invoice WHERE  TradeopsFileID = '" + IRFile_ID + "' and(StepDate IS NOT NULL) and   F7_MIS = 0 GROUP BY StepDate, Dealer_Name, Dealer_Code, IMEX_DEAL_NUMBER, StepDate Union select count(*) as Inv_Count, 'TOTAL' as Dealer_Name, '' as Dealer_Code, '' as IMEX_DEAL_NUMBER, StepDate, SUM(CAST(Invoice_Amount AS FLOAT)) AS Inv_Amt From Invoice WHERE TradeopsFileID = '" + IRFile_ID + "' and(StepDate IS NOT NULL) and   F7_MIS = 0 group by stepdate) order by stepdate ,Dealer_Code desc").ToList();
                if (dtR.Rows.Count > 0)
                {
                    if (Directory.Exists(Directory.GetCurrentDirectory() + "\\FCC"))
                    {
                    }
                    else
                    {
                        Directory.CreateDirectory(Directory.GetCurrentDirectory() + "\\FCC");
                    }

                    DateTime dt = DateTime.Now;
                    string FCC_Date = dt.ToString("ddMMyyyyHHmmss");

                    if (Directory.Exists(Directory.GetCurrentDirectory() + "\\FCC\\" + FCC_Date))
                    {
                    }
                    else
                    {
                        Directory.CreateDirectory(Directory.GetCurrentDirectory() + "\\FCC\\" + FCC_Date);
                    }
                    CreateExcel(objBaseClass.FCC_FileSheet, Directory.GetCurrentDirectory() + "\\FCC\\" + FCC_Date + "\\" + FCC_Date + ".xlsx");
                    //File.Copy(ExcelMISReportMails + "\\Blank_MISReportMails.xlsx", Directory.GetCurrentDirectory() + "\\FCC\\" + FCC_Date + "\\" + FCC_Date + ".xlsx", true);

                    Int32 Row_Cnt = 0;
                    double TotalInv_Amt = 0;

                    using (var fs = new FileStream(Directory.GetCurrentDirectory() + "\\FCC\\" + FCC_Date + "\\" + FCC_Date + ".xlsx", FileMode.Open, FileAccess.Read))
                    {
                        XSSFWorkbook workbook = new XSSFWorkbook(fs);
                        ISheet excelSheet = workbook.GetSheet("Sheet1");
                        fs.Close();
                        try
                        {
                            for (int i = 0; i < dtR.Rows.Count; i++)
                            {

                                TotalInv_Amt = TotalInv_Amt + Convert.ToDouble(dtR.Rows[i]["Inv_Amt"].ToString());
                                if (dtR.Rows[i]["Dealer_Name"].ToString().ToUpper() == "TOTAL")
                                {
                                    Row_Cnt = Row_Cnt + 1;
                                }
                                else
                                {
                                    Row_Cnt = Row_Cnt + 1;
                                }

                                IRow row = excelSheet.CreateRow(i + 1);
                                row.CreateCell(0).SetCellValue(i + 1);
                                row.CreateCell(1).SetCellValue(dtR.Rows[i]["Dealer_Code"].ToString());
                                row.CreateCell(2).SetCellValue(dtR.Rows[i]["Dealer_Name"].ToString());
                                row.CreateCell(3).SetCellValue(dtR.Rows[i]["IMEX_DEAL_NUMBER"].ToString());
                                row.CreateCell(4).SetCellValue(dtR.Rows[i]["Inv_Count"].ToString());
                                row.CreateCell(5).SetCellValue(dtR.Rows[i]["Inv_Amt"].ToString());
                                row.CreateCell(6).SetCellValue(dtR.Rows[i]["stepdate"].ToString());

                                //workbook.Write(fs,true);
                            }
                            using (var file = new FileStream(Directory.GetCurrentDirectory() + "\\FCC\\" + FCC_Date + "\\" + FCC_Date + ".xlsx", FileMode.Create, FileAccess.Write))
                            {
                                workbook.Write(file, true);
                                file.Close();
                                file.Dispose();
                            }
                            SendEmail(sttg.FCC_EMail, "", "", Directory.GetCurrentDirectory() + "\\FCC\\" + FCC_Date + "\\" + FCC_Date + ".xlsx", "FCC", "FYI");

                        }
                        catch (Exception ex)
                        {

                            clserr.WriteErrorToTxtFile(ex.Message + "", "FCC", "File ID" + IRFile_ID);
                            workbook.Close();
                            workbook.Dispose();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message + "", "FCC", "File ID" + IRFile_ID);
            }
        }
        public void MatchPaymentFiles()
        {
            try
            {
                string CheckType = "", OrderID = "";
                //clserr.LogEntry("Entered MatchPaymentFiles Function", false);
                string filename_RTGS = sttg.Payment_Rejection_Date + System.DateTime.Now.Date.ToString("ddMMyyyy") + "\\Payment_Rejection_RTGS_" + System.DateTime.Now.Date.ToString("ddMMyyyyHHmmss") + ".xlsx";
                string filename_NEFT = sttg.Payment_Rejection_Date + System.DateTime.Now.Date.ToString("ddMMyyyy") + "\\Payment_Rejection_NEFT_" + System.DateTime.Now.Date.ToString("ddMMyyyyHHmmss") + ".xlsx";
                string filename_FT = sttg.Payment_Rejection_Date + System.DateTime.Now.Date.ToString("ddMMyyyy") + "\\Payment_Rejection_FT_" + System.DateTime.Now.Date.ToString("ddMMyyyyHHmmss") + ".xlsx";
                string dirname = sttg.Payment_Rejection_Date + System.DateTime.Now.Date.ToString("ddMMyyyy");
                if (Directory.Exists(dirname) == false)
                    Directory.CreateDirectory(dirname);
                //Chk Type
                //A=Order Number/Virtual Account No. Match and Order Amt/Trans Amt match.
                //B=Order Number/Virtual Account No. Match and Order Amt/Trans Amt Not match.
                //C=Order Number/Virtual Account No. Not Found.
                //E=Order Number/Virtual Account No. Match and Payment Type Invalid
                //F=Order Number/Virtual Account No. Match and UTR No Duplicate
                DataTable dt = service.getDetails("Get_PaymantStatus", "", "", "", "", "", "", "");
                if (dt.Rows.Count > 0)
                {
                    // create excels for NEFT, FT, RTGS
                    CreateExcel(objBaseClass.Create_NEFT_FileSheet, filename_NEFT);
                    CreateExcel(objBaseClass.Create_FT_FileSheet, filename_FT);
                    CreateExcel(objBaseClass.Create_RTGS_FileSheet, filename_RTGS);

                    for (int i = 0; i <= dt.Rows.Count - 1; i++)
                    {
                        CheckType = "";
                        DataTable dtP = service.getDetails("Get_PaymantStatusASper_VirtuAcc", dt.Rows[i]["Virtual_Account"].ToString(), "", "", "", "", "", "");
                        if (dtP.Rows.Count > 0)
                            CheckType = "D";
                        if (CheckType != "D")
                        {
                            DataTable dtOrdr = service.getDetails("Get_Order_Desc", dt.Rows[i]["Virtual_Account"].ToString(), "", "", "", "", "", "");
                            if (dtOrdr.Rows.Count > 0)
                            {
                                OrderID = dtOrdr.Rows[0]["Order_ID"].ToString();
                                if (Convert.ToDouble(dt.Rows[i]["Transaction_Amount"].ToString()) == Convert.ToDouble(dtOrdr.Rows[0]["Order_amount"].ToString()) && (dt.Rows[i]["CashOps_FileType"].ToString() == "NEFT" || dt.Rows[i]["CashOps_FileType"].ToString() == "RTGS" || dt.Rows[i]["CashOps_FileType"].ToString() == "FT" || dt.Rows[i]["CashOps_FileType"].ToString() == "FUND TRANS"))
                                {
                                    DataTable dtUTRDpt = service.getDetails("GetUTRDetailsForPaymentUpload", dt.Rows[i]["UTR_No"].ToString(), dt.Rows[i]["Virtual_Account"].ToString(), dt.Rows[i]["Transaction_Amount"].ToString(), "", "", "", "");
                                    if (dtUTRDpt.Rows.Count > 0)
                                    {
                                        CheckType = "F";    ////New Added F condition 03-04-2023
                                        clserr.LogEntry("Found UTR No, Do Number, Amount and Payment status Duplicate", false);
                                    }
                                    else
                                        CheckType = "A";
                                }
                                else if (Convert.ToDouble(dt.Rows[i]["Transaction_Amount"].ToString()) == Convert.ToDouble(dtOrdr.Rows[0]["Order_amount"].ToString()) && (dt.Rows[i]["CashOps_FileType"].ToString() != "NEFT" || dt.Rows[i]["CashOps_FileType"].ToString() != "RTGS" || dt.Rows[i]["CashOps_FileType"].ToString() != "FT" || dt.Rows[i]["CashOps_FileType"].ToString() != "FUND TRANS"))
                                    CheckType = "E";      ////New Added
                                else
                                    CheckType = "B";
                            }
                            else
                            {
                                CheckType = "C";
                            }
                        }
                        // Update data as per checkType at backend
                        DataTable dtUpdate = service.UpdateDetails("Update_OrderDesc_AsPerCheckType", CheckType, dt.Rows[i]["Cash_Ops_ID"].ToString(), dt.Rows[i]["Virtual_Account"].ToString(), dt.Rows[i]["UTR_No"].ToString(), OrderID, "", "");
                        if (CheckType == "A")
                        {// Confirmation Order
                            clserr.LogEntry("DRC Generated process start for UTR No :" + dt.Rows[i]["UTR_No"].ToString() + " and DO NO :" + dt.Rows[i]["Virtual_Account"].ToString(), false);
                            WriteConfirmationOrder(dt.Rows[i]["UTR_No"].ToString(), dt.Rows[i]["Virtual_Account"].ToString());
                        }
                        // Rejection Files
                        if (CheckType == "D" || CheckType == "B")
                        {
                            // As per File Type Generate Excel
                            Rejection_Excels(dt.Rows[i]["CashOps_FileType"].ToString().ToUpper(), dt.Rows[i]["Cash_Ops_ID"].ToString(), filename_NEFT, filename_RTGS, filename_FT);
                        }
                    }


                }

                if (Rows_NEFT > 0)
                {
                    SendEmail(sttg.Payment_Rejection_EMail, "", "", filename_NEFT, "Rejection NEFT", "FYI");
                    clserr.LogEntry("Rejection NEFT File mail sent to-" + sttg.Payment_Rejection_EMail, false);
                    Rows_NEFT = 0;
                }
                else
                {
                    File.Delete(filename_NEFT);
                }
                if (Rows_FT > 0)
                {
                    SendEmail(sttg.Payment_Rejection_EMail, "", "", filename_FT, "Rejection FT", "FYI");
                    clserr.LogEntry("Rejection FT File mail sent to-" + sttg.Payment_Rejection_EMail, false);
                    Rows_FT = 0;
                }
                else
                {
                    File.Delete(filename_FT);
                }
                if (Rows_RTGS > 0)
                {
                    SendEmail(sttg.Payment_Rejection_EMail, "", "", filename_RTGS, "Rejection RTGS", "FYI");
                    clserr.LogEntry("Rejection RTGS File mail sent to - " + sttg.Payment_Rejection_EMail, false);
                    Rows_RTGS = 0;
                }
                else
                {
                    File.Delete(filename_RTGS);
                }
                //clserr.LogEntry("Exiting function MatchPaymentFiles", false);
            }
            catch (Exception ex)
            { clserr.WriteErrorToTxtFile(ex.Message + "", "MatchPaymentFiles", ""); }
        }
        public void CreateExcel(string[] ExclColumns, string FileName)
        {

            //clserr.LogEntry("CreateExcel 1", false);
            using (var fs = new FileStream(FileName, FileMode.Create, FileAccess.Write))
            {
                XSSFWorkbook workbook = new XSSFWorkbook();
                ISheet excelSheet = workbook.CreateSheet("Sheet1");
                IRow rowHeader = excelSheet.CreateRow(0);
                try
                {
                    for (int cc = 0; cc < ExclColumns.Length; cc++)
                    {
                        rowHeader.CreateCell(cc).SetCellValue(ExclColumns[cc].ToString());
                    }
                    workbook.Write(fs, true);
                    fs.Close();
                    workbook.Dispose();
                    fs.Dispose();
                    //clserr.LogEntry("CreateExcel 2", false);
                }
                catch (Exception ex)
                {
                    clserr.WriteErrorToTxtFile(ex.Message + "", "CreateExcel", FileName);
                    workbook.Write(fs, true);
                    fs.Close();
                    workbook.Dispose();
                    fs.Dispose();
                }
            }

        }
        public void ReadFormatMail(out StringBuilder Header, out StringBuilder Body, out StringBuilder Footer)
        {
            Header = new StringBuilder();
            Body = new StringBuilder();
            Footer = new StringBuilder();
            try
            {
                StreamReader objStrmReader = new StreamReader(sttg.Document_Release_Template, Encoding.UTF8);
                string Read_Line = "";

                string Header_Line_Temp = "";
                while (!objStrmReader.EndOfStream)
                {
                ReadL: Read_Line = objStrmReader.ReadLine();
                    if ("<Header>".ToUpper() == Read_Line.ToUpper())
                    {
                    ReadNextHeader: Header_Line_Temp = objStrmReader.ReadLine();
                        if (Header_Line_Temp != null)
                        {
                            if ("</Header>".ToUpper() == Header_Line_Temp.ToUpper())
                                goto ReadL;
                            else if ("</Header>".ToUpper() != Header_Line_Temp.ToUpper())
                            {
                                Header.Append("<br/>");
                                Header.Append(Header_Line_Temp);
                                goto ReadNextHeader;
                                //Header_Line = Header_Line + Header_Line.Append("<br/>")+ Header_Line_Temp;
                            }
                        }
                    }
                    if ("<Body>".ToUpper() == Read_Line.ToUpper())
                    {
                    ReadNextBody: Header_Line_Temp = objStrmReader.ReadLine();
                        if (Header_Line_Temp != null)
                        {
                            if ("</Body>".ToUpper() == Header_Line_Temp.ToUpper())
                                goto ReadL;
                            else if ("</Body>".ToUpper() != Header_Line_Temp.ToUpper())
                            {
                                if (Header_Line_Temp.Contains("Goods Receipt Note �"))
                                    Header_Line_Temp = Header_Line_Temp.Replace("�", "-");
                                if (Header_Line_Temp.Contains("�We hereby certify "))
                                    Header_Line_Temp = Header_Line_Temp.Replace("�We hereby certify ", "\"We hereby certify ");
                                if (Header_Line_Temp.Contains("this invoice�"))
                                    Header_Line_Temp = Header_Line_Temp.Replace("this invoice�", "this invoice\"");
                                Body.Append("<br/>");
                                Body.Append(Header_Line_Temp);


                                goto ReadNextBody;
                                //Header_Line = Header_Line + Header_Line.Append("<br/>")+ Header_Line_Temp;
                            }
                        }
                    }
                    if ("<Footer>".ToUpper() == Read_Line.ToUpper())
                    {
                    ReadNextFooter: Header_Line_Temp = objStrmReader.ReadLine();
                        if (Header_Line_Temp != null)
                        {
                            if ("</Footer>".ToUpper() == Header_Line_Temp.ToUpper())
                                goto ReadL;
                            else if ("</Footer>".ToUpper() != Header_Line_Temp.ToUpper())
                            {
                                Footer.Append("<br/>");
                                Footer.Append(Header_Line_Temp);
                                goto ReadNextFooter;
                                //Header_Line = Header_Line + Header_Line.Append("<br/>")+ Header_Line_Temp;
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { clserr.WriteErrorToTxtFile(ex.Message + "", "ReadFormatMail", ""); }
        }
        public void WriteConfirmationOrder(string UtrNo, string VANo)
        {
            try
            {
                int Cnt_Transporter_Name1, LInt, CnEmail = 0;
                StringBuilder Body_Line = new StringBuilder();
                StringBuilder Header_Line = new StringBuilder();
                StringBuilder Footer_Line = new StringBuilder();

                string Dealer_Code = "", Transporter_Name = "", Payment_Amt = "", Transporter_Name1 = "", STp = "", STp_Full = "";
                string Middle_line1, Middle_line2, Header_name = "", vehicleNo = "", FinancerName = "", OrderNo = "", Invoice_date = "", InvoiceNo = "", AmtRecvd = "", AmtRecvddate = "", Goodlorry = "";
                string E_Body = "", AllEmail = "", ToEmail = "", BCCEmails = "", strrecord = "", strtemplate;
                string[] Spl_Transporter_Name1, Split_String, EmailArr, ArrFooterLine;
                Boolean Flag_Transporter_Name1;
                double Inv_Amt;

                sttg = readWriteAppSettings.ReadGetSectionAppSettings();
                ReadFormatMail(out Header_Line, out Body_Line, out Footer_Line);
                clserr.LogEntry("Entered WriteConfirmationOrder Function", false);
                if (File.Exists(sttg.Document_Release_Template) == false)
                {
                    clserr.LogEntry("Document Release Confirmation Template Not Found", false);
                    return;
                }
                else
                {

                }
                DataTable dtInv = service.getDetails("Get_InvoiceSum", UtrNo, VANo, "", "", "", "", "");
                if (dtInv.Rows.Count > 0)
                {
                    Dealer_Code = dtInv.Rows[0]["Dealer_Name"].ToString() + "-" + dtInv.Rows[0]["Dealer_Code"].ToString();
                    Transporter_Name = dtInv.Rows[0]["Transporter_Name"].ToString();
                    Payment_Amt = dtInv.Rows[0]["PaymentAmt"].ToString();
                }
                DataTable dtDT = service.getDetails("Get_DealerTransporter", UtrNo, VANo, "", "", "", "", "");
                if (dtDT.Rows.Count > 0)
                {
                    for (int i = 0; i < dtDT.Rows.Count; i++)
                    {
                        Flag_Transporter_Name1 = false;
                        Spl_Transporter_Name1 = Transporter_Name1.Split(",");
                        for (Cnt_Transporter_Name1 = 0; Cnt_Transporter_Name1 < Spl_Transporter_Name1.Length; Cnt_Transporter_Name1++)
                        {
                            if (dtDT.Rows[i]["Transporter_Name"].ToString().ToUpper() == Spl_Transporter_Name1[Cnt_Transporter_Name1].ToUpper())
                                Flag_Transporter_Name1 = true;
                        }
                        if (Flag_Transporter_Name1 == false)
                        {
                            if (Transporter_Name1 == "")
                                Transporter_Name1 = dtDT.Rows[i]["Transporter_Name"].ToString();
                            else

                                Transporter_Name1 = Transporter_Name1 + "," + dtDT.Rows[i]["Transporter_Name"].ToString();
                        }
                    }

                }
                //Transporter_Name1 = Transporter_Name1.Substring(1, Transporter_Name1.Length - 1);

                Split_String = Transporter_Name1.Split();
                STp = "";
                STp_Full = "";
                for (LInt = 0; LInt < Transporter_Name1.Length; LInt++)
                {
                    if (STp.Length == 100)
                    {
                        if (STp_Full == "")
                            STp_Full = STp;
                        else
                            STp_Full = STp_Full + System.Environment.NewLine + STp;
                        STp = "";
                    }
                    STp = STp + Transporter_Name1.Substring(LInt, 1);

                }
                if (STp_Full == "")
                    STp_Full = STp;
                else
                    STp_Full = STp_Full + System.Environment.NewLine + STp;
                Transporter_Name1 = STp_Full;

                Body_Line = Body_Line.Replace("<Transporter Name from Inv table>", "       " + Transporter_Name1);
                Body_Line = Body_Line.Replace("<Dealer Name from Inv table>", Dealer_Code);
                //ArrBodyLine = Body_Line.ToString().Split(System.Environment.NewLine);

                clserr.LogEntry("Start - Generate of DRC File of UTR No: " + UtrNo + " and DO No: " + VANo, false);
                Inv_Amt = 0;

                DataTable dtDetails = service.getDetails("Get_DealerDetails", UtrNo, VANo, "", "", "", "", "");
                if (dtDetails.Rows.Count > 0)
                {
                    for (int k = 0; k < dtDetails.Rows.Count; k++)
                    {
                        //Table records
                        CultureInfo provider = CultureInfo.InvariantCulture;
                        OrderNo = dtDetails.Rows[k]["DO_number"].ToString();
                        InvoiceNo = "" + dtDetails.Rows[k]["Invoice_Number"].ToString();
                        Invoice_date = dtDetails.Rows[k]["Transport_Date"].ToString(); //"" + DateTime.ParseExact(dtDetails.Rows[k]["Transport_Date"].ToString(), "000000", provider);
                        AmtRecvd = "" + dtDetails.Rows[k]["Order_Inv_Amount"].ToString();
                        string date = Convert.ToDateTime(dtDetails.Rows[k]["payment_received_date"]).ToString("dd/MM/yyyy").Replace('-', '/');
                        AmtRecvddate = Convert.ToDateTime(dtDetails.Rows[k]["payment_received_date"]).ToString("dd/MM/yyyy").Replace('-', '/');
                        Goodlorry = "" + dtDetails.Rows[k]["Transporter_Code"].ToString();
                        vehicleNo = "" + dtDetails.Rows[k]["Vehical_ID"].ToString();
                        FinancerName = "" + dtDetails.Rows[k]["FNCR_name"].ToString();
                        AllEmail = dtDetails.Rows[k]["Email_IDs"].ToString().Replace(((Char)34).ToString(), "");  //or use "\""
                        AllEmail = AllEmail.Replace(",", ";");
                        EmailArr = AllEmail.Split(";");
                        BCCEmails = "";
                        ToEmail = "";

                        for (CnEmail = 0; CnEmail < EmailArr.Length; CnEmail++)
                        {
                            if (EmailArr[CnEmail].ToString().ToUpper().Contains("@maruti.co.in".ToUpper())) //Means "@maruti.co.in" found then BCC.MSIL Ids in BCC
                                BCCEmails = BCCEmails + ";" + EmailArr[CnEmail];
                            else
                                ToEmail = ToEmail + ";" + EmailArr[CnEmail];
                        }
                        strrecord += "<tr><td width:auto>" + OrderNo + "</td><td width:auto>" + InvoiceNo + "</td><td width:auto> " + Invoice_date + "</td><td width:auto>" + AmtRecvd + "</td><td width:auto>" + AmtRecvddate + "</td><td width:auto>" + Goodlorry + "</td><td width:auto>" + vehicleNo + "</td><td  width: 500px; >" + FinancerName + "</td> </tr>";
                        // strrecord = OrderNo + "".PadRight(5) + InvoiceNo + "".PadRight(5) + Invoice_date + "".PadRight(10) + AmtRecvd + "".PadRight(10) + AmtRecvddate + "".PadRight(10) + Goodlorry + "".PadRight(10) + vehicleNo + "".PadRight(7) + FinancerName + System.Environment.NewLine + strrecord;
                        Inv_Amt = Inv_Amt + Convert.ToDouble("" + dtDetails.Rows[k]["Order_Inv_Amount"].ToString());
                    }
                    //Table header
                    Middle_line1 = "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------";
                    //Header_name = "".PadRight(3) + "Order No" + "".PadRight(5) + "Invoice No" + "".PadRight(5) + "Invoice Date" + "".PadRight(5) + "Amt Received(INR)" + "".PadRight(5) + "Amt Received Date" + "".PadRight(5) + "Good/Lorry Receipt no:" + "".PadRight(5) + "Vehicle Identification no" + "".PadRight(5) + "Financier Name";
                    //Header_name = " <table border=" + 1 + " cellpadding=" + 0 + " cellspacing=" + 0 + " width = " + 1000 + " alignment><tr bgcolor='#4da6ff' ><td width:auto><b>Order No</b></td> <td width:auto> <b> Invoice No</b> </td> <td width:auto> <b> Invoice Date</b> </td><td width:auto> <b> Amt Received(INR)</b> </td><td width:auto> <b> Amt Received Date</b> </td><td width:auto> <b> Good/Lorry Receipt no:</b> </td><td width:auto> <b> Vehicle Identification no</b> </td><td width:auto> <b>Financier Name</b> </td></tr>";
                    Header_name = " <table border=" + 1 + " cellpadding=" + 0 + " cellspacing=" + 0 + " width = " + 1000 + " alignment><tr><td width:auto><b>Order No</b></td> <td width:auto> <b> Invoice No</b> </td> <td width:auto> <b> Invoice Date</b> </td><td width:auto> <b> Amt Received(INR)</b> </td><td width:auto> <b> Amt Received Date</b> </td><td width:auto> <b> Good/Lorry Receipt no:</b> </td><td width:auto> <b> Vehicle Identification no</b> </td><td width:auto> <b>Financier Name</b> </td></tr>";
                    strrecord += "</table>";
                    Middle_line2 = "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------";
                    //strtemplate = Header_Line + Body_Line + Middle_line1 + "\r\n" + Header_name + "\r\n" + Middle_line2 + "\r\n" + strrecord + "\r\n" + "\n" + Footer_Line + "\r\n";

                    if (ToEmail != "") { ToEmail = ToEmail.Substring(1, ToEmail.Length - 1); }
                    if (BCCEmails != "") { BCCEmails = BCCEmails.Substring(1, BCCEmails.Length - 1); }


                    Footer_Line = Footer_Line.Replace("<dealer name invoice table>", Dealer_Code);
                    Footer_Line = Footer_Line.Replace("<Payment amount received>", Inv_Amt.ToString("0.00"));
                    ArrFooterLine = Footer_Line.ToString().Split("\r\n");
                    string str = Header_Line.ToString();
                    StringBuilder sb = new StringBuilder();
                    sb.Append("<b>\t\t\t\t" + Header_Line + "</b>");
                    sb.Append("<br/>");
                    sb.Append(Body_Line);
                    sb.Append("<br/>");
                    sb.Append(Header_name);
                    sb.Append("<br/>");
                    //sb.Append(Middle_line2);
                    sb.Append("<br/>");
                    sb.Append(strrecord);
                    sb.Append("<br/>");
                    sb.Append("<br/>");
                    sb.Append(Footer_Line);
                    sb.Append("<br/>");
                    strtemplate = sb.ToString();
                    // strtemplate = Header_Line + "\r\n" + Body_Line + Middle_line1 + "\r\n" + Header_name + "\r\n" + Middle_line2 + "\r\n" + strrecord + "\r\n" + "\n" + Footer_Line + "\r\n";
                    E_Body = sb.ToString();
                    clserr.LogEntry("End - Generate of DRC File", false);
                    DataTable dt_Stockyard = service.getDetails("Get_StockyardDetailsAsPerDO", VANo.ToUpper().Trim().Substring(0, 5), VANo.ToUpper().Trim().Substring(0, 6), "", "", "", "", "");
                    if (dt_Stockyard.Rows.Count > 0)
                    {
                        if (dt_Stockyard.Rows[0]["Email"].ToString().Trim() != "")
                        {
                            SendEmailDRC(ToEmail + ";" + dt_Stockyard.Rows[0]["Email"].ToString(), "", BCCEmails, "", "Document Release Confirmation(" + VANo + ")", E_Body);
                            clserr.LogEntry("Mail send for " + dt_Stockyard.Rows[0]["Name"].ToString() + " ToEmail=" + ToEmail + " and " + dt_Stockyard.Rows[0]["Email"].ToString(), false);
                        }
                        else
                        {
                            SendEmailDRC(ToEmail, "", BCCEmails, "", "Document Release Confirmation(" + VANo + ")", E_Body);
                            SendEmailDRC(dt_Stockyard.Rows[0]["Email"].ToString(), "", "", "", "Document Release Confirmation(" + VANo + ")", E_Body);
                            clserr.LogEntry("Mail send for " + dt_Stockyard.Rows[0]["Name"].ToString() + "DRC_EMAIL & ToEmail= " + dt_Stockyard.Rows[0]["Email"].ToString() + " and " + ToEmail, false);
                        }
                    }

                    //clserr.LogEntry("Exiting WriteConfirmationOrder Function", false);

                }
                else
                { clserr.LogEntry("Get_DealerDetails of UTR No " + UtrNo + " and DO No: " + VANo + " not found.", false); }
            }

            catch (Exception ex) { clserr.WriteErrorToTxtFile(ex.Message + "", "WriteConfirmationOrder", "UtrNo: " + UtrNo + "VANo: " + VANo); }
        }
        public void WriteConfirmationOrderOld(string UtrNo, string VANo)
        {
            try
            {
                int cnt = 0, Cnt_Transporter_Name1, LInt, CnEmail = 0;
                StringBuilder Body_Line = new StringBuilder();
                StringBuilder Header_Line = new StringBuilder();
                StringBuilder Footer_Line = new StringBuilder();
                string Header_Line_Temp = "", Body_Line_Temp = "", Footer_Line_Temp = "";
                string Dealer_Code = "", Transporter_Name = "", Payment_Amt = "", Transporter_Name1 = "", STp = "", STp_Full = "";
                string Middle_line1, Middle_line2, Header_name = "", vehicleNo = "", FinancerName = "", OrderNo = "", Invoice_date = "", InvoiceNo = "", AmtRecvd = "", AmtRecvddate = "", Goodlorry = "";
                string E_Body = "", AllEmail = "", ToEmail = "", BCCEmails = "", strrecord = "", strtemplate;
                string[] Spl_Transporter_Name1, Split_String, ArrBodyLine, EmailArr, ArrFooterLine;
                Boolean Flag_Transporter_Name1;
                double Inv_Amt;

                sttg = readWriteAppSettings.ReadGetSectionAppSettings();

                clserr.LogEntry("Entered WriteConfirmationOrder Function", false);
                if (File.Exists(sttg.Document_Release_Template) == false)
                {
                    clserr.LogEntry("Document Release Confirmation Template Not Found", false);
                    return;
                }
                else
                {
                   
                    StreamReader objStrmReader = new StreamReader(sttg.Document_Release_Template);
                    string Read_Line = "";
                    while (!objStrmReader.EndOfStream)
                    {
                        Read_Line = objStrmReader.ReadLine();
                        if ("<Header>".ToUpper() == Read_Line.ToUpper())
                        {
                            for (cnt = 1; cnt <= 1000; cnt++)
                            {
                                Header_Line_Temp = objStrmReader.ReadLine();
                                if (Header_Line_Temp != null)
                                {
                                    if ("</Header>".ToUpper() == Header_Line_Temp.ToUpper())
                                    {
                                    }
                                    else
                                    {
                                        Header_Line.Append("<br/>");
                                        Header_Line.Append(Header_Line_Temp);
                                        //Header_Line = Header_Line + Header_Line.Append("<br/>")+ Header_Line_Temp;
                                    }
                                }
                            }
                        }
                        else if ("<Body>".ToUpper() == Read_Line.ToUpper())
                        {
                            for (cnt = 1; cnt <= 1000; cnt++)
                            {
                                Body_Line_Temp = objStrmReader.ReadLine();
                                if ("</Body>".ToUpper() == Body_Line_Temp.ToUpper())
                                {
                                }
                                else
                                {
                                    Body_Line.Append("<br/>");
                                    Body_Line.Append(Body_Line_Temp);
                                    Body_Line.Append("<br/>");
                                    //Body_Line = Body_Line + Body_Line_Temp + System.Environment.NewLine;
                                }
                            }
                        }
                        else if ("<Footer>".ToUpper() == Read_Line.ToUpper())
                        {
                            for (cnt = 1; cnt <= 1000; cnt++)
                            {
                                Footer_Line_Temp = objStrmReader.ReadLine();
                                if ("</Body>".ToUpper() == Footer_Line_Temp.ToUpper())
                                {
                                }
                                else
                                {
                                    Footer_Line.Append(Footer_Line_Temp);
                                    Footer_Line.Append("<br/>");
                                    //Footer_Line = Footer_Line + Footer_Line_Temp + System.Environment.NewLine;
                                }
                            }
                        }
                    }
                }
                DataTable dtInv = service.getDetails("Get_InvoiceSum", UtrNo, VANo, "", "", "", "", "");
                if (dtInv.Rows.Count > 0)
                {
                    Dealer_Code = dtInv.Rows[0]["Dealer_Name"].ToString() + "-" + dtInv.Rows[0]["Dealer_Code"].ToString();
                    Transporter_Name = dtInv.Rows[0]["Transporter_Name"].ToString();
                    Payment_Amt = dtInv.Rows[0]["PaymentAmt"].ToString();
                }
                DataTable dtDT = service.getDetails("Get_DealerTransporter", UtrNo, VANo, "", "", "", "", "");
                if (dtDT.Rows.Count > 0)
                {
                    Flag_Transporter_Name1 = false;
                    Spl_Transporter_Name1 = Transporter_Name1.Split(",");
                    for (Cnt_Transporter_Name1 = 0; Cnt_Transporter_Name1 < Spl_Transporter_Name1.Length; Cnt_Transporter_Name1++)
                    {
                        if (dtDT.Rows[0]["Transporter_Name"].ToString().ToUpper() == Spl_Transporter_Name1[Cnt_Transporter_Name1].ToUpper())
                            Flag_Transporter_Name1 = true;
                    }
                    if (Flag_Transporter_Name1 == false)
                        Transporter_Name1 = Transporter_Name1 + "," + dtDT.Rows[0]["Transporter_Name"].ToString();
                }
                Transporter_Name1 = Transporter_Name1.Substring(1, Transporter_Name1.Length - 1);

                Split_String = Transporter_Name1.Split();
                STp = "";
                STp_Full = "";
                for (LInt = 1; LInt < Transporter_Name1.Length; LInt++)
                {
                    if (STp.Length == 100)
                    {
                        if (STp_Full == "")
                            STp_Full = STp;
                        else
                            STp_Full = STp_Full + System.Environment.NewLine + STp;
                        STp = "";
                        STp = STp + Transporter_Name1.Substring(LInt, 1);
                    }
                }
                if (STp_Full == "")
                    STp_Full = STp;
                else
                    STp_Full = STp_Full + System.Environment.NewLine + STp;
                Transporter_Name1 = STp_Full;

                Body_Line = Body_Line.Replace("<Transporter Name from Inv table>", Transporter_Name1);
                Body_Line = Body_Line.Replace("<Dealer Name from Inv table>", Dealer_Code);
                ArrBodyLine = Body_Line.ToString().Split(System.Environment.NewLine);

                clserr.LogEntry("Start - Generate of DRC File", false);
                Inv_Amt = 0;

                DataTable dtDetails = service.getDetails("Get_DealerDetails", UtrNo, VANo, "", "", "", "", "");
                if (dtDetails.Rows.Count > 0)
                {
                    for (int k = 0; k < dtDetails.Rows.Count - 1; k++)
                    {
                        //Table records
                        CultureInfo provider = CultureInfo.InvariantCulture;
                        OrderNo = dtDetails.Rows[k]["DO_number"].ToString();
                        InvoiceNo = "" + dtDetails.Rows[k]["Invoice_Number"].ToString();
                        Invoice_date = dtDetails.Rows[k]["Transport_Date"].ToString(); //"" + DateTime.ParseExact(dtDetails.Rows[k]["Transport_Date"].ToString(), "000000", provider);
                        AmtRecvd = "" + dtDetails.Rows[k]["Order_Inv_Amount"].ToString();
                        string date = Convert.ToDateTime(dtDetails.Rows[k]["payment_received_date"]).ToString("dd/MM/yyyy").Replace('-', '/');
                        AmtRecvddate = Convert.ToDateTime(dtDetails.Rows[k]["payment_received_date"]).ToString("dd/MM/yyyy").Replace('-', '/');
                        Goodlorry = "" + dtDetails.Rows[k]["Transporter_Code"].ToString();
                        vehicleNo = "" + dtDetails.Rows[k]["Vehical_ID"].ToString();
                        FinancerName = "" + dtDetails.Rows[k]["FNCR_name"].ToString();
                        AllEmail = dtDetails.Rows[k]["Email_IDs"].ToString().Replace(((Char)34).ToString(), "");  //or use "\""
                        AllEmail = AllEmail.Replace(",", ";");
                        EmailArr = AllEmail.Split(";");
                        BCCEmails = "";
                        ToEmail = "";

                        for (CnEmail = 0; CnEmail < EmailArr.Length; CnEmail++)
                        {
                            if (EmailArr[CnEmail].ToString().ToUpper().Contains("@maruti.co.in".ToUpper())) //Means "@maruti.co.in" found then BCC.MSIL Ids in BCC
                                BCCEmails = BCCEmails + ";" + EmailArr[CnEmail];
                            else
                                ToEmail = ToEmail + ";" + EmailArr[CnEmail];
                        }

                        strrecord = OrderNo + "".PadRight(5) + InvoiceNo + "".PadRight(5) + Invoice_date + "".PadRight(10) + AmtRecvd + "".PadRight(10) + AmtRecvddate + "".PadRight(10) + Goodlorry + "".PadRight(10) + vehicleNo + "".PadRight(7) + FinancerName + System.Environment.NewLine + strrecord;
                        Inv_Amt = Inv_Amt + Convert.ToDouble("" + dtDetails.Rows[k]["Order_Inv_Amount"].ToString());
                        StringBuilder sb_Details = new StringBuilder();
                        sb_Details.Append(OrderNo);
                        sb_Details.Append("\t");
                        sb_Details.Append(InvoiceNo);
                        sb_Details.Append("\t");
                        sb_Details.Append(Invoice_date);
                        sb_Details.Append("\t");
                        sb_Details.Append(AmtRecvd);
                        sb_Details.Append("\t");
                        sb_Details.Append(AmtRecvddate);
                        sb_Details.Append("\t");
                        sb_Details.Append(Goodlorry);
                        sb_Details.Append("\t");
                        sb_Details.Append(vehicleNo);
                        sb_Details.Append("\t");
                        sb_Details.Append(FinancerName);
                        sb_Details.Append("\t");
                        strrecord = sb_Details.ToString();

                        //Table header
                        Middle_line1 = "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------";
                        Header_name = "".PadRight(3) + "Order No" + "".PadRight(5) + "Invoice No" + "".PadRight(5) + "Invoice Date" + "".PadRight(5) + "Amt Received(INR)" + "".PadRight(5) + "Amt Received Date" + "".PadRight(5) + "Good/Lorry Receipt no:" + "".PadRight(5) + "Vehicle Identification no" + "".PadRight(5) + "Financier Name";
                        StringBuilder sb_HeaderName = new StringBuilder();
                        sb_HeaderName.Append("Order No");
                        sb_HeaderName.Append("\t");
                        sb_HeaderName.Append("Invoice No");
                        sb_HeaderName.Append("\t");
                        sb_HeaderName.Append("Invoice Date");
                        sb_HeaderName.Append("\t");
                        sb_HeaderName.Append("Amt Received(INR)");
                        sb_HeaderName.Append("\t");
                        sb_HeaderName.Append("Amt Received Date");
                        sb_HeaderName.Append("\t");
                        sb_HeaderName.Append("Good/Lorry Receipt no:");
                        sb_HeaderName.Append("\t");
                        sb_HeaderName.Append("Vehicle Identification no");
                        sb_HeaderName.Append("\t");
                        sb_HeaderName.Append("Financier Name");
                        sb_HeaderName.Append("\t");
                        Header_name = sb_HeaderName.ToString();
                        Middle_line2 = "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------";
                        //strtemplate = Header_Line + Body_Line + Middle_line1 + "\r\n" + Header_name + "\r\n" + Middle_line2 + "\r\n" + strrecord + "\r\n" + "\n" + Footer_Line + "\r\n";
                        StringBuilder sb1 = new StringBuilder();
                        sb1.Append(Header_Line);
                        sb1.Append("<br/>");
                        sb1.Append(Body_Line + Middle_line1);
                        sb1.Append("<br/>");
                        sb1.Append(Header_name);
                        sb1.Append("<br/>");
                        sb1.Append(Middle_line2);
                        sb1.Append("<br/>");
                        sb1.Append(strrecord);
                        sb1.Append("<br/>");
                        sb1.Append("<br/>");
                        sb1.Append(Footer_Line);
                        sb1.Append("<br/>");
                        strtemplate = sb1.ToString();
                        if (ToEmail != "") { ToEmail = ToEmail.Substring(1, ToEmail.Length - 1); }
                        if (BCCEmails != "") { BCCEmails = BCCEmails.Substring(1, BCCEmails.Length - 1); }


                        Footer_Line = Footer_Line.Replace("<dealer name invoice table>", Dealer_Code);
                        Footer_Line = Footer_Line.Replace("<Payment amount received>", Inv_Amt.ToString("0.00"));
                        ArrFooterLine = Footer_Line.ToString().Split("\r\n");
                        string str = Header_Line.ToString();
                        StringBuilder sb = new StringBuilder();
                        sb.Append(Header_Line);
                        sb.Append("<br/>");
                        sb.Append(Body_Line + Middle_line1);
                        sb.Append("<br/>");
                        sb.Append(Header_name);
                        sb.Append("<br/>");
                        sb.Append(Middle_line2);
                        sb.Append("<br/>");
                        sb.Append(strrecord);
                        sb.Append("<br/>");
                        sb.Append("<br/>");
                        sb.Append(Footer_Line);
                        sb.Append("<br/>");

                        // strtemplate = Header_Line + "\r\n" + Body_Line + Middle_line1 + "\r\n" + Header_name + "\r\n" + Middle_line2 + "\r\n" + strrecord + "\r\n" + "\n" + Footer_Line + "\r\n";
                        E_Body = sb.ToString();
                        clserr.LogEntry("End - Generate of DRC File", false);
                    }
                    if (((VANo.ToUpper().Trim().Substring(0, 6))) == "HDFCIN")
                    {
                        if (sttg.DRC_EMail.Trim() != "")
                        {
                            SendEmail(ToEmail + ";" + sttg.DRC_EMail, "", BCCEmails, "", "Document Release Confirmation(" + VANo + ")", E_Body);
                            clserr.LogEntry("Mail send ToEmail=" + ToEmail + " and " + sttg.DRC_EMail, false);
                        }
                        else
                        {
                            SendEmail(ToEmail, "", BCCEmails, "", "Document Release Confirmation(" + VANo + ")", E_Body);
                            SendEmail(sttg.DRC_EMail, "", "", "", "Document Release Confirmation(" + VANo + ")", E_Body);
                            clserr.LogEntry("Mail send DRC_EMAIL & ToEmail= " + sttg.DRC_EMail + " and " + ToEmail, false);

                        }
                    }
                    else
                    {
                        if (sttg.DRC_EMail_BNGR.Trim() != "")
                        {
                            SendEmail(sttg.DRC_EMail_BNGR + ";" + ToEmail, "", BCCEmails, "", "Document Release Confirmation(" + VANo + ")", E_Body);
                            clserr.LogEntry("Mail send ToEmail=" + ToEmail + " and " + sttg.DRC_EMail_BNGR, false);
                        }
                        else
                        {
                            SendEmail(ToEmail, "", BCCEmails, "", "Document Release Confirmation(" + VANo + ")", E_Body);
                            SendEmail(sttg.DRC_EMail_BNGR, "", "", "", "Document Release Confirmation(" + VANo + ")", E_Body);
                            clserr.LogEntry("Mail send DRC_EMail_BNGR & ToEmail= " + ToEmail + " and " + sttg.DRC_EMail_BNGR, false);

                        }
                    }
                    clserr.LogEntry("Exiting WriteConfirmationOrder Function", false);

                }

            }

            catch (Exception ex) { clserr.WriteErrorToTxtFile(ex.Message + "", "WriteConfirmationOrder", "UtrNo: " + UtrNo + "VANo: " + VANo); }
        }
        public void SendEmailDRC(string SendTo, string SendCC, string SendBCC, string FileAttach, string subject, string Message)
        {
            try
            {
                // **************************************************************************
                //Advanced Properties, change only if you have a good reason to do so.
                //**************************************************************************
                // .ConnectTimeout = 10                      ' Optional, default = 10
                // .ConnectRetry = 5                         ' Optional, default = 5
                //.MessageTimeout = 60                      ' Optional, default = 60
                //.PersistentSettings = True                ' Optional, default = TRUE
                //Optional, default = 25
                MailMessage mail = new MailMessage();
                Settings sttg = new Settings();
                ReadWriteAppSettings readWriteAppSettings = new ReadWriteAppSettings();
                sttg = readWriteAppSettings.ReadGetSectionAppSettings();
                SmtpClient SmtpServer = new SmtpClient(sttg.SMTP_HOST, Convert.ToInt32(sttg.Port));
                mail.From = new MailAddress(sttg.Email_FromID);

                Attachment attachment;

                string[] ArrSendTo;
                ArrSendTo = SendTo.ToString().Trim().Split(";");
                for (int i = 0; i <= ArrSendTo.Length - 1; i++)
                {
                    if (ArrSendTo[i].ToString() != "")
                        mail.To.Add(new MailAddress(ArrSendTo[i]));
                }
                if (SendCC != "")
                {
                    ArrSendTo = SendCC.ToString().Trim().Split(";");
                    for (int i = 0; i <= ArrSendTo.Length - 1; i++)
                    {
                        if (ArrSendTo[i].ToString() != "")
                            mail.CC.Add(new MailAddress(ArrSendTo[i]));
                    }
                    //mail.CC.Add(SendCC);
                }
                ArrSendTo = SendBCC.ToString().Trim().Split(";");
                if (SendBCC != "")
                {
                    ArrSendTo = SendBCC.ToString().Trim().Split(";");
                    for (int i = 0; i <= ArrSendTo.Length - 1; i++)
                    {
                        if (ArrSendTo[i].ToString() != "")
                            mail.Bcc.Add(new MailAddress(ArrSendTo[i]));
                    }
                    //mail.CC.Add(SendCC);
                }
                //mail.Bcc.Add(SendBCC);
                mail.Subject = subject;
                // string strMailBody;
                mail.IsBodyHtml = true;
                StringBuilder sb = new StringBuilder();

                sb.Append("<br/>");
                sb.Append(Message);
                sb.Append("<br/>");
                sb.Append("<br/>");
                if (FileAttach != "")
                {
                    attachment = new Attachment(FileAttach);
                    mail.Attachments.Add(attachment);
                    sb.Append(" Please find attachment of " + subject);
                }
                mail.Body = sb.ToString();

                SmtpServer.Credentials = new System.Net.NetworkCredential(sttg.Email_FromID.ToString(), sttg.PWD);
                if (mail.To.ToString() == "")
                    clserr.WriteErrorToTxtFile("Sent To is blank for Subject : " + subject , " SendEmail ", FileAttach);
                else
                    SmtpServer.Send(mail);
                mail.Dispose();
                SmtpServer.Dispose();
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();
                //clserr.LogEntry("Exiting WriteConfirmationOrder Function", false);
                clserr.LogEntry("Exiting WriteConfirmationOrder Function. Subject: " + subject, false);
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message + "", "SendEmail", FileAttach);
                clserr.Handle_Error(ex, "SendEmail", "SendEmail");
            }



        }
        public void SendEmail(string SendTo, string SendCC, string SendBCC, string FileAttach, string subject, string Message)
        {
            try
            {
                // **************************************************************************
                //Advanced Properties, change only if you have a good reason to do so.
                //**************************************************************************
                // .ConnectTimeout = 10                      ' Optional, default = 10
                // .ConnectRetry = 5                         ' Optional, default = 5
                //.MessageTimeout = 60                      ' Optional, default = 60
                //.PersistentSettings = True                ' Optional, default = TRUE
                //Optional, default = 25
                MailMessage mail = new MailMessage();
                Settings sttg = new Settings();
                ReadWriteAppSettings readWriteAppSettings = new ReadWriteAppSettings();
                sttg = readWriteAppSettings.ReadGetSectionAppSettings();
                SmtpClient SmtpServer = new SmtpClient(sttg.SMTP_HOST, Convert.ToInt32(sttg.Port));
                mail.From = new MailAddress(sttg.Email_FromID);

                Attachment attachment;

                string[] ArrSendTo;
                ArrSendTo = SendTo.ToString().Trim().Split(";");
                for (int i = 0; i <= ArrSendTo.Length - 1; i++)
                {
                    if (ArrSendTo[i].ToString() != "")
                        mail.To.Add(new MailAddress(ArrSendTo[i]));
                }
                if (SendCC != "")
                {
                    ArrSendTo = SendCC.ToString().Trim().Split(";");
                    for (int i = 0; i <= ArrSendTo.Length - 1; i++)
                    {
                        if (ArrSendTo[i].ToString() != "")
                            mail.CC.Add(new MailAddress(ArrSendTo[i]));
                    }
                    //mail.CC.Add(SendCC);
                }
                ArrSendTo = SendBCC.ToString().Trim().Split(";");
                if (SendBCC != "")
                {
                    ArrSendTo = SendBCC.ToString().Trim().Split(";");
                    for (int i = 0; i <= ArrSendTo.Length - 1; i++)
                    {
                        if (ArrSendTo[i].ToString() != "")
                            mail.Bcc.Add(new MailAddress(ArrSendTo[i]));
                    }
                    //mail.CC.Add(SendCC);
                }

                mail.Subject = subject;
                // string strMailBody;
                mail.IsBodyHtml = true;
                StringBuilder sb = new StringBuilder();
                sb.Append("Respected,");
                sb.Append("<br/>");
                sb.Append("<br/>");
                sb.Append(Message);
                sb.Append("<br/>");
                sb.Append("<br/>");
                if (FileAttach != "")
                {
                    attachment = new Attachment(FileAttach);
                    mail.Attachments.Add(attachment);
                    sb.Append(" Please find attachment of " + subject);
                }
                sb.Append("<br/>");
                sb.Append("<br/>");
                sb.Append("********This is system Generated Mail. Do not reply.********");
                sb.Append("<br/>");
                sb.Append("<br/>");
                sb.Append("Thanks And Regards,");
                sb.Append("<br/>");
                sb.Append("     HDFC Bank     ");
                mail.Body = sb.ToString();
                SmtpServer.Credentials = new System.Net.NetworkCredential(sttg.Email_FromID.ToString(), sttg.PWD);
                if(mail.To.ToString()=="")
                    clserr.WriteErrorToTxtFile("Sent To is blank for Subject : " + subject, " SendEmail ", FileAttach);
                else
                SmtpServer.Send(mail);
                mail.Dispose();
                SmtpServer.Dispose();
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();
                //clserr.LogEntry("Exiting WriteConfirmationOrder Function", false);
                clserr.LogEntry("Exiting SendMail Function.  Subject : " + subject, false);
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message + "", "SendEmail", FileAttach);
                clserr.Handle_Error(ex, "SendEmail", "SendEmail");
            }
        }
        public void Rejection_Excels(string FileType, string Cash_Ops_ID, string NEFT_FileName, string RTGS_FileName, string FT_FileName)
        {
            try
            {
                string dirname = sttg.Payment_Rejection_Date + System.DateTime.Now.Date.ToString("ddMMyyyy");
                if (Directory.Exists(dirname) == false)
                    Directory.CreateDirectory(dirname);
                if (FileType == "NEFT")
                {
                    clserr.LogEntry("Entered function NEFT_Excel", false);

                    using (var fs = new FileStream(NEFT_FileName, FileMode.Open, FileAccess.ReadWrite))
                    {
                        XSSFWorkbook workbook = new XSSFWorkbook(fs);
                        ISheet excelSheet = workbook.GetSheet("Sheet1");
                        try
                        {
                            DataTable DT_NEFT = service.getDetails("Get_NEFTDetails", Cash_Ops_ID, "", "", "", "", "", "");
                            if (DT_NEFT.Rows.Count > 0)
                            {
                                for (int i = 0; i < DT_NEFT.Rows.Count; i++)
                                {
                                    Rows_NEFT = Rows_NEFT + 1;
                                    IRow row = excelSheet.CreateRow(Rows_NEFT);
                                    row.CreateCell(0).SetCellValue(DT_NEFT.Rows[i]["Product"].ToString());
                                    row.CreateCell(1).SetCellValue(DT_NEFT.Rows[i]["Party_Code"].ToString());
                                    row.CreateCell(2).SetCellValue(DT_NEFT.Rows[i]["Party_Name"].ToString());
                                    row.CreateCell(3).SetCellValue(DT_NEFT.Rows[i]["RemittingBank"].ToString());
                                    row.CreateCell(4).SetCellValue(DT_NEFT.Rows[i]["UTR_No"].ToString());
                                    row.CreateCell(5).SetCellValue(DT_NEFT.Rows[i]["Transaction_Amount"].ToString());
                                    row.CreateCell(6).SetCellValue(DT_NEFT.Rows[i]["IFSC_Code"].ToString());
                                    row.CreateCell(7).SetCellValue(DT_NEFT.Rows[i]["Virtual_Account"].ToString());
                                    row.CreateCell(8).SetCellValue(DT_NEFT.Rows[i]["Payment_Status"].ToString());

                                    //row.CreateCell(0).SetCellValue(DT_NEFT.Rows[i]["Releated_Ref_No"].ToString());
                                    //row.CreateCell(1).SetCellValue(DT_NEFT.Rows[i]["Transaction_Amount"].ToString());
                                    //row.CreateCell(2).SetCellValue(System.DateTime.Now.Date.ToString("ddmmyy"));
                                    //row.CreateCell(3).SetCellValue("0599");
                                    //row.CreateCell(4).SetCellValue("11");
                                    //row.CreateCell(5).SetCellValue(DT_NEFT.Rows[i]["Party_Code"].ToString());
                                    //row.CreateCell(6).SetCellValue(DT_NEFT.Rows[i]["Party_Name"].ToString());
                                    //row.CreateCell(7).SetCellValue(DT_NEFT.Rows[i]["IFSC_code"].ToString());
                                    //row.CreateCell(8).SetCellValue(DT_NEFT.Rows[i]["Virtual_Account"].ToString());
                                    //row.CreateCell(9).SetCellValue("1");
                                    //row.CreateCell(10).SetCellValue(DT_NEFT.Rows[i]["Dealer_Account_No"].ToString());
                                    //row.CreateCell(11).SetCellValue(DT_NEFT.Rows[i]["Dealer_Name"].ToString());
                                    //row.CreateCell(12).SetCellValue(DT_NEFT.Rows[i]["Payment_Status"].ToString());
                                    //row.CreateCell(13).SetCellValue("1");
                                    //row.CreateCell(14).SetCellValue("HDFC Bank Ltd.");

                                }
                                var fsNew = new FileStream(NEFT_FileName, FileMode.Create, FileAccess.Write);
                                workbook.Write(fsNew, true);
                                fs.Close();
                                fs.Dispose();
                                fsNew.Close();
                                workbook.Dispose();
                                fsNew.Dispose();


                            }
                        }
                        catch (Exception ex1)
                        {
                            clserr.WriteErrorToTxtFile(ex1.Message + "", "Rejection_Excels", FileType + ": " + NEFT_FileName + RTGS_FileName + FT_FileName);
                            fs.Close();
                            fs.Dispose();
                            workbook.Close();
                            workbook.Dispose();
                        }
                    }

                    clserr.LogEntry("Exiting function NEFT_Excel", false);
                }

                if (FileType == "FT" || FileType == "FUND TRANS")
                {
                    clserr.LogEntry("Entered function FT_Excel", false);

                    using (var fs = new FileStream(FT_FileName, FileMode.Open, FileAccess.ReadWrite))
                    {
                        XSSFWorkbook workbook = new XSSFWorkbook(fs);
                        ISheet excelSheet = workbook.GetSheet("Sheet1");
                        try
                        {
                            DataTable DT_FT = service.getDetails("Get_FTDetails", Cash_Ops_ID, "", "", "", "", "", "");
                            if (DT_FT.Rows.Count > 0)
                            {
                                DataTable DT_FTord = service.getDetails("Get_FTOrderDetails", DT_FT.Rows[0]["Virtual_Account"].ToString(), "", "", "", "", "", "");

                                //for (int i = 0; i < DT_FTord.Rows.Count; i++) ////Change in field on discuss with hdfc on 14-12-2023
                                for (int i = 0; i < DT_FT.Rows.Count; i++)
                                {
                                    Rows_FT = Rows_FT + 1;
                                    IRow row = excelSheet.CreateRow(Rows_FT);
                                    //row.CreateCell(0).SetCellValue(DT_FTord.Rows[i]["Dealer_Code"].ToString() + "," + DT_FTord.Rows[i]["Dealer_Name"].ToString()); ////Change in field on discuss with hdfc on 14-12-2023
                                    row.CreateCell(0).SetCellValue(DT_FT.Rows[i]["Product"].ToString());
                                    row.CreateCell(1).SetCellValue(DT_FT.Rows[i]["Party_Code"].ToString());
                                    row.CreateCell(2).SetCellValue(DT_FT.Rows[i]["Party_Name"].ToString());
                                    row.CreateCell(3).SetCellValue(DT_FT.Rows[i]["RemittingBank"].ToString());
                                    row.CreateCell(4).SetCellValue(DT_FT.Rows[i]["UTR_No"].ToString());
                                    row.CreateCell(5).SetCellValue(DT_FT.Rows[i]["Transaction_Amount"].ToString());
                                    row.CreateCell(6).SetCellValue(DT_FT.Rows[i]["IFSC_Code"].ToString());
                                    row.CreateCell(7).SetCellValue(DT_FT.Rows[i]["Virtual_Account"].ToString());
                                    row.CreateCell(8).SetCellValue(DT_FT.Rows[i]["Payment_Status"].ToString());

                                    // change logic for all rejected reports to set like payment upload file rejection report.
                                    //row.CreateCell(0).SetCellValue(DT_FT.Rows[i]["UTR_No"].ToString());
                                    ////row.CreateCell(1).SetCellValue(DT_FT.Rows[0]["Dealer_Account_No"].ToString());   ////Change in field on discuss with hdfc on 14-12-2023
                                    //row.CreateCell(1).SetCellValue(DT_FT.Rows[0]["Dealer_Account_No"].ToString());
                                    //row.CreateCell(2).SetCellValue(DT_FT.Rows[0]["Transaction_Amount"].ToString());
                                    ////row.CreateCell(3).SetCellValue(DT_FTord.Rows[0]["IMEX_DEAL_NUMBER"].ToString()); ////Change in field on discuss with hdfc on 14-12-2023
                                    //row.CreateCell(3).SetCellValue(DT_FT.Rows[0]["Virtual_Account"].ToString());
                                    //row.CreateCell(4).SetCellValue(DT_FT.Rows[0]["Payment_Status"].ToString());
                                }


                                var fsNew = new FileStream(FT_FileName, FileMode.Create, FileAccess.Write);
                                workbook.Write(fsNew, true);
                                fs.Close();
                                fs.Dispose();
                                fsNew.Close();
                                workbook.Dispose();
                                fsNew.Dispose();

                            }
                        }
                        catch (Exception ex1)
                        {
                            clserr.WriteErrorToTxtFile(ex1.Message + "", "Rejection_Excels", FileType + ": " + NEFT_FileName + RTGS_FileName + FT_FileName);
                            fs.Close();
                            fs.Dispose();
                            workbook.Close();
                            workbook.Dispose();
                        }
                    }

                    clserr.LogEntry("Exiting function FT_Excel", false);
                }

                if (FileType == "RTGS")
                {
                    clserr.LogEntry("Entered function RTGS_Excel", false);

                    using (var fs = new FileStream(RTGS_FileName, FileMode.Open, FileAccess.ReadWrite))
                    {
                        XSSFWorkbook workbook = new XSSFWorkbook(fs);
                        ISheet excelSheet = workbook.GetSheet("Sheet1");
                        try
                        {
                            DataTable DT_RTGS = service.getDetails("Get_RTGSDetails", Cash_Ops_ID, "", "", "", "", "", "");
                            if (DT_RTGS.Rows.Count > 0)
                            {

                                for (int i = 0; i < DT_RTGS.Rows.Count; i++)
                                {
                                    Rows_RTGS = Rows_RTGS + 1;
                                    IRow row = excelSheet.CreateRow(Rows_RTGS);

                                    row.CreateCell(0).SetCellValue(DT_RTGS.Rows[i]["Product"].ToString());
                                    row.CreateCell(1).SetCellValue(DT_RTGS.Rows[i]["Party_Code"].ToString());
                                    row.CreateCell(2).SetCellValue(DT_RTGS.Rows[i]["Party_Name"].ToString());
                                    row.CreateCell(3).SetCellValue(DT_RTGS.Rows[i]["RemittingBank"].ToString());
                                    row.CreateCell(4).SetCellValue(DT_RTGS.Rows[i]["UTR_No"].ToString());
                                    row.CreateCell(5).SetCellValue(DT_RTGS.Rows[i]["Transaction_Amount"].ToString());
                                    row.CreateCell(6).SetCellValue(DT_RTGS.Rows[i]["IFSC_Code"].ToString());
                                    row.CreateCell(7).SetCellValue(DT_RTGS.Rows[i]["Virtual_Account"].ToString());
                                    row.CreateCell(8).SetCellValue(DT_RTGS.Rows[i]["Payment_Status"].ToString());

                                    //row.CreateCell(0).SetCellValue(DT_RTGS.Rows[i]["Releated_Ref_No"].ToString());
                                    //row.CreateCell(1).SetCellValue(DT_RTGS.Rows[i]["IFSC_code"].ToString());
                                    //row.CreateCell(2).SetCellValue("11");
                                    //row.CreateCell(3).SetCellValue("0599");
                                    //row.CreateCell(4).SetCellValue(DT_RTGS.Rows[i]["Virtual_Account"].ToString());
                                    //row.CreateCell(5).SetCellValue(DT_RTGS.Rows[i]["Transaction_Amount"].ToString());
                                    ////row.CreateCell(5).SetCellValue(DT_RTGS.Rows[i]["Party_Code"].ToString());
                                    //row.CreateCell(6).SetCellValue(DT_RTGS.Rows[i]["UTR_No"].ToString() + "-" + DT_RTGS.Rows[i]["Payment_Status"].ToString() + DT_RTGS.Rows[i]["Dealer_Account_No"].ToString() + DT_RTGS.Rows[i]["Dealer_Name"].ToString());


                                }
                                var fsNew = new FileStream(RTGS_FileName, FileMode.Create, FileAccess.Write);
                                workbook.Write(fsNew, true);
                                fs.Close();
                                fs.Dispose();
                                fsNew.Close();
                                workbook.Dispose();
                                fsNew.Dispose();
                            }
                        }
                        catch (Exception ex1)
                        {
                            clserr.WriteErrorToTxtFile(ex1.Message + "", "Rejection_Excels", FileType + ": " + NEFT_FileName + RTGS_FileName + FT_FileName);
                            fs.Close();
                            fs.Dispose();
                            workbook.Close();
                            workbook.Dispose();
                        }
                    }
                    clserr.LogEntry("Exiting function RTGS_Excel", false);
                }

            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message + "", "Rejection_Excels", FileType + ": " + NEFT_FileName + RTGS_FileName + FT_FileName);
            }
        }
        private void FillF1()
        {
            string filename = Application.StartupPath + "\\MIS_Temp" + "\\Payment_Received_" + System.DateTime.Now.Date.ToString("ddMMyyyyHHmmss") + ".xlsx";
            try
            {
                clserr.LogEntry("Entered FillF1 Function", false);
                int cnt = 2;
                DataTable dt = service.getDetails("Get_InvoiceDetails_PayRec", "", "", "", "", "", "", "");
                if (dt.Rows.Count > 0)
                {
                    string paths = Application.StartupPath + "MIS_Temp";
                    if (!Directory.Exists(Application.StartupPath + "MIS_Temp"))
                        Directory.CreateDirectory(Application.StartupPath + "MIS_Temp");
                    using (var fs = new FileStream(filename, FileMode.Create, FileAccess.Write))
                    {
                        XSSFWorkbook workbook = new XSSFWorkbook();
                        try
                        {
                            ISheet excelSheet = workbook.CreateSheet("Sheet1");
                            IRow rowHeader = excelSheet.CreateRow(0);
                            rowHeader.CreateCell(0).SetCellValue("Deal Number");
                            rowHeader.CreateCell(1).SetCellValue("Ccy");
                            rowHeader.CreateCell(2).SetCellValue("Amount");
                            rowHeader.CreateCell(3).SetCellValue("Invoice Number");
                            rowHeader.CreateCell(4).SetCellValue("RTGS number");
                            rowHeader.CreateCell(5).SetCellValue("DATE");

                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                cnt = cnt + 1;
                                IRow row = excelSheet.CreateRow(i + 1);
                                row.CreateCell(0).SetCellValue(dt.Rows[i]["IMEX_DEAL_NUMBER"].ToString());
                                row.CreateCell(1).SetCellValue("INR");
                                row.CreateCell(2).SetCellValue(dt.Rows[i]["Invoice_Amount"].ToString());
                                row.CreateCell(3).SetCellValue(dt.Rows[i]["Invoice_Number"].ToString());
                                row.CreateCell(4).SetCellValue(dt.Rows[i]["Utr_Number"].ToString());
                                row.CreateCell(5).SetCellValue(Convert.ToDateTime(dt.Rows[i]["payment_received_date"].ToString()).ToString("ddmmyy"));
                                DataTable dtu = service.UpdateDetails("Update_InvoiceF", "F1_MIS", dt.Rows[i]["Invoice_Number"].ToString(), "", "", "", "", "");
                                if (Convert.ToInt32(dtu.Rows[0][0]) > 0)
                                { }
                                else
                                    clserr.LogEntry("Record Not Updated for Update_InvoiceF and invoice :" +dt.Rows[i]["Invoice_Number"].ToString() + " FillF1 ", false);
                            }
                            workbook.Write(fs, true);
                            fs.Close();
                            workbook.Close();
                            workbook.Dispose();
                            fs.Dispose();
                            if (dt.Rows.Count > 0)
                            {
                                SendEmail(sttg.PAYREC_TRADE_EMAIL, "", "", filename, "MSIL Payments Received", "FYI");
                                clserr.LogEntry("MSIL Payments Received mail sent to: " + sttg.PAYREC_TRADE_EMAIL + "  FillF1", false);
                                Thread.Sleep(100);
                            }
                        }
                        catch (Exception ex)
                        {
                            clserr.WriteErrorToTxtFile(ex.Message, "Payment Received-F1", filename);
                            fs.Close();
                            workbook.Close();
                            workbook.Dispose();
                        }
                    }
                    System.GC.Collect();
                    System.GC.WaitForPendingFinalizers();
                    if (Directory.Exists(Application.StartupPath + "MIS_Temp"))
                        Directory.Delete(Application.StartupPath + "MIS_Temp", true);
                }
            }

            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message, "Payment Received-F1", filename);
                clserr.WriteErrorToTxtFile("EOD Error : Payment Received-F1. " + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss"), "FillF1", filename);
                clserr.WriteErrorToTxtFile("Exiting function FillF1", "FillF1", filename);
            }

        }
        public string Decrypt(string Name, long Key)
        {
            string DecryptRet = default;
            try
            {
                long v;
                string c1;
                var z = default(string);

                var loopTo = (long)Strings.Len(Name);
                for (v = 1L; v <= loopTo; v++)
                {
                    string str = Name.Substring(1, 1);
                    c1 = Strings.Asc(Strings.Mid(Name, (int)v, 1)).ToString();
                    c1 = Conversions.ToString(Strings.Chr((int)Math.Round(Conversions.ToDouble(c1) - Key)));
                    z = z + c1;
                }
                DecryptRet = z;
            }
            catch (Exception EX)
            {
            }
            return DecryptRet;
        }
        public string Encrypt(string Name, long Key)
        {
            string EncryptRet = default;
            long v;
            string c1;
            var z = default(string);

            var loopTo = (long)Strings.Len(Name);
            for (v = 1L; v <= loopTo; v++)
            {
                c1 = Strings.Asc(Strings.Mid(Name, (int)v, 1)).ToString();
                c1 = Conversions.ToString(Strings.Chr((int)Math.Round(Conversions.ToDouble(c1) + Key)));
                z = z + c1;
            }
            EncryptRet = z;
            return EncryptRet;
        }

        private void FillF2()
        {
            string strTempFolder = Application.StartupPath + "Payconf";
            string FileLastPart = System.DateTime.Now.ToString("ddMMyyyyHHmmss");
            string OpLine = "";
            string OutFileName = strTempFolder + "\\BNGRITKOIN01.PHYINV." + FileLastPart + ".CSV";
            string OutFileName1 = sttg.IntraDayPath + "\\BNGRITKOIN01.PHYINV." + FileLastPart + ".CSV";
            //string OutFileName = strTempFolder + "\\ITKOIN01.PHYINV." + FileLastPart + ".CSV";
            //string OutFileName1 = sttg.IntraDayPath + "\\ITKOIN01.PHYINV." + FileLastPart + ".CSV";
            try
            {
                clserr.LogEntry("Entered FillF2 Function", false);
                if (!Directory.Exists(sttg.IntraDayPath))
                    Directory.CreateDirectory(sttg.IntraDayPath);
                if (!Directory.Exists(strTempFolder))
                    Directory.CreateDirectory(strTempFolder);
                DataTable dt = service.getDetails("Get_InvoiceDetailsF2", "", "", "", "", "", "", "");
                OpLine = "Invoice Number";

                if (dt.Rows.Count > 0)
                {
                    StreamWriter ts_in = new StreamWriter(OutFileName);
                    try
                    {
                        ts_in.WriteLine(OpLine);
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            OpLine = dt.Rows[i]["Invoice_Number"].ToString();
                            ts_in.WriteLine(OpLine);
                            DataTable dtu = service.UpdateDetails("Update_InvoiceF", "F2_MIS", dt.Rows[i]["Invoice_Number"].ToString(), "", "", "", "", "");
                            if (Convert.ToInt32(dtu.Rows[0][0]) > 0)
                            { }
                            else
                                clserr.LogEntry("Record Not Updated for Update_InvoiceF and invoice :" + dt.Rows[i]["Invoice_Number"].ToString() + " FillF2 ", false);

                        }
                        ts_in.Close();
                    }
                    catch (Exception ex)
                    {
                        ts_in.Close();
                        clserr.WriteErrorToTxtFile(ex.Message, "PHYSICAL INVOICE RECEIVED-F2", OutFileName1);
                    }
                }
                clserr.LogEntry("EOD END: PHYSICAL_INV_REC. " + System.DateTime.Now.ToString("dd / MMM / yyyy HH: MM:SS") + "   FillF2", false);

                if (dt.Rows.Count > 0)
                {
                    SendEmail(sttg.PHY_INV, "", "", OutFileName, "PHYSICAL INVOICE RECEIVED", "FYI");
                    clserr.LogEntry("PHYSICAL INVOICE RECEIVED mail sent-" + sttg.PHY_INV + "   PHYSICAL INVOICE RECEIVED", false);
                    System.GC.Collect();
                    System.GC.WaitForPendingFinalizers();
                    File.Move(OutFileName, OutFileName1);
                }

            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message, "PHYSICAL INVOICE RECEIVED-F2", OutFileName1);
                clserr.WriteErrorToTxtFile("EOD Error : Payment Received-F2. " + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss"), "FillF2", OutFileName1);
                clserr.WriteErrorToTxtFile("Exiting function FillF2", "FillF2", OutFileName);
            }

        }
        private void FillF3(string stockYard_code, string StockYardName, string DoNumber, string StockYardDRC_EMail)
        {
            int temp = 0; string OutFileName = "";
            try
            {
                //'Unique File Sr No|UTR No|"VCF" + Dealer Code|Dealer For Code|Dealer Outlet Code|Transaction Type|Amount|Fund Receipt Date (DD/MM/YYYY)|Fix|Account Type|Financier Code|Instrument Type|Issuing Bank|16 Digit DO Number|End of the line indicator - Fix
                int cnt = 1;
                string OpLine = "", UtrNo = "", OutFileName1 = "", FileLastPart = "", strArchivePayconf = "", strIFSC, strIFSCPart, strVA, strVAPart, strFinCode;

                Boolean ChkRecord = false;
                strArchivePayconf = System.IO.Path.GetDirectoryName(Application.StartupPath) + "\\Payconf" + "\\PayconfBackup";
                FileLastPart = System.DateTime.Now.ToString("ddMMyyyyHHmmss");
                if (!Directory.Exists(System.IO.Path.GetDirectoryName(Application.StartupPath) + "\\Payconf"))
                    Directory.CreateDirectory(System.IO.Path.GetDirectoryName(Application.StartupPath) + "\\Payconf");
                OutFileName = System.IO.Path.GetDirectoryName(Application.StartupPath) + "\\Payconf\\MSILIN01.PAYCONF" + FileLastPart + ".CSV";
                OutFileName1 = sttg.MISOutputFilePath + "\\MSILIN01.PAYCONF" + FileLastPart + ".CSV";
                if (!Directory.Exists(sttg.MISOutputFilePath))
                    Directory.CreateDirectory(sttg.MISOutputFilePath);

                if (File.Exists(OutFileName) == true)
                {
                    if (!Directory.Exists(strArchivePayconf))
                        Directory.CreateDirectory(strArchivePayconf);
                }
                else
                    ts_inGlobal = new StreamWriter(OutFileName);

                DataTable dt = service.getDetails("Get_SeqNo_FileMail", "", "", "", "", "", "", "");
                temp = Convert.ToInt32(dt.Rows[0][0].ToString());
                cnt = temp + 1;
                DataTable dtOI = service.getDetails("Get_OrderInv_FileMail", DoNumber, "", "", "", "", "", "");
                if (dtOI.Rows.Count > 0)
                {
                    clserr.LogEntry("Writing Started for  " + StockYardName + " Paycon: MSILIN01.PAYCONF. " + System.DateTime.Now.Date.ToString("dd/MMM/yyyy HH:MM:ss"), false);
                    UtrNo = "";
                    for (int kk = 0; kk < dtOI.Rows.Count; kk++)
                    {
                        if (UtrNo.ToUpper() != dtOI.Rows[kk]["Utr_Number"].ToString().ToUpper())
                        {
                            ChkRecord = true;
                            UtrNo = dtOI.Rows[kk]["Utr_Number"].ToString();
                            OpLine = "";
                            OpLine = ("HDFC" + System.DateTime.Now.ToString("ddMMyy") + Convert.ToInt32(cnt).ToString("000")).ToString() + ",";
                            clserr.LogEntry("Before Counter Value Update", false);
                            DataTable dtu = service.UpdateDetails("Update_FileMail", cnt.ToString(), "", "", "", "", "", "");
                            clserr.LogEntry("Updating the File_mail table with Date and SeqNo = " + System.DateTime.Now.Date + " and " + cnt + "File Name [ " + OutFileName1 + "]", false);

                            OpLine = OpLine + dtOI.Rows[kk]["Utr_Number"].ToString() + ",";
                            if (dtOI.Rows[kk]["Dealer_Code"].ToString().Length > 10)
                                OpLine = OpLine + dtOI.Rows[kk]["Dealer_Code"].ToString().Substring(0, 10) + ",";
                            else
                                OpLine = OpLine + dtOI.Rows[kk]["Dealer_Code"].ToString() + ",";

                            clserr.LogEntry("VCF Dealer_code = " + OpLine + " Payconf file for " + StockYardName, false);
                            OpLine = OpLine + dtOI.Rows[kk]["Dealer_Destination_code"].ToString() + ",";

                            strIFSC = "";
                            strIFSCPart = "";
                            strVA = "";
                            strVAPart = "";
                            strFinCode = "";
                            strIFSC = dtOI.Rows[kk]["IFSC_code"].ToString();
                            strVA = dtOI.Rows[kk]["FNCR_Virtual_Account"].ToString();
                            if (strVA.Length > 1)
                                strVAPart = strVA.ToString().Substring(strVA.Length - 1, 1);
                            strIFSCPart = strIFSC.Substring(0, 4);
                            DataTable dt_Finance = service.getDetails("Get_FinancerDetailsAsPer", strVAPart, strIFSCPart, "", "", "", "", "");
                            if (dt_Finance.Rows.Count > 0)
                                strFinCode = dt_Finance.Rows[0]["FCode"].ToString();

                            OpLine = OpLine + dtOI.Rows[kk]["Dealer_Outlet_Code"].ToString() + ","; //Convert.ToDouble(dtOI.Rows[kk]["Dealer_Outlet_Code"].ToString().Substring(0, 1)).ToString("00") + ","; /////On call make hard code on call discuss 26062024
                            OpLine = OpLine + "TVP-NAG-RTGS-HDFC" + ",";// 'TRANSACTION_Type
                            OpLine = OpLine + Convert.ToDouble("" + dtOI.Rows[kk]["Transaction_Amount"].ToString()).ToString("0.00") + ",";
                            OpLine = OpLine + Convert.ToDateTime(dtOI.Rows[kk]["Cash_ops_Date"].ToString()).ToString("dd-MMM-yyyy") + ",";
                            OpLine = OpLine + "NA".ToString().Substring(0, 1) + ",";// 'Fix
                            OpLine = OpLine + "V".ToString().Substring(0, 1) + ",";// 'Account_Type
                            if (strFinCode.Length > 49)
                                OpLine = OpLine + strFinCode.ToString().Substring(0, 49) + ",";// 'Financier_Code
                            else
                                OpLine = OpLine + strFinCode.ToString() + ",";// 'Financier_Code                                                                              
                            OpLine = OpLine + "TT".ToString() + ",";// 'Instrument_Type                                                                   
                            OpLine = OpLine + "OTH".ToString() + ",";// 'Issuing_Bank

                            OpLine = OpLine + dtOI.Rows[kk]["DO_Number"].ToString() + ",";
                            OpLine = OpLine + "/" + ",";
                            ts_inGlobal.WriteLine(OpLine);
                        }
                        DataTable dtinu = service.UpdateDetails("Update_InvoiceF3", "F3_MIS", dtOI.Rows[kk]["Invoice_Number"].ToString(), "", "", "", "", "");
                        if (Convert.ToInt32(dtinu.Rows[0][0]) > 0)
                        { }
                        else
                            clserr.LogEntry("Record Not Updated for Update_InvoiceF3 for invoice No: "+ dtOI.Rows[kk]["Invoice_Number"].ToString() + " FillF3 ", false);

                        //DataTable dtinu = service.UpdateDetails("Update_InvoiceF", "F3_MIS", dtOI.Rows[kk]["Invoice_Number"].ToString(), "", "", "", "", "");
                    }
                    clserr.LogEntry("END of writing : MSILIN01.PAYCONF. " + System.DateTime.Now.Date.ToString("dd/MMM/yyyy HH:MM:ss") + "  FillF3_" + StockYardName, false);
                }
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();
                ts_inGlobal.Close();
                ts_inGlobal = null;
                if (ChkRecord == false)
                    File.Delete(OutFileName); // it delete the file from Payconf folder
                else
                {
                    //-----------------------Payconf mail will be sent from DRC Emailid--07/05/2019 For Nagpur Loc------------------------
                    if (StockYardDRC_EMail == "")
                        clserr.LogEntry("Email ID Not Maintained." + "  DRC mail", false);
                    else
                    {
                        SendEmail(StockYardDRC_EMail, "", "", OutFileName, "MSIL PAYCON", "FYI");
                        clserr.LogEntry("DRC_Email mail sent -" + StockYardDRC_EMail + "  FillF3_" + StockYardName, false);
                    }
                    System.GC.Collect();
                    System.GC.WaitForPendingFinalizers();
                    File.Move(OutFileName, OutFileName1);// ' moving the file from payconf to MIS folder
                    clserr.LogEntry("Payconf File moving from OutFileName-" + OutFileName + "   FillF3_" + StockYardName, false);
                    clserr.LogEntry("Payconf File moved from OutFileName to OutFileName1-" + OutFileName1 + "  FillF3_" + StockYardName, false);
                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message + "", "FillF3_" + StockYardName, OutFileName);
                ts_inGlobal.Close();
                ts_inGlobal = null;
            }
        }
        private void FillF4()
        {
            string OutFileName = "";
            try
            {
                //Unique File Sr No|UTR No|"VCF" + Dealer Code|Dealer For Code|Dealer Outlet Code|Transaction Type|Amount|Fund Receipt Date (DD/MM/YYYY)|Fix|Account Type|Financier Code|Instrument Type|Issuing Bank|16 Digit DO Number|End of the line indicator - Fix
                int cnt = 1;

                string OpLine = "", UtrNo = "", MISOutputFilePath = "";
                MISOutputFilePath = Application.StartupPath + "MIS"; //'Payment Received
                if (!Directory.Exists(MISOutputFilePath))
                    Directory.CreateDirectory(MISOutputFilePath);
                OutFileName = MISOutputFilePath + "\\MSILIN01.PAYCONF.CONSOLIDATE" + System.DateTime.Now.ToString("ddMMyyyyHHmmss") + ".CSV";

                DataTable dt = service.getDetails("Get_CashOPSDetails_F4", "", "", "", "", "", "", "");

                if (dt.Rows.Count > 0)
                {
                    StreamWriter ts_in = new StreamWriter(OutFileName);
                    try
                    {
                        clserr.LogEntry("EOD Start : PAYCONF.CONSOLIDATE. " + System.DateTime.Now.ToString("dd /MMM/yyyy HH:MM:ss") + "    FillF4", false);
                        UtrNo = "";
                        for (int jj = 0; jj < dt.Rows.Count; jj++)
                        {
                            if (UtrNo != dt.Rows[jj]["Utr_Number"].ToString().ToUpper())
                            {
                                UtrNo = dt.Rows[jj]["Utr_Number"].ToString();
                                OpLine = "";
                                OpLine = "\"" + ("HDFC" + System.DateTime.Now.ToString("ddMMyyyyHHmmss") + cnt.ToString("000")) + "\"" + ",";

                                if (dt.Rows[jj]["Utr_Number"].ToString().Length > 30)
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["Utr_Number"].ToString().Substring(0, 30) + "\"" + ",";
                                else
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["Utr_Number"].ToString() + "\"" + ",";

                                if (dt.Rows[jj]["Dealer_Code"].ToString().Length > 10)
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["Dealer_Code"].ToString().Substring(0, 10) + "\"" + ",";
                                else
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["Dealer_Code"].ToString() + "\"" + ",";

                                OpLine = OpLine + "\"" + dt.Rows[jj]["Dealer_Destination_code"].ToString().Substring(0, 2) + "\"" + ",";
                                if (dt.Rows[jj]["Dealer_Outlet_Code"].ToString().Length > 3)
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["Dealer_Outlet_Code"].ToString().Substring(0, 3) + "\"" + ",";
                                else
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["Dealer_Outlet_Code"].ToString() + "\"" + ",";


                                if (dt.Rows[jj]["TRANSACTION_Type"].ToString().Length > 30)
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["TRANSACTION_Type"].ToString().Substring(0, 30) + "\"" + ",";
                                else
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["TRANSACTION_Type"].ToString() + "\"" + ",";

                                OpLine = OpLine + Convert.ToDouble("" + dt.Rows[jj]["Transaction_Amount"]).ToString("0.00") + ",";
                                OpLine = OpLine + Convert.ToDateTime(dt.Rows[jj]["Cash_ops_Date"]).ToString("dd/MMM/yyyy") + ",";

                                if (dt.Rows[jj]["Fix"].ToString().Length > 1)
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["Fix"].ToString().Substring(0, 2) + "\"" + ",";
                                else
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["Fix"].ToString() + "\"" + ",";

                                if (dt.Rows[jj]["Account_Type"].ToString().Length > 1)
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["Account_Type"].ToString().Substring(0, 2) + "\"" + ",";
                                else
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["Account_Type"].ToString() + "\"" + ",";

                                if (dt.Rows[jj]["Financier_Code"].ToString().Length > 50)
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["Financier_Code"].ToString().Substring(0, 50) + "\"" + ",";
                                else
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["Financier_Code"].ToString() + "\"" + ",";

                                if (dt.Rows[jj]["Instrument_Type"].ToString().Length > 50)
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["Instrument_Type"].ToString().Substring(0, 50) + "\"" + ",";
                                else
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["Instrument_Type"].ToString() + "\"" + ",";
                                if (dt.Rows[jj]["Issuing_Bank"].ToString().Length > 50)
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["Issuing_Bank"].ToString().Substring(0, 50) + "\"" + ",";
                                else
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["Issuing_Bank"].ToString() + "\"" + ",";
                                if (dt.Rows[jj]["DO_Number"].ToString().Length > 50)
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["DO_Number"].ToString().Substring(0, 50) + "\"" + ",";
                                else
                                    OpLine = OpLine + "\"" + dt.Rows[jj]["DO_Number"].ToString() + "\"" + ",";
                                OpLine = OpLine + "/" + ",";
                                ts_in.WriteLine(OpLine);
                                DataTable dtinu = service.UpdateDetails("Update_InvoiceF", "F4_MIS", dt.Rows[0]["Invoice_Number"].ToString(), "", "", "", "", "");
                                if (Convert.ToInt32(dtinu.Rows[0][0]) > 0)
                                { }
                                else
                                    clserr.LogEntry("Record Not Updated for Update_InvoiceF for " + " FillF4 ", false);

                            }
                        }
                        clserr.LogEntry("EOD END : PAYCONF.CONSOLIDATE. " + System.DateTime.Now.ToString("dd /MMM/yyyy HH:MM:ss") + "    FillF4", false);
                        ts_in.Close();
                        SendEmail(sttg.EOD_MIS, "", "", OutFileName, "CONSOLIDATED END OF DAY FILE", "Please find Consolidated Collection Report attached");
                    }
                    catch (Exception ex)
                    {
                        clserr.WriteErrorToTxtFile("EOD Error : PAYCONF.CONSOLIDATE. " + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss"), "FillF4", OutFileName);
                        ts_in.Close();
                        ts_in = null;
                    }
                }

            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile("EOD Error : PAYCONF.CONSOLIDATE. " + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss"), "FillF4", OutFileName);
                //ts_in.Close();
                ts_inGlobal = null;
            }
        }
        private void FillEODACC()
        {
            string OutFileName = "";
            try
            {
                clserr.LogEntry("Entered FillEODACC Function" + "   FillEODACC", false);
                DataTable dt = service.getDetails("Get_CashOPS_PendingCrd", "", "", "", "", "", "", "");
                if (dt.Rows.Count > 0)
                {
                    string MISOutputFilePath = "", strTempFolder = "", OutFileName1 = "", FileLastPart = "";
                    strTempFolder = Application.StartupPath + "DailyExcel";
                    if (!Directory.Exists(sttg.IntraDayPath))
                        Directory.CreateDirectory(sttg.IntraDayPath);
                    if (!Directory.Exists(strTempFolder))
                        Directory.CreateDirectory(strTempFolder);
                    MISOutputFilePath = sttg.IntraDayPath;// 'Payment Received
                    FileLastPart = System.DateTime.Now.ToString("ddMMyyyyHHmmss");
                    OutFileName = MISOutputFilePath + "\\DailyExcel_" + FileLastPart + ".xlsx";
                    OutFileName1 = strTempFolder + "\\DailyExcel_" + FileLastPart + ".xlsx";
                    clserr.LogEntry("EOD Start : EOD FILE. " + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss") + "   FillEODACC", false);
                    for (int kk = 0; kk < dt.Rows.Count; kk++)
                    {
                        using (var fs = new FileStream(OutFileName1, FileMode.Create, FileAccess.Write))
                        {
                            XSSFWorkbook workbook = new XSSFWorkbook();
                            ISheet excelSheet = workbook.CreateSheet("Sheet1");
                            try
                            {
                                IRow rowHeader = excelSheet.CreateRow(0);
                                rowHeader.CreateCell(0).SetCellValue("Dealer Code,Name");
                                rowHeader.CreateCell(1).SetCellValue("Dealer account number");
                                rowHeader.CreateCell(2).SetCellValue("Amount");
                                rowHeader.CreateCell(3).SetCellValue("DO Number");
                                rowHeader.CreateCell(4).SetCellValue("Ref no");
                                rowHeader.CreateCell(5).SetCellValue("Status");

                                for (int i = 0; i < dt.Rows.Count; i++)
                                {
                                    string Dealer_Code = "";
                                    string Dealer_Name = "";
                                    IRow row = excelSheet.CreateRow(i + 1);
                                    DataTable dtAcc = service.getDetails("Get_AccountDetails", "", "", "", "", "", "", "");
                                    if (dtAcc.Rows.Count > 0)
                                        row.CreateCell(0).SetCellValue(dt.Rows[i]["CR_Account_No"].ToString());
                                    row.CreateCell(1).SetCellValue(Convert.ToDateTime(dt.Rows[i]["Cash_ops_Date"].ToString()).ToString("dd/MM/yyyy"));

                                    DataTable dtInvoice = service.getDetails("Get_InvoiceDetails_AsCashID", dt.Rows[i]["Cash_Ops_ID"].ToString(), "", "", "", "", "", "");
                                    if (dtInvoice.Rows.Count > 0)
                                    {
                                        row.CreateCell(2).SetCellValue(dtInvoice.Rows[0]["IMEX_DEAL_NUMBER"].ToString());
                                        Dealer_Code = dtInvoice.Rows[0]["Dealer_Code"].ToString();
                                        Dealer_Name = dtInvoice.Rows[0]["Dealer_Name"].ToString();
                                    }

                                    row.CreateCell(3).SetCellValue(dt.Rows[i]["Virtual_Account"].ToString() + dt.Rows[i]["UTR_No"].ToString());
                                    row.CreateCell(4).SetCellValue(dt.Rows[i]["CashOps_FileType"].ToString());
                                    row.CreateCell(5).SetCellValue(dt.Rows[i]["Virtual_Account"].ToString() + dt.Rows[i]["UTR_No"].ToString() + Dealer_Code + Dealer_Name);

                                    row.CreateCell(6).SetCellValue(Convert.ToDateTime(dt.Rows[i]["Cash_ops_Date"].ToString()).ToString("dd/MM/yyyy"));
                                    if (dt.Rows[i]["Payment_Status"].ToString().ToUpper() == "CREDIT MSIL")
                                        row.CreateCell(8).SetCellValue(dt.Rows[i]["Transaction_Amount"].ToString());
                                    else
                                        row.CreateCell(7).SetCellValue(dt.Rows[i]["Transaction_Amount"].ToString());

                                    DataTable dtu = service.UpdateDetails("Update_CashUpload_EODFlag", dt.Rows[i]["Cash_Ops_ID"].ToString(), "", "", "", "", "", "");
                                    if (Convert.ToInt32(dtu.Rows[0][0]) > 0)
                                    { }
                                    else
                                        clserr.LogEntry("Record Not Updated for Update_CashUpload_EODFlag  " + "   SendInvoiceConfirmation", false);

                                }

                                clserr.LogEntry("EOD END : EOD FILE. " + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss") + "   FillEODACC", false);

                                if (dt.Rows.Count > 0)
                                {
                                    SendEmail(sttg.EOD_File_EMail, "", "", OutFileName1, "MSIL Daily EOD Report", "FYI");
                                    clserr.LogEntry("MSIL Payments Received mail sent to: " + sttg.PAYREC_TRADE_EMAIL + "  FillF1", false);
                                }

                                workbook.Write(fs, true);
                                fs.Close();
                                workbook.Dispose();
                                fs.Dispose();
                            }
                            catch (Exception ex)
                            {
                                clserr.WriteErrorToTxtFile(ex.Message, "Payment Received-F1" + "   FillEODACC", OutFileName);
                                workbook.Close();
                                fs.Close();
                                workbook.Dispose();
                                fs.Dispose();
                            }
                        }
                    }
                    System.GC.Collect();
                    System.GC.WaitForPendingFinalizers();
                    File.Move(OutFileName1, OutFileName);
                    if (Directory.Exists(Application.StartupPath + "MIS_Temp"))
                        Directory.Delete(Application.StartupPath + "MIS_Temp");
                }

            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message, "Payment Received-F1" + "   FillEODACC", OutFileName);
                clserr.WriteErrorToTxtFile("EOD Error : Payment Received-F1. " + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss"), "FillEODACC", OutFileName);
                clserr.WriteErrorToTxtFile("Exiting function FillEODACC", " FillEODACC", OutFileName);
            }
        }
        private void CancelReport()
        {
            string OutFileName = "";
            try
            {
                clserr.LogEntry("Entered CancelReport Function" + "   CancelReport", false);
                DataTable dt = service.getDetails("Get_CancelDetails", "", "", "", "", "", "", "");
                if (dt.Rows.Count > 0)
                {
                    string MISOutputFilePath = "", strTempFolder = "", OutFileName1 = "", FileLastPart = "", OpLine = "";
                    strTempFolder = Application.StartupPath + "Payconf";
                    if (!Directory.Exists(sttg.IntraDayPath))
                        Directory.CreateDirectory(sttg.IntraDayPath);
                    if (!Directory.Exists(strTempFolder))
                        Directory.CreateDirectory(strTempFolder);
                    MISOutputFilePath = sttg.IntraDayPath;// 'Payment Received
                    FileLastPart = System.DateTime.Now.ToString("ddMMyyyyHHmmss");
                    OutFileName = strTempFolder + "\\INVCAN.INVDATA.Success." + FileLastPart + ".CSV";
                    OutFileName1 = MISOutputFilePath + "\\INVCAN.INVDATA.Success." + FileLastPart + ".CSV";

                    clserr.LogEntry("Successfully Invoice Cancel Data . " + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss") + "   CancelReport", false);
                    StreamWriter ts_in = new StreamWriter(OutFileName);
                    try
                    {
                        OpLine = "Invoice Number";
                        ts_in.WriteLine(OpLine);
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            OpLine = dt.Rows[i]["Invoice_Number"].ToString();
                            ts_in.WriteLine(OpLine);
                            DataTable dtu = service.UpdateDetails("Update_CancelInvoices", dt.Rows[i]["Invoice_Number"].ToString(), "", "", "", "", "", "");
                        }
                        ts_in.Close();
                        clserr.LogEntry("End of Successfully Invoice Cancel. " + System.DateTime.Now.ToString("dd / MMM / yyyy HH: MM:SS") + "   CancelReport", false);

                        if (dt.Rows.Count > 0)
                        {
                            SendEmail(sttg.Invoice_Cancel_Email, "", "", OutFileName, "Succesfully Invoice records cancelled", "FYI");
                            clserr.LogEntry("Mail sent to -" + sttg.Invoice_Cancel_Email + OutFileName, false);
                        }
                    }
                    catch (Exception ex)
                    {
                        clserr.WriteErrorToTxtFile(ex.Message, "Cancellation Invoice-F2" + "   CancelReport", OutFileName);
                        ts_in.Close();
                    }
                    System.GC.Collect();
                    System.GC.WaitForPendingFinalizers();
                    File.Move(OutFileName, OutFileName1);
                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message, "Cancellation Invoice-F2" + "   CancelReport", OutFileName);
                clserr.WriteErrorToTxtFile("Error : Succesfully INVOICE Cancel. " + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss"), "CancelReport", OutFileName);
                clserr.WriteErrorToTxtFile("Exiting function CancelReport", " CancelReport", OutFileName);
            }
        }
        private void UnsucessCancelReport()
        {
            string OutFileName = "";
            try
            {
                clserr.LogEntry("Entered UnsucessCancelReport Function" + "   UnsucessCancelReport", false);
                DataTable dt = service.getDetails("Get_UnsuccessCancelDetails", "", "", "", "", "", "", "");
                if (dt.Rows.Count > 0)
                {
                    string MISOutputFilePath = "", strTempFolder = "", OutFileName1 = "", FileLastPart = "", OpLine = "";
                    strTempFolder = Application.StartupPath + "Payconf";
                    if (!Directory.Exists(sttg.IntraDayPath))
                        Directory.CreateDirectory(sttg.IntraDayPath);
                    if (!Directory.Exists(strTempFolder))
                        Directory.CreateDirectory(strTempFolder);
                    MISOutputFilePath = sttg.IntraDayPath;// 'Payment Received
                    FileLastPart = System.DateTime.Now.ToString("ddMMyyyyHHmmss");
                    OutFileName = strTempFolder + "\\UNSUC.INVCAN.INVDATA." + FileLastPart + ".CSV";
                    OutFileName1 = MISOutputFilePath + "\\UNSUC.INVCAN.INVDATA." + FileLastPart + ".CSV";

                    clserr.LogEntry("Successfully Invoice Cancel Data . " + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss") + "   UnsucessCancelReport", false);
                    StreamWriter ts_in = new StreamWriter(OutFileName);
                    try
                    {
                        OpLine = "Invoice Number";
                        ts_in.WriteLine(OpLine);
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            OpLine = dt.Rows[i]["Invoice_Number"].ToString();
                            ts_in.WriteLine(OpLine);
                            DataTable dtu = service.UpdateDetails("Update_CancelInvoices", dt.Rows[i]["Invoice_Number"].ToString(), "", "", "", "", "", "");
                        }
                        ts_in.Close();
                        clserr.LogEntry("EOD END : UnSuccessfully Invoice records Cancelled. " + System.DateTime.Now.ToString("dd / MMM / yyyy HH: MM:SS") + "   UnsucessCancelReport", false);

                        if (dt.Rows.Count > 0)
                        {
                            SendEmail(sttg.Invoice_Cancel_Email, "", "", OutFileName, "UnSuccessfully Invoice records cancelled", "FYI");
                            clserr.LogEntry("Mail sent to -" + sttg.Invoice_Cancel_Email + OutFileName, false);
                        }
                        System.GC.Collect();
                        System.GC.WaitForPendingFinalizers();
                        File.Move(OutFileName, OutFileName1);
                    }
                    catch (Exception ex)
                    {
                        ts_in.Close();
                        ts_in.Dispose();
                        clserr.WriteErrorToTxtFile(ex.Message, "Unsucces Cancellation Invoice-F2" + "   UnsucessCancelReport", OutFileName);
                    }
                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message, "Unsucces Cancellation Invoice-F2" + "   UnsucessCancelReport", OutFileName);
                clserr.WriteErrorToTxtFile("EOD Error : Unsuccess CANCEL INVOICE. " + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss"), "UnsucessCancelReport", OutFileName);
                clserr.WriteErrorToTxtFile("Exiting function UnsucessCancelReport", " UnsucessCancelReport", OutFileName);
            }
        }
        private void Cancel_DO_Report()
        {
            string OutFileName = "";
            try
            {
                clserr.LogEntry("Entered Cancel_DO_Report Function" + "   Cancel_DO_Report", false);
                DataTable dt = service.getDetails("Get_Cancel_DO_Details", "", "", "", "", "", "", "");
                if (dt.Rows.Count > 0)
                {
                    string MISOutputFilePath = "", strTempFolder = "", OutFileName1 = "", FileLastPart = "", OpLine = "";
                    strTempFolder = Application.StartupPath + "Payconf";
                    if (!Directory.Exists(sttg.IntraDayPath))
                        Directory.CreateDirectory(sttg.IntraDayPath);
                    if (!Directory.Exists(strTempFolder))
                        Directory.CreateDirectory(strTempFolder);
                    MISOutputFilePath = sttg.IntraDayPath;// 'Payment Received
                    FileLastPart = System.DateTime.Now.ToString("ddMMyyyyHHmmss");
                    OutFileName = strTempFolder + "\\DOCAN.ORDATA.Success." + FileLastPart + ".CSV";
                    OutFileName1 = MISOutputFilePath + "\\DOCAN.ORDATA.Success." + FileLastPart + ".CSV";

                    clserr.LogEntry("Successfully DO Cancel Data . " + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss") + "   Cancel_DO_Report", false);
                    StreamWriter ts_in = new StreamWriter(OutFileName);
                    try
                    {
                        OpLine = "Order Number";
                        ts_in.WriteLine(OpLine);
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            OpLine = dt.Rows[i]["DO_number"].ToString();
                            ts_in.WriteLine(OpLine);
                            DataTable dtu = service.UpdateDetails("Update_DO_Desc", dt.Rows[i]["DO_number"].ToString(), "", "", "", "", "", "");
                        }
                        ts_in.Close();
                        clserr.LogEntry("End of Successfully DO Cancel. " + System.DateTime.Now.ToString("dd / MMM / yyyy HH: MM:SS") + "   Cancel_DO_Report", false);

                        if (dt.Rows.Count > 0)
                        {
                            SendEmail(sttg.DO_Cancel_Email, "", "", OutFileName, "Succesfully DO records cancelled", "FYI");
                            clserr.LogEntry("Mail sent to -" + sttg.DO_Cancel_Email + OutFileName, false);
                        }
                        System.GC.Collect();
                        System.GC.WaitForPendingFinalizers();
                        File.Move(OutFileName, OutFileName1);
                    }
                    catch (Exception ex)
                    {
                        clserr.WriteErrorToTxtFile(ex.Message, "Cancellation DO - F2" + "   Cancel_DO_Report", OutFileName);
                        ts_in.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message, "Cancellation DO - F2" + "   Cancel_DO_Report", OutFileName);
                clserr.WriteErrorToTxtFile("EOD Error : Succesfully DO Cancel. " + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss"), "Cancel_DO_Report", OutFileName);
                clserr.WriteErrorToTxtFile("Exiting function Cancel_DO_Report", " Cancel_DO_Report", OutFileName);
            }
        }
        private void UnsucCancel_DO_Report()
        {
            string OutFileName = "";
            try
            {
                clserr.LogEntry("Entered UnsucCancel_DO_Report Function" + "   UnsucCancel_DO_Report", false);
                DataTable dt = service.getDetails("Get_UnsucCancel_DO_Report", "", "", "", "", "", "", "");
                if (dt.Rows.Count > 0)
                {
                    string MISOutputFilePath = "", strTempFolder = "", OutFileName1 = "", FileLastPart = "", OpLine = "";
                    strTempFolder = Application.StartupPath + "Payconf";
                    if (!Directory.Exists(sttg.IntraDayPath))
                        Directory.CreateDirectory(sttg.IntraDayPath);
                    if (!Directory.Exists(strTempFolder))
                        Directory.CreateDirectory(strTempFolder);
                    MISOutputFilePath = sttg.IntraDayPath;// 'Payment Received
                    FileLastPart = System.DateTime.Now.ToString("ddMMyyyyHHmmss");

                    OutFileName = strTempFolder + "\\DOCAN.ORDATA.UnSuccess." + FileLastPart + ".CSV";
                    OutFileName1 = MISOutputFilePath + "\\DOCAN.ORDATA.UnSuccess." + FileLastPart + ".CSV";


                    clserr.LogEntry("EOD Start : Unsuccess Cancel Invoice." + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss") + "   UnsucCancel_DO_Report", false);
                    StreamWriter ts_in = new StreamWriter(OutFileName);
                    try
                    {
                        OpLine = "Order Number";
                        ts_in.WriteLine(OpLine);
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            OpLine = dt.Rows[i]["DO_number"].ToString();
                            ts_in.WriteLine(OpLine);
                            DataTable dtu = service.UpdateDetails("Update_DO_Desc", dt.Rows[i]["DO_number"].ToString(), "", "", "", "", "", "");
                        }
                        ts_in.Close();
                        clserr.LogEntry("End of UnSuccessful DO data cancel." + System.DateTime.Now.ToString("dd / MMM / yyyy HH: MM:SS") + "   UnsucCancel_DO_Report", false);

                        if (dt.Rows.Count > 0)
                        {
                            SendEmail(sttg.DO_Cancel_Email, "", "", OutFileName, "UnSuccessful DO record Cancelled", "FYI");
                            clserr.LogEntry("Mail sent to -" + sttg.DO_Cancel_Email + OutFileName, false);
                        }
                        System.GC.Collect();
                        System.GC.WaitForPendingFinalizers();
                        File.Move(OutFileName, OutFileName1);
                    }
                    catch (Exception ex)
                    {
                        ts_in.Close();
                        clserr.WriteErrorToTxtFile(ex.Message, "UnCancellation DO-F2" + "   UnsucCancel_DO_Report", OutFileName);
                    }
                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message, "UnCancellation DO-F2" + "   UnsucCancel_DO_Report", OutFileName);
                clserr.WriteErrorToTxtFile("EOD Error : UnSuccesfully DO Cancel. " + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss"), "UnsucCancel_DO_Report", OutFileName);
                clserr.WriteErrorToTxtFile("Exiting function Cancel_DO_Report", " UnsucCancel_DO_Report", OutFileName);
            }
        }
        private void Cancel_DO_Invoice_Report()
        {
            string OutFileName = "";
            try
            {
                clserr.LogEntry("Entered Cancel_DO_Invoice_Report Function" + "   Cancel_DO_Invoice_Report", false);
                DataTable dt = service.getDetails("Get_Cancel_DO_Invoice_Report", "", "", "", "", "", "", "");
                if (dt.Rows.Count > 0)
                {
                    string MISOutputFilePath = "", strTempFolder = "", OutFileName1 = "", FileLastPart = "", OpLine = "";
                    strTempFolder = Application.StartupPath + "Payconf";
                    if (!Directory.Exists(sttg.IntraDayPath))
                        Directory.CreateDirectory(sttg.IntraDayPath);
                    if (!Directory.Exists(strTempFolder))
                        Directory.CreateDirectory(strTempFolder);
                    MISOutputFilePath = sttg.IntraDayPath;// 'Payment Received
                    FileLastPart = System.DateTime.Now.ToString("ddMMyyyyHHmmss");
                    OutFileName = strTempFolder + "\\DOINVCAN.ORDATA.Success." + FileLastPart + ".CSV";
                    OutFileName1 = MISOutputFilePath + "\\DOINVCAN.ORDATA.Success." + FileLastPart + ".CSV";
                    clserr.LogEntry("Successfully DO and Invoice Cancel Data." + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss") + "   Cancel_DO_Invoice_Report", false);
                    StreamWriter ts_in = new StreamWriter(OutFileName);
                    try
                    {
                        OpLine = "Order Number" + ";" + "Invoice Number";
                        ts_in.WriteLine(OpLine);
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            OpLine = dt.Rows[i]["DO_number"].ToString() + ";" + dt.Rows[i]["Invoice_Number"].ToString();
                            ts_in.WriteLine(OpLine);
                            DataTable dtu = service.UpdateDetails("Update_Cancel_DO_Invoice_Report", dt.Rows[i]["DO_number"].ToString(), "", "", "", "", "", "");
                        }
                        ts_in.Close();
                        clserr.LogEntry("End of Successfully DO and Invoice Cancel. " + System.DateTime.Now.ToString("dd / MMM / yyyy HH: MM:SS") + "   Cancel_DO_Invoice_Report", false);

                        if (dt.Rows.Count > 0)
                        {
                            SendEmail(sttg.DO_Invoice_Cancel_Email, "", "", OutFileName, "SUCCESSFUL DO RECORD CANCELLATION ALONG WITH INVOICE", "FYI");
                            clserr.LogEntry("Sucessfully Do and Invoice records cancelled. Mail sent to -" + sttg.DO_Invoice_Cancel_Email + OutFileName, false);
                        }
                        System.GC.Collect();
                        System.GC.WaitForPendingFinalizers();
                        File.Move(OutFileName, OutFileName1);
                    }
                    catch (Exception ex)
                    {
                        clserr.WriteErrorToTxtFile(ex.Message, "Cancellation DO and Invoice-F2" + "   Cancel_DO_Invoice_Report", OutFileName);
                        ts_in.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message, "Cancellation DO and Invoice-F2" + "   Cancel_DO_Invoice_Report", OutFileName);
                clserr.WriteErrorToTxtFile("EOD Error : Succesfully DO and Invoice Cancel. " + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss"), "Cancel_DO_Invoice_Report", OutFileName);
                clserr.WriteErrorToTxtFile("Exiting function Cancel_DO_Invoice_Report", " Cancel_DO_Invoice_Report", OutFileName);
            }
        }
        private void UnCancel_DO_Invoice_Report()
        {
            string OutFileName = "";
            try
            {
                clserr.LogEntry("Entered UnCancel_DO_Invoice_Report Function" + "   UnCancel_DO_Invoice_Report", false);
                DataTable dt = service.getDetails("Get_UnCancel_DO_Invoice_Report", "", "", "", "", "", "", "");
                if (dt.Rows.Count > 0)
                {
                    string MISOutputFilePath = "", strTempFolder = "", OutFileName1 = "", FileLastPart = "", OpLine = "";
                    strTempFolder = Application.StartupPath + "Payconf";
                    if (!Directory.Exists(sttg.IntraDayPath))
                        Directory.CreateDirectory(sttg.IntraDayPath);
                    if (!Directory.Exists(strTempFolder))
                        Directory.CreateDirectory(strTempFolder);
                    MISOutputFilePath = sttg.IntraDayPath;// 'Payment Received
                    FileLastPart = System.DateTime.Now.ToString("ddMMyyyyHHmmss");
                    OutFileName = strTempFolder + "\\DOINVCAN.ORDATA.UnSuccess." + FileLastPart + ".CSV";
                    OutFileName1 = MISOutputFilePath + "\\DOINVCAN.ORDATA.UnSuccess." + FileLastPart + ".CSV";
                    clserr.LogEntry("Successfully DO and Invoice Cancel Data." + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss") + "   UnCancel_DO_Invoice_Report", false);
                    StreamWriter ts_in = new StreamWriter(OutFileName);
                    try
                    {
                        OpLine = "Order Number" + ";" + "Invoice Number";
                        ts_in.WriteLine(OpLine);
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            OpLine = dt.Rows[i]["DO_number"].ToString() + ";" + dt.Rows[i]["Invoice_Number"].ToString();
                            ts_in.WriteLine(OpLine);
                            DataTable dtu = service.UpdateDetails("Update_Cancel_DO_Invoice_Report", dt.Rows[i]["DO_number"].ToString(), "", "", "", "", "", "");
                        }
                        ts_in.Close();
                        clserr.LogEntry("End of UnSuccessfully DO and Invoice Cancel. " + System.DateTime.Now.ToString("dd / MMM / yyyy HH: MM:SS") + "   UnCancel_DO_Invoice_Report", false);

                        if (dt.Rows.Count > 0)
                        {
                            SendEmail(sttg.DO_Invoice_Cancel_Email, "", "", OutFileName, "SUCCESSFUL DO RECORD CANCELLATION ALONG WITH INVOICE", "FYI");
                            clserr.LogEntry("UnSucessfully Do and Invoice records cancelled.Mail sent to -" + sttg.DO_Invoice_Cancel_Email + OutFileName, false);
                        }
                        System.GC.Collect();
                        System.GC.WaitForPendingFinalizers();
                        File.Move(OutFileName, OutFileName1);
                    }
                    catch (Exception ex) { ts_in.Close(); clserr.WriteErrorToTxtFile(ex.Message, "UnCancellation DO and Invoice-F2" + "   UnCancel_DO_Invoice_Report", OutFileName); }
                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message, "UnCancellation DO and Invoice-F2" + "   UnCancel_DO_Invoice_Report", OutFileName);
                clserr.WriteErrorToTxtFile("EOD Error : UnSuccesfully DO and Invoice Cancel. " + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss"), "UnCancel_DO_Invoice_Report", OutFileName);
                clserr.WriteErrorToTxtFile("Exiting function UnCancel_DO_Invoice_Report", " UnCancel_DO_Invoice_Report", OutFileName);
            }
        }
        private void SendInvoiceConfirmation()
        {
            string OutFileName = "";
            try
            {
                clserr.LogEntry("Entered SendInvoiceConfirmation Function" + "   SendInvoiceConfirmation", false);
                DataTable dtF = service.getDetails("Get_SendInvoiceConfirmation", "", "", "", "", "", "", "");
                if (dtF.Rows.Count > 0)
                {
                    string OpLine = "";
                    if (!Directory.Exists(Application.StartupPath + "\\Invoice_Conf"))
                        Directory.CreateDirectory(Application.StartupPath + "\\Invoice_Conf");
                    OutFileName = Application.StartupPath + "\\Invoice_Conf\\" + sttg.INV_CONF_FNAME + "INV_CONFIRMATION" + System.DateTime.Now.ToString("ddMM") + ".CSV";
                    clserr.LogEntry("EOD Start : INVOICE CONFIRMATION. " + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss") + "   SendInvoiceConfirmation", false);
                    ts_inGlobal = new StreamWriter(OutFileName);
                    //OpLine = "Order Number" + ";" + "Invoice Number";
                    //ts_in.WriteLine(OpLine);

                    for (int kk = 0; kk < dtF.Rows.Count; kk++)
                    {
                        DataTable dt = service.getDetails("Get_SendInvoiceConfirmationDetails", "", "", "", "", "", "", "");
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            OpLine = dt.Rows[i]["Invoice_Number"].ToString() + ",";
                            for (int hh = 3; hh <= 17; hh++)
                                OpLine = OpLine + "" + dt.Rows[i][hh].ToString().Replace(",", "") + ",";

                            ts_inGlobal.WriteLine(OpLine);
                        }
                        DataTable dtu = service.UpdateDetails("Update_SendInvoiceConfirmation", "", "", "", "", "", "", "");
                        if (Convert.ToInt32(dtu.Rows[0][0]) > 0)
                        { }
                        else
                            clserr.LogEntry("Record Not Updated for Update_SendInvoiceConfirmation  "+ "   SendInvoiceConfirmation", false);

                    }
                    ts_inGlobal.Close();
                    clserr.LogEntry("EOD END : INVOICE CONFIRMATION.  " + System.DateTime.Now.ToString("dd / MMM / yyyy HH: MM:SS") + "   SendInvoiceConfirmation", false);

                    if (dtF.Rows.Count > 1)
                    {
                        SendEmail(sttg.INV_CONF_EMAIL, "", "", Application.StartupPath + "\\Invoice_Conf\\" + sttg.INV_CONF_FNAME + "INV_CONFIRMATION" + System.DateTime.Now.ToString("ddMM") + ".CSV", "INVOICE Confirmation DATA", "FYI");
                        //clserr.LogEntry("UnSucessfully Do and Invoice records cancelled.Mail sent to -" + sttg.DO_Invoice_Cancel_Email + OutFileName, false);
                    }
                    System.GC.Collect();
                    System.GC.WaitForPendingFinalizers();
                    //File.Move(OutFileName, OutFileName1);
                }
            }
            catch (Exception ex)
            {
               
                clserr.WriteErrorToTxtFile(ex.Message, "INV_CONFIRMATION-Invoice Confirmation" + "   SendInvoiceConfirmation", OutFileName);
                clserr.WriteErrorToTxtFile("Exiting function SendInvoiceConfirmation", " SendInvoiceConfirmation", OutFileName);
            }
        }
        private void OtherReports()
        {
            clserr.LogEntry("Entered OtherReports Function", false);
            Order_Delete_Report2();//  'order delete
            Old_Invoice_Report();//    'Invoice status is Physical pending
            Order_NotRec_Report();//   'order is pending but physical inv received
            Payment_NotRec_Report();// 'Payment not received but order received and physical received
            MIS3_Report();
            clserr.LogEntry("Exiting OtherReports Function", false);
        }
        private void Order_Delete_Report2()
        {
            string OutFileName = "";
            try
            {
                clserr.LogEntry("Entered Order_Delete_Report2 Function" + "   Order_Delete_Report2", false);
                DataTable dt = service.getDetails("Get_Order_Delete_Report2", "", "", "", "", "", "", "");
                if (dt.Rows.Count > 0)
                {
                    string strTempFolder = "", FileLastPart = "", OpLine = "";
                    strTempFolder = Application.StartupPath + "Payconf";
                    if (!Directory.Exists(Application.StartupPath + "ORD_DEL"))
                        Directory.CreateDirectory(Application.StartupPath + "ORD_DEL");
                    if (!Directory.Exists(strTempFolder))
                        Directory.CreateDirectory(strTempFolder);

                    FileLastPart = System.DateTime.Now.ToString("ddMMyyyyHHmmss");
                    OutFileName = Application.StartupPath + "ORD_DEL" + "\\MSILIN01.ORDDEL." + System.DateTime.Now.ToString("ddMMyyyyHHmmss") + ".CSV";
                    clserr.LogEntry("EOD Start : Order Delete Staus." + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss") + "   Order_Delete_Report2", false);
                    StreamWriter ts_in = new StreamWriter(OutFileName);
                    try
                    {
                        OpLine = "Order_no" + ";" + "Order_Del_Status";
                        ts_in.WriteLine(OpLine);
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            OpLine = dt.Rows[i]["Order_no"].ToString() + "," + dt.Rows[i]["Order_Del_Status"].ToString() + ",";
                            ts_in.WriteLine(OpLine);
                            DataTable dtu = service.UpdateDetails("Update_Order_Delete_Report2", dt.Rows[i]["Order_Del_ID"].ToString(), "", "", "", "", "", "");
                            if (Convert.ToInt32(dtu.Rows[0][0]) > 0)
                            {}
                            else
                                clserr.LogEntry(" Record not Updated :For Order No and Order Delete ID:" + dt.Rows[i]["Order_no"].ToString()+"," + dt.Rows[i]["Order_Del_ID"].ToString()  , false);

                        }
                        ts_in.Close();
                        clserr.LogEntry("EOD End: Order Delete Staus. " + System.DateTime.Now.ToString("dd / MMM / yyyy HH: MM:SS") + "   Order_Delete_Report2", false);

                        if (dt.Rows.Count > 0)
                        {
                            SendEmail(sttg.ORD_DEL, "", "", OutFileName, "Order Delete Staus", "Order Delete Staus.");
                            clserr.LogEntry("Order Delete Report Mail sent to -" + sttg.DO_Invoice_Cancel_Email + OutFileName, false);
                        }
                        //if (Directory.Exists(Application.StartupPath + "ORD_DEL"))
                        //    Directory.Delete(Application.StartupPath + "ORD_DEL");
                    }
                    catch(Exception ex)
                    {
                        ts_in.Close();
                        clserr.WriteErrorToTxtFile(ex.Message, "Order Delete Data" + "   Order_Delete_Report2", OutFileName);
                    }
                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message, "Order Delete Data" + "   Order_Delete_Report2", OutFileName);
                clserr.WriteErrorToTxtFile("Exiting function Order_Delete_Report2", " Order_Delete_Report2", OutFileName);
            }
        }
        private void Old_Invoice_Report()
        {
            string OutFileName = "";
            try
            {
                clserr.LogEntry("Entered Old_Invoice_Report Function" + "   Old_Invoice_Report", false);
                DataTable dt = service.getDetails("Get_Old_Invoice_Report", "", "", "", "", "", "", "");
                if (dt.Rows.Count > 0)
                {
                    string OpLine = "";

                    if (!Directory.Exists(Application.StartupPath + "OLD_INV"))
                        Directory.CreateDirectory(Application.StartupPath + "OLD_INV");

                    OutFileName = Application.StartupPath + "\\OLD_INV" + "\\INV_PHY_PENDING" + System.DateTime.Now.ToString("ddMMyyyyHHmmss") + ".CSV";
                    clserr.LogEntry("EOD Start : PHYSICAL PENDING." + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss") + "   Old_Invoice_Report", false);
                    StreamWriter ts_in = new StreamWriter(OutFileName);
                    try
                    {
                        OpLine = "Invoice Number";
                        ts_in.WriteLine(OpLine);
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            OpLine = dt.Rows[i]["Invoice_Number"].ToString();
                            ts_in.WriteLine(OpLine);
                        }
                        ts_in.Close();
                        clserr.LogEntry("EOD End: PHYSICAL PENDING. " + System.DateTime.Now.ToString("dd / MMM / yyyy HH:MM:SS") + "   Old_Invoice_Report", false);

                        if (dt.Rows.Count > 0)
                        {
                            SendEmail(sttg.INV_PHY_NTRC, "", "", OutFileName, "Physical Pending", "Physical Pending.");
                        }
                    }
                    catch(Exception ex)
                    {
                        clserr.WriteErrorToTxtFile(ex.Message, "Old Invoice Report" + "   Old_Invoice_Report", OutFileName);
                        ts_in.Close();
                    }
                    //if (Directory.Exists(Application.StartupPath + "OLD_INV"))
                    //    Directory.Delete(Application.StartupPath + "OLD_INV");
                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message, "Old Invoice Report" + "   Old_Invoice_Report", OutFileName);
                clserr.WriteErrorToTxtFile("Exiting function Old_Invoice_Report", " Old_Invoice_Report", OutFileName);
            }
        }
        private void Order_NotRec_Report()
        {
            string OutFileName = "";
            try
            {
                clserr.LogEntry("Entered Order_NotRec_Report Function" + "   Old_Invoice_Report", false);
                DataTable dt = service.getDetails("Get_Order_NotRec_Report", "", "", "", "", "", "", "");
                if (dt.Rows.Count > 0)
                {
                    string OpLine = "";

                    if (!Directory.Exists(Application.StartupPath + "OLD_INV"))
                        Directory.CreateDirectory(Application.StartupPath + "OLD_INV");

                    OutFileName = Application.StartupPath + "\\OLD_INV" + "\\Order_PENDING" + System.DateTime.Now.ToString("ddMMyyyyHHmmss") + ".CSV";
                    clserr.LogEntry("EOD Start : Order_PENDING." + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss") + "   Order_NotRec_Report", false);
                    StreamWriter ts_in = new StreamWriter(OutFileName);
                    try
                    {
                        //OpLine = "Invoice Number";
                        ts_in.WriteLine(OpLine);
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            OpLine = dt.Rows[i]["Invoice_Number"].ToString() + ",";
                            OpLine = OpLine + dt.Rows[i]["IMEX_DEAL_NUMBER"].ToString() + ",";
                            OpLine = OpLine + Convert.ToDouble(dt.Rows[i]["Invoice_Amount"]).ToString("0.00") + ",";
                            OpLine = OpLine + dt.Rows[i]["Dealer_Code"].ToString() + ",";
                            ts_in.WriteLine(OpLine);
                        }
                        ts_in.Close();
                        clserr.LogEntry("EOD End: Order_PENDING. " + System.DateTime.Now.ToString("dd / MMM / yyyy HH: MM:SS") + "   Order_NotRec_Report", false);

                        if (dt.Rows.Count > 0)
                        {
                            SendEmail(sttg.INV_PHY_NTRC, "", "", OutFileName, "Order Pending", "Order Pending.");
                        }
                    }
                    catch(Exception ex) { ts_in.Close(); clserr.WriteErrorToTxtFile(ex.Message, "Order Not Received Report" + "   Order_NotRec_Report", OutFileName); }
                    //if (!Directory.Exists(Application.StartupPath + "OLD_INV"))
                    //    Directory.Delete(Application.StartupPath + "OLD_INV");
                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message, "Order Not Received Report" + "   Order_NotRec_Report", OutFileName);
                clserr.WriteErrorToTxtFile("Exiting function Order_NotRec_Report", " Order_NotRec_Report", OutFileName);
            }
        }
        private void Payment_NotRec_Report()
        {
            string OutFileName = "";
            try
            {
                clserr.LogEntry("Entered Payment_NotRec_Report Function" + "   Old_Invoice_Report", false);
                DataTable dt = service.getDetails("Get_Payment_NotRec_Report", "", "", "", "", "", "", "");
                if (dt.Rows.Count > 0)
                {
                    string OpLine = "";

                    if (!Directory.Exists(Application.StartupPath + "OLD_INV"))
                        Directory.CreateDirectory(Application.StartupPath + "OLD_INV");

                    OutFileName = Application.StartupPath + "\\OLD_INV" + "\\Payment_PENDING" + System.DateTime.Now.ToString("ddMMyyyyHHmmss") + ".CSV";
                    clserr.LogEntry("EOD Start : Payment_PENDING." + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss") + "   Payment_NotRec_Report", false);
                    StreamWriter ts_in = new StreamWriter(OutFileName);
                    try
                    {
                        //OpLine = "Invoice Number";
                        ts_in.WriteLine(OpLine);
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            OpLine = dt.Rows[i]["Invoice_Number"].ToString() + ",";
                            OpLine = OpLine + dt.Rows[i]["order_number"].ToString() + ",";
                            OpLine = OpLine + dt.Rows[i]["IMEX_DEAL_NUMBER"].ToString() + ",";
                            OpLine = OpLine + Convert.ToDouble(dt.Rows[i]["Invoice_Amount"]).ToString("0.00") + ",";
                            OpLine = OpLine + dt.Rows[i]["Dealer_Code"].ToString() + ",";
                            ts_in.WriteLine(OpLine);
                        }
                        ts_in.Close();
                        clserr.LogEntry("EOD End: Payment_PENDING. " + System.DateTime.Now.ToString("dd / MMM / yyyy HH:MM:SS") + "   Payment_NotRec_Report", false);

                        if (dt.Rows.Count > 0)
                        {
                            SendEmail(sttg.INV_PHY_NTRC, "", "", OutFileName, "Payment Pending", "Payment Pending.");
                        }
                    }
                    catch (Exception ex) { ts_in.Close(); clserr.WriteErrorToTxtFile(ex.Message, "Payment Pending Report" + "   Payment_NotRec_Report", OutFileName); }
                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message, "Payment Pending Report" + "   Payment_NotRec_Report", OutFileName);
                clserr.WriteErrorToTxtFile("Exiting function Payment_NotRec_Report", " Payment_NotRec_Report", OutFileName);
            }
        }
        private void MIS3_Report()
        {
            string OutFileName = "";
            try
            {
                clserr.LogEntry("Entered MIS3_Report Function" + "   Old_Invoice_Report", false);
                DataTable dt = service.getDetails("Get_MIS3_Report", "", "", "", "", "", "", "");
                if (dt.Rows.Count > 0)
                {
                    string OpLine = "";

                    if (!Directory.Exists(Application.StartupPath + "OLD_INV"))
                        Directory.CreateDirectory(Application.StartupPath + "OLD_INV");

                    OutFileName = Application.StartupPath + "\\OLD_INV" + "\\MIS3_LEVEL_DATA-DO Cancellation And Retain Invoice Report_" + System.DateTime.Now.ToString("ddMMyyyyHHmmss") + ".CSV";
                    clserr.LogEntry("EOD Start : DO Cancellation and retain invoice done." + System.DateTime.Now.ToString("dd/MMM/yyyy HH:MM:ss") + "   MIS3_Report", false);
                    StreamWriter ts_in = new StreamWriter(OutFileName);
                    try
                    {
                        //OpLine = "Invoice Number";
                        ts_in.WriteLine(OpLine);
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            OpLine = dt.Rows[i]["Invoice No"].ToString() + ",";
                            OpLine = OpLine + dt.Rows[i]["Invoice Amount"].ToString() + ",";
                            OpLine = OpLine + Convert.ToDouble(dt.Rows[i]["Invoice Amount"]).ToString("0.00") + ",";
                            OpLine = OpLine + dt.Rows[i]["Physically Receipt Date"].ToString() + ",";
                            OpLine = OpLine + dt.Rows[i]["Trade Reference No"].ToString() + ",";
                            if (dt.Rows[i]["Delivery Order No"].ToString() == null)
                                OpLine = OpLine + "-" + ",";
                            else
                                OpLine = OpLine + dt.Rows[i]["Delivery Order No"].ToString() + ",";
                            OpLine = OpLine + dt.Rows[i]["Dealer Code"].ToString() + ",";
                            OpLine = OpLine + dt.Rows[i]["Dealer Name"].ToString() + ",";
                            OpLine = OpLine + dt.Rows[i]["Location"].ToString() + ",";
                            ts_in.WriteLine(OpLine);
                        }
                        ts_in.Close();
                        clserr.LogEntry("EOD End: DO Cancellation and retain invoice done." + System.DateTime.Now.ToString("dd / MMM / yyyy HH: MM:SS") + "   MIS3_Report", false);

                        if (dt.Rows.Count > 0)
                        {
                            SendEmail(sttg.DO_Invoice_Cancel_Email, "", "", OutFileName, "Sucessfully DO Cancelled and retained invoice", "DO Cancelled and Retained Invoice");
                        }
                    }
                    catch(Exception ex) { ts_in.Close(); clserr.WriteErrorToTxtFile(ex.Message, "Payment Pending Report" + "   MIS3_Report", OutFileName); }
                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message, "Payment Pending Report" + "   MIS3_Report", OutFileName);
                clserr.WriteErrorToTxtFile("Exiting function MIS3_Report", " MIS3_Report", OutFileName);
            }
        }
    }

}
