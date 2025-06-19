using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Linq;
using Microsoft.EntityFrameworkCore;
using _001TN0173.Entities;
using System.Data.Common;
using Microsoft.Data.SqlClient;

namespace _001TN0173.Shared
{
    class service
    {
        public static ClsErrorLog clserr = new ClsErrorLog();
        public static bool InvoiceFileCheck(string FileType, string File_Date)
        {           
            try
            {
                File_Date = File_Date.ToString().Split(' ')[0];
                using (var db = new Entities.DatabaseContext())
                {
                    var rec = db.Set<InvoiceFileCheck>().FromSqlRaw("Select * from File_Desc where FileType='" + FileType + "' AND File_Date='" + File_Date + "'").ToList();
                    // var rec = db.InvoiceFileChecks.Where(a => a.FileType == FileType && a.File_Date == File_Date).FirstOrDefault();

                    if (rec != null && rec.Count!=0)
                    {
                        return true;
                    }

                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message + "Service", "InvoiceFileCheck", "");
                //clserr.Handle_Error(ex, "Form1", "Generate_SettingFile");
            }
            return false;
        }
        public static bool File_MailF5(string File_Mail_Name, string File_Mail_Date)
        {
            try
            {
                File_Mail_Date = File_Mail_Date.ToString().Split(' ')[0];

                using (var db = new Entities.DatabaseContext())
                {
                    var rec = db.Set<File_MailF5>().FromSqlRaw("Select * from File_Mail where File_Mail_Name='F5' AND File_Mail_Date='" + File_Mail_Date + "'").ToList();

                    //var rec = db.File_MailF5s.Where(a => a.File_Mail_Name == File_Mail_Name && a.File_Mail_Date == File_Mail_Date).FirstOrDefault();

                    if (rec != null && rec.Count != 0)
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message + "Service", "File_MailF5", "");
                return false;
            }
            return false;
        }
        public static DataTable Insert_InvoiceDetails(string Task, string[] details)
        {
            try
            {
                DataTable dataTable = new DataTable();
                //using var cmd = db.LoginMSTs.FromSqlRaw($"SP_FTPayment");
                using (var db = new Entities.DatabaseContext())
                {
                    DbConnection connection = db.Database.GetDbConnection();
                    using var cmd = db.Database.GetDbConnection().CreateCommand();
                    DbProviderFactory dbFactory = DbProviderFactories.GetFactory(connection);
                    cmd.CommandText = "SP_Insert_InvoiceDetails";
                    //common
                    cmd.CommandType = CommandType.StoredProcedure;
                    if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();
                    cmd.Parameters.Add(new SqlParameter("@Task", SqlDbType.VarChar) { Value = Task });
                    cmd.Parameters.Add(new SqlParameter("@Invoice_ID", SqlDbType.VarChar) { Value = details[0] });
                    cmd.Parameters.Add(new SqlParameter("@Sr_No", SqlDbType.VarChar) { Value = details[1] });
                    cmd.Parameters.Add(new SqlParameter("@Invoice_Number", SqlDbType.VarChar) { Value = details[2] });
                    cmd.Parameters.Add(new SqlParameter("@Invoice_Amount", SqlDbType.VarChar) { Value = details[3] });
                    cmd.Parameters.Add(new SqlParameter("@Currency", SqlDbType.VarChar) { Value = details[4] });
                    cmd.Parameters.Add(new SqlParameter("@Vehical_ID", SqlDbType.VarChar) { Value = details[5] });
                    cmd.Parameters.Add(new SqlParameter("@DueDate", SqlDbType.VarChar) { Value = details[6] });
                    cmd.Parameters.Add(new SqlParameter("@Dealer_Name", SqlDbType.VarChar) { Value = details[7] });
                    cmd.Parameters.Add(new SqlParameter("@Dealer_Address1", SqlDbType.VarChar) { Value = details[8] });
                    cmd.Parameters.Add(new SqlParameter("@Dealer_City", SqlDbType.VarChar) { Value = details[9] });
                    cmd.Parameters.Add(new SqlParameter("@Transporter_Name", SqlDbType.VarChar) { Value = details[10] });
                    cmd.Parameters.Add(new SqlParameter("@Transport_Number", SqlDbType.VarChar) { Value = details[11] });
                    cmd.Parameters.Add(new SqlParameter("@Transport_Date", SqlDbType.VarChar) { Value = details[12] });
                    cmd.Parameters.Add(new SqlParameter("@Dealer_Code", SqlDbType.VarChar) { Value = details[13] });
                    cmd.Parameters.Add(new SqlParameter("@Transporter_Code", SqlDbType.VarChar) { Value = details[14] });
                    cmd.Parameters.Add(new SqlParameter("@Dealer_Address2", SqlDbType.VarChar) { Value = details[15] });
                    cmd.Parameters.Add(new SqlParameter("@Dealer_Address3", SqlDbType.VarChar) { Value = details[16] });
                    cmd.Parameters.Add(new SqlParameter("@Dealer_Address4", SqlDbType.VarChar) { Value = details[17] });
                    cmd.Parameters.Add(new SqlParameter("@Invoice_data_Received_Date", SqlDbType.VarChar) { Value = System.DateTime.Now.Date.ToString("dd/MMM/yyyy") });
                    cmd.Parameters.Add(new SqlParameter("@Physical_Invoice_Received_Date", SqlDbType.VarChar) { Value = details[19] });
                    cmd.Parameters.Add(new SqlParameter("@IMEX_DEAL_NUMBER", SqlDbType.VarChar) { Value = details[20] });
                    cmd.Parameters.Add(new SqlParameter("@StepDate", SqlDbType.VarChar) { Value = details[21] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Data_Received_Date", SqlDbType.VarChar) { Value = details[22] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Number", SqlDbType.VarChar) { Value = details[23] });
                    cmd.Parameters.Add(new SqlParameter("@Payment_Received_Date", SqlDbType.VarChar) { Value = details[24] });
                    cmd.Parameters.Add(new SqlParameter("@Utr_Number", SqlDbType.VarChar) { Value = details[25] });
                    cmd.Parameters.Add(new SqlParameter("@Invoice_Status", SqlDbType.VarChar) { Value = details[26] });
                    cmd.Parameters.Add(new SqlParameter("@TradeOp_Selected_Invoice_Flag", SqlDbType.VarChar) { Value = details[27] });
                    cmd.Parameters.Add(new SqlParameter("@TradeOp_Selected_Invoice_Date", SqlDbType.VarChar) { Value = details[28] });
                    cmd.Parameters.Add(new SqlParameter("@FileID", SqlDbType.VarChar) { Value = details[29] });
                    cmd.Parameters.Add(new SqlParameter("@Status", SqlDbType.VarChar) { Value = details[30] });
                    cmd.Parameters.Add(new SqlParameter("@Ord_Inv_ID", SqlDbType.VarChar) { Value = details[31] });
                    cmd.Parameters.Add(new SqlParameter("@Cash_Ops_ID", SqlDbType.VarChar) { Value = details[32] });
                    cmd.Parameters.Add(new SqlParameter("@F1_MIS", SqlDbType.VarChar) { Value = details[33] });
                    cmd.Parameters.Add(new SqlParameter("@F2_MIS", SqlDbType.VarChar) { Value = details[34] });
                    cmd.Parameters.Add(new SqlParameter("@F3_MIS", SqlDbType.VarChar) { Value = details[35] });
                    cmd.Parameters.Add(new SqlParameter("@F4_MIS", SqlDbType.VarChar) { Value = details[36] });
                    cmd.Parameters.Add(new SqlParameter("@Trade_OPs_Remarks", SqlDbType.VarChar) { Value = details[37] });
                    cmd.Parameters.Add(new SqlParameter("@LoginID_TradeOps", SqlDbType.VarChar) { Value = details[38] });
                    
                    cmd.Parameters.Add(new SqlParameter("@LoginID_CashOps", SqlDbType.VarChar) { Value = details[39] });
                    cmd.Parameters.Add(new SqlParameter("@F7_MIS", SqlDbType.VarChar) { Value = details[40] });
                    cmd.Parameters.Add(new SqlParameter("@TradeopsFileID", SqlDbType.VarChar) { Value = details[41] });
                    
                    //cmd.ExecuteReader();  
                    using (DbDataAdapter adapter = dbFactory.CreateDataAdapter())
                    {
                        adapter.SelectCommand = cmd;
                        adapter.Fill(dataTable);
                    }

                    cmd.Connection.Close();
                    return dataTable;
                }
            }
            catch (Exception ex) { clserr.WriteErrorToTxtFile(ex.Message + "Service", "Insert_InvoiceDetails", ""); return null; }
        }
        public static DataTable Insert_InvoiceCancel_DescDetails(string Task, string[] details)
        {
            DataTable dataTable = new DataTable();
            try
            {                
                //using var cmd = db.LoginMSTs.FromSqlRaw($"SP_FTPayment");
                using (var db = new Entities.DatabaseContext())
                {
                    DbConnection connection = db.Database.GetDbConnection();
                    using var cmd = db.Database.GetDbConnection().CreateCommand();
                    DbProviderFactory dbFactory = DbProviderFactories.GetFactory(connection);
                    cmd.CommandText = "SP_Insert_InvoiceCancel_DescDetails";
                    //common
                    cmd.CommandType = CommandType.StoredProcedure;
                    if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();

                    cmd.Parameters.Add(new SqlParameter("@Task", SqlDbType.VarChar) { Value = Task });
                    cmd.Parameters.Add(new SqlParameter("@Invoice_ID", SqlDbType.VarChar) { Value = details[0] });
                    cmd.Parameters.Add(new SqlParameter("@Sr_No", SqlDbType.VarChar) { Value = details[1] });
                    cmd.Parameters.Add(new SqlParameter("@Invoice_Number", SqlDbType.VarChar) { Value = details[2] });
                    cmd.Parameters.Add(new SqlParameter("@Invoice_Amount", SqlDbType.VarChar) { Value = details[3] });
                    cmd.Parameters.Add(new SqlParameter("@Requested_Date", SqlDbType.VarChar) { Value = System.DateTime.Now.Date.ToString("dd/MMM/yyyy") });
                    cmd.Parameters.Add(new SqlParameter("@Updated_Date", SqlDbType.VarChar) { Value = details[5] });
                    cmd.Parameters.Add(new SqlParameter("@DO_number", SqlDbType.VarChar) { Value = details[6] });
                    cmd.Parameters.Add(new SqlParameter("@Order_ID", SqlDbType.VarChar) { Value = details[7] });
                    cmd.Parameters.Add(new SqlParameter("@Cancelled_Flag", SqlDbType.VarChar) { Value = details[8] });
                    cmd.Parameters.Add(new SqlParameter("@Reason", SqlDbType.VarChar) { Value = details[9] });
                    cmd.Parameters.Add(new SqlParameter("@F2_MIS", SqlDbType.VarChar) { Value = details[10] });
                    cmd.Parameters.Add(new SqlParameter("@DO_IN_Flag", SqlDbType.VarChar) { Value = details[11] });
                    cmd.Parameters.Add(new SqlParameter("@Authorize_Flag", SqlDbType.VarChar) { Value = details[12] });
                    cmd.Parameters.Add(new SqlParameter("@Deleted_By", SqlDbType.VarChar) { Value = details[13] });
                    cmd.Parameters.Add(new SqlParameter("@Deleted_On", SqlDbType.VarChar) { Value = details[14] });
                    cmd.Parameters.Add(new SqlParameter("@Authorized_By", SqlDbType.VarChar) { Value = details[15] });
                    cmd.Parameters.Add(new SqlParameter("@Authorized_On", SqlDbType.VarChar) { Value = details[16] });
                    
           
                    //cmd.ExecuteReader();  
                    using (DbDataAdapter adapter = dbFactory.CreateDataAdapter())
                    {
                        adapter.SelectCommand = cmd;
                        adapter.Fill(dataTable);
                    }

                    cmd.Connection.Close();
                    
                }
            }
            catch (Exception ex) { clserr.WriteErrorToTxtFile(ex.Message + "Service", "Insert_InvoiceCancel_DescDetails", ""); }
            return dataTable;
        }
        public static DataTable Insert_Order_RejectedDetails(string Task, string[] details)
        {
            try
            {
                DataTable dataTable = new DataTable();
                //using var cmd = db.LoginMSTs.FromSqlRaw($"SP_FTPayment");
                using (var db = new Entities.DatabaseContext())
                {
                    DbConnection connection = db.Database.GetDbConnection();
                    using var cmd = db.Database.GetDbConnection().CreateCommand();
                    DbProviderFactory dbFactory = DbProviderFactories.GetFactory(connection);
                    cmd.CommandText = "SP_Insert_Order_RejectedDetails";
                    //common
                    cmd.CommandType = CommandType.StoredProcedure;
                    if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();

                    cmd.Parameters.Add(new SqlParameter("@Task", SqlDbType.VarChar) { Value = Task });
                    cmd.Parameters.Add(new SqlParameter("@Order_ID_Rej", SqlDbType.VarChar) { Value = details[0] });
                    cmd.Parameters.Add(new SqlParameter("@Order_ID", SqlDbType.VarChar) { Value = details[1] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Date", SqlDbType.VarChar) { Value = details[2] });
                    cmd.Parameters.Add(new SqlParameter("@Record_Identifier", SqlDbType.VarChar) { Value = details[3] });
                    cmd.Parameters.Add(new SqlParameter("@DO_Number", SqlDbType.VarChar) { Value = details[4]});
                    cmd.Parameters.Add(new SqlParameter("@Do_Date", SqlDbType.VarChar) { Value = details[5] });
                    cmd.Parameters.Add(new SqlParameter("@Dealer_Code", SqlDbType.VarChar) { Value = details[6] });
                    cmd.Parameters.Add(new SqlParameter("@Dealer_Destination_Code", SqlDbType.VarChar) { Value = details[7] });
                    cmd.Parameters.Add(new SqlParameter("@Dealer_Outlet_Code", SqlDbType.VarChar) { Value = details[8] });
                    cmd.Parameters.Add(new SqlParameter("@Financier_Code", SqlDbType.VarChar) { Value = details[9] });
                    cmd.Parameters.Add(new SqlParameter("@Financier_Name", SqlDbType.VarChar) { Value = details[10] });
                    cmd.Parameters.Add(new SqlParameter("@Email_IDs", SqlDbType.VarChar) { Value = details[11] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Amount", SqlDbType.VarChar) { Value = details[12] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Status", SqlDbType.VarChar) { Value = details[13] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Data_Received_On", SqlDbType.VarChar) { Value = details[14] });
                    cmd.Parameters.Add(new SqlParameter("@Rejected_Reson", SqlDbType.VarChar) { Value = details[15] });
                    cmd.Parameters.Add(new SqlParameter("@EmailStatus", SqlDbType.VarChar) { Value = details[16] });


                    //cmd.ExecuteReader();  
                    using (DbDataAdapter adapter = dbFactory.CreateDataAdapter())
                    {
                        adapter.SelectCommand = cmd;
                        adapter.Fill(dataTable);
                    }

                    cmd.Connection.Close();
                    return dataTable;
                }
            }
            catch (Exception ex) { clserr.WriteErrorToTxtFile(ex.Message + "Service", "Insert_Order_RejectedDetails", ""); return null; }
        }
        public static DataTable Insert_Order_Invoice_RejectedDetails(string Task, string[] details)
        {
            try
            {
                DataTable dataTable = new DataTable();
                //using var cmd = db.LoginMSTs.FromSqlRaw($"SP_FTPayment");
                using (var db = new Entities.DatabaseContext())
                {
                    DbConnection connection = db.Database.GetDbConnection();
                    using var cmd = db.Database.GetDbConnection().CreateCommand();
                    DbProviderFactory dbFactory = DbProviderFactories.GetFactory(connection);
                    cmd.CommandText = "SP_Insert_Order_Invoice_RejectedDetails";
                    //common
                    cmd.CommandType = CommandType.StoredProcedure;
                    if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();
                    
                    cmd.Parameters.Add(new SqlParameter("@Task", SqlDbType.VarChar) { Value = Task });
                    cmd.Parameters.Add(new SqlParameter("@Ord_Inv_ID_Rej", SqlDbType.VarChar) { Value = details[0] });
                    cmd.Parameters.Add(new SqlParameter("@Ord_Inv_ID", SqlDbType.VarChar) { Value = details[1] });
                    cmd.Parameters.Add(new SqlParameter("@Order_ID", SqlDbType.VarChar) { Value = details[2] });
                    cmd.Parameters.Add(new SqlParameter("@Record_Identifier", SqlDbType.VarChar) { Value = details[3] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Inv_Number", SqlDbType.VarChar) { Value = details[4] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Inv_Amount", SqlDbType.VarChar) { Value = details[5] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Inv_Status", SqlDbType.VarChar) { Value = details[6] });
                    cmd.Parameters.Add(new SqlParameter("@Order_ID_Rej", SqlDbType.VarChar) { Value = details[7] });
                    
                    //cmd.ExecuteReader();  
                    using (DbDataAdapter adapter = dbFactory.CreateDataAdapter())
                    {
                        adapter.SelectCommand = cmd;
                        adapter.Fill(dataTable);
                    }

                    cmd.Connection.Close();
                    return dataTable;
                }
            }
            catch (Exception ex) { clserr.WriteErrorToTxtFile(ex.Message + "Service", "Insert_Order_Invoice_RejectedDetails", ""); return null; }
        }
        public static DataTable Insert_Order_DeleteDetails(string Task, string[] details)
        {
            try
            {
                DataTable dataTable = new DataTable();
                //using var cmd = db.LoginMSTs.FromSqlRaw($"SP_FTPayment");
                using (var db = new Entities.DatabaseContext())
                {
                    DbConnection connection = db.Database.GetDbConnection();
                    using var cmd = db.Database.GetDbConnection().CreateCommand();
                    DbProviderFactory dbFactory = DbProviderFactories.GetFactory(connection);
                    cmd.CommandText = "SP_Insert_Order_DeleteDetails";
                    //common
                    cmd.CommandType = CommandType.StoredProcedure;
                    if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();

                    cmd.Parameters.Add(new SqlParameter("@Task", SqlDbType.VarChar) { Value = Task });
                    cmd.Parameters.Add(new SqlParameter("@Order_Del_ID", SqlDbType.VarChar) { Value = details[0] });
                    cmd.Parameters.Add(new SqlParameter("@Order_No", SqlDbType.VarChar) { Value = details[1] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Delete_Date", SqlDbType.VarChar) { Value = details[2] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Delete_Time", SqlDbType.VarChar) { Value = details[3] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Del_Status", SqlDbType.VarChar) { Value = System.DateTime.Now.Date.ToString("dd/MMM/yyyy") });
                    cmd.Parameters.Add(new SqlParameter("@FileID", SqlDbType.VarChar) { Value = details[5] });
                    cmd.Parameters.Add(new SqlParameter("@Email_Flag", SqlDbType.VarChar) { Value = details[6] });
                    
                    //cmd.ExecuteReader();  
                    using (DbDataAdapter adapter = dbFactory.CreateDataAdapter())
                    {
                        adapter.SelectCommand = cmd;
                        adapter.Fill(dataTable);
                    }

                    cmd.Connection.Close();
                    return dataTable;
                }
            }
            catch (Exception ex) { clserr.WriteErrorToTxtFile(ex.Message + "Service", "Insert_Order_DeleteDetails", ""); return null; }
        }
        public static DataTable Insert_DO_DescDetails(string Task, string[] details)
        {
            try
            {
                DataTable dataTable = new DataTable();
                //using var cmd = db.LoginMSTs.FromSqlRaw($"SP_FTPayment");
                using (var db = new Entities.DatabaseContext())
                {
                    DbConnection connection = db.Database.GetDbConnection();
                    using var cmd = db.Database.GetDbConnection().CreateCommand();
                    DbProviderFactory dbFactory = DbProviderFactories.GetFactory(connection);
                    cmd.CommandText = "SP_Insert_DO_DescDetails";
                    //common
                    cmd.CommandType = CommandType.StoredProcedure;
                    if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();

                    cmd.Parameters.Add(new SqlParameter("@Task", SqlDbType.VarChar) { Value = Task });
                    cmd.Parameters.Add(new SqlParameter("@Order_ID", SqlDbType.VarChar) { Value = details[0] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Date", SqlDbType.VarChar) { Value = details[1] });
                    cmd.Parameters.Add(new SqlParameter("@Record_Identifier", SqlDbType.VarChar) { Value = details[2] });
                    cmd.Parameters.Add(new SqlParameter("@DO_Number", SqlDbType.VarChar) { Value = details[3] });
                    cmd.Parameters.Add(new SqlParameter("@Do_Date", SqlDbType.VarChar) { Value = details[4] });
                    cmd.Parameters.Add(new SqlParameter("@Dealer_Code", SqlDbType.VarChar) { Value = details[5] });
                    cmd.Parameters.Add(new SqlParameter("@Dealer_Destination_Code", SqlDbType.VarChar) { Value = details[6] });
                    cmd.Parameters.Add(new SqlParameter("@Dealer_Outlet_Code", SqlDbType.VarChar) { Value = details[7] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Amount", SqlDbType.VarChar) { Value = details[8] });
                    cmd.Parameters.Add(new SqlParameter("@DO_Requested_date", SqlDbType.VarChar) { Value = details[9] });
                    cmd.Parameters.Add(new SqlParameter("@DO_Updated_date", SqlDbType.VarChar) { Value = details[10] });
                    cmd.Parameters.Add(new SqlParameter("@Delete_Flag", SqlDbType.VarChar) { Value = details[11] });
                    cmd.Parameters.Add(new SqlParameter("@Reason", SqlDbType.VarChar) { Value = details[12] });
                    cmd.Parameters.Add(new SqlParameter("@DO_Rej_Flag", SqlDbType.VarChar) { Value = details[13] });
                    cmd.Parameters.Add(new SqlParameter("@F2_MIS", SqlDbType.VarChar) { Value = details[14] });
                    cmd.Parameters.Add(new SqlParameter("@DO_IN_Flag", SqlDbType.VarChar) { Value = details[15] });
                    cmd.Parameters.Add(new SqlParameter("@Authorize_Flag", SqlDbType.VarChar) { Value = details[16] });
                    cmd.Parameters.Add(new SqlParameter("@Deleted_By", SqlDbType.VarChar) { Value = details[17] });
                    cmd.Parameters.Add(new SqlParameter("@Deleted_On", SqlDbType.VarChar) { Value = details[18] });
                    cmd.Parameters.Add(new SqlParameter("@Authorized_By", SqlDbType.VarChar) { Value = details[19] });
                    cmd.Parameters.Add(new SqlParameter("@Authorized_On", SqlDbType.VarChar) { Value = details[20] });
                   

                    //cmd.ExecuteReader();  
                    using (DbDataAdapter adapter = dbFactory.CreateDataAdapter())
                    {
                        adapter.SelectCommand = cmd;
                        adapter.Fill(dataTable);
                    }

                    cmd.Connection.Close();
                    return dataTable;
                }
            }
            catch (Exception ex) { clserr.WriteErrorToTxtFile(ex.Message + "Service", "Insert_DO_DescDetails", ""); return null; }
        }
        public static DataTable Insert_Order_Desc_Details(string Task, string[] details)
        {
            try
            {
                DataTable dataTable = new DataTable();
                //using var cmd = db.LoginMSTs.FromSqlRaw($"SP_FTPayment");
                using (var db = new Entities.DatabaseContext())
                {
                    DbConnection connection = db.Database.GetDbConnection();
                    using var cmd = db.Database.GetDbConnection().CreateCommand();
                    DbProviderFactory dbFactory = DbProviderFactories.GetFactory(connection);
                    cmd.CommandText = "SP_Insert_Order_Desc_Details";
                    //common
                    cmd.CommandType = CommandType.StoredProcedure;
                    if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();              

                    cmd.Parameters.Add(new SqlParameter("@Task", SqlDbType.VarChar) { Value = Task });
                    cmd.Parameters.Add(new SqlParameter("@Order_ID", SqlDbType.VarChar) { Value = details[0] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Date", SqlDbType.VarChar) { Value = details[1] });
                    cmd.Parameters.Add(new SqlParameter("@Record_Identifier", SqlDbType.VarChar) { Value = details[2] });
                    cmd.Parameters.Add(new SqlParameter("@DO_Number", SqlDbType.VarChar) { Value = details[3] });
                    cmd.Parameters.Add(new SqlParameter("@Do_Date", SqlDbType.VarChar) { Value = details[4] });
                    cmd.Parameters.Add(new SqlParameter("@Dealer_Code", SqlDbType.VarChar) { Value = details[5] });
                    cmd.Parameters.Add(new SqlParameter("@Dealer_Destination_Code", SqlDbType.VarChar) { Value = details[6] });
                    cmd.Parameters.Add(new SqlParameter("@Dealer_Outlet_Code", SqlDbType.VarChar) { Value = details[7] });
                    cmd.Parameters.Add(new SqlParameter("@Financier_Code", SqlDbType.VarChar) { Value = details[8] });
                    cmd.Parameters.Add(new SqlParameter("@Financier_Name", SqlDbType.VarChar) { Value = details[9] });
                    cmd.Parameters.Add(new SqlParameter("@Email_IDs", SqlDbType.VarChar) { Value = details[10] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Amount", SqlDbType.VarChar) { Value = details[11] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Status", SqlDbType.VarChar) { Value = details[12] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Data_Received_On", SqlDbType.VarChar) { Value = details[13] });
                    cmd.Parameters.Add(new SqlParameter("@Cash_Ops_ID", SqlDbType.VarChar) { Value = details[14] });
                    cmd.Parameters.Add(new SqlParameter("@Ord_Rej_Flag", SqlDbType.VarChar) { Value = details[15] });
                    cmd.Parameters.Add(new SqlParameter("@ORD_Rej_Reason", SqlDbType.VarChar) { Value = details[16] });
                    //cmd.ExecuteReader();  
                    using (DbDataAdapter adapter = dbFactory.CreateDataAdapter())
                    {
                        adapter.SelectCommand = cmd;
                        adapter.Fill(dataTable);
                    }

                    cmd.Connection.Close();
                    return dataTable;
                }
            }
            catch (Exception ex) { clserr.WriteErrorToTxtFile(ex.Message + "Service", "Insert_Order_Desc_Details", ""); return null; }
        }
        public static DataTable Insert_Order_Invoice_Details(string Task, string[] details)
        {
            try
            {
                DataTable dataTable = new DataTable();
                //using var cmd = db.LoginMSTs.FromSqlRaw($"SP_FTPayment");
                using (var db = new Entities.DatabaseContext())
                {
                    DbConnection connection = db.Database.GetDbConnection();
                    using var cmd = db.Database.GetDbConnection().CreateCommand();
                    DbProviderFactory dbFactory = DbProviderFactories.GetFactory(connection);
                    cmd.CommandText = "SP_Insert_Order_Invoice_Details";
                    //common
                    cmd.CommandType = CommandType.StoredProcedure;
                    if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();

                    cmd.Parameters.Add(new SqlParameter("@Task", SqlDbType.VarChar) { Value = Task });
                    cmd.Parameters.Add(new SqlParameter("@Ord_Inv_ID", SqlDbType.VarChar) { Value = details[0] });
                    cmd.Parameters.Add(new SqlParameter("@Order_ID", SqlDbType.VarChar) { Value = details[1] });
                    cmd.Parameters.Add(new SqlParameter("@Record_Identifier", SqlDbType.VarChar) { Value = details[2] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Inv_Number", SqlDbType.VarChar) { Value = details[3] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Inv_Amount", SqlDbType.VarChar) { Value = details[4] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Inv_Status", SqlDbType.VarChar) { Value = details[5] });
                    

                    //cmd.ExecuteReader();  
                    using (DbDataAdapter adapter = dbFactory.CreateDataAdapter())
                    {
                        adapter.SelectCommand = cmd;
                        adapter.Fill(dataTable);
                    }

                    cmd.Connection.Close();
                    return dataTable;
                }
            }
            catch (Exception ex) { clserr.WriteErrorToTxtFile(ex.Message + "Service", "Insert_Order_Invoice_Details", ""); return null; }
        }
        public static DataTable Insert_Order_Desc_DelDetails(string Task, string[] details)
        {
            try
            {
                DataTable dataTable = new DataTable();
                //using var cmd = db.LoginMSTs.FromSqlRaw($"SP_FTPayment");
                using (var db = new Entities.DatabaseContext())
                {
                    DbConnection connection = db.Database.GetDbConnection();
                    using var cmd = db.Database.GetDbConnection().CreateCommand();
                    DbProviderFactory dbFactory = DbProviderFactories.GetFactory(connection);
                    cmd.CommandText = "SP_Insert_Order_Desc_Details";
                    //common
                    cmd.CommandType = CommandType.StoredProcedure;
                    if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();                   

                    cmd.Parameters.Add(new SqlParameter("@Task", SqlDbType.VarChar) { Value = Task });
                    cmd.Parameters.Add(new SqlParameter("@Order_ID_Del", SqlDbType.VarChar) { Value = details[0] });
                    cmd.Parameters.Add(new SqlParameter("@Order_ID", SqlDbType.VarChar) { Value = details[1] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Date", SqlDbType.VarChar) { Value = details[2] });
                    cmd.Parameters.Add(new SqlParameter("@Record_Identifier", SqlDbType.VarChar) { Value = details[3] });
                    cmd.Parameters.Add(new SqlParameter("@DO_Number", SqlDbType.VarChar) { Value = System.DateTime.Now.Date.ToString("dd/MMM/yyyy") });
                    cmd.Parameters.Add(new SqlParameter("@Do_Date", SqlDbType.VarChar) { Value = details[5] });
                    cmd.Parameters.Add(new SqlParameter("@Dealer_Code", SqlDbType.VarChar) { Value = details[6] });
                    cmd.Parameters.Add(new SqlParameter("@Dealer_Destination_Code", SqlDbType.VarChar) { Value = details[7] });
                    cmd.Parameters.Add(new SqlParameter("@Dealer_Outlet_Code", SqlDbType.VarChar) { Value = details[8] });
                    cmd.Parameters.Add(new SqlParameter("@Financier_Code", SqlDbType.VarChar) { Value = details[9] });
                    cmd.Parameters.Add(new SqlParameter("@Financier_Name", SqlDbType.VarChar) { Value = details[10] });
                    cmd.Parameters.Add(new SqlParameter("@Email_IDs", SqlDbType.VarChar) { Value = details[11] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Amount", SqlDbType.VarChar) { Value = details[12] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Status", SqlDbType.VarChar) { Value = details[13] });
                    cmd.Parameters.Add(new SqlParameter("@Order_Data_Received_On", SqlDbType.VarChar) { Value = details[14] });
                    cmd.Parameters.Add(new SqlParameter("@Cash_Ops_ID", SqlDbType.VarChar) { Value = details[15] });
                    //cmd.ExecuteReader();  
                    using (DbDataAdapter adapter = dbFactory.CreateDataAdapter())
                    {
                        adapter.SelectCommand = cmd;
                        adapter.Fill(dataTable);
                    }

                    cmd.Connection.Close();
                    return dataTable;
                }
            }
            catch (Exception ex) { clserr.WriteErrorToTxtFile(ex.Message + "Service", "Insert_Order_Desc_DelDetails", ""); return null; }
        }
        public static bool File_MailUpdate(string File_Mail_Time, string File_Mail_Date,string File_Mail_Name)
        {
            try
            {
                using (var db = new Entities.DatabaseContext())
                {

                    var rec = db.Set<File_MailUpdate>().FromSqlRaw("update File_Mail set File_Mail_Time='" + File_Mail_Time + "',File_Mail_Date='" + File_Mail_Date + "' where File_Mail_Name='F5'"); //.ToList();

                    //var rec = db.File_MailUpdates.SingleOrDefault(a => a.File_Mail_Names == File_Mail_Name);

                    if (rec != null)
                    {
                        //rec.File_Mail_Time = File_Mail_Time;
                        //rec.File_Mail_Dates = File_Mail_Date;
                        //db.SaveChanges();
                        return true;
                    }

                }
            }
            catch(Exception ex)
            {
                clserr.WriteErrorToTxtFile(ex.Message + "Service", "File_MailUpdate", "");
                return false;
            }
            return false;
        }
        public static DataTable getDetails(string Task, string Search1, string Search2, string Search3, string Search4, string Search5, string Search6, string Search7)
        {
            try
            {
                DataTable dataTable = new DataTable();
                //using var cmd = db.LoginMSTs.FromSqlRaw($"SP_FTPayment");
                using (var db = new Entities.DatabaseContext())
                {
                    DbConnection connection = db.Database.GetDbConnection();
                    using var cmd = db.Database.GetDbConnection().CreateCommand();
                    DbProviderFactory dbFactory = DbProviderFactories.GetFactory(connection);
                    cmd.CommandText = "SP_GetDetails";
                    //common
                    cmd.CommandType = CommandType.StoredProcedure;
                    if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();
                    cmd.Parameters.Add(new SqlParameter("@Task", SqlDbType.VarChar) { Value = Task });
                    cmd.Parameters.Add(new SqlParameter("@Search1", SqlDbType.VarChar) { Value = Search1 });
                    cmd.Parameters.Add(new SqlParameter("@Search2", SqlDbType.VarChar) { Value = Search2 });
                    cmd.Parameters.Add(new SqlParameter("@Search3", SqlDbType.VarChar) { Value = Search3 });
                    cmd.Parameters.Add(new SqlParameter("@Search4", SqlDbType.VarChar) { Value = Search4 });
                    cmd.Parameters.Add(new SqlParameter("@Search5", SqlDbType.VarChar) { Value = Search5 });
                    cmd.Parameters.Add(new SqlParameter("@Search6", SqlDbType.VarChar) { Value = Search6 });
                    cmd.Parameters.Add(new SqlParameter("@Search7", SqlDbType.VarChar) { Value = Search7 });
                    //cmd.ExecuteReader();  
                    using (DbDataAdapter adapter = dbFactory.CreateDataAdapter())
                    {
                        adapter.SelectCommand = cmd;
                        adapter.Fill(dataTable);
                    }
                    cmd.Connection.Close();
                    return dataTable;
                }
            }
            catch (Exception ex) { clserr.WriteErrorToTxtFile(ex.Message + "Service", "getDetails", ""); return null; }
        }
        public static DataTable InsertDetails(string Task, string Search1, string Search2, string Search3, string Search4, string Search5, string Search6, string Search7)
        {
            try
            {
                DataTable dataTable = new DataTable();
                //using var cmd = db.LoginMSTs.FromSqlRaw($"SP_FTPayment");
                using (var db = new Entities.DatabaseContext())
                {

                    DbConnection connection = db.Database.GetDbConnection();
                    using var cmd = db.Database.GetDbConnection().CreateCommand();
                    DbProviderFactory dbFactory = DbProviderFactories.GetFactory(connection);
                    cmd.CommandText = "SP_InsertDetails";

                    //common
                    cmd.CommandType = CommandType.StoredProcedure;
                    if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();
                    cmd.Parameters.Add(new SqlParameter("@Task", SqlDbType.VarChar) { Value = Task });
                    cmd.Parameters.Add(new SqlParameter("@Search1", SqlDbType.VarChar) { Value = Search1 });
                    cmd.Parameters.Add(new SqlParameter("@Search2", SqlDbType.VarChar) { Value = Search2 });
                    cmd.Parameters.Add(new SqlParameter("@Search3", SqlDbType.VarChar) { Value = Search3 });
                    cmd.Parameters.Add(new SqlParameter("@Search4", SqlDbType.VarChar) { Value = Search4 });
                    cmd.Parameters.Add(new SqlParameter("@Search5", SqlDbType.VarChar) { Value = Search5 });
                    cmd.Parameters.Add(new SqlParameter("@Search6", SqlDbType.VarChar) { Value = Search6 });
                    cmd.Parameters.Add(new SqlParameter("@Search7", SqlDbType.VarChar) { Value = Search7 });
                    //cmd.ExecuteReader();  
                    using (DbDataAdapter adapter = dbFactory.CreateDataAdapter())
                    {
                        adapter.SelectCommand = cmd;
                        adapter.Fill(dataTable);
                    }

                    cmd.Connection.Close();
                    return dataTable;
                }
            }
            catch (Exception ex) { clserr.WriteErrorToTxtFile(ex.Message + "Service", "InsertDetails", ""); return null; }
        }
        public static DataTable UpdateDetails(string Task, string Search1, string Search2, string Search3, string Search4, string Search5, string Search6, string Search7)
        {
            try
            {
                DataTable dataTable = new DataTable();
                using (var db = new Entities.DatabaseContext())
                {
                    DbConnection connection = db.Database.GetDbConnection();
                    using var cmd = db.Database.GetDbConnection().CreateCommand();
                    DbProviderFactory dbFactory = DbProviderFactories.GetFactory(connection);
                    cmd.CommandText = "SP_UpdateDetails";

                    cmd.CommandType = CommandType.StoredProcedure;
                    if (cmd.Connection.State != ConnectionState.Open) cmd.Connection.Open();
                    cmd.Parameters.Add(new SqlParameter("@Task", SqlDbType.VarChar) { Value = Task });
                    cmd.Parameters.Add(new SqlParameter("@Search1", SqlDbType.VarChar) { Value = Search1 });
                    cmd.Parameters.Add(new SqlParameter("@Search2", SqlDbType.VarChar) { Value = Search2 });
                    cmd.Parameters.Add(new SqlParameter("@Search3", SqlDbType.VarChar) { Value = Search3 });
                    cmd.Parameters.Add(new SqlParameter("@Search4", SqlDbType.VarChar) { Value = Search4 });
                    cmd.Parameters.Add(new SqlParameter("@Search5", SqlDbType.VarChar) { Value = Search5 });
                    cmd.Parameters.Add(new SqlParameter("@Search6", SqlDbType.VarChar) { Value = Search6 });
                    cmd.Parameters.Add(new SqlParameter("@Search7", SqlDbType.VarChar) { Value = Search7 });
                    using (DbDataAdapter adapter = dbFactory.CreateDataAdapter())
                    {
                        adapter.SelectCommand = cmd;
                        adapter.Fill(dataTable);
                    }

                    cmd.Connection.Close();
                    return dataTable;
                }
            }
            catch (Exception ex) { clserr.WriteErrorToTxtFile(ex.Message + "Service", "UpdateDetails", ""); return null; }
        }

    }
}
