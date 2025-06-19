using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;

namespace _001TN0173.Entities
{
    class DatabaseContext : DbContext
    {
       
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            FunctionsCl fncl = new FunctionsCl();
            var builder = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory()).AddJsonFile("appsettings.json");
            var configuration = builder.Build();
            // string ConnectionString = configuration["connectionStrings:DefaultConnection"];
            // int arrayindx = ConnectionString.IndexOf("Password=");
            // int len = ConnectionString.Length;
            // string Password = ConnectionString.Substring(arrayindx+9,ConnectionString.Length- (arrayindx+9));
            //// string encname = fncl.Encrypt(Password[1], Password[1].ToString().Length);
            // string decname = fncl.Decrypt(Password, Password.ToString().Length);          
            var configuration1 = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build();
             DataService _dataService = new DataService(configuration.GetConnectionString("DefaultConnection"));
            optionsBuilder.UseSqlServer(configuration["connectionStrings:DefaultConnection"]);
           
            //with encryption
            //optionsBuilder.UseSqlServer(ConnectionString);

        }
        private static string GetParameters()
        {
            var builder = new ConfigurationBuilder()
                                .SetBasePath(Directory.GetCurrentDirectory())
                                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

            var InputFilePath = builder.Build().GetSection("MSILSettings").GetSection("InputFilePath").Value;
            var OutputFilePath = builder.Build().GetSection("MSILSettings").GetSection("OutputFilePath").Value;
            var BackupFilePath = builder.Build().GetSection("MSILSettings").GetSection("BackupFilePath").Value;
            var NonConvertedFile = builder.Build().GetSection("MSILSettings").GetSection("NonConvertedFile").Value;

            return $"The values of parameters are: {InputFilePath} and {OutputFilePath} and {BackupFilePath} and {NonConvertedFile}";
        }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<InvoiceFileCheck>().HasNoKey();
            modelBuilder.Entity<File_MailF5>().HasNoKey();
            modelBuilder.Entity<File_MailUpdate>().HasNoKey();
            modelBuilder.Entity<RectifyUpdateEr>().HasNoKey();

        }
        public DbSet<InvoiceFileCheck> InvoiceFileChecks { get; set; }
        public DbSet<File_MailF5> File_MailF5s { get; set; }

        public DbSet<File_MailUpdate> File_MailUpdates { get; set; }

        public DbSet<RectifyUpdateEr> RectifyUpdateErs { get; set; }
    }
}
