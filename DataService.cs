using System;
using System.IO;
using System.Reflection;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using Microsoft.Extensions.Hosting;

namespace dotnet_core_popgmail {
    public class DataService : IHostedService, IDisposable
    {
        private Timer _timer;
      
        public Task StartAsync(CancellationToken cancellationToken)
        {
            Logger.WriteLog("Start Service");

            _timer = new Timer(
                (e) => Processing(), 
                null, 
                TimeSpan.Zero, 
                TimeSpan.FromMinutes(5));
            return Task.CompletedTask;
        }

        private void Processing(){
            
            MailSettings mSettings = new MailSettings();
            MailClient _mailClientObj = new MailClient(mSettings.Read());
            _mailClientObj.Read("EXP_NMB_WMLOT");

            DbSettings dbSettings = new DbSettings();
            InvoiceExp mInvoiceExp = new InvoiceExp(dbSettings.GetConnectionString());
            mInvoiceExp.ReadExcelToDb();
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            Logger.WriteLog("Stop Service");

            _timer?.Change(Timeout.Infinite, 0);
            return Task.CompletedTask;
        }

          public void Dispose()
        {
            _timer?.Dispose();
        }

    }
}

