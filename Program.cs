using System;
using System.IO;

using MailKit.Net.Imap;
using MailKit.Search;
using MailKit;
using MimeKit;

namespace dotnet_core_popgmail
{
    class Program
    {

        static private MailClient _mailClientObj;
        static void Main(string[] args)
        {

            MailSettings mSettings = new MailSettings();
            _mailClientObj = new MailClient(mSettings.Read());
            _mailClientObj.Read("EXP_NMB_WMLOT");

            DbSettings dbSettings = new DbSettings();
            InvoiceExp mInvoiceExp = new InvoiceExp(dbSettings.GetConnectionString());
            mInvoiceExp.ReadExcelToDb();
            
        }
    }
}
