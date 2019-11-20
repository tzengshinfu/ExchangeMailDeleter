using Microsoft.Exchange.WebServices.Data;
using System;

namespace ExchangeMailDeleter {
    class Program {
        /// <summary>
        /// 批次刪除的筆數
        /// </summary>
        const int pageCount = 100;

        static void Main(string[] args) {
            try {
                if (args.Length != 3) {
                    Console.WriteLine("命令列必須是[保留天數,如=>30] [帳號,如=>domain\\userName或userName] [密碼]");

                    Environment.ExitCode = -1;
                    return;
                }

                var domain = "";
                var userName = "";
                var accountInfo = args[1].Split('\\');
                if (accountInfo.Length == 2) {
                    domain = accountInfo[0];
                    userName = accountInfo[1];
                }
                else if (accountInfo.Length == 1) {
                    userName = accountInfo[0];
                }
                else {
                    Console.WriteLine("帳號必須是合法格式,如=>domain\\userName或userName");

                    Environment.ExitCode = -1;
                    return;
                }
                var password = args[2];

                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010);
                service.Credentials = new WebCredentials(userName, password, domain);
                service.Url = new Uri(@"<Exchange主機web service asmx位址>");

                var remainingDays = int.Parse(args[0]) * -1;
                var searchFilter = new SearchFilter.IsLessThan(ItemSchema.DateTimeReceived, DateTime.Now.AddDays(remainingDays));
                var inboxMailCount = service.FindItems(WellKnownFolderName.Inbox, searchFilter, new ItemView(1)).TotalCount;
                var totalRunCount = (inboxMailCount / pageCount) + 1;

                for (var currentRunCount = 0; currentRunCount < totalRunCount; currentRunCount += 1) {
                    var findResults = service.FindItems(WellKnownFolderName.Inbox, searchFilter, new ItemView(pageCount));
                    foreach (var findResult in findResults) {
                        findResult.Delete(DeleteMode.HardDelete);
                    }
                }
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);

                Environment.ExitCode = -1;
            }            
        }
    }
}