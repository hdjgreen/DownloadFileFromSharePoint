using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Security;
using System.Net;
using System.IO;
using Windows.Security.Credentials;

namespace DownloadFileFromSharePoint
{
    class Program
    {
        static int Main(string[] args)
        {
            string sharePointSiteURL;
            string fileURL;
            string userName;
            string password;
            string desPath;

            if (5 != args.Length)
            {
                Console.WriteLine("The number of parameters is wrong, please check and try again!");
                Console.WriteLine("The correct format should be:");
                Console.WriteLine("DownloadFileFromSharePoint.exe <sharePointSiteURL> <fileURL> <userName> <password> <desPath>");
                Console.WriteLine("For example: ");
                Console.WriteLine(@"DownloadFileFromSharePoint.exe ""https://microsoft.sharepoint.com"" ""/teams/XXXX/filename.xlsx"" <your account, email name> <your password> ""E:\temp""");
                return 1;  //return 1 if the number of parameter is wrong.
            }
            else
            {
                sharePointSiteURL = args[0];
                fileURL = args[1];
                userName = args[2];
                password = args[3];
                desPath = args[4];
            }

            try
            {
                //using (var ctx = GetSPOContext(new Uri(sharePointSiteURL), userName, password))
                using (var ctx = GetSPOContext(new Uri(sharePointSiteURL)))
                {
                    var web = ctx.Web;
                    DownloadFile(web, fileURL, desPath);
                    Console.WriteLine("Download completed!");
                    return 0;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error Occured: " + ex.Message);
                return 2;  // return 2 if encountered exception.
            }
        }

        private static void DownloadFile(Web web, string fileUrl, string targetPath)
        {
            var ctx = (ClientContext)web.Context;
            ctx.ExecuteQuery();
            using (var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, fileUrl))
            {
                var fileName = Path.Combine(targetPath, Path.GetFileName(fileUrl));
                using (var fileStream = System.IO.File.Create(fileName))
                {
                    fileInfo.Stream.CopyTo(fileStream);
                }
            }
        }

        private static ClientContext GetSPOContext(Uri webUri, string userName, string password)
        {
            var securePassword = new SecureString();
            foreach (var ch in password) securePassword.AppendChar(ch);
            return new ClientContext(webUri) { Credentials = new SharePointOnlineCredentials(userName, securePassword) };
        }

        private static ClientContext GetSPOContext(Uri webUri)
        {

            //SharePointOnlineCredentials userCredentials = System.Net.CredentialCache.DefaultCredentials as SharePointOnlineCredentials;
            //return new ClientContext(webUri) { Credentials = userCredentials };

            NetworkCredential networkCredentials = GetCredential("v-dejhua@microsoft.com");
            if (networkCredentials != null && !string.IsNullOrEmpty(networkCredentials.UserName))
            {   // works only when stored in the 'Web Credentials' not as Windows Credentials :(
                return new ClientContext(webUri) { Credentials = new SharePointOnlineCredentials(networkCredentials.UserName, networkCredentials.SecurePassword) };
            }
            else
            {   // default code from MSDN, does not work for SharePoint Online
                return null;
            }
        }

        public static NetworkCredential GetCredential(string userName)
        {
            PasswordCredential credential = GetCredentialFromLocker(userName);
            if (credential == null)
                return null;

            credential.RetrievePassword();

            var networkCred = new NetworkCredential(credential.UserName, credential.Password);
            return networkCred;
        }

        private static PasswordCredential GetCredentialFromLocker(string userName)
        {
            PasswordCredential credential = null;
            IReadOnlyList<PasswordCredential> credentialList = null;

            var vault = new PasswordVault();
            try
            {
                credentialList = vault.FindAllByUserName(userName);
            }
            catch
            {
                // log error  
            }
            if (credentialList == null)
                credentialList = vault.RetrieveAll();

            if (credentialList != null && credentialList.Count > 0)
            {
                if (credentialList.Count == 1)
                {
                    credential = credentialList[0];
                }
                else
                {
                    // manage issue when multiple user names
                }
            }
            return credential;
        }

    }
}
