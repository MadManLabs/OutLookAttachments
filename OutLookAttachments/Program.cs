using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using RatiocinationLibrary;

/* Save all attachments from email;
 * output all files into directory for post-processing.
 * https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.aspx
 */

namespace OutLookAttachments
{
    class Program
    {
        public static string savepath = "D:\\attach";

        static void Main(string[] args)
        {
            Log.Initialize();
            _Application olApp = new Outlook.Application();

            Directory.CreateDirectory(savepath);

            foreach (Store store in olApp.Session.Stores)
            {
                Folder root =
                store.GetRootFolder() as Folder;
                EnumerateFolders(root);
                Log.TextEntry(Environment.NewLine);
            }

            olApp = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Log.TextEntry("Press Enter to Exit", ConsoleColor.Red);
            Console.ReadLine();
        }
        
        // Uses recursion to enumerate Outlook subfolders.
        private static void EnumerateFolders(Folder folder)
        {
            Folders childFolders =
                folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Folder childFolder in childFolders)
                {
                    // Call EnumerateFolders using childFolder.
                    EnumerateFolders(childFolder);
                    ParseFolder(childFolder);
                }
            }
        }

        private static void ParseFolder(Folder folder)
        {
            // Write the folder path.
            Log.TextEntry(folder.FolderPath, ConsoleColor.Green);
            Outlook.Items attachItems = folder.Items;

            string filter = "@SQL=" + "urn:schemas:httpmail:hasattachment = True";
            attachItems = attachItems.Restrict(filter);
            if (attachItems.Count > 0)
            {
                try
                {
                    foreach (Outlook.MailItem mail in attachItems)
                    {
                        try
                        {
                            foreach (Outlook.Attachment attach in mail.Attachments)
                            {
                                try
                                {
                                    if (attach.Size > 0)
                                    {
                                        saveFile(mail, attach);
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    Log.Exception(ex);
                                }
                            }
                        }
                        catch(System.Exception ex)
                        {
                            Log.Exception(ex);
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    Log.Exception(ex);
                }
            }
            attachItems = null;
            folder = null;
        }

        private static void saveFile(MailItem mail, Attachment attach)
        {
            //SAVE FILE
            Log.TextEntry(mail.Subject, ConsoleColor.Cyan);
            Log.TextEntry(attach.FileName, ConsoleColor.Cyan);
            var fileName = mail.Subject + " " + attach.FileName;
            foreach (var c in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(c, '-');
            }
            attach.SaveAsFile(Path.Combine(savepath, fileName));
        }
    }
}


