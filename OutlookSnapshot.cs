using System.Diagnostics;
using MSOutlook = Microsoft.Office.Interop.Outlook;


namespace OutlookToMariadb
{
    public class Mail
    {
        public MSOutlook.MAPIFolder Folder;
        public MSOutlook.MailItem MailItem;

        public Mail(MSOutlook.MAPIFolder folder, MSOutlook.MailItem mailItem)
        {
            Folder = folder;
            MailItem = mailItem;
        }

        public override bool Equals(object? obj)
        {
            return obj is Mail mail &&
                Folder.StoreID == mail.Folder.StoreID &&
                MailItem.EntryID == mail.MailItem.EntryID;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Folder.EntryID, MailItem.EntryID);
        }
    }


    public class OutlookSnapshot
    {
        public MSOutlook.Application Outlook;
        
        // all folders in outlook
        public List<MSOutlook.MAPIFolder> OutlookFolders;

        // folders to be sync to mariadb
        public List<MSOutlook.MAPIFolder> SyncFolders;

        // all emails indexed by storeId and mailid
        public Dictionary<string, Dictionary<string, Mail>> Idx_StoreId_EntryId = new();

        public Dictionary<string, Dictionary<string, Mail>> Idx_FolderPath_EntryId = new ();

        public List<Mail> Mails = new();

        public readonly string Filter;

        public int Total;
        public int Errors;

        public OutlookSnapshot(MSOutlook.Application outlook, List<string> syncFolderNames, string filter)
        {
            Outlook = outlook;
            OutlookFolders = GetFolders();
            SyncFolders = OutlookFolders.Where(e => syncFolderNames.Contains(e.FullFolderPath)).ToList();
            Filter = filter;
            Total = SyncFolders.Sum(e=>e.Items.Count);
        }

        public Mail? FindMailByStoreIdEntryId(string storeid, string entryid)
        {
            if (Idx_StoreId_EntryId.TryGetValue(storeid, out var map))
            {
                if (map.TryGetValue(entryid, out var mail))
                    return mail;
                else
                    return null;
            }
            else
                return null;
        }

        public void AddMail(MSOutlook.MAPIFolder folder, MSOutlook.MailItem mail)
        {
            var newMail = new Mail(folder, mail);
            if (Idx_FolderPath_EntryId.TryGetValue(folder.FullFolderPath, out var smap))
                smap.Add(mail.EntryID, newMail);
            else
                Idx_FolderPath_EntryId.Add(folder.FullFolderPath, 
                    new Dictionary<string, Mail>() { {mail.EntryID, newMail}, } );

            if (Idx_StoreId_EntryId.TryGetValue(folder.StoreID, out var smap1))
                smap1.Add(mail.EntryID, newMail);
            else
                Idx_StoreId_EntryId.Add(folder.StoreID,
                    new Dictionary<string, Mail>() { { mail.EntryID, newMail}, });

            Mails.Add(newMail);
        }

        public int LoadEmails(TraceSource Logger)
        {
            int current = 0;
            foreach (var f in SyncFolders)
            {
                var items = string.IsNullOrEmpty(Filter) ? f.Items : f.Items.Restrict(Filter);
                foreach (var it in items)
                {
                    if (it is not MSOutlook.MailItem)
                    {
                        Console.WriteLine("");
                        Errors++;
                        current++;
                        Logger.TraceEvent(TraceEventType.Information, 1, $"Item {current} is not mail");
                        continue;
                    }

                    // 内存保留一份mail instance，方便后续操作
                    var mailitem = it as MSOutlook.MailItem;
                    AddMail(f, mailitem);
                    current++;

                    // not use PrintStatus for not line wrap
                    var s = $"\r{current} of {Total}, Errors：{Errors}";
                    Console.Write(s);
                }
            }
            Console.WriteLine("");
            return Total;
        }

        private static void AddFolder(List<MSOutlook.MAPIFolder> Folders, MSOutlook.MAPIFolder currentFolder)
        {
            Folders.Add(currentFolder);

            foreach (MSOutlook.MAPIFolder folder in currentFolder.Folders)
                AddFolder(Folders, folder);
        }

        public List<MSOutlook.MAPIFolder> GetFolders()
        {
            var Folders = new List<MSOutlook.MAPIFolder>();

            // get all the folder and subfolders.
            foreach (MSOutlook.MAPIFolder folder in Outlook.Session.Folders)
                AddFolder(Folders, folder);

            foreach (var f in Folders)
                Console.WriteLine(f.FullFolderPath);

            return Folders;
        }
    };
}
