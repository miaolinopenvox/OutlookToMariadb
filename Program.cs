// See https://aka.ms/new-console-template for more information

using IniParser;
using MySql.Data.MySqlClient;
using OutlookToMariadb;
using System.Diagnostics;
using MSOutlook = Microsoft.Office.Interop.Outlook;

// Init System.Diagnostics.TraceSource
static TraceSource InitLogger(string AppName)
{
    var Logger = new TraceSource(AppName);
    Logger.Listeners.Remove("Default");

    Logger.Switch = new SourceSwitch("sourceSwitch", "Verbose");

    TextWriterTraceListener textListener = new($"{AppName}.log");
    textListener.Filter = new EventTypeFilter(SourceLevels.Verbose);
    Logger.Listeners.Add(textListener);

    ConsoleTraceListener console = new(true);
    console.Filter = new EventTypeFilter(SourceLevels.Verbose);
    console.Name = "console";
    Logger.Listeners.Add(console);

    return Logger;
}

//always try create database and all tables
static void CreateDb(MySqlServer svr)
{
    var conn = svr.Connect();
    var sql = File.ReadAllText("outlook.sql");
    var script = new MySqlScript(conn);
    script.Query = sql;
    int count = script.Execute();
}

// Update folders in Outlook to Mariadb
static void UpdateFolders(OutlookSnapshot sn, MySqlServer db)
{
    db.TruncateTable("folders");

    foreach(var f in sn.OutlookFolders)
        db.ExecNonQuery($"insert into folders values('{f.FullFolderPath}','{f.StoreID}','{f.EntryID}')");
}



var AppName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
var Logger = InitLogger(AppName);

// Load config from Ini file
var IniParser = new FileIniDataParser();
IniParser.Parser.Configuration.CommentString = "#";
var Ini = IniParser.ReadFile($"{AppName}.ini");
Logger.TraceEvent(TraceEventType.Information, 1, "Config File:");
Logger.TraceEvent(TraceEventType.Information, 2, $"{Environment.NewLine}{Ini}");

var SyncFolderNames = Ini["outlook"]["SyncFolders"].Split(";").Select(e=>e.Trim()).ToList();
Logger.TraceEvent(TraceEventType.Information, 1, $"SyncFolders:{Environment.NewLine}{string.Join($"{Environment.NewLine}", SyncFolderNames)}");
var Filter = Ini["outlook"]["Filter"] ?? "";
Filter = Filter.Trim();

// Connect to mysql
var mySqlServer = new MySqlServer(
    Ini["mysql"]["server"], 
    Ini["mysql"]["user"], 
    Ini["mysql"]["pass"], 
    Ini["mysql"]["database"]);

var conn = mySqlServer.Connect();
Logger.TraceEvent(TraceEventType.Information, 1, $"Mysql connected");

// try create database everytime.
Logger.TraceEvent(TraceEventType.Information, 1, $"Try create database");
CreateDb(mySqlServer);


// Conect to outlook
var Outlook = new MSOutlook.Application();
Logger.TraceEvent(TraceEventType.Information, 1, $"Outlook connected");

var syncInterval = Int32.Parse(Ini["system"]["SyncInterval"] ?? "30");
syncInterval = syncInterval < 1 ? 1: syncInterval;  // minimal sync interval 1 minutes
var getActionInterval = Int32.Parse(Ini["system"]["SyncInterval"] ?? "1");
getActionInterval = getActionInterval < 1 ? 1 : getActionInterval;  // minimal get action interval 1 minutes

/*
 * Loop sync email to database
 * Delete from database while mails not exists.
 * Insert new mails to database.
 */

var exit = false;
while (!exit)
{
    Logger.TraceEvent(TraceEventType.Information, 1, $"Sync Start at {DateTime.Now}, Pres 'X' to exit");

    Logger.TraceEvent(TraceEventType.Information, 1, $"Load Outlook emails");
    var snapshot = new OutlookSnapshot(Outlook, SyncFolderNames, Filter);
    Logger.TraceEvent(TraceEventType.Information, 1, $"{snapshot.OutlookFolders.Count} Folders Found");
    UpdateFolders(snapshot, mySqlServer);

    Logger.TraceEvent(TraceEventType.Information, 1, $"Load Emails from Outlook");
    snapshot.LoadEmails(Logger);

    Logger.TraceEvent(TraceEventType.Information, 1, $"Get emails from Mariadb");
    var tableEmail = new TableEmail(mySqlServer);
    var emailsInDb = tableEmail.GetKeyOfEmails();
    Logger.TraceEvent(TraceEventType.Information, 1, $"{emailsInDb.Count} emails in Mariadb");

    var emailsToDelete = emailsInDb.Where(e => !snapshot.Idx_StoreId_EntryId.Keys.Contains(e.Item1)
        && !snapshot.Idx_StoreId_EntryId[e.Item1].Keys.Contains(e.Item2)).ToList();
    Logger.TraceEvent(TraceEventType.Information, 1, $"Delete {emailsToDelete.Count()} mails from  Mariadb");
    tableEmail.DeleteByStoreIdAndEntryIds(emailsToDelete);

    Logger.TraceEvent(TraceEventType.Information, 1, $"Finding new mails to insert into Mariadb");
    var emailsToInsert = snapshot.Mails.Where(e => !emailsInDb.Contains(new Tuple<string, string>(e.Folder.StoreID, e.MailItem.EntryID))).ToList();
    Logger.TraceEvent(TraceEventType.Information, 1, $"{emailsToInsert.Count()} mails will be insert to Mariadb");
    tableEmail.BatchInsertEmails(emailsToInsert);

    Logger.TraceEvent(TraceEventType.Information, 1, $"{Environment.NewLine}Wait for {syncInterval} seconds, Pres 'X' to exit");
    var syncIntervalSeconds = syncInterval * 60;
    var oldt = DateTime.Now;
    while(!exit)
    {
        var actions = tableEmail.GetActions();
        if(actions.Any())
        {
            Logger.TraceEvent(TraceEventType.Information, 1, $"Processing {actions.Count} actions");
            int cnt = 0;
            foreach(var a in actions)
            {
                var alist = a.Item3.Split(" ").Select(e=>e.Trim()).ToArray();
                Logger.TraceEvent(TraceEventType.Information, 1, $"Handling Command {cnt}:");
                switch (alist[0])
                {
                    case "delete":
                        {
                            var mail = snapshot.FindMailByStoreIdEntryId(a.Item1, a.Item2);
                            if (mail == null)
                                Logger.TraceEvent(TraceEventType.Information, 1, $"delete Error: Mail not found. storeid={a.Item1}, entryid={a.Item2}, action={a.Item3}");
                            else
                            {
                                mail.MailItem.Delete();
                                Logger.TraceEvent(TraceEventType.Information, 1, $"delete Success: storeid={a.Item1}, entryid={a.Item2}, action={a.Item3}");
                            }
                        }
                        break;

                    case "markreaded":
                        {
                            var mail = snapshot.FindMailByStoreIdEntryId(a.Item1, a.Item2);
                            if (mail == null)
                                Logger.TraceEvent(TraceEventType.Information, 1, $"markread Error: Mail not found. storeid={a.Item1}, entryid={a.Item2}, action={a.Item3}");
                            else
                            {
                                mail.MailItem.Delete();
                                Logger.TraceEvent(TraceEventType.Information, 1, $"markread Success: storeid={a.Item1}, entryid={a.Item2}, action={a.Item3}");
                            }
                        }
                        break;

                    default:
                        Logger.TraceEvent(TraceEventType.Information, 1, $"Unknown command {a.Item3}, storeid={a.Item1}, entryid={a.Item2}");
                        break;
                }
                cnt++;
            }
            oldt = DateTime.Now;
        }

        var interval = DateTime.Now - oldt;
        if (interval.TotalSeconds > syncIntervalSeconds)
        {
            Console.WriteLine("");
            Console.WriteLine("Start Sync mails");
            break;
        }

        Thread.Sleep(100);
        Console.Write($"\r{syncIntervalSeconds - interval.Seconds} seconds before next sync...");
        while(Console.KeyAvailable)
        {
            var k = Console.ReadKey();
            if (k.KeyChar == 'X')
            {
                exit = true;
                break;
            }
            else
                Console.WriteLine($"{Environment.NewLine}Press 'X' to exit");
        }
    }
    Console.WriteLine("");
}





