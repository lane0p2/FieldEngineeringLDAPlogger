using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Diagnostics.Eventing;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Threading;
using Microsoft.Win32;
using System.Net;
using System.DirectoryServices.ActiveDirectory;
using System.Data.SqlClient;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Security.Cryptography;


namespace FieldEngineeringIPLog
{
    class Program
    {

        static int timeInt = 0;
        static RegistryKey rk = Registry.LocalMachine;
        static RegistryKey FE15;
        static RegistryKey FE152;
        static DateTime dtStart;
        static bool verbose = false;
        static bool veryVerbose = false;
        static bool ResetDefaults = false;
        static string server = "";
        static StreamWriter writer;
        static StreamWriter ErrWriter;
        static StreamWriter statwriter;
        static String TimeString = DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Year.ToString() + DateTime.Now.Hour.ToString().PadLeft(2, '0') + DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');
        static bool killLoop = false;
        static int LogVal = 0;
        static bool AllDCs = false;
        static Domain domain;
        static DomainControllerCollection DCList;
        static string SQLConnStr = "";
        static SqlConnection conn;
        static List<string> ServerList = new List<string>();
        static List<string> IPList = new List<string>();
        static List<string> SiteList = new List<string>();
        static ArrayList SkipDownDC = new ArrayList();
        static string oneDCName = "";
        static string oneDCSite = "";
        static string ExecutionPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        static private Object screenLock = new Object();

        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Usage();
            }
            else
            {
                if (args[0].ToString().StartsWith("/?"))
                {
                    Usage();
                }
                else
                {
                    try
                    {
                        Console.WriteLine("");
                        // Reset defaults------------------------------------------------------
                        if (args[0].ToString().ToUpper() == "/RESETDEFAULTS")
                        {
                            if (args[1].ToUpper().StartsWith("S:"))
                            {
                                server = args[1].Substring(2, (args[1].Length - 2));
                                if (server.ToUpper() == "DOMAIN")
                                {
                                    AllDCs = true;
                                    domain = Domain.GetComputerDomain();
                                    Console.WriteLine("Resetting all Domain Controllers in domain " + domain.Name);
                                }
                                else
                                { Console.WriteLine("Resetting server: " + server.ToString() + "\n"); }
                            }
                            ResetDefaults = true;
                            ResetFEValue("OFF");
                        }
                        // Reset defaults------------------------------------------------------

                        if (args.Length != 5)
                        {
                            if (ResetDefaults == false)
                            { Usage(); }
                        }
                        else
                        {
                            foreach (string item in args)
                            {
                                if (item.ToUpper().StartsWith("S:"))
                                {
                                    server = item.Substring(2, (item.Length - 2));
                                    
                                    if (server.Contains("."))
                                    {
                                        server = server.Split('.')[0].ToString();
                                    }

                                    if (server.ToUpper() == "DOMAIN")
                                    {
                                        AllDCs = true;
                                        domain = Domain.GetComputerDomain();
                                        OutputToScreen("Testing all Domain Controllers in domain: " + domain.Name, ConsoleColor.Green);
                                        CreateSkipDownDCList();
                                    }
                                    else
                                    { OutputToScreen("Testing Domain Controller: " + server.ToString() + "\n", ConsoleColor.Green); }
                                }

                                if (item.ToUpper().StartsWith("FILE:"))
                                {
                                    try
                                    {
                                        string fileName = Path.GetFileName(item.Substring(5, (item.Length - 5)));
                                        string OutputFileName = fileName.Split('.')[0] + "_" + TimeString + ".txt";
                                        CreateLogs(OutputFileName);
                                    }
                                    catch (Exception exp)
                                    { OutputToScreen("Error creating file: " + exp.Message.ToString(), ConsoleColor.Red); }
                                }

                                if (item.ToUpper().StartsWith("SQL:"))
                                {
                                    try
                                    {
                                        Console.WriteLine("\nCreating SQL Connection");
                                        SQLConnStr = item.Substring(4, (item.Length - 4));
                                        conn = new SqlConnection();
                                        conn.ConnectionString = SQLConnStr;
                                        conn.Open();
                                        Console.WriteLine("SQL Connection State: " + conn.State.ToString());
                                    }
                                    catch (Exception exp)
                                    { OutputToScreen("Error creating SQL connection: " + exp.Message.ToString(), ConsoleColor.Red); }
                                }

                                if (item.ToUpper() == "V")
                                { verbose = true; }

                                if (item.ToUpper() == "VV")
                                {
                                    verbose = true;
                                    veryVerbose = true;
                                }

                                if (item.ToUpper().StartsWith("L:"))
                                {
                                    bool b2 = Int32.TryParse((item.Substring(2, (item.Length - 2))), out LogVal);
                                    if (b2)
                                    {
                                        if (LogVal < 1)
                                        { LogVal = 1; }

                                        if (LogVal > 5)
                                        { LogVal = 5; }
                                    }
                                    else
                                    { OutputToScreen("Logging level must be a value between 1 and 5!", ConsoleColor.Red); }

                                }
                            }


                            bool b = Int32.TryParse(args[0].ToString(), out timeInt);
                            if (b)
                            {
                                Console.TreatControlCAsInput = false;
                                Console.CancelKeyPress += delegate
                                {
                                    killLoop = true;
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("\nCleaning up ...\n");
                                    Console.ResetColor();
                                    ResetFEValue("OFF");
                                    Thread.Sleep(5000);
                                    CalculateStats();
                                    CloseLogs();                                    
                                };

                                ResetFEValue("ON");

                                if (killLoop == false)
                                {
                                    dtStart = DateTime.Now;
                                    Console.WriteLine("");
                                    for (int i = timeInt; i > -1; i--)
                                    {
                                        if (killLoop == false)
                                        {
                                            Console.SetCursorPosition(0, Console.CursorTop);
                                            Console.Write("Sleeping for " + i.ToString() + " seconds...");
                                            Thread.Sleep(1000);
                                        }
                                    }
                                    Console.SetCursorPosition(0, Console.CursorTop - 1);
                                    Console.WriteLine("");

                                    ResetFEValue("OFF");
                                    ReadEventLogs();
                                    CloseLogs();
                                    CalculateStats();
                                }

                                OutputToScreen("\nFinished!\n", ConsoleColor.Green);
                            }
                            else
                            { OutputToScreen("\nThe time interval is not an integer!\n", ConsoleColor.Red); }

                        }

                    }
                    catch (Exception exp)
                    {
                        ResetFEValue("OFF");
                        OutputToScreen("Error in Main " + exp.Message.ToString(), ConsoleColor.Red);
                    }
                }
            }
            
        }
                
        static void CalculateStats()
        {
            
            try
            {
                string statfile = "LDAPQuertyStats_" + TimeString + ".log";
                File.Create(statfile).Close();
                statwriter = new StreamWriter(statfile);
                statwriter.AutoFlush = true;
            }
            catch (Exception exp)
            {
                OutputToScreen("Error creating stats log - " + exp.Message.ToString(), ConsoleColor.Red);
                ErrWriter.WriteLine("Error creating stats log - " + exp.Message.ToString());
            }


            try
            {
                Console.WriteLine("");
                Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                statwriter.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                Console.WriteLine("Server Occurrences:");
                statwriter.WriteLine("Server Occurrences:");

                if (AllDCs)
                {
                    DCList = domain.FindAllDiscoverableDomainControllers();
                    foreach (DomainController dc in DCList)
                    {
                        int total = ServerList.FindAll(s => s.Equals(dc.Name.Split('.')[0])).Count;
                        Console.WriteLine(dc.Name.Split('.')[0] + " x " + total.ToString());
                        statwriter.WriteLine(dc.Name.Split('.')[0] + " x " + total.ToString());
                    }
                }
                else
                {
                    int total = ServerList.FindAll(s => s.Equals(server)).Count;
                    Console.WriteLine(server + " x " + total.ToString());
                    statwriter.WriteLine(server + " x " + total.ToString());
                }

                Console.WriteLine("");
                statwriter.WriteLine("");
                Console.WriteLine("IP Occurrences:");
                statwriter.WriteLine("IP Occurrences:");

                ArrayList IPArray = new ArrayList();
                foreach (string IPItem in IPList)
                {
                    if (!IPArray.Contains(IPItem))
                    {
                        IPArray.Add(IPItem);
                        int total = IPList.FindAll(s => s.Equals(IPItem)).Count;
                        Console.WriteLine(IPItem + " x " + total.ToString());
                        statwriter.WriteLine(IPItem + " x " + total.ToString());
                    }
                }

                Console.WriteLine("");
                statwriter.WriteLine("");
                Console.WriteLine("Site Occurrences");
                statwriter.WriteLine("Site Occurrences");

                ArrayList SiteArray = new ArrayList();
                foreach (string site in SiteList)
                {
                    if (!SiteArray.Contains(site))
                    {
                        SiteArray.Add(site);
                        int total = SiteList.FindAll(s => s.Equals(site)).Count;
                        Console.WriteLine(site + " x " + total.ToString());
                        statwriter.WriteLine(site + " x " + total.ToString());
                    }
                }

                Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                statwriter.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                Console.WriteLine("");

                statwriter.Flush();
                statwriter.Close();
                statwriter.Dispose();
            }
            catch (Exception exp)
            {
                OutputToScreen("Error calculating stats - " + exp.Message.ToString(), ConsoleColor.Red);
                ErrWriter.WriteLine("Error calculating stats - " + exp.Message.ToString());
            }
        }
                        
        static void ReadEventLogs()
        {
            if (AllDCs)
            {
                try
                {
                    DCList = domain.FindAllDiscoverableDomainControllers();
                    foreach (DomainController dc in DCList)
                    {
                        if (SkipDownDC.Contains(dc.Name))
                        {
                            OutputToScreen("Skipping DC - " + dc.Name, ConsoleColor.Red);
                        }
                        else
                        {
                            DoReadEventLogs(timeInt, dc.Name.Split('.')[0], dc.SiteName);
                        }
                    }
                }
                catch (Exception exp)
                {
                    OutputToScreen("Error - Cannot find all Domain Controllers - " + exp.Message.ToString(), ConsoleColor.Red);
                }
            }
            else
            {
                try
                {
                    domain = Domain.GetComputerDomain();
                    DCList = domain.FindAllDiscoverableDomainControllers();

                    foreach (DomainController dc in DCList)
                    {
                        if (dc.Name.Split('.')[0].ToUpper() == server.ToUpper())
                        {
                            oneDCName = dc.Name;
                            oneDCSite = dc.SiteName;
                            DoReadEventLogs(timeInt, oneDCName, oneDCSite);
                            break;
                        }
                    }
                }
                catch (Exception exp)
                {
                    OutputToScreen("Error - Cannot find one Domain Controller - " + exp.Message.ToString(), ConsoleColor.Red);
                }
                
            }
        }

        static void CreateLogs(string OutputFileName)
        {
            try
            {
                Console.WriteLine("\nCreating output file");
                File.Create(OutputFileName).Close();
                writer = new StreamWriter(OutputFileName);
                writer.AutoFlush = true;
                writer.WriteLine('\u0022' + "LDAPServer" + '\u0022' + "," + '\u0022' + "LDAPServerSite" + '\u0022' + "," + '\u0022' + "ClientName" + '\u0022' + "," + '\u0022' + "TimeGenerated" + '\u0022' + "," + '\u0022' + "ClientIP" + '\u0022' + "," + '\u0022' + "ClientPort" + '\u0022' + "," + '\u0022' + "UserName" + '\u0022' + "," + '\u0022' + "StartingNode" + '\u0022' + "," + '\u0022' + "Filter" + '\u0022' + "," + '\u0022' + "SearchScope" + '\u0022' + "," + '\u0022' + "AttributeSelection" + '\u0022' + "," + '\u0022' + "ServerControls" + '\u0022' + "," + '\u0022' + "VisitedEntries" + '\u0022' + "," + '\u0022' + "ReturnedEntries" + '\u0022' + "," + '\u0022' + "UsedIndexes" + '\u0022' + "," + '\u0022' + "PagesReferenced" + '\u0022' + "," + '\u0022' + "PagesReadFromDisk" + '\u0022' + "," + '\u0022' + "PagesPreReadFromDisk" + '\u0022' + "," + '\u0022' + "CleanPagesModified" + '\u0022' + "," + '\u0022' + "DirtyPagesModified" + '\u0022' + "," + '\u0022' + "SearchTimeMS" + '\u0022' + "," + '\u0022' + "AttributesPreventingOptimization" + '\u0022');

                Console.WriteLine("Creating error file");
                File.Create("FieldEngineeringIPErrorLog.txt").Close();
                ErrWriter = new StreamWriter("FieldEngineeringIPErrorLog.txt");
                ErrWriter.WriteLine(DateTime.Now.ToString() + " Error log created");
                ErrWriter.AutoFlush = true;
            }
            catch (Exception exp)
            {
                OutputToScreen("Error creating logs - " + exp.Message.ToString(), ConsoleColor.Red);
            }
        }

        static void CloseLogs()
        {
            try
            {
                writer.Flush();
                writer.Close();
                writer.Dispose();
            }
            catch (Exception exp)
            {
                OutputToScreen("Error closing output log - " + exp.Message.ToString(), ConsoleColor.Red);
            }

            try
            {
                ErrWriter.WriteLine(DateTime.Now.ToString() + " Error log closed");
                ErrWriter.Flush();
                ErrWriter.Close();
                ErrWriter.Dispose();
            }
            catch (Exception exp)
            {
                OutputToScreen("Error closing error log - " + exp.Message.ToString(), ConsoleColor.Red);
            }
        }

        static void CreateSkipDownDCList()
        {
            DCList = domain.FindAllDiscoverableDomainControllers();
            foreach (DomainController dc in DCList)
            {
                Console.Write("Connecting ... " + dc.Name);
                
                try
                {
                    using (TcpClient client = new TcpClient())
                    {
                        client.ReceiveTimeout = 2;
                        client.SendTimeout = 2;                        
                        IAsyncResult result = client.BeginConnect(dc.Name, 135, null, null);
                        bool isConnected = result.AsyncWaitHandle.WaitOne(TimeSpan.FromSeconds(2));
                        
                        if (isConnected)
                        {
                            Console.WriteLine(" ... Success!");
                            client.Close();
                        }
                        else
                        {
                            Console.WriteLine(" ... Failure!");
                            SkipDownDC.Add(dc.Name);
                        }                        
                    }

                }
                catch (Exception exp)
                {
                    Console.WriteLine(" ... Failure!");
                    SkipDownDC.Add(dc.Name);
                    ErrWriter.WriteLine("Error creating skip list for " + dc.Name + " - " + exp.Message.ToString());
                }
            }

        }

        static void DoReadEventLogs(int s, string svr, string siteName)
        {
            try
            {       
                
                Console.WriteLine("\nStarting to read eventlogs on " + svr + "\n");
                int LogCount = 0;
                TimeSpan ts = new TimeSpan(0, 0, s);
                EventLog el = new EventLog("Directory Service", svr, "NTDS General");
                EventLogEntryCollection entries = el.Entries;
                EventLogEntry entry;
                Stack<EventLogEntry> stack = new Stack<EventLogEntry>();
                string csvOutString = "";              

                for (int i = 0; i < entries.Count; i++)
                {
                    if (killLoop)
                    { break; }

                    entry = entries[i];
                    if (entry.Category.ToString().ToUpper() == "FIELD ENGINEERING")
                    {
                        if (killLoop)
                        { break; }

                        if (DateTime.Compare(dtStart, entry.TimeGenerated) != 1)
                        {
                            ServerList.Add(svr);
                            SiteList.Add(siteName);
                            stack.Push(entry);
                        }
                    }
                }                

                LogCount = stack.Count;
                while (stack.Count > 0 && killLoop == false)
                {
                    Console.Write(".");
                    string[] message = stack.Peek().Message.ToString().Split('\n');
                    int count = 0;

                    foreach (string st in message)
                    {                       

                        if (st.ToUpper().Contains("CLIENT:"))
                        {
                            string RawIPandPort = message[count + 1].Replace("\n", "").Trim().ToString();
                            string Port = RawIPandPort.Split(new char[] { ':' }).Last().Trim().ToString();
                            string IP = RawIPandPort.Replace(':'+ Port, "");
                            string IPString = ReturnNiceIP(message[count + 1].Replace("\n", "").Trim().ToString());
                            IPList.Add(IPString);
                            string ResolvedIP = ResolveIPToName(IPString);
                            csvOutString = '\u0022' + ResolvedIP + '\u0022' + "," + '\u0022' + stack.Peek().TimeGenerated.ToShortDateString() + " " + stack.Peek().TimeGenerated.ToLongTimeString() +  '\u0022' + "," + '\u0022' + IP  + '\u0022' + "," + '\u0022' + Port + '\u0022' + ","  + '\u0022' +  stack.Peek().UserName.ToString() + '\u0022' + ",";
                        }

                        if (st.ToUpper().Contains("STARTING NODE"))
                        {
                            csvOutString += '\u0022' + message[count + 1].Replace("\n", "").Trim() + '\u0022' + ",";
                        }

                        if (st.ToUpper().Contains("FILTER"))
                        {
                            csvOutString += '\u0022' + message[count + 1].Replace("\n", "").Trim() + '\u0022' + ",";
                        }

                        if (st.ToUpper().Contains("SEARCH SCOPE"))
                        {
                            csvOutString += '\u0022' +  message[count + 1].Replace("\n", "").Trim() + '\u0022' + ",";
                        }

                        if (st.ToUpper().Contains("ATTRIBUTE SELECTION"))
                        {
                            string newString = "";
                            if (string.IsNullOrWhiteSpace(message[count + 1].Replace("\n", "").Trim()))
                            { newString = "none"; }
                            else
                            { newString = message[count + 1].Replace("\n", "").Replace("\t", "").Replace("\r", "").Replace(" ","").Trim(); }

                            csvOutString += '\u0022' + newString + '\u0022' + ",";
                        }

                        if (st.ToUpper().Contains("SERVER CONTROLS"))
                        {
                            string newString = "";
                            if (string.IsNullOrWhiteSpace(message[count + 1].Replace("\n", "").Trim()))
                            { newString = "none"; }
                            else
                            { newString =  message[count + 1].Replace("\n", "").Replace("\t", "").Replace("\r", "").Trim(); }

                            csvOutString += '\u0022' + newString + '\u0022' + ",";
                        }

                        if (st.ToUpper().Contains("VISITED ENTRIES"))
                        {
                            csvOutString += '\u0022' +  message[count + 1].Replace("\n", "").Replace("\t", "").Replace("\r", "").Trim() + '\u0022' + ",";
                        }

                        if (st.ToUpper().Contains("RETURNED ENTRIES"))
                        {
                            csvOutString += '\u0022' +  message[count + 1].Replace("\n", "").Replace("\t", "").Replace("\r", "").Trim() + '\u0022' + ",";
                        }

                        if (st.ToUpper().Contains("USED INDEXES"))
                        {
                            csvOutString += '\u0022' +  message[count + 1].Replace("\n", "").Replace("\t", "").Replace("\r", "").Trim() + '\u0022' + ",";
                        }

                        if (st.ToUpper().Contains("PAGES REFERENCED"))
                        {
                            csvOutString += '\u0022' +  message[count + 1].Replace("\n", "").Replace("\t", "").Replace("\r", "").Trim() + '\u0022' + ",";
                        }

                        if (st.ToUpper().Contains("PAGES READ FROM DISK"))
                        {
                            csvOutString += '\u0022' +  message[count + 1].Replace("\n", "").Replace("\t", "").Replace("\r", "").Trim() + '\u0022' + ",";
                        }

                        if (st.ToUpper().Contains("PAGES PREREAD FROM DISK"))
                        {
                            csvOutString += '\u0022' +  message[count + 1].Replace("\n", "").Replace("\t", "").Replace("\r", "").Trim() + '\u0022' + ",";
                        }

                        if (st.ToUpper().Contains("CLEAN PAGES MODIFIED"))
                        {
                            csvOutString += '\u0022' +  message[count + 1].Replace("\n", "").Replace("\t", "").Replace("\r", "").Trim() + '\u0022' + ",";
                        }

                        if (st.ToUpper().Contains("DIRTY PAGES MODIFIED"))
                        {
                            csvOutString += '\u0022' +  message[count + 1].Replace("\n", "").Replace("\t", "").Replace("\r", "").Trim() + '\u0022' + ",";
                        }

                        if (st.ToUpper().Contains("SEARCH TIME (MS)"))
                        {
                            csvOutString += '\u0022' + message[count + 1].Replace("\n", "").Replace("\t", "").Replace("\r", "").Trim() + '\u0022' + ",";
                        }

                        if (st.ToUpper().Contains("ATTRIBUTES PREVENTING OPTIMIZATION"))
                        {
                            csvOutString += '\u0022' +  message[count + 1].Replace("\n", "").Replace("\t", "").Replace("\r", "").Trim() + '\u0022' + ",";
                        }

                        count = count + 1;                            
                        
                    }

                    writer.WriteLine('\u0022' + svr + '\u0022' + "," + '\u0022' + siteName + '\u0022' + "," + csvOutString);
                    csvOutString = string.Empty;
                    stack.Pop();
                }
                Console.WriteLine("\n");
                Console.WriteLine("Number of events: " + LogCount.ToString());
            }
            catch (IndexOutOfRangeException exp)
            {
                OutputToScreen("Error - Out of Range on " + svr + " - " + exp.Message.ToString(), ConsoleColor.Red);
                ErrWriter.WriteLine("Error - Out of Range on " + svr + " - " + exp.Message.ToString());
            }
            catch (Exception exp)
            {
                OutputToScreen("Error parsing events on " + svr + " - " + exp.Message.ToString(), ConsoleColor.Red);
                ErrWriter.WriteLine("Error parsing events on " + svr + " - " + exp.Message.ToString());
            }  
        }

        static string ReturnNiceIP(string IPAddr)
        {
            try
            {
                IPAddress IP;
                bool TrueIP = System.Net.IPAddress.TryParse(IPAddr, out IP);
                if (TrueIP)
                { return IP.ToString(); }
                else
                {
                    if (IPAddr.StartsWith("["))
                    { return IPAddr.Split(']')[0].TrimStart('['); }

                    if (IPAddr.Contains(":"))
                    { return IPAddr.Split(':')[0]; }
                }
            }
            catch (Exception exp)
            {
                ErrWriter.WriteLine("Error returning nice IP for " + IPAddr + " - " + exp.Message.ToString());
                OutputToScreen("Error returning nice IP for " + IPAddr + " - " + exp.Message.ToString(), ConsoleColor.Red);
            }
            return IPAddr;
        }

        static void OutputToScreen(string message, ConsoleColor cc)
        {
            lock (screenLock)
            {
                Console.ForegroundColor = cc;
                Console.WriteLine(message);
                Console.ResetColor();
            }            
        }

        static void OutputToSQL()
        {
            Console.WriteLine("Not implemented yet!");
        }

        static void ResetFEValue(string s)
        {            
            if(AllDCs)
            {
                DCList = domain.FindAllDiscoverableDomainControllers();                
                foreach (DomainController dc in DCList)
                {
                    if (SkipDownDC.Contains(dc.Name))
                    {
                        OutputToScreen("Skipping DC - " + dc.Name, ConsoleColor.Red);
                    }
                    else
                    {
                        DoTurnFEValueOnOrOff(s, dc.Name.Split('.')[0]);
                        if (veryVerbose)
                        { SetSearchThresholdValues(s, dc.Name.Split('.')[0]); }

                        if (ResetDefaults)
                        { SetSearchThresholdValues(s, dc.Name.Split('.')[0]); }
                    }
                }
            }
            else
            {
                DoTurnFEValueOnOrOff(s, server);
                if(veryVerbose)
                { SetSearchThresholdValues(s, server); }

                if (ResetDefaults)
                { SetSearchThresholdValues(s, server); }
            }           
        }

        static void DoTurnFEValueOnOrOff(string s, string svr)
        {
            try
            {
                FE15 = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, svr).OpenSubKey("SYSTEM\\CurrentControlSet\\Services\\NTDS\\Diagnostics", true);

                if (s.ToUpper() == "OFF")
                {
                    OutputToScreen("Resetting Field Engineering value to 0 for " + svr, ConsoleColor.Green);

                    FE15.SetValue("15 Field Engineering", "0", RegistryValueKind.DWord);
                    FE15.Close();
                    Thread.Sleep(200);

                    FE152 = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, svr).OpenSubKey("SYSTEM\\CurrentControlSet\\Services\\NTDS\\Diagnostics", false);
                    string FE15val2 = FE152.GetValue("15 Field Engineering").ToString();
                    
                    if (FE15val2 == "0")
                    {  
                        OutputToScreen("Field Engineering value set to 0 on " + svr, ConsoleColor.Green);
                    }
                    else
                    {  
                        OutputToScreen("Error setting Field Engineering value set to 0 on " + svr, ConsoleColor.Red);
                        ErrWriter.WriteLine("Error setting Field Engineering value set to 0 on " + svr);
                    }
                }

                if (s.ToUpper() == "ON")
                {
                    OutputToScreen("Setting Field Engineering value to " + LogVal.ToString() + " for " + svr, ConsoleColor.Green);

                    FE15.SetValue("15 Field Engineering", LogVal.ToString(), RegistryValueKind.DWord);
                    FE15.Close();
                    Thread.Sleep(200);
                    FE152 = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, svr).OpenSubKey("SYSTEM\\CurrentControlSet\\Services\\NTDS\\Diagnostics", false);

                    string FE15val2 = FE152.GetValue("15 Field Engineering").ToString();
                    if (FE15val2 == LogVal.ToString())
                    {   
                        OutputToScreen("Field Engineering value set on " + svr, ConsoleColor.Green);
                    }
                    else
                    {   
                        OutputToScreen("Error setting Field Engineering value on " + svr, ConsoleColor.Red);
                        ErrWriter.WriteLine("Error setting Field Engineering value on " + svr);
                    }
                }
            }
            catch (Exception exp)
            {
                ErrWriter.WriteLine("Error setting Field Engineering on " + svr + " - " + exp.Message.ToString());
                OutputToScreen("Error setting Field Engineering value on " + svr + " - " + exp.Message.ToString(), ConsoleColor.Red);  
            }

        }
    
        static void SetSearchThresholdValues(string s, string svr)
        {
            try
            {
                RegistryKey searththreshold = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, svr).OpenSubKey("SYSTEM\\CurrentControlSet\\Services\\NTDS\\Parameters", true);
                bool b1 = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, svr).OpenSubKey("SYSTEM\\CurrentControlSet\\Services\\NTDS\\Parameters", true).GetValue("Expensive Search Results Threshold", "NOT EXIST").ToString() != "NOT EXIST";
                bool b2 = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, svr).OpenSubKey("SYSTEM\\CurrentControlSet\\Services\\NTDS\\Parameters", true).GetValue("Inefficient Search Results Threshold", "NOT EXIST").ToString() != "NOT EXIST";

                if (s.ToUpper() == "ON")
                {
                    OutputToScreen("Setting Threshold values for " + svr, ConsoleColor.Green);

                    searththreshold.SetValue("Expensive Search Results Threshold", "1", RegistryValueKind.DWord);
                    searththreshold.SetValue("Inefficient Search Results Threshold", "1", RegistryValueKind.DWord);
                    
                    OutputToScreen("Threshold values set to 1", ConsoleColor.Green);
                }                            

                if (s.ToUpper() == "OFF")
                {
                    OutputToScreen("Removing Threshold values for " + svr, ConsoleColor.Green);

                    if (b1)
                    { searththreshold.DeleteValue("Expensive Search Results Threshold"); }

                    if (b2)
                    { searththreshold.DeleteValue("Inefficient Search Results Threshold"); }

                    OutputToScreen("Threshold values removed", ConsoleColor.Green);
                    
                }

                searththreshold.Flush();
                searththreshold.Close(); 
            }
            catch (Exception exp)
            {
                ErrWriter.WriteLine("Error setting Search Threasholds on " + svr + " - " + exp.Message.ToString());    
                OutputToScreen("Error setting Search Thresholds on " + svr + " - " + exp.Message.ToString(), ConsoleColor.Red);   
            }
        }

        static string ResolveIPToName(string IP)
        {
            try
            {
                IPHostEntry ResolvedIP = Dns.GetHostEntry(IPAddress.Parse(IP));                
                return ResolvedIP.HostName.ToString();
            }
            catch (Exception exp)
            {
                ErrWriter.WriteLine("Error resolving IP - " + IP + " - " + exp.Message.ToString());    
                return IP;     
            }            
        }

        static void Usage()
        {            
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("----------------------------------------------------------------------------------------------------------");
            Console.WriteLine("\nFieldEngineeringIPLog:");
            Console.WriteLine("\nWritten by Phil Lane to log details of inefficient & expensive LDAP searches.");
            Console.WriteLine("");
            Console.WriteLine("Any questions or comments to phlane@microsoft.com");
            Console.WriteLine("V.13\n");
            Console.WriteLine("----------------------------------------------------------------------------------------------------------");
            Console.WriteLine("\nFieldEngineeringIPLog: ");
            Console.WriteLine("\t <interval>              - time in seconds to turnh on logging");
            Console.WriteLine("\t s:<DC Name>             - Domain Controller to query. \"domain\" can be used to query all DC's in the domain");
            Console.WriteLine("\t l:<int>                 - Logging level 1 - minimal, 5 - maximum");
            Console.WriteLine("\t v                       - verbose output");
            Console.WriteLine("\t vv                      - very verbose output (resource intensive!)");
            Console.WriteLine("\t file:<FileName.txt>     - output to file is in csv format");
            //Console.WriteLine("\t sql:<connection string> - output to an SQL server.");
            Console.WriteLine("\t /ResetDefaults          - Reset the default registry values");
            Console.WriteLine("\nExample: FieldEngineeringIPLog 15 l:5 s:brisdc1 vv file:brisdc1_output.txt");
            Console.WriteLine("The above line will enable Field Engineering logging at its maximum level for 15 seconds on server brisdc1");
            Console.WriteLine("and output the results to a file called brisdc1_output.txt\n");
            Console.WriteLine("\nExample: FieldEngineeringIPLog /ResetDefaults s:domain");
            Console.WriteLine("The above example will reset all the registry values to the default setting for Domain Controllers in the domain.");
            Console.WriteLine("\n\n");
            Console.WriteLine("Active Directory Diagnostic Logging - http://technet.microsoft.com/en-gb/library/cc961809.aspx");
            Console.WriteLine("Creating More Efficient Microsoft Active Directory-Enabled Applications - http://msdn.microsoft.com/en-us/library/ms808539.aspx");

            Console.WriteLine("\n");
            Console.WriteLine("When running the tool via a Scheduled Task it will produce a lot of csv files, these can be merged with the following PowerShell:");
            Console.WriteLine("Dir *.csv | ? {$_.basename -like '*output*'} | Select -ExpandProperty Name | Import-Csv | Export-Csv -Path MergedOutputs.csv -NoTypeInformation");
            Console.WriteLine("----------------------------------------------------------------------------------------------------------");
            Console.ResetColor();
        }
    }
}
