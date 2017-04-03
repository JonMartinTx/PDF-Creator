using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Report;
using TemplatePrinting;
using static System.Configuration.ConfigurationSettings;

namespace AS400Report
{
    #region Program

    public class Program
    {
        private static readonly List<Task> PrinterThreads = new List<Task>();
        public static bool BReport { get; }

        #region Main

        static int Main(string[] args)
        {
            var reportType = args[0].ToUpper();

            var reports = "Reports"; reports = reports.ToUpper();
            var statements = "Statements"; statements = statements.ToUpper();
            var interest = "Interest"; interest = interest.ToUpper();
            var rehab = "Rehab"; rehab = rehab.ToUpper();
            var perkinsReassignment = "PerkinsReassignment"; perkinsReassignment = perkinsReassignment.ToUpper();
            var all = "All"; all = all.ToUpper();


            if (reportType == reports || reportType == all)
            {
                var configPath = @"C:\AS400Report\Print Queue\prtq\";
                const string searchPattern = "*.*";
                var directorySingleton = PrintQueue.DirectoryInstance;


                if (!Directory.Exists(configPath))
                {
                    Console.WriteLine($"(0) is not a valid file or directory");
                    return -1;
                }

                //Before Receiving files from the AS400, delete the old files from the last print job.
                //ClearPrtq();                          //Commented out until we're ready to implement the sockets.

                //For All Reports generated in the previous month, Zip them into a zip file then Delete the pdf versions to save space.
                //ArchiveFiles();

                //FTP all files from the AS400 to the prtq folder
                //CommandPromptScript_Files();

                //Process the directory (compiles a list of files to print and gets their config files), then Print the files.
                PrintQueue.ProcessDirectory(directorySingleton, @"C:\AS400Report\Print Queue\prtq\", searchPattern);
                Print();
            }
            if (reportType == statements || reportType == all)
            {

                var csvStatements = new PrintStatement();

                //var PrintingTask = new Task(csvStatements.PrintStatementPDF);
                //PrinterThreads.Add(PrintingTask);

                ////Might need to comment out.
                //Parallel.ForEach(PrinterThreads, t =>
                //{

                //    t.Start();

                //});
                //Task.WaitAll(PrinterThreads.ToArray());


                csvStatements.PrintStatementPdf();

            }

            if (reportType == interest)
            {
                var interestPaidMailer = new InterestPaid();
                interestPaidMailer.PrintInterestPaid();

            }

            if (reportType == perkinsReassignment || reportType == all)
            {
                var reassignment = new PerkinsReassignment();
                reassignment.PrintPerkinsAssignment();

            }

            if (reportType == rehab || reportType == all)
            {

            }

            else
            {
                Console.Write("The AS400 sockets argument was incorrectly handled.\n");
            }

            return (0);
        }

        #endregion

        #region PrintDocument

        // Print a document
        public static void PrintDocument(ReportConfig cfg, string fileName)
        {

            // Create an instance of the printer class
            using (var wtr = new ReportWriter("FCFC"))
            {

                // Set the font
                wtr.PrinterFont = new System.Drawing.Font(cfg.FontName, cfg.FontSize);
                wtr.LandScape = cfg.LandScape;
                if (cfg.Printer.Length > 0)
                    wtr.PrinterName = cfg.Printer;
                else
                    wtr.PrinterName = PrintQueue.DirectoryInstance.ToUserDirectory + AppSettings["DefaultPrinter"];
                wtr.LeftMargin = cfg.LeftMargin;
                wtr.RightMargin = cfg.RightMargin;
                wtr.TopMargin = cfg.TopMargin;
                wtr.BottomMargin = cfg.BottomMargin;

                // Set text to print
                using (var sr = new StreamReader(fileName))
                {
                    wtr.TextToPrint = sr.ReadToEnd();
                }
                //wtr.Print();
            }

        }

        #endregion

        #region Print

        private static void Print()
        {
            //Creates instances of my 3 singletons.
            //Instance of my Template Singleton, handles choosing which template to print.
            var templateInstance = TemplateSingleton.Instance;

            //Instance of my Print Queue Singleton, which helps manage Directory information
            var directorySingleton = PrintQueue.DirectoryInstance;

            //Instance of my ObjectLibrary. This lets me store my instances globally.
            var objLibInstance = ObjectLibrary.ObjLib;

            var filesProcessed = 0;

            if (directorySingleton.FileEntries.Length > 0)
            {
                //Unavoidable that I have to loop through the directory twice. This loop stores a new ReportConfig() for each Document.
                templateInstance.CfgList = new List<ReportConfig>();
                for (var i = 0; i < directorySingleton.FileEntries.Length; i++)
                {
                    var cfg = directorySingleton.GetConfigFile(i);
                    //TemplateInstance.CFGList.Add(cfg);
                    //If the File is to be a hard copy file, just send to the hard copy print method.
                    if (cfg.BPrint)
                    {
                        //Commented out during testing. This will print the Document as a hard copy.
                        //Program.PrintDocument(cfg, DirectorySingleton.FileEntries[i]);
                    }
                }

                //Pair my list of Reports and Config Files.
                templateInstance.PairFileLists();

                //Linq statement gives me a list of Unique Report Types.
                var uniqueCfgType = templateInstance.CfgList.GroupBy(x => x.Title).Select(y => y.First()).ToList();

                var listByReportType = new List<List<string>>(uniqueCfgType.Count);
                directorySingleton.DirectoriesByReportType = listByReportType;

                foreach (var s in uniqueCfgType)
                {
                    //Get a list of all the files that share the same config type.
                    //var sharedConfigList = GetValuefromdictionary(DirectorySingleton.ThisConfig, TemplateInstance.FileCFGReportPairs);
                    //ListByReportType.Add(sharedConfigList);
                    listByReportType.Add(new List<string>(GetValuefromdictionary(s, templateInstance.FileCfgReportPairs)));

                }
                foreach (var filesWithSharedConfig in listByReportType)
                {
                    //Retrieve the Config File from my list.
                    directorySingleton.ThisConfig = templateInstance.CfgList[filesProcessed];

                    //If we are printing with pdf...
                    if (directorySingleton.ThisConfig.BPdf)
                    {


                        //New PDFCreator object; populate it with the config data; then store it in my Object Library.
                        var pdf = new PdfCreator();

                        if (filesProcessed != 0)
                        {
                            filesProcessed++;
                            //Making absolutely sure that my Config Files are synced across my classes.
                            pdf.ConfigFile = directorySingleton.ThisConfig = templateInstance.CfgList[filesProcessed];
                        }

                        else
                        {
                            pdf.ConfigFile = directorySingleton.ThisConfig;
                        }

                        //Storing this PDF in my Object Library. (??? Do I need this ???)
                        objLibInstance.Pdf = pdf;

                        //Deleted once testing is done.
                        directorySingleton.ThisConfig.BPdf = true;

                        // Set the path and file name properties for the output PDF.
                        objLibInstance.Pdf.PathName = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + directorySingleton.ThisConfig.PdfPath;

                        pdf.ConfigFile = directorySingleton.ThisConfig = templateInstance.CfgList[filesProcessed];

                        var printingTask = new Task(() => pdf.Print_PDF(filesWithSharedConfig));
                        PrinterThreads.Add(printingTask);
                        if (filesProcessed == 0)
                        {
                            filesProcessed++;
                        }
                    }
                }
            }

            else
            {
                Console.Write("There are no files in the Directory that match the search pattern.");
            }

            Parallel.ForEach(PrinterThreads, t =>
            {

                t.Start();

            });
            Task.WaitAll(PrinterThreads.ToArray());

            var archiveUtility = new FileMover();
            archiveUtility.ArchiveFiles();

            //This is where I would push the Reports to Webster.

        }

        #endregion

        #region Detect Operating System

        //It's here if we need it.
        public enum Platform
        {
            Windows,
            Linux,
            Mac
        }

        public static Platform RunningPlatform()
        {
            Contract.Ensures(Contract.Result<Platform>() != null);
            switch (Environment.OSVersion.Platform)
            {
                case PlatformID.Unix:
                    {

                        // Looking For root folders to find out which OS is being used

                        return Directory.Exists("/Applications")
                            & Directory.Exists("/System")
                            & Directory.Exists("/Users")
                            & Directory.Exists("/Volumes") ? Platform.Mac : Platform.Linux;
                    }
                case PlatformID.MacOSX:
                    {
                        return Platform.Mac;
                    }
                default:
                    {
                        return Platform.Windows;
                    }
            }
        }

        #endregion

        #region LinQ -> Get all Files for Each Config File

        //Searches my Dictionary of KeyValuePairs. For each file of a given ReportConfig Type, the file is added to a list to be printed. This allows me to print by file type.
        private static List<string> GetValuefromdictionary(ReportConfig cfgFile, Dictionary<ReportConfig, string> cfgPairs)
        {
            var configName = cfgFile.Title;
            var selectedValues = TemplateSingleton.Instance.CfgList
                     .Where(x => x.Title == configName)
                     .Select(x => cfgPairs[x])

                     .ToList();

            return selectedValues;
        }

        #endregion

        #region Zip Old Files

        public static void ArchiveFiles()
        {
            //Later, I'll send a List of files (so all of Taft's Reports, for example).
            var d = DateTime.Now;
            d = d.AddMonths(-1);
            var lastMonth = d.ToString("MM-yyyy");
            var lastMonthVar2 = d.ToString("MM/dd/yyyy");
            var lastMonthVar3 = d.ToString("MM.dd.yyyy");
            var lastMonthVar4 = d.ToString("yyyyMMdd");
            var lastMonthVar5 = d.ToString("dd-MM-yyyy");
            var lastMonthAsText = d.ToString("MMMM, yyyy");
            const string archivePath = @"C:\AS400Report\ArchivedFiles\";
            const string path = @"C:\AS400Report\Reports\";
            var zipName = archivePath + lastMonthAsText + "'s Reports";


            var deleteFiles = new DirectoryInfo(path);
            deleteFiles.GetFiles("*.*")
                .Where(p => p.Extension == "*.*").ToArray();

            //Files in Directory is a list of all files found in the folder.
            var filesInDirectory = Directory.GetFiles(path, "*", SearchOption.AllDirectories).ToList();

            //Files to Zip are the ones ready to be archived.
            var filesToZip = new List<string>(filesInDirectory.Count);

            //For each month old report, add them to the list of files to Zip.
            foreach (var file in filesInDirectory)
            {
                var extensionIndex = file.LastIndexOf(".");
                var extensionChars = file.Substring(extensionIndex);
                var fileExtension = file.Substring(extensionIndex);

                if (
                    file.Contains("Batchsheets") ||
                    file.Contains("Deposit Reciepts")
                    )

                {
                    continue;
                }

                //Only grab files if they are pdf. I don't want to collect .zip files.
                else if (file.Contains(lastMonth) && fileExtension == extensionChars
                         || (file.Contains(lastMonthVar2) && fileExtension == extensionChars)
                         || (file.Contains(lastMonthVar3) && fileExtension == extensionChars)
                         || (file.Contains(lastMonthVar4) && fileExtension == extensionChars)
                         || (file.Contains(lastMonthVar5) && fileExtension == extensionChars)
                    /*||  testcounter > 34*/
                )
                    filesToZip.Add(file);
            }

            //For each file to zip, zip them into the archive.
            foreach (var file in filesToZip)
            {
                using (var zip = new Ionic.Zip.ZipFile())
                {
                    zip.AddDirectory(file, zipName);
                    zip.Save(zipName);
                }

            }

            ClearPrtq();
        }

        #endregion

        #region Delete Items in prtq

        private static void ClearPrtq()
        {
            var prtq = @"C:\AS400Report\Print Queue\";
            var oldFiles = Directory.GetFiles(prtq, "*", SearchOption.AllDirectories).ToList();

            //Delete files in prtq before receiving new files from the AS400
            foreach (var file in oldFiles)
            {
                var deleteFile = new FileInfo(file)
                {
                    Attributes = FileAttributes.Normal
                };
                File.Delete(deleteFile.FullName);
            }

        }

        #endregion

        #region GetShortPath

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern int GetShortPathName(
        [MarshalAs(UnmanagedType.LPTStr)]
         string path,
        [MarshalAs(UnmanagedType.LPTStr)]
         StringBuilder shortPath,
        int shortPathLength
        );

        private static string GetShortPath(string path)
        {
            const int maxPath = 255;
            var shortPath = new StringBuilder(maxPath);
            GetShortPathName(path, shortPath, maxPath);
            return shortPath.ToString();

        }

        #endregion

        #region Collect FTP Statements

        //This should Asyncronously run the script (in the string command) to execute the ftp. The window won't pop up because that would look dumb.
        static void CommandPromptScript_Statements()
        {
            const string statementString = @"C:\AS400Report\Print Queue\CSV Files";

            var shortPath = new StringBuilder(255);
            GetShortPathName(statementString, shortPath, shortPath.Capacity);

            #region SecureString

            var s = new SecureString();
            s.AppendChar('g');
            s.AppendChar('s');
            s.AppendChar('l');
            s.AppendChar('t');
            s.AppendChar('w');
            s.AppendChar('1');
            s.AppendChar('t');
            s.AppendChar('h');

            #endregion


            //Process CMDScript = new Process();
            //var processInfo = new System.Diagnostics.ProcessStartInfo();
            //processInfo.FileName = "cmd.exe";
            //processInfo.RedirectStandardInput = true;
            //processInfo.CreateNoWindow = true;
            //processInfo.UseShellExecute = false;
            //processInfo.RedirectStandardError = true;
            //processInfo.RedirectStandardOutput = true;
            ////processInfo.UserName = "ftpadmin";
            ////processInfo.Password = s;

            var ftpRequest = (FtpWebRequest)WebRequest.Create("ftp://192.168.1.5//Home/Export/");
            //ftpRequest.Credentials = new NetworkCredential("ftpadmin", s);
            ftpRequest.Method = WebRequestMethods.Ftp.ListDirectory;

            ftpRequest.Credentials = new NetworkCredential("ftpadmin", s);

            var response = (FtpWebResponse)ftpRequest.GetResponse();

            var responseStream = response.GetResponseStream();
            if (responseStream == null) return;
            using (var reader = new StreamReader(responseStream))
            {

                var directories = new List<string>();

                var line = reader.ReadLine();

                while (!string.IsNullOrEmpty(line))
                {
                    directories.Add(line);
                    line = reader.ReadLine();
                }
                reader.Close();

                using (var ftpClient = new WebClient())
                {
                    ftpClient.Credentials = new NetworkCredential("ftpadmin", s);

                    for (var i = 0; i <= directories.Count - 1; i++)
                    {
                        if (directories[i].Contains("statements.csv"))
                        {

                            var path = "ftp://192.168.1.5//Home/Export/" + directories[i];
                            var trnsfrpth = statementString;
                            ftpClient.DownloadFile(path, trnsfrpth);
                        }
                    }
                }
            }
            //CMDScript.StartInfo = processInfo;
            //CMDScript.Start();

            //using (StreamWriter sw = CMDScript.StandardInput)
            //{
            //    if (sw.BaseStream.CanWrite)
            //    {
            //        sw.WriteLine("cd" + shortPath.ToString().Substring(2));
            //        sw.WriteLine("del /Q *");
            //        sw.WriteLine("ftp 192.168.1.5");
            //        sw.WriteLine("ftpadmin");
            //        sw.WriteLine(s);
            //        sw.WriteLine("cd /home/export");
            //        sw.WriteLine("lcd" + shortPath.ToString().Substring(2));
            //        sw.WriteLine("get statements.csv");
            //        sw.WriteLine("bye");
            //    }
            //}



            //CMDScript.OutputDataReceived += (object sender, System.Diagnostics.DataReceivedEventArgs e) =>
            //    Console.WriteLine("output>>" + e.Data);
            //CMDScript.BeginOutputReadLine();

            //CMDScript.ErrorDataReceived += (object sender, System.Diagnostics.DataReceivedEventArgs e) =>
            //    Console.WriteLine("error>>" + e.Data);
            //CMDScript.BeginErrorReadLine();

            //CMDScript.WaitForExit();

            //Console.WriteLine("ExitCode: {0}", CMDScript.ExitCode);
            //CMDScript.Close();
        }

        #endregion

        #region Collect FTP Files

        //This should Asyncronously run the script (in the string command) to execute the ftp. The window won't pop up because that would look dumb.
        static void CommandPromptScript_Files()
        {
            const string command = "cd/C:\\Users\\Owner\\Educational Data Systems, Inc\\EDSIReports - Documents\\prtq " +
       "del /Q *.* " +
       "open 192.168.1.5 " +
       "ftpadmin " +
       "gsltw1th " +
       "cd / home / export " +
       "lcd C:\\Users\\Owner\\Educational Data Systems, Inc\\EDSIReports - Documents\\prtq " +
       "mget *.* " +
       "bye";

            var processInfo = new System.Diagnostics.ProcessStartInfo("cmd.exe", "/c " + command)
            {
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardError = true,
                RedirectStandardOutput = true
            };

            var process = System.Diagnostics.Process.Start(processInfo);

            if (process == null) return;
            process.OutputDataReceived += (object sender, System.Diagnostics.DataReceivedEventArgs e) =>
                Console.WriteLine("output>>" + e.Data);
            process.BeginOutputReadLine();

            process.ErrorDataReceived += (object sender, System.Diagnostics.DataReceivedEventArgs e) =>
                Console.WriteLine("error>>" + e.Data);
            process.BeginErrorReadLine();

            process.WaitForExit();

            Console.WriteLine($"ExitCode: {process.ExitCode}");
            process.Close();
        }

        #endregion
    }

    #endregion

    #region Clean

    public class FileMover
    {
        public string ToUserDirectory { get; set; }

        #region Constructor

        public FileMover()
        {
            ToUserDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        }

        #endregion

        #region ArchiveFiles

        public void ArchiveFiles()
        {
            const string csvPath = @"C:\AS400Report\Print Queue\CSV Files\";
            const string oldCsvPath = @"C:\AS400Report\ArchivedFiles\OldCSVFiles\";

            var csvFolder = new DirectoryInfo(csvPath);
            var csvFiles = csvFolder.GetFiles("*.csv");

            string csvArchiveName;

            foreach (var file in csvFiles)
            {
                csvArchiveName = oldCsvPath + file.Name;
                //Clever way to check if a file exists. If it does, it automatically adds a _n where n in the number sequence.
                if (File.Exists(csvArchiveName))
                {
                    var counter = 2;
                    var fileLenBfrExtnsn = csvArchiveName.Length - 4;
                    var sequence = counter.ToString();
                    var withoutExtension = csvArchiveName.Substring(0, fileLenBfrExtnsn);
                    var extension = csvArchiveName.Substring(fileLenBfrExtnsn);
                    csvArchiveName = withoutExtension + "_" + sequence + extension;
                    while (File.Exists(csvArchiveName))
                    {
                        counter++;
                        sequence = counter.ToString();
                        csvArchiveName = withoutExtension + "_" + sequence + extension;
                    }
                }

                File.Move(csvPath + file.Name, csvArchiveName);
            }

            const string toPrtq = @"C:\AS400Report\Print Queue\prtq\";
            const string toArchivePrtq = @"C:\AS400Report\ArchivedFiles\ArchivedPrtq\";

            var prtqFolder = new DirectoryInfo(toPrtq);
            var prtq = prtqFolder.GetFiles();

            foreach (var splfile in prtq)
            {
                var splfArchiveName = toArchivePrtq + splfile.Name;
                //Clever way to check if a file exists. If it does, it automatically adds a _n where n in the number sequence.
                if (File.Exists(splfArchiveName))
                {
                    var counter = 2;
                    var sequence = counter.ToString();
                    splfArchiveName = splfile.Name + "_" + sequence;
                    while (File.Exists(splfArchiveName))
                    {
                        counter++;
                        sequence = counter.ToString();
                        splfArchiveName = splfile.Name + "_" + sequence;
                    }
                }
                File.Move(toPrtq + splfile.Name, splfArchiveName);
            }


            CompressArchive();
        }

        #endregion

        #region Zip Old Files

        public static void CompressArchive()
        {
            //Later, I'll send a List of files (so all of Taft's Reports, for example).
            var yesterday = DateTime.Now.AddDays(-1);
            var monthAsText = yesterday.ToString("MMMM, yyyy");

            //Path Variables
            var path = @"C:\AS400Report\ArchivedFiles\";
            var pathCsv = path + @"OldCSVFiles\";
            var pathPrtq = path + @"ArchivedPrtq\";

            //Zip Names
            var zipName = monthAsText + "'s Reports";
            var zipArchiveName = path + zipName + ".zip";


            var archiveFolder = new FileInfo(zipName);
            archiveFolder.Directory.Create(); // If the directory already exists, this method does nothing.

            var zipDirectory = new DirectoryInfo(path);

            //Files to Zip are the ones ready to be archived.
            var zipfiles = zipDirectory.GetFiles("*", SearchOption.AllDirectories)
                .Where(p => p.CreationTime <= yesterday)
                .OrderBy(p => p.CreationTime).ToList();

            if (Directory.Exists(zipArchiveName))
            {
                using (var zip = new Ionic.Zip.ZipFile())
                {
                    foreach (var file in zipfiles)
                    {
                        zip.AddDirectory(pathCsv);
                        zip.AddDirectory(pathPrtq);
                        ZipFile.CreateFromDirectory(path, zipArchiveName + ".zip");
                        zip.Save(file.Name + yesterday.ToLongDateString());
                    }
                }
            }

            else if (!Directory.Exists(zipArchiveName + ".zip"))
            {
                var zipEnumerable = zipDirectory.GetFiles("*", SearchOption.AllDirectories)
                    .Where(p => p.CreationTime <= yesterday)
                    .OrderBy(p => p.CreationTime)
                    .Select(x => x.Name);

                var dirInfos = zipDirectory.GetDirectories("*", SearchOption.AllDirectories)
                    //.Where(p => p.CreationTime <= yesterday)
                    .OrderBy(p => p.CreationTime);

                int number;

                using (var zip = new Ionic.Zip.ZipFile())
                {
                    foreach (var file in zipfiles)
                    {
                        zip.AddFile(file.FullName);
                    }
                    //zip.AddFiles(ZipEnumerable, false, "");

                    //zip.ParallelDeflateThreshold = -1;



                    zip.Save(zipArchiveName);
                }
            }



            //For each file to zip, zip them into the archive.
            if (File.Exists(zipArchiveName + ".zip") && zipfiles.Count > 0)
            {


                using (var zipToOpen = new FileStream(zipArchiveName + ".zip", FileMode.Open))
                {
                    using (var archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                    {
                        var zipList = zipfiles.Select(x => archive.CreateEntry(x.Name))
                            .ToList();

                        //using (StreamWriter writer = new StreamWriter(ZipList.ForEach(x=> x.Open()))
                        //{
                        //    writer.WriteLine("Information about this package.");
                        //    writer.WriteLine("========================");
                        //}
                    }
                }

            }


            //Delete Each file after it's been archived.
            foreach (var file in zipfiles)
            {
                file.Delete();
            }
        }

        #endregion
    }

    #endregion

    #region PrintQueue Singleton

    public sealed class PrintQueue : Program
    {
        #region Variables

        #region Static Instance Variable

        private static readonly PrintQueue _directoryInstance = new PrintQueue();

        #endregion

        #region Public Variables

        public PdfCreator Pdf;
        //From FCFC Scanner
        public string Text { get; set; }            // text to scan (replaced before start of each new page
        public int TextLen { get; set; }            // original length of _text
        public string Page { get; set; }            // page to print
        public int CurPageOffset { get; set; }      // current byte count (offset) for original _text. Object user must set once initially via constructor
        public int PageLen { get; set; }            // length of _page after NextPage method
        public bool MorePages { get; set; }         // true if _curPageOffset < _textLen
        public int NumLines { get; set; }           // number of lines in the current page
        public List<int> WordsPerEachPage { get; set; }
        public List<CommandLineArgs> CmdArgList { get; set; }
        public int NumPaired { get; set; }
        public List<string> OutputPaths { get; set; }
        public string ToUserDirectory { get; set; }
        public List<List<string>> DirectoriesByReportType { get; set; }

        #endregion

        #region Accessors Mutators

        public static PrintQueue DirectoryInstance
        {
            get
            {
                return _directoryInstance;
            }
        }
        public ReportConfig ThisConfig { get; set; }
        public string FilePath { get; set; }
        public string ConfigPath { get; set; }
        public string SearchPattern { get; set; }
        public string[] FileEntries { get; set; }
        public string CfgFile { get; set; }
        private int NumberOfReports { get; set; }
        public int NumberLeft { get; set; }
        public int NumberCompleted { get; set; }

        #endregion

        #endregion

        #region CommandLine Arg Struct

        public struct CommandLineArgs
        {
            readonly string _filePath;
            readonly string _searchPattern;


            public CommandLineArgs(string path, string pattern)
            {
                _filePath = path;
                _searchPattern = pattern;
            }
        }

        #endregion



        #region Constructor

        private PrintQueue()
        {
            NumberOfReports = 0;
            NumberLeft = 0;
            NumberCompleted = 0;
            NumPaired = 0;

        }

        #endregion

        #region Methods

        #region Record Tallies

        private void RecordTallies()
        {
            NumberLeft = (NumberOfReports - NumberCompleted);
        }

        #endregion

        #region Store Command Line Args

        //Update for more sophisticated args.
        private void StoreCmdLnArgs(string fileDir, string searchpattern)
        {
            FilePath = fileDir;
            SearchPattern = searchpattern;
            var argStorage = new CommandLineArgs(fileDir, searchpattern);
            CmdArgList.Add(argStorage);
        }

        #endregion

        #region ProcessDirectory

        // Process files in the directory
        public static void ProcessDirectory(PrintQueue directorySingleton, string targetDirectory, string searchPattern)
        {
            directorySingleton.ToUserDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            directorySingleton.ConfigPath = @"\PrintSetUp\";
            var fileDirectory = targetDirectory;

            //An Array of File Names equal to the number of files in the target print queue.
            directorySingleton.FileEntries = Directory.GetFiles(fileDirectory, searchPattern);

            //A list of command line arguments. Will be used once the AS400 sends Doc Data as JSON.
            directorySingleton.CmdArgList = new List<CommandLineArgs>();
            directorySingleton.OutputPaths = new List<string>(directorySingleton.FileEntries.Length);

            //For each File in the Directory...
            foreach (var fileName in directorySingleton.FileEntries)
            {
                directorySingleton.NumberOfReports++;
                //Add the commandline args to my struct
                directorySingleton.StoreCmdLnArgs(fileName, searchPattern);
            }

            directorySingleton.RecordTallies();
        }

        #endregion

        #region Get Config File

        public ReportConfig GetConfigFile(int counter)
        {
            var filePath = FileEntries[counter];

            // Load configuration file for report
            var fileOnly = Path.GetFileName(filePath);
            var configFile = fileOnly.Substring(0, 4) + "_config.xml";
            DirectoryInstance.ConfigPath = @"C:\AS400Report\PrintSetUp\";
            var configPath = DirectoryInstance.ConfigPath/* + ConfigurationSettings.AppSettings["ConfigPath"]*/ + configFile;
            if (File.Exists(configPath))
            {
                ThisConfig = new ReportConfig(configPath, filePath);

                ConfigPath = configPath;

                return ThisConfig;
            }
            else
            {
                //var buildCFG = new ConfigBuilder.CFGBuilder(FilePath);
                //buildCFG.GenerateXML();
                //System.Environment.Exit(10);
                return null;
            }
        }

        public ReportConfig GetConfigFile(string manualConfigPrefix, string manualFilePath)
        {
            string configFile, configPath;
            configFile = manualConfigPrefix + "_config.xml";
            configPath = DirectoryInstance.ConfigPath + AppSettings["ConfigPath"] + configFile;
            ThisConfig = new ReportConfig(configPath, manualFilePath);
            ConfigPath = configPath;

            return ThisConfig;
        }

        #endregion

        #endregion
    }

    #endregion

    #region Class-Object Singleton

    public class ObjectLibrary
    {
        #region Variables

        static private ObjectLibrary _objLib;
        private int CurrentPosition { get; set; }
        public string FullOutputPath { get; set; }

        #endregion

        #region Mutators Accessors

        //The instantiation of this document class.
        public static ObjectLibrary ObjLib
        {
            get
            {
                if (_objLib == null)
                {
                    _objLib = new ObjectLibrary();
                }
                return _objLib;
            }
        }

        public FcfcScanner Fcfc { get; set; }
        public PdfCreator Pdf { get; set; }
        public Templates MyTemplate { get; set; }
        private string FullInputPath { get; set; }
        private string InputPath { get; set; }
        public string InputFileName { get; set; }
        public string CsvPath { get; set; }
        public string TemplatePath { get; set; }

        #endregion

        #region Constructor

        private ObjectLibrary()
        {
            InputPath = PrintQueue.DirectoryInstance.ToUserDirectory + @"\Google Drive\PrintingProgram\CompletedTemplates\";
            InputFileName = "PDFStatement.pdf";
            FullInputPath = InputPath + InputFileName;
            CsvPath = @"C:\AS400Report\Print Queue\CSV Files";
            TemplatePath = PrintQueue.DirectoryInstance.ToUserDirectory + @"\Google Drive\PrintingProgram\CompletedTemplates\";

        }

        #endregion

    }
    #endregion
}
