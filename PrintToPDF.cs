using System;
using System.Text;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Drawing;
using Report;
using System.Linq;
using System.Collections.Generic;
using TemplatePrinting;
using System.Threading;
using AS400Report;
using static System.IO.File;
using static AS400Report.PrintQueue;
using static TemplatePrinting.PrintStatement;

namespace Report
{
    #region PDFCreator

    /************************************************************
    / Class:    PDFCreator  									*
    / Purpose:  Serves as Main() for PDF Documents. This class  *
    /               Sends each file list to the appropriate     *
    /               Class to print according to its template.   *
    / Methods:  ConfigParser, Print_PDF, GetAddress, & GetLines.*
    /               and CheckForTemplate.                       *
    *************************************************************/
    public class PdfCreator
    {
        #region Class_Properties

        public List<string> PageArray { get; set; }
        protected FileStream Fs { get; set; }
        public string InputPath;
        public string InputFileName;
        public string FullInputPath;
        public int NumLines { get; set; }
        public int CurrLine { get; set; }

        // set the text to print
        public string TextToPrint { get; set; }

        // get/set the path to destination PDF
        public string PathName { get; set; }

        // get/set the font size
        public int FontSize { get; set; }
        // get/set top print offset override
        public float TopPrintOffset { get; set; }
        public float LineMultiplier { get; set; }
        public bool FromTemplate { get; set; }
        public ReportConfig ConfigFile { get; set; }
        //public FileStream Fs
        //{
        //    get { return Fs; }
        //}
        protected FcfcScanner ScannerText { get; set; }

        public int PageNumber { get; set; }



        #endregion

        #region Class_Contsructor

        public PdfCreator()
        {
            FromTemplate = false;
            PageNumber = 0;

        }

        #endregion

        #region Methods

        #region Print_PDF

        // ------------------------------------------------------------------------------
        // Create PDF document
        // ------------------------------------------------------------------------------
        public int Print_PDF(List<string> filesToPrint)
        {
            //Instance of my Print Queue Singleton, which helps manage Directory information
            var directorySingleton = DirectoryInstance;

            //Instance of my Template Singleton, handles choosing which template to print.

            //Instance of my ObjectLibrary. This lets me store my instances globally.
            var objLibInstance = ObjectLibrary.ObjLib;

            //testing, figure something else out here.
            var pdfContentByteText = true;

            //Assigning my Config File for the Batch of Reports that use it.
            var cfg = TemplateSingleton.Instance.FileCfgReportPairs
                   .FirstOrDefault(x => x.Value.Contains(filesToPrint[0]))
                   .Key;

            //Only one template type, since it depends on the config file. Moved up here so I'm not executing the function once per file in FilesToPrint.
            var myTemplate = ConfigParser(cfg);

            //The name of this template is tested in the foreach block below.
            var templateName = myTemplate.ToString();


            //For Each File in the Directory...
            //foreach (string file in FilesToPrint)
            //{
            //Retrieve the Config File from my list.
            ConfigFile = cfg;
            objLibInstance.Pdf.PathName = directorySingleton.ThisConfig.PdfPath;


#pragma warning disable CC0021 // Use nameof
            if (templateName == "Remittance" || templateName == "Invoice" || templateName == "PDFStatement" || templateName == "StudentJournal" || templateName == "BlankStatement")

            {
                if (templateName == "Remittance" || templateName == "Invoice")
                {
                    myTemplate = Templates.Invoice;
                    var temp = new Remittance();
                    temp.PrintRemittancePdf(filesToPrint);
                    pdfContentByteText = true;
                }
                if (templateName == "Statement")
                {
                    myTemplate = Templates.Statement;
                    var statementPrinting = new PrintStatement();
                    statementPrinting.PrintStatementPdf(filesToPrint);
                    pdfContentByteText = false;
                }
                if (templateName == "StudentJournal")
                {
                    myTemplate = Templates.StudentJournal;
                    var temp = new StudentJournal();
                    temp.PrintStudentJournalPdf(filesToPrint);
                    pdfContentByteText = true;
                }
                if (templateName == "BlankStatement")
                {
                    myTemplate = Templates.Blank;
                    pdfContentByteText = false;
                }

                //... continue to the rest of the reports
            }
            else
            {
                myTemplate = Templates.Default;
                var temp = new Default();
                temp.PrintDefaultPdf(filesToPrint);
                //PdfContentByteText = true;
            }



            if (pdfContentByteText)
            {

            }

            //DirectorySingleton.NumberCompleted++;


#pragma warning restore CC0021 // Use nameof

            Console.WriteLine("Report Type: " + ConfigFile.Title + " Completed.");
            return 0;
        }


        #endregion

        #region GetAddress

        public static string GetAddress(List<string> lines)
        {
            try
            {
                string addressBlock;
                const int maxElements = 11;
                var temp = new string[maxElements];
                temp[0] = lines[1];
                addressBlock = temp[0] + "\n";

                for (int i = 1, j = 4; i < maxElements; i++, j++)
                {
                    if (String.IsNullOrWhiteSpace(lines[j]))
                        continue;

                    temp[i] = lines[j];

                    if (i == (maxElements - 1))
                    {
                        break;
                    }


                    if (j == 4 || j == 5 || j == 11)
                        temp[i] = temp[i] + " ";

                    else if (j == 10)
                    {
                        temp[i] = temp[i] + ", ";
                    }

                    else
                    {
                        temp[i] = temp[i] + "\n";
                    }
                }

                foreach (var s in temp.Skip(1))
                {
                    addressBlock = addressBlock + s;
                }

                return addressBlock;
            }
            catch (Exception err)
            {
                Console.Write(err);
                return null;
            }
        }

        #endregion

        #region GetLines

        public static IEnumerable<string> ReadLines(Func<Stream> streamProvider,
                                         Encoding encoding)
        {
            StreamReader reader;
            using (var stream = streamProvider?.Invoke())
                if (stream != null)
                    using (reader = new StreamReader(stream, encoding))
                    {
                        string line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            yield return line;
                        }
                    }
        }

        #endregion

        #region Config Parser

        /************************************************************
        / Method:   Config Parser   								*
        / Purpose:  Reads the Config File name to determine the type*
        /			    of report the program is processing. This 	*
        /			    Determines the Template Type.       		*
        / Returns:  Template.	                                    *
        *************************************************************/

        public static Templates ConfigParser(ReportConfig cfg)
        {
            Templates myTemplate;
            var reportType = (cfg.Title != null) ? cfg.Title.Substring(0, 4) : "0";
            reportType = reportType.ToUpper();


            try
            {
                myTemplate = (Templates)reportType;
            }

            catch
            {
                myTemplate = (Templates)"0";
                Console.Write("Config File Type " + reportType + " Not Found.\n");
            }

            if (myTemplate == Templates.Error)
                Console.WriteLine("Error Template Selected. Check ConfigParser class for errors.\n");

            return myTemplate;
        }

        #endregion

        #endregion

    }

#endregion
}


namespace TemplatePrinting
{

    #region Document Print Strategies

    /************************************************************
    / ClassGroup:   Document Print Strategies       			*
    / Purpose:  Each Type of Report prints to pdf in its own    *
    /               class. Allows for multi threading and spec- *
    /               ific pdf templates.                         *
    / Classes:  Statement, BlankStatement, Remittance, Student  *
    /               Journal, and Default. (Add as list grows).  *
     *************************************************************/
    #region Statements

    public class PrintStatement : PdfCreator
    {
        #region Variables

        public string _outputPath { get; set; }
        public string FormName { get; set; }   // form name
        public string Title { get; set; }      // title
        public PdfSmartCopy Copy { get; set; }
        public Document Doc { get; set; }
        public FileStream NewFileStream { get; set; }
        public string Text { get; set; }
        public PdfReader FileReader { get; set; }
        string AccountNumber { get; set; }
        public List<string> OutputFileList { get; set; }
        public string InputFile { get; set; }
        public string CsvFile { get; set; }
        public string OutputPath { get; set; }
        public string OutputFileName { get; set; }
        public string TodaysDate { get; set; }
        public string FullOutputPath { get; set; }
        public Boolean BSuccess;

        #endregion

        #region PrintPDF

        #region From List of Strings

        public void PrintStatementPdf(List<string> filesToPrint)
        {
            #region From CSV

            //CSV File Reader
            using (var sr = new StreamReader(@"C:\AS400Report\Print Queue\CSV Files\statements.csv"))
            //Stream reader to read the Delimited File
            {
                var bytes = ReadAllBytes(DirectoryInstance.ToUserDirectory + @"\Google Drive\PrintingProgram\CompletedTemplates\PDFStatement.pdf");
                var bytes1 = bytes;
                TemplateSingleton.Lock(@"C:\AS400Report\Print Queue\CSV Files\statements.csv",
                    (f) =>
                    {
                        try
                        {
                            f.Write(bytes1, 0, bytes1.Length);
                        }
                        catch (IOException ioe)
                        {
                            Console.WriteLine(ioe);
                        }
                    });

                //2d Array. Left dimension is the page number, Right dimension is the text field number.
                var fields = new List<List<string>>();

                var readLine = sr.ReadLine();
                while (!sr.EndOfStream)
                {
                    if (readLine == null)
                    {
                        continue;
                    }

                    var field = readLine.Split('\t').ToList();
                    fields.Add(field);
                }

                #endregion

                foreach (var file in filesToPrint)
                {
                    var cfg = TemplateSingleton.Instance.CfgList[DirectoryInstance.NumberCompleted];
                    FormName = cfg.FormName;
                    Title = cfg.Title;
                    FullOutputPath = cfg.PdfPath;

                    //Clever way to check if a file exists. If it does, it automatically adds a _n where n in the number sequence.
                    if (Exists(FullOutputPath))
                    {
                        var counter = 2;
                        var fileLenBfrExtnsn = FullOutputPath.Length - 4;
                        var sequence = counter.ToString();
                        var withoutExtension = FullOutputPath.Substring(0, fileLenBfrExtnsn);
                        var extension = FullOutputPath.Substring(fileLenBfrExtnsn);
                        FullOutputPath = withoutExtension + "_" + sequence + extension;
                        while (Exists(FullOutputPath))
                        {
                            counter++;
                            sequence = counter.ToString();
                            FullOutputPath = withoutExtension + "_" + sequence + extension;
                        }
                    }

                    //The File Reader reads the template document; then I get the template Document's dimensions.
                    FileReader = new PdfReader(DirectoryInstance.ToUserDirectory + @"\Google Drive\PrintingProgram\CompletedTemplates\PDFStatement.pdf");
                    FileReader.GetPageSize(1);

                    bytes = ReadAllBytes(DirectoryInstance.ToUserDirectory + @"\Google Drive\PrintingProgram\CompletedTemplates\PDFStatement.pdf");

                    var bytes2 = bytes;
                    TemplateSingleton.Lock(DirectoryInstance.ToUserDirectory + @"\Google Drive\PrintingProgram\CompletedTemplates\PDFStatement.pdf",
                        (f) =>
                        {
                            try
                            {
                                f.Write(bytes2, 0, bytes2.Length);
                            }
                            catch (IOException ioe)
                            {
                                Console.WriteLine(ioe);
                            }
                        });


                    //Creates the Document from the template's dimensions in a using block to save memory.
                    using (Doc = new Document(FileReader.GetPageSize(1)))
                    {
                        //NewFileStream is the output PDF, PDFSmartCopy copies the template and pastes it to the PDFDoc. [later in code]
                        using (NewFileStream = new FileStream(FullOutputPath, FileMode.Create, FileAccess.Write))
                        {
                            using (Copy = new PdfSmartCopy(Doc, NewFileStream))
                            {
                                //Open the Document.
                                Doc.Open();
                                var counter = 0;
                                // generate one page per statement

                                ++counter;

                                // replace this with your PDF form template
                                FileReader = new PdfReader(DirectoryInstance.ToUserDirectory + @"\Google Drive\PrintingProgram\CompletedTemplates\PDFStatement.pdf");
                                using (var ms = new MemoryStream())
                                    //The stamper copies the contents of the old file onto a memory stream. I can then add text and proceed to print the stream to the Doc.
                                using (var stamper = new PdfStamper(FileReader, ms))
                                {
                                    //generate one page per statement
                                    foreach (var field in fields.Skip(1))
                                    {
                                        ++counter;

                                        //Each field in my pdf is called "form" [they're called form fields]
                                        var form = stamper.AcroFields;

                                        //TopRight Block
                                        form.SetField("Payoff-Thru#0", field[22]);
                                        form.SetField("Total-Due#0", field[14]);
                                        form.SetField("Account#0", field[0]);
                                        AccountNumber = field[0];
                                        //End TopRight Block

                                        //Address Block
                                        var addressBlock = GetAddress(fields[counter]);
                                        form.SetField(
                                            "School, Name1, Name2, Name3, Address1, Address2, Address3, City, State, Zip, Country",
                                            addressBlock);
                                        //End Address Block

                                        //Comments Block
                                        form.SetField("Msg1", field[31]);
                                        form.SetField("Msg2", field[32]);
                                        form.SetField("Msg3", field[33]);
                                        form.SetField("Msg4", field[34]);
                                        form.SetField("Msg5", field[35]);
                                        form.SetField("Msg6", field[36]);
                                        form.SetField("Msg7", field[37]);
                                        //End Comments Block

                                        /*Remove this before deploying the code to make statements
                                         * This is only here to test the comment box
                                         * with an example announcement of going paperless.*/
                                        field[39] =
                                            /*"We are now offering the opportunity to go paperless! \nIf you'd like to receive E-Statements, \nplease include your e-mail address on the return stub."*/
                                            "";
                                        form.SetField("Msg9", field[39]);
                                        //End E-Statement Block

                                        //EDSI Phone Number Block
                                        form.SetField("Tel-1, Tel-2", field[41] + " / " + field[42]);
                                        //End EDSI Phone Number Block

                                        //Left "Loan Amount" Block
                                        form.SetField("Orig-Amt#1", field[15]);
                                        form.SetField("Prin-Ltd", field[17]);
                                        form.SetField("Prin-Can", field[18]);
                                        form.SetField("Prin-Bal", field[16]);
                                        form.SetField("Int-Ltd", field[19]);
                                        form.SetField("Int-Can", field[20]);
                                        form.SetField("Payoff", field[21]);
                                        form.SetField("Payoff-Thru#1", field[22]);
                                        form.SetField("Date-Last-Pay", field[23]);
                                        form.SetField("Amt-Last-Pay", field[24]);
                                        //End Left "Loan Amount" Block

                                        //Right "Loan Balance" Block
                                        form.SetField("Orig-Amt#0", field[16]);
                                        form.SetField("Total-Due#1", field[14]);
                                        form.SetField("Pay-Past", field[26]);
                                        form.SetField("LC-Due", field[27]);
                                        form.SetField("Fee-Due", field[28]);
                                        form.SetField("Total-Due#2", field[14]);
                                        //End Right "Loan Balance" Block

                                        //Stub Right Block
                                        form.SetField("Account#1", field[0]);
                                        //Name Sub-Block
                                        form.SetField("Name#1", field[4] + " " + field[5] + " " + field[6] + "\n");
                                            //Name
                                        //End Name Sub-Block
                                        form.SetField("Payoff-Thru#2", field[22]);
                                        //End Stub Right Block

                                        //Begin Extra Stub Block
                                        /***********************************************************************************************************************************************************************************
                                        *form.SetField("Name#2", Field[4] + " " + Field[5] + " " + Field[6] + "\n");                                      //Name                                                              *
                                        *form.SetField("Address1, Address2, Address3", Field[7] + "\n" +                                                //Address 1                                                         *
                                        *    Field[8] + "\n" +                                                                                          //Address 2                                                         *
                                        *    Field[9] + "\n");                                                                                          //Address 3                                                         *
                                        *form.SetField("City, State, Zip", Field[10] + ", " + Field[11] + " " + Field[12] + "\n" + Field[13]);             //City, State Zip                                                   *
                                        *if (Field[41] != "" && Field[42] != "") { form.SetField("Phone", Field[41] + " / " + Field[42]); }                //If student has both phone numbers                                 *
                                        *else if (Field[41] != "") form.SetField("Phone", Field[41]);                                                    //If student has Telephone 1 ONLY                                   *
                                        *else if (Field[42] != "") form.SetField("Phone", Field[42]);                                                    //If student has Telephone 2 ONLY                                   *
                                        *else if (Field[41] == "" && Field[42] == "") { form.SetField("Phone", "No Contact info on file."); }            //Small phrase that prompts students to provide phone number.       *
                                        ************************************************************************************************************************************************************************************/
                                        //End Extra Stub Block

                                        //"Flatten" the form so it wont be editable/usable anymore
                                        stamper.FormFlattening = true;
                                    }
                                    //Add this page to my memory stream array.
                                    FileReader = new PdfReader(ms.ToArray());

                                    //Copy 1 page of the template to the output
                                    Copy.AddPage(Copy.GetImportedPage(FileReader, 1));
                                }
                            }
                            //Add meta data
                            Doc.AddAuthor("Educational Data Systems, Inc");
                            Doc.AddCreator("EDSI");
                            Doc.AddSubject(FormName);
                            Doc.AddTitle(Title);
                        }
                    }

                    DirectoryInstance.NumberCompleted++;
                }
            }
        }

        #endregion

        #region No Args

        //Using a CSV File
        public void PrintStatementPdf()
        {

            if (DirectoryInstance.ToUserDirectory == null)
            {
                DirectoryInstance.ToUserDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            }


            /**********************************************************************
                *                Delete All Statements Before Starting               *
                **********************************************************************/



            var csvList = new List<string>(10);

            Array.ForEach(Directory.GetFiles(@"C:\AS400Report\Print Queue\CSV Files", "state*.*"), csvList.Add);

            for (var csvCounter = 0; csvCounter < csvList.Count; csvCounter++)
            {
                CsvFile = csvList[csvCounter];
                InputFile = DirectoryInstance.ToUserDirectory + @"\Google Drive\PrintingProgram\CompletedTemplates\PDFStatement.pdf";
                OutputPath = @"C:\AS400Report\Reports\Statements\\";

                var bytes = ReadAllBytes(CsvFile);

                TemplateSingleton.Lock(CsvFile,
                    (f) =>
                    {
                        {
                            f.Write(bytes, 0, bytes.Length);
                        }
                    });


                using (var sr = new StreamReader(CsvFile))
                {
                    //2d Array. Left dimension is the page number, Right dimension is the text field number.
                    var fields = new List<List<string>>();
                    while (!sr.EndOfStream)
                    {
                        var field = sr.ReadLine().Split('\t').ToList();
                        fields.Add(field);
                    }

                    //var BlankVersion = new PrintBlankStatement();
                    //using (var PrintingTask = new Task(() => BlankVersion.PrintBlankStatementPDF(Fields)))
                    //{
                    //    PrintingTask.Start();
                    //    await Task.WhenAll(PrintingTask);

                    //}

                    TodaysDate = fields[1][22].Replace("/", "-");

                    //BlankVersion.PrintBlankStatementPDF(Fields);

                    OutputFileList = new List<string>(fields.Count - 1);
                    for (var i = 1; i < fields.Count; i++)
                    {
                        //Statement Naming
                        AccountNumber = fields[i][0];
                        OutputFileName = "statements_for_" + AccountNumber + "_" + TodaysDate + ".pdf";
                        FullOutputPath = OutputPath + OutputFileName;
                        OutputFileList.Add(FullOutputPath);
                    }

                    var counter = 0;

                    //generate one page per statement
                    foreach (var field in fields)
                    {
                        if (counter == 0)
                        {
                            counter++;
                            continue;
                        }

                        if (counter >= fields.Count)
                        {
                            Console.Write("What's happening here?");
                        }
                        try
                        {
                            PdfReader fileReader;
                            PdfStamper stamper;

                            FullOutputPath = OutputFileList[counter - 1];

                            //Meta Data
                            AccountNumber = field[0];
                            FormName = AccountNumber + "'s Statement";
                            Title = AccountNumber + "'s Statement Due: " + field[22] + ".";
                            //Clever way to check if a file exists. If it does, it automatically adds a _n where n in the number sequence.
                            if (Exists(FullOutputPath))
                            {
                                var counterVar = 2;
                                var fileLenBfrExtnsn = FullOutputPath.Length - 4;
                                var sequence = counterVar.ToString();
                                var withoutExtension = FullOutputPath.Substring(0, fileLenBfrExtnsn);
                                var extension = FullOutputPath.Substring(fileLenBfrExtnsn);
                                FullOutputPath = withoutExtension + "_" + sequence + extension;
                                while (Exists(FullOutputPath))
                                {
                                    counterVar++;
                                    sequence = counterVar.ToString();
                                    FullOutputPath = withoutExtension + "_" + sequence + extension;
                                }
                            }



                            //The File Reader reads the template document; then I get the template Document's dimensions.
                            fileReader = new PdfReader(InputFile);
                            fileReader.GetPageSize(1);
                            var doc = new Document(fileReader.GetPageSize(1));
                            NewFileStream = new FileStream(FullOutputPath, FileMode.Create, FileAccess.Write);
                            Copy = new PdfSmartCopy(doc, NewFileStream);

                            //I've added the using block to dispose of the stamper. If this breaks something, just remove the using(...){...} syntax.
                            using (stamper = new PdfStamper(fileReader, NewFileStream))
                            {
                                doc.Open();


                                //Each field in my pdf is called "form" [they're called form fields]
                                var form = stamper.AcroFields;

                                //TopRight Block
                                BSuccess = form.SetField("Payoff-Thru#0", field[22]);
                                BSuccess = form.SetField("Total-Due#0", field[14]);
                                BSuccess = form.SetField("Account#0", field[0]);
                                AccountNumber = field[0];
                                //End TopRight Block

                                //Address Block
                                var addressBlock = GetAddress(fields[counter]);
                                BSuccess = form.SetField("School, Name1, Name2, Name3, Address1, Address2, Address3, City, State, Zip, Country", addressBlock);
                                //End Address Block

                                //Comments Block
                                BSuccess = form.SetField("Msg1", field[31]);
                                BSuccess = form.SetField("Msg2", field[32]);
                                BSuccess = form.SetField("Msg3", field[33]);
                                BSuccess = form.SetField("Msg4", field[34]);
                                BSuccess = form.SetField("Msg5", field[35]);
                                BSuccess = form.SetField("Msg6", field[36]);
                                BSuccess = form.SetField("Msg7", field[37]);
                                //End Comments Block

                                /*Remove this before deploying the code to make statements
                                    * This is only here to test the comment box
                                    * with an example announcement of going paperless.*/
                                field[39] = "";
                                BSuccess = form.SetField("Msg9", field[39]);
                                //End E-Statement Block

                                //EDSI Phone Number Block
                                BSuccess = form.SetField("Tel-1, Tel-2", field[41] + " / " + field[42]);
                                //End EDSI Phone Number Block

                                //Left "Loan Amount" Block
                                BSuccess = form.SetField("Orig-Amt#1", field[15]);
                                BSuccess = form.SetField("Prin-Ltd", field[17]);
                                BSuccess = form.SetField("Prin-Can", field[18]);
                                BSuccess = form.SetField("Prin-Bal", field[16]);
                                BSuccess = form.SetField("Int-Ltd", field[19]);
                                BSuccess = form.SetField("Int-Can", field[20]);
                                BSuccess = form.SetField("Payoff", field[21]);
                                BSuccess = form.SetField("Payoff-Thru#1", field[22]);
                                BSuccess = form.SetField("Date-Last-Pay", field[23]);
                                BSuccess = form.SetField("Amt-Last-Pay", field[24]);
                                //End Left "Loan Amount" Block

                                //Right "Loan Balance" Block
                                BSuccess = form.SetField("Orig-Amt#0", field[16]);
                                BSuccess = form.SetField("Total-Due#1", field[14]);
                                BSuccess = form.SetField("Pay-Past", field[26]);
                                BSuccess = form.SetField("LC-Due", field[27]);
                                BSuccess = form.SetField("Fee-Due", field[28]);
                                BSuccess = form.SetField("Total-Due#2", field[14]);
                                //End Right "Loan Balance" Block

                                //Stub Right Block
                                BSuccess = form.SetField("Account#1", field[0]);
                                //Name Sub-Block
                                BSuccess = form.SetField("Name#1", field[4] + " " + field[5] + " " + field[6] + "\n");                        //Name
                                                                                                                                              //End Name Sub-Block
                                BSuccess = form.SetField("Amount Due", field[14]);
                                BSuccess = form.SetField("Payoff-Thru#2", field[22]);
                                //End Stub Right Block

                                //Begin Extra Stub Block
                                /****************************************************************************************************************************************************************************************************
                                *bSuccess = form.SetField("Name#2", Field[4] + " " + Field[5] + " " + Field[6] + "\n");                                         //Name                                                              *
                                *bSuccess = form.SetField("Address1, Address2, Address3", Field[7] + "\n" +                                                     //Address 1                                                         *
                                *    Field[8] + "\n" +                                                                                                          //Address 2                                                         *
                                *    Field[9] + "\n");                                                                                                          //Address 3                                                         *
                                *bSuccess = form.SetField("City, State, Zip", Field[10] + ", " + Field[11] + " " + Field[12] + "\n" + Field[13]);               //City, State Zip                                                   *
                                *if (Field[41] != "" && Field[42] != "") { bSuccess = form.SetField("Phone", Field[41] + " / " + Field[42]); }                  //If student has both phone numbers                                 *
                                *else if (Field[41] != "") bSuccess = form.SetField("Phone", Field[41]);                                                        //If student has Telephone 1 ONLY                                   *
                                *else if (Field[42] != "") bSuccess = form.SetField("Phone", Field[42]);                                                        //If student has Telephone 2 ONLY                                   *
                                *else if (Field[41] == "" && Field[42] == "") { bSuccess = form.SetField("Phone", "No Contact info on file."); }                //Small phrase that prompts students to provide phone number.       *
                                *****************************************************************************************************************************************************************************************************/
                                //End Extra Stub Block

                                //"Flatten" the form so it wont be editable/usable anymore
                                stamper.FormFlattening = true;

                                //Copy 1 page of the template to the output
                                Copy.AddPage(Copy.GetImportedPage(fileReader, 1));

                            }

                            //Add meta data
                            doc.AddAuthor("Educational Data Systems, Inc");
                            doc.AddCreator("EDSI");
                            doc.AddSubject(FormName);
                            doc.AddTitle(Title);

                            counter++;
                        }
                        catch (Exception err)
                        {
                            if (FileReader != null)
                                FileReader.Close();

                            if (Doc != null)
                                Doc.Close();

                            OutputFileName = null;

                            Console.WriteLine(err);
                        }

                        finally
                        {
                            if (FileReader != null)
                                FileReader.Close();

                            if (Doc != null)
                                Doc.Close();

                            OutputFileName = null;
                        }
                    }




                    IEnumerable<string> filenames = OutputFileList;
                    var mergedPdfName = OutputPath + "MonthlyStatements\\statements_for_" + TodaysDate + ".pdf";

                    //Clever way to check if a file exists. If it does, it automatically adds a _n where n in the number sequence.
                    if (Exists(mergedPdfName))
                    {
                        var counterVar = 2;
                        var fileLenBfrExtnsn = mergedPdfName.Length - 4;
                        var sequence = counterVar.ToString();
                        var withoutExtension = mergedPdfName.Substring(0, fileLenBfrExtnsn);
                        var extension = mergedPdfName.Substring(fileLenBfrExtnsn);
                        mergedPdfName = withoutExtension + "_" + sequence + extension;
                        while (Exists(mergedPdfName))
                        {
                            counterVar++;
                            sequence = counterVar.ToString();
                            mergedPdfName = withoutExtension + "_" + sequence + extension;
                        }
                    }
                    BSuccess = MergePdFs(filenames, mergedPdfName);


                    var choatestatements = OutputFileList.Where(s => s.Contains("12090"));
                    mergedPdfName = OutputPath + "ChoateStatements\\statements_for_" + TodaysDate + ".pdf";

                    //Clever way to check if a file exists. If it does, it automatically adds a _n where n in the number sequence.
                    if (Exists(mergedPdfName))
                    {
                        var counterVar = 2;
                        var fileLenBfrExtnsn = mergedPdfName.Length - 4;
                        var sequence = counterVar.ToString();
                        var withoutExtension = mergedPdfName.Substring(0, fileLenBfrExtnsn);
                        var extension = mergedPdfName.Substring(fileLenBfrExtnsn);
                        mergedPdfName = withoutExtension + "_" + sequence + extension;
                        while (Exists(mergedPdfName))
                        {
                            counterVar++;
                            sequence = counterVar.ToString();
                            mergedPdfName = withoutExtension + "_" + sequence + extension;
                        }
                    }

                    var fileNames = choatestatements as string[] ?? choatestatements.ToArray();
                    if (fileNames.Any())
                        BSuccess = MergePdFs(fileNames, mergedPdfName);
                }
            }
        }

        #endregion

        #endregion

        #region Combine Statements

        public static bool MergePdFs(IEnumerable<string> fileNames, string targetPdf)
        {
            var merged = true;
            using (var stream = new FileStream(targetPdf, FileMode.Create))
            {
                var document = new Document();

                //added a using block here to dispose of PdfCopy, if something breaks, remove it and revert it to var pdf = new PdfCopy(document, stream);
                using (var pdf = new PdfCopy(document, stream))
                {
                    PdfReader reader = null;
                    try
                    {
                        document.Open();
                        foreach (var file in fileNames)
                        {
                            reader = new PdfReader(file);
                            pdf.AddDocument(reader);
                            reader.Close();
                        }
                    }

                    catch (Exception)
                    {
                        merged = false;
                        reader?.Close();
                    }
                    finally
                    {
                        document.Close();
                    }
                }
                return merged;
            }
        }

        #endregion

    }
    #endregion

    #region Blank Statements (used when printing on your perforated paper)

    public class PrintBlankStatement : PdfCreator
    {
        #region Variables

        public string FullOutputPath { get; set; }
        public string OutputPath { get; set; }
        public string FormName { get; set; }   // form name
        public string Title { get; set; }      // title
        public PdfSmartCopy Copy { get; set; }
        public Document Doc { get; set; }
        public FileStream NewFileStream { get; set; }
        public string Text { get; set; }
        public PdfReader FileReader { get; set; }

        #endregion

        #region PrintPDF

        public void PrintBlankStatementPdf(List<List<string>> fields)
        {
            string outputFileName, accountNumber, fullOutputPath;
            var todaysDate = DateTime.Now.ToString("dd-MM-yyyy");


            var inputFile = DirectoryInstance.ToUserDirectory + @"\Google Drive\PrintingProgram\CompletedTemplates\BLANK_TEMPLATE.pdf";
            var outputPath = @"C:\AS400Report\Reports\BlankStatements\";

            if (!Exists(inputFile))
            {
                Console.WriteLine("Error at Line 937 - Print To PDF");
                inputFile = DirectoryInstance.ToUserDirectory + @"\Google Drive\PrintingProgram\CompletedTemplates\BLANK_TEMPLATE.pdf";
            }

            if (!Exists(outputPath))
            {
                Console.WriteLine("Error at Line 942 - Print To PDF");
                outputPath = @"C:\AS400Report\Reports\BlankStatements\";
            }


            var outputFileList = new List<string>(fields.Count - 1);
            for (var i = 1; i < fields.Count; i++)
            {
                //Statement Naming
                accountNumber = fields[i][0];
                outputFileName = "Statement_" + accountNumber + "_" + todaysDate + "BLANK.pdf";
                fullOutputPath = outputPath + outputFileName;


                //Clever way to check if a file exists. If it does, it automatically adds a _n where n in the number sequence.
                if (Exists(fullOutputPath))
                {
                    var counterVar = 2;
                    var fileLenBfrExtnsn = fullOutputPath.Length - 4;
                    var sequence = counterVar.ToString();
                    var withoutExtension = fullOutputPath.Substring(0, fileLenBfrExtnsn);
                    var extension = fullOutputPath.Substring(fileLenBfrExtnsn);
                    fullOutputPath = withoutExtension + "_" + sequence + extension;
                    while (Exists(fullOutputPath))
                    {
                        counterVar++;
                        sequence = counterVar.ToString();
                        fullOutputPath = withoutExtension + "_" + sequence + extension;
                    }
                }


                outputFileList.Add(fullOutputPath);
            }

            var counter = 0;

            //generate one page per statement
            foreach (var field in fields)
            {
                if (counter == 0)
                {
                    counter++;
                    continue;
                }

                if (counter >= fields.Count)
                {
                    Console.Write("What's happening here?");
                }
                try
                {
                    PdfStamper stamper;

                    fullOutputPath = outputFileList[counter - 1];

                    //Meta Data
                    accountNumber = field[0];
                    FormName = accountNumber + "'s Statement";
                    Title = accountNumber + "'s Statement Due: " + field[22] + ".";

                    //The File Reader reads the template document; then I get the template Document's dimensions.
                    var fileReader = new PdfReader(inputFile);
                    fileReader.GetPageSize(1);
                    var doc = new Document(fileReader.GetPageSize(1));
                    NewFileStream = new FileStream(fullOutputPath, FileMode.Create, FileAccess.Write);
                    Copy = new PdfSmartCopy(doc, NewFileStream);

                    //added a using block here to dispose of PdfStamper, if something breaks, remove it and revert it to Stamper = new PdfStamper(FileReader, NewFileStream);
                    using (stamper = new PdfStamper(fileReader, NewFileStream))
                    {
                        doc.Open();

                        //Each field in my pdf is called "form" [they're called form fields]
                        var form = stamper.AcroFields;

                        //TopRight Block
                        form.SetField("Payoff-Thru#0", field[22]);
                        form.SetField("Total-Due#0", field[14]);
                        form.SetField("Account#0", field[0]);
                        //End TopRight Block

                        //Address Block
                        var addressBlock = GetAddress(fields[counter]);
                        form.SetField("School, Name1, Name2, Name3, Address1, Address2, Address3, City, State, Zip, Country", addressBlock);
                        //End Address Block

                        //Comments Block
                        form.SetField("Msg1", field[31]);
                        form.SetField("Msg2", field[32]);
                        form.SetField("Msg3", field[33]);
                        form.SetField("Msg4", field[34]);
                        form.SetField("Msg5", field[35]);
                        form.SetField("Msg6", field[36]);
                        form.SetField("Msg7", field[37]);
                        //End Comments Block

                        /*Remove this before deploying the code to make statements
                         * This is only here to test the comment box
                         * with an example announcement of going paperless.*/
                        field[39] = "";
                        form.SetField("Msg9", field[39]);
                        //End E-Statement Block

                        //EDSI Phone Number Block
                        form.SetField("Tel-1, Tel-2", field[41] + " / " + field[42]);
                        //End EDSI Phone Number Block

                        //Left "Loan Amount" Block
                        form.SetField("Orig-Amt#1", field[15]);
                        form.SetField("Prin-Ltd", field[17]);
                        form.SetField("Prin-Can", field[18]);
                        form.SetField("Prin-Bal", field[16]);
                        form.SetField("Int-Ltd", field[19]);
                        form.SetField("Int-Can", field[20]);
                        form.SetField("Payoff", field[21]);
                        form.SetField("Payoff-Thru#1", field[22]);
                        form.SetField("Date-Last-Pay", field[23]);
                        form.SetField("Amt-Last-Pay", field[24]);
                        //End Left "Loan Amount" Block

                        //Right "Loan Balance" Block
                        form.SetField("Orig-Amt#0", field[16]);
                        form.SetField("Total-Due#1", field[14]);
                        form.SetField("Pay-Past", field[26]);
                        form.SetField("LC-Due", field[27]);
                        form.SetField("Fee-Due", field[28]);
                        form.SetField("Total-Due#2", field[14]);
                        //End Right "Loan Balance" Block

                        //Stub Right Block
                        form.SetField("Account#1", field[0]);
                        //Name Sub-Block
                        form.SetField("Name#1", field[4] + " " + field[5] + " " + field[6] + "\n");                        //Name
                                                                                                                                      //End Name Sub-Block
                        form.SetField("Amount Due", field[14]);
                        form.SetField("Payoff-Thru#2", field[22]);
                        //End Stub Right Block

                        //Begin Extra Stub Block
                        /****************************************************************************************************************************************************************************************************
                        *bSuccess = form.SetField("Name#2", Field[4] + " " + Field[5] + " " + Field[6] + "\n");                                         //Name                                                              *
                        *bSuccess = form.SetField("Address1, Address2, Address3", Field[7] + "\n" +                                                     //Address 1                                                         *
                        *    Field[8] + "\n" +                                                                                                          //Address 2                                                         *
                        *    Field[9] + "\n");                                                                                                          //Address 3                                                         *
                        *bSuccess = form.SetField("City, State, Zip", Field[10] + ", " + Field[11] + " " + Field[12] + "\n" + Field[13]);               //City, State Zip                                                   *
                        *if (Field[41] != "" && Field[42] != "") { bSuccess = form.SetField("Phone", Field[41] + " / " + Field[42]); }                  //If student has both phone numbers                                 *
                        *else if (Field[41] != "") bSuccess = form.SetField("Phone", Field[41]);                                                        //If student has Telephone 1 ONLY                                   *
                        *else if (Field[42] != "") bSuccess = form.SetField("Phone", Field[42]);                                                        //If student has Telephone 2 ONLY                                   *
                        *else if (Field[41] == "" && Field[42] == "") { bSuccess = form.SetField("Phone", "No Contact info on file."); }                //Small phrase that prompts students to provide phone number.       *
                        *****************************************************************************************************************************************************************************************************/
                        //End Extra Stub Block

                        //"Flatten" the form so it wont be editable/usable anymore
                        stamper.FormFlattening = true;

                        //Copy 1 page of the template to the output
                        Copy.AddPage(Copy.GetImportedPage(fileReader, 1));

                    }

                    //Add meta data
                    doc.AddAuthor("Educational Data Systems, Inc");
                    doc.AddCreator("EDSI");
                    doc.AddSubject(FormName);
                    doc.AddTitle(Title);


                    counter++;
                }
                catch (Exception err)
                {
                    FileReader?.Close();

                    Doc?.Close();

                    Console.WriteLine(err);
                }

                finally
                {
                    FileReader?.Close();

                    Doc?.Close();
                }
            }

            //Task.WaitAll();

            IEnumerable<string> filenames = outputFileList;
            var mergedPdfName = outputPath + "MonthlyStatements\\Statements_" + todaysDate + "_BLANK.pdf";

            //Clever way to check if a file exists. If it does, it automatically adds a _n where n in the number sequence.
            if (Exists(mergedPdfName))
            {
                var counterVar = 2;
                var fileLenBfrExtnsn = mergedPdfName.Length - 4;
                var sequence = counterVar.ToString();
                var withoutExtension = mergedPdfName.Substring(0, fileLenBfrExtnsn);
                var extension = mergedPdfName.Substring(fileLenBfrExtnsn);
                mergedPdfName = withoutExtension + "_" + sequence + extension;
                while (Exists(mergedPdfName))
                {
                    counterVar++;
                    sequence = counterVar.ToString();
                    mergedPdfName = withoutExtension + "_" + sequence + extension;
                }
            }

            MergePdFs(filenames, mergedPdfName);
        }
        #endregion
    }

    #endregion

    #region Remittance

    public class Remittance : PdfCreator
    {
        #region Variables

        public string FullOutputPath { get; set; }
        public string OutputPath { get; set; }
        public string FormName { get; set; }   // form name
        public string Title { get; set; }      // title
        public PdfSmartCopy Copy { get; set; }
        public Document Doc { get; set; }
        public FileStream NewFileStream { get; set; }
        public string Text { get; set; }
        public PdfReader FileReader { get; set; }

        #endregion

        #region PrintPDF

        public void PrintRemittancePdf(List<string> filesToPrint)
        {
            // generate one page per statement
            foreach (var file in filesToPrint)
            {
                //The File Reader reads the template document; then I get the template Document's dimensions.
                var templateFile = DirectoryInstance.ToUserDirectory +
                                      @"\Google Drive\PrintingProgram\CompletedTemplates\RemittanceTemplate.pdf";
                FileReader = new PdfReader(templateFile);
                FileReader.GetPageSize(1);

                var cfg = TemplateSingleton.Instance.FileCfgReportPairs
                   .FirstOrDefault(x => x.Value.Contains(file))
                   .Key;
                FormName = cfg.FormName;
                Title = cfg.Title;

                FullOutputPath = cfg.PdfPath;

                //Clever way to check if a file exists. If it does, it automatically adds a _n where n in the number sequence.
                if (Exists(FullOutputPath))
                {
                    var counterVar = 2;
                    var fileLenBfrExtnsn = FullOutputPath.Length - 4;
                    var sequence = counterVar.ToString();
                    var withoutExtension = FullOutputPath.Substring(0, fileLenBfrExtnsn);
                    var extension = FullOutputPath.Substring(fileLenBfrExtnsn);
                    FullOutputPath = withoutExtension + "_" + sequence + extension;
                    while (Exists(FullOutputPath))
                    {
                        counterVar++;
                        sequence = counterVar.ToString();
                        FullOutputPath = withoutExtension + "_" + sequence + extension;
                    }
                }

                //Use a StreamReader to read the file into a string.
                using (var sr = new StreamReader(file))
                {
                    //The string is equal to the stream reader read to end of file.
                    Text = sr.ReadToEnd();
                }

                //FCFC Scanner Object
                ScannerText = new FcfcScanner(Text);

                //Send FCFC Scanner the text from the file, PageArray is filled with the return value.
                PageArray = new List<string>(100);

                //Fill PageArray with the returned List<string> from FCFC Scanner.
                PageArray = ScannerText.NextPage(Text);
                PageArray.TrimExcess();
                PageNumber = 0;
                var bDone = false;

                //Creates the Document from the template's dimensions in a using block to save memory.
                using (Doc = new Document(FileReader.GetPageSize(1)))
                {
                    //NewFileStream is the output PDF, PDFSmartCopy copies the template and pastes it to the PDFDoc. [later in code]
                    using (NewFileStream = new FileStream(FullOutputPath, FileMode.Create, FileAccess.Write))
                    {
                        using (Copy = new PdfSmartCopy(Doc, NewFileStream))
                        {
                            //Open the Document.
                            Doc.Open();

                            while (!bDone)
                            {
                                PageNumber++;
                                if (PageArray.Count == 0)
                                {
                                    bDone = true;
                                    continue;
                                }


                                //The File Reader reads the template document; then I get the template Document's dimensions.
                                FileReader = new PdfReader(templateFile);
                                FileReader.GetPageSize(1);

                                using (var ms = new MemoryStream())
                                {
                                    //The stamper copies the contents of the old file onto a memory stream. I can then add text and proceed to print the stream to the Doc.
                                    using (var stamper = new PdfStamper(FileReader, ms))
                                    {

                                        //The text to be added to the page.
                                        var fieldText = "";

                                        //Retrieve the Template's form fields. I put the text inside the fields.
                                        var formfield = stamper.AcroFields;

                                        //Determines how much text I place in each page before adding a new page.
                                        FontSize = 9;
                                        float maxLines = PageArray.Count - 1;
                                        var i = 0;

                                        foreach (var line in PageArray)
                                        {
                                            if (PageArray[i] == "\r\n")
                                            {
                                                if (PageArray[i + 1] == "\r\n")
                                                {
                                                    i++;
                                                    continue;
                                                }
                                            }
                                            if (i >= maxLines)
                                            {
                                                PageArray = ScannerText.NextPage(Text);
                                            }

                                            fieldText = fieldText + line + "\r\n";
                                            i++;
                                        }

                                        // replace this with your field data for each page
                                        formfield.SetField("Text", fieldText);
                                        stamper.FormFlattening = true;
                                    }
                                    FileReader = new PdfReader(ms.ToArray());

                                    //Add the info from my stamper onto the page.
                                    Copy.AddPage(Copy.GetImportedPage(FileReader, 1));

                                }
                            }
                        }
                        //Add meta data
                        Doc.AddAuthor("Educational Data Systems, Inc");
                        Doc.AddCreator("EDSI");
                        Doc.AddSubject(FormName);
                        Doc.AddTitle(Title);
                    }
                }

                DirectoryInstance.NumberCompleted++;
            }
        }
    }

    #endregion


    #endregion

    #region Student Journal

    // ------------------------------------------------------------------------------
    // Creates a Student Journal PDF document
    // ------------------------------------------------------------------------------
    public class StudentJournal : PdfCreator
    {
        #region Variables

        public string FullOutputPath { get; set; }
        public string OutputPath { get; set; }
        public string FormName { get; set; }   // form name
        public string Title { get; set; }      // title
        public PdfSmartCopy Copy { get; set; }
        public Document Doc { get; set; }
        public PdfWriter Writer { get; set; }
        public FileStream NewFileStream { get; set; }
        public string Text { get; set; }
        public PdfReader FileReader { get; set; }
        public string Hex;
        public string Hex2;
        public Color Color;
        public Color Color2;
        public byte A;
        public byte A2;
        public byte R;
        public byte R2;
        public byte G;
        public byte G2;
        public byte B;
        public byte B2;
        public int Left;
        public int Right;
        public int Top;
        public int Bottom;
        public BaseFont FoundationFont;
        public BaseFont BookFont;
        public BaseFont AgMediumBookFont;

        #endregion

        #region PrintStudentJournalPDF

        public void PrintStudentJournalPdf(List<string> filesToPrint)
        {
            //Setting The Font
            const string fontpath = @"C:\Windows\Fonts\";
            var monoFont = BaseFont.CreateFont(fontpath + "Consola.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);

            //Foundation Roman(Used in Letterhead)
            var foundationFont = BaseFont.CreateFont(fontpath + "FoundationRoman-Italic.otf", BaseFont.CP1252, BaseFont.EMBEDDED);

            //AG Book, used everywhere else. (Medium is used for certain bolded portions).
            BaseFont.CreateFont(fontpath + "AGRoundedBook Book.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            BaseFont.CreateFont(fontpath + "AGSchoolbookMediumABook Book.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);

            //instance.CloseAndFinish();
            var letterhead = DirectoryInstance.ToUserDirectory + @"\Google Drive\PrintingProgram\TestDocs\Watermark\Letterhead.jpg";

            foreach (var file in filesToPrint)
            {

                var cfg = TemplateSingleton.Instance.FileCfgReportPairs
                   .FirstOrDefault(x => x.Value.Contains(file))
                   .Key;
                FormName = cfg.FormName;
                Title = cfg.Title;
                FullOutputPath = cfg.PdfPath;

                //Clever way to check if a file exists. If it does, it automatically adds a _n where n in the number sequence.
                if (Exists(FullOutputPath))
                {
                    var counterVar = 2;
                    var fileLenBfrExtnsn = FullOutputPath.Length - 4;
                    var sequence = counterVar.ToString();
                    var withoutExtension = FullOutputPath.Substring(0, fileLenBfrExtnsn);
                    var extension = FullOutputPath.Substring(fileLenBfrExtnsn);
                    FullOutputPath = withoutExtension + "_" + sequence + extension;
                    while (Exists(FullOutputPath))
                    {
                        counterVar++;
                        sequence = counterVar.ToString();
                        FullOutputPath = withoutExtension + "_" + sequence + extension;
                    }
                }

                ObjectLibrary.ObjLib.FullOutputPath = FullOutputPath;
                DirectoryInstance.OutputPaths.Add(FullOutputPath);
                var lineCounter = 0;

                //Creates the Document from the template's dimensions in a using block to save memory.
                using (Doc = new Document(
                            PageSize.A4.Rotate(),
                            Left,
                            Right,
                            Top,
                            Bottom))
                {
                    //NewFileStream is the output PDF, PDFSmartCopy copies the template and pastes it to the PDFDoc. [later in code]
                    using (NewFileStream = new FileStream(FullOutputPath, FileMode.Create, FileAccess.Write))
                    {
                        using (Writer = PdfWriter.GetInstance(Doc, NewFileStream))
                        {
                            Writer.CloseStream = false;

                            //Open the Document.
                            Doc.Open();

                            //Use a StreamReader to read the file into a string.
                            using (var sreader = new StreamReader(file))
                            {
                                //The string is equal to the stream reader read to end of file.
                                Text = sreader.ReadToEnd();
                            }

                            //FCFC Scanner Object; Store it in my ObjectLibrary.
                            ScannerText = new FcfcScanner(Text);
                            ObjectLibrary.ObjLib.Fcfc = ScannerText;

                            //Send FCFC Scanner the text from the file, PageArray is filled with the return value.
                            PageArray = new List<string>(100);

                            //Fill PageArray with the returned List<string> from FCFC Scanner.
                            PageArray = ObjectLibrary.ObjLib.Fcfc.NextPage(Text);
                            PageArray.TrimExcess();

                            while (ScannerText.PageLength > 0)
                            {
                                //Making a PDF Contentbyte for coloring blue rectangles and absolute positioned text
                                const float headerLines = 7;     // specifically student journal
                                const float lineMultiplier = 4;  // specifically student journal
                                var yOffset = 11.6f;
                                const float xPosRc = 40f;         //x-Position for the Rectangle
                                var heightRc = 452.8f;    //Height of Rectangle
                                var yPosRc = (heightRc) - (yOffset * lineMultiplier);               //y-Position for the Rectangle
                                const float widthRc = 810F;       //Width of Rectangle
                                const float minY = 13.2f;
                                const float yMax = 555.8f;
                                var cb = Writer.DirectContent;

                                //Applying the EDSI Letterhead
                                var letterHead = iTextSharp.text.Image.GetInstance(letterhead);
                                letterHead.SetAbsolutePosition(650, 530);
                                letterHead.ScalePercent(30);

                                //Alignment
                                letterHead.Alignment = iTextSharp.text.Image.UNDERLYING;
                                //letterHead.RotationDegrees = 270;

                                //Add the Letterhead
                                cb.AddImage(letterHead);

                                heightRc = (yMax - (headerLines * yOffset));

                                //Colors (Alice Blue, Royal Blue)
                                Hex = "#D1D8E5";
                                Hex2 = "#57718A";
                                //Color of Row Fill
                                Color = ColorTranslator.FromHtml(Hex);
                                A = Color.A;
                                R = Color.R;
                                G = Color.G;
                                B = Color.B;
                                //Color of Border
                                Color2 = ColorTranslator.FromHtml(Hex2);
                                A2 = Color2.A;
                                R2 = Color2.R;
                                G2 = Color2.G;
                                B2 = Color2.B;


                                //Drawing the Border
                                cb.MoveTo(xPosRc - 1, heightRc + 1);
                                cb.SetLineWidth(1);
                                cb.LineTo(widthRc + 1, heightRc + 1);
                                cb.SetLineWidth(1);
                                cb.LineTo(widthRc + 1, minY);
                                cb.SetLineWidth(1);
                                cb.LineTo(xPosRc - 1, minY);
                                cb.SetLineWidth(1);
                                cb.LineTo(xPosRc - 1, heightRc + 1.5);
                                cb.SetColorStroke(
                                    new BaseColor(
                                        R2,
                                        G2,
                                        B2,
                                        A2));
                                cb.Stroke();

                                for (var i = 0; i < PageArray.Count; i++)
                                {
                                    //Skipping the Header
                                    if (PageArray[i] == "\r\n")
                                    {
                                        i++;
                                        continue;
                                    }
                                    //Setting Background Color for Text
                                    //Even Lines Should be Colored Blue
                                    if ((i % 2 == 0) && (yPosRc > minY))
                                    {
                                        //Setting the Alice Blue Color and Opacity
                                        var gs1 = new PdfGState
                                        {
                                            FillOpacity = 1.0f
                                        };
                                        cb.SetGState(gs1);
                                        cb.SetColorStroke(
                                            new BaseColor(
                                                R,
                                                G,
                                                B,
                                                A));
                                        cb.SetColorFill(
                                            new BaseColor(
                                                R,
                                                G,
                                                B,
                                                A));
                                        gs1.FillOpacity = .99999f;
                                        cb.SetGState(gs1);
                                        yPosRc = (heightRc) - (yOffset * lineMultiplier);


                                        //Drawing a Rectangle and Filling it.
                                        cb.MoveTo(xPosRc, heightRc);
                                        cb.LineTo(widthRc, heightRc);
                                        cb.LineTo(widthRc, yPosRc);
                                        cb.LineTo(xPosRc, yPosRc);
                                        cb.LineTo(xPosRc, heightRc);
                                        cb.ClosePathFillStroke();

                                        //Moving the location
                                        heightRc -= (yOffset * lineMultiplier);
                                        yPosRc -= (yOffset * lineMultiplier);
                                        continue;
                                    }

                                    else if ((i % 2 > 0) && (yPosRc > minY))
                                    {
                                        var whiteOffset = yOffset;
                                        yPosRc = (heightRc) - (yOffset * lineMultiplier);
                                        //Drawing the Rectangle with lines
                                        cb.MoveTo(xPosRc, yPosRc);
                                        cb.LineTo(widthRc, yPosRc);
                                        cb.LineTo(widthRc, heightRc);
                                        cb.LineTo(xPosRc, heightRc);
                                        cb.LineTo(xPosRc, yPosRc);

                                        //Coloring the Rectangle and moving the location.
                                        cb.SetColorStroke(BaseColor.WHITE);
                                        cb.SetColorFill(BaseColor.WHITE);
                                        cb.ClosePathFillStroke();
                                        yPosRc -= (whiteOffset * lineMultiplier);
                                        heightRc -= (whiteOffset * lineMultiplier);
                                        continue;
                                    }

                                    if (yPosRc <= minY)
                                        i = PageArray.Count;

                                    if (i != PageArray.Count)
                                    {
                                        if (PageArray[i] == null)
                                            break;
                                    }
                                }

                                //Making a super cool header file
                                //Border Color and Font Size/Type
                                var tableColor = new BaseColor(
                                    R2, G2, B2, A2);
                                //Font for Header
                                cb.SetFontAndSize(foundationFont, 20);
                                cb.SetColorFill(tableColor);
                                cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "Student Loan Journal", 430, 550, 0);
                                cb.EndText();

                                //Making a PDF Contentbyte for building a header headerTable

                                var headerTable = new PdfPTable(13)
                                {
                                    //Actual width of headerTable in points
                                    TotalWidth = 770f,
                                    //Fixing the absolute width of the headerTable
                                    LockedWidth = true
                                };

                                //Setting up a content byte table size and placement
                                const float xTable = 39f;             //x-Position for the Rectangle
                                var heightTable = 475.6f;     //Height of Rectangle
                                const float widthTable = 811F;        //Width of Rectangle
                                                                      //float minYTable = 13.2f;
                                                                      //The height of my table's cells; the table has 4 rows so the table is 4 * yTableOffset; the width = the rightmost x coordinate - the left most x coordinate.
                                const float yTableOffset = 10f;
                                heightTable = heightTable + (4 * yTableOffset);
                                headerTable.TotalWidth = (widthTable - xTable);

                                //Relative column widths in proportions - 1/10 and 2/10
                                var widths = new[] { 2.06f, 1.25f, 1.03f, 0.86f, 1.03f, 0.90f, 0.90f, 0.90f, 0.90f, 0.46f, 0.91f, 0.90f, 0.90f };
                                headerTable.SetWidths(widths);
                                headerTable.HorizontalAlignment = 0;
                                //Leaving a gap before and after the headerTable
                                headerTable.SpacingBefore = 0f;
                                headerTable.SpacingAfter = 0f;
                                var headerArray = new[,] {
                                { "Borrower", "Acct Nbr - Seq", "Plan    Amount", "Loan Amount", "Accrued Intr", "Prin Due", "Prin Past Due", "Next Due Date",
                                    "Prin Coll LTD", "T    N", "Prin Cancel", "Defer Amount", "Credit Bureau" },
                                { "Address 1", "Telephone #      TC", "Status", "Current Bal", "Daily Intr Factor", "Intr Due", "Intr Past Due", "Last Pay Date",
                                    "Intr Coll LTD", "Y    B", "Intr Cancel", "Defer Date", "Last Reported"},
                                { "Address 2", "Intr rt    Grace", "BT  LT  Address", "Separation Dt", "Payoff Amount", "Late/Pen Due", "Days Past", "Last Pay Amt",
                                    "Late/Pen LTD", "P    R", "Cancel End", "Defer Code", "Last Rpt Code" },
                                { "City, State Zip", "Comment", "", "","Payoff Date", "Coll/Lit Due", "Total Due", "Last Activity",
                                    "Coll Cost LTD", "E     ", "", "", ""}
                                };


                                var fontPdf = new iTextSharp.text.Font(monoFont, 6, iTextSharp.text.Font.NORMAL);
                                fontPdf.SetColor(R2, G2, B2);

                                for (var i = 0; i < 4; i++)
                                {
                                    for (var j = 0; j < 13; j++)
                                    {
                                        var c = new PdfPCell(new Phrase(headerArray[i, j], fontPdf))
                                        {

                                            FixedHeight = 10f,
                                            BorderColor = tableColor,
                                            BorderWidth = 1f,
                                            HorizontalAlignment = Element.ALIGN_CENTER,
                                            VerticalAlignment = Element.ALIGN_MIDDLE
                                        };
                                        headerTable.AddCell(c);
                                    }
                                }

                                headerTable.WriteSelectedRows(0, -1, xTable, heightTable, cb);

                                //Setting up a content byte for absolute positioned text
                                const float xPos = 42f;
                                var yPos = 548.8F;
                                yOffset = 11.6f;
                                yPos = yPos - (2 * yOffset);
                                cb = Writer.DirectContent;
                                FontSize = cfg.FontSize;



                                var index = PageArray[7].TakeWhile(char.IsWhiteSpace).Count();

                                index = index - PageArray[4].Length;

                                var replaced = PageArray[4];
                                var released = PageArray[7].Substring(index);


                                PageArray[2] = replaced + released;
                                PageArray[3] = " ";
                                PageArray[4] = " ";
                                PageArray[7] = " ";

                                for (var i = 0; i < ScannerText.NumLines; i++)
                                {

                                    if ((PageArray[i] == "\r\n" || PageArray[i] == "\n" || PageArray[i] == "\r"))
                                    {
                                        continue;
                                    }

                                    cb.BeginText();
                                    cb.SetFontAndSize(monoFont, FontSize);
                                    cb.SetTextMatrix(xPos, yPos);  //(xPos, yPos)
                                    cb.SetColorFill(BaseColor.BLACK);
                                    //cb.SaveState();
                                    cb.ShowTextAligned(Element.ALIGN_JUSTIFIED, PageArray[i], xPos, yPos, 0);
                                    cb.EndText();
                                    yPos -= yOffset;
                                    lineCounter++;
                                    if (i == 2)
                                    {
                                        yPos -= yOffset;
                                    }

                                }

                                if (ScannerText.MorePages)
                                    Doc.NewPage();
                                PageArray = ScannerText.NextPage(Text);
                            }

                            Doc.Close();
                            Writer.Close();
                        }

                        NewFileStream.Close();

                    }

                    Doc = null;
                    Writer = null;
                    NewFileStream = null;
                }

                DirectoryInstance.NumberCompleted++;
            }
        }

        #endregion
    }

    #endregion

    #region Interest Paid

    public class InterestPaid
    {
        #region Variables

        public string _outputPath { get; set; }
        public string FormName { get; set; }   // form name
        public string Title { get; set; }      // title
        public PdfSmartCopy Copy { get; set; }
        public Document MyDoc { get; set; }
        public FileStream NewFileStream { get; set; }
        public string Text { get; set; }
        public PdfReader FileReader { get; set; }
        string SchoolId { get; set; }
        public List<string> OutputFileList { get; set; }
        public string InputFile { get; set; }
        public string CsvFile { get; set; }
        public string OutputPath { get; set; }
        public string OutputFileName { get; set; }
        public string TodaysDate { get; set; }
        public string FullOutputPath { get; set; }
        public Boolean BSuccess;

        #endregion

        #region PrintInterestPaid

        //Using a CSV File
        public void PrintInterestPaid()
        {
            #region Manage Directory

            if (DirectoryInstance.ToUserDirectory == null)
            {
                DirectoryInstance.ToUserDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            }


            /**********************************************************************
                *                Delete All Statements Before Starting               *
                **********************************************************************/

            TodaysDate = DateTime.Now.ToString("yyyy");

            var csvList = new List<string>(10);

            Array.ForEach(Directory.GetFiles(@"C:\AS400Report\Print Queue\CSV Files"), csvList.Add);

            foreach (var t1 in csvList)
            {
                Array.ForEach(Directory.GetFiles(@"C:\AS400Report\Print Queue\CSV Files", "yd*.*"), Delete);

                CsvFile = t1;
                var inputMasterFile = DirectoryInstance.ToUserDirectory + @"\Google Drive\PrintingProgram\CompletedTemplates\InterestPaid_MasterList.pdf";
                InputFile = DirectoryInstance.ToUserDirectory + @"\Google Drive\PrintingProgram\CompletedTemplates\InterestPaid_StudentDetails.pdf";
                OutputPath = DirectoryInstance.ToUserDirectory + @"C:\AS400Report\Reports\InterestPaid\";


                var bytes = ReadAllBytes(CsvFile);

                TemplateSingleton.Lock(CsvFile,
                    (f) =>
                    {
                        {
                            f.Write(bytes, 0, bytes.Length);
                        }
                    });

                using (var sr = new StreamReader(CsvFile))
                {
                    //2d Array. Left dimension is the page number, Right dimension is the text field number.
                    var fields = new List<List<string>>();
                    while (!sr.EndOfStream)
                    {
                        var field = sr.ReadLine().Split('\t').ToList();
                        fields.Add(field);
                    }

                    var duplicateSchoolIDs = fields.Skip(1).GroupBy(x => x[0])
                        .Select(group => group.First())
                        .ToList();

                    OutputFileList = new List<string>(duplicateSchoolIDs.Count);
                    foreach (var t in duplicateSchoolIDs)
                    {
                        //Student Details Naming
                        SchoolId = t[0];
                        OutputFileName = SchoolId + "_" + "InterestPaid_" + TodaysDate + ".pdf";
                        FullOutputPath = OutputPath + OutputFileName;
                        OutputFileList.Add(FullOutputPath);
                    }

                    var counter = 0;



                    var j = 0;

                    #endregion

                    #region Print Reports

                    var k = 1; var m = 0;
                    var groupList = new List<List<string>>(fields.Count);

                    for (var i = 1; i <= duplicateSchoolIDs.Count; i++)
                    {
                        MyDoc = null;
                        FileReader = null;
                        groupList = null;
                        m = 0;
                        counter = 0;


                        foreach (var field in fields.Skip(k))
                        {

                            if (groupList == null)
                            {
                                groupList = new List<List<string>>(fields.Count)
                                {
                                    fields[k]
                                };
                                k++;

                            }

                            else if (field[0] == groupList[m][0])
                            {
                                groupList.Add(field);
                                k++;
                                m++;
                            }

                            else
                            {
                                break;
                            }

                            j++;
                        }


                        FullOutputPath = OutputFileList[i - 1];



                        //Meta Data
                        var scid = fields[i][0];
                        SchoolId = scid;
                        FormName = SchoolId + "'s Master Student List";
                        Title = SchoolId + "'s Master Student Mailing List.";
                        //Clever way to check if a file exists. If it does, it automatically adds a _n where n in the number sequence.
                        if (Exists(FullOutputPath))
                        {
                            var counterVar = 2;
                            var fileLenBfrExtnsn = FullOutputPath.Length - 4;
                            var sequence = counterVar.ToString();
                            var withoutExtension = FullOutputPath.Substring(0, fileLenBfrExtnsn);
                            var extension = FullOutputPath.Substring(fileLenBfrExtnsn);
                            FullOutputPath = withoutExtension + "_" + sequence + extension;
                            while (Exists(FullOutputPath))
                            {
                                counterVar++;
                                sequence = counterVar.ToString();
                                FullOutputPath = withoutExtension + "_" + sequence + extension;
                            }
                        }

                        foreach (var field in groupList)
                        {
                            //if (counter == 0)
                            //{

                            //    counter++;
                            //    continue;
                            //}

                            if (counter >= fields.Count)
                            {
                                Console.Write("What's happening here?");
                            }

                            try
                            {

                                /****************************************************************
                                * Build the Master Mailer List:                                *
                                *      Contains a list of student names and account numbers    *
                                *      that paid $600 or more in interest.                     *
                                ****************************************************************/


                                PdfStamper stamper;
                                AcroFields form;
                                if (counter == 0)
                                {
                                    //The File Reader reads the template document; then I get the template Document's dimensions.
                                    using (FileReader = new PdfReader(inputMasterFile))
                                    {
                                        FileReader.GetPageSize(1);
                                        MyDoc = new Document(FileReader.GetPageSize(1));


                                        using (NewFileStream = new FileStream(FullOutputPath, FileMode.Create, FileAccess.Write))
                                        using (stamper = new PdfStamper(FileReader, NewFileStream))
                                        {
                                            Copy = new PdfSmartCopy(MyDoc, NewFileStream);
                                            MyDoc.Open();


                                            //Each field in my pdf is called "form" [they're called form fields]
                                            form = stamper.AcroFields;

                                            var q = 0;
                                            foreach (var student in groupList)
                                            {
                                                BSuccess = form.SetField("StudentName." + q, groupList[q][3]);
                                                BSuccess = form.SetField("AccountNumber." + q, groupList[q][2]);
                                                q++;
                                            }

                                            //"Flatten" the form so it wont be editable/usable anymore
                                            stamper.FormFlattening = true;

                                            //Copy 1 page of the template to the output
                                            Copy.AddPage(Copy.GetImportedPage(FileReader, 1));
                                        }
                                    }

                                    counter++;
                                }


                                /****************************************************************
                                 * Build the Student Details                                    *
                                 *      Contains details on student such as Account Number, SSN,*
                                 *      Amount Paid, and Address.                               *
                                 ****************************************************************/

                                FullOutputPath = OutputFileList[i - 1];

                                if (Exists(FullOutputPath))
                                {
                                    var counterVar = 2;
                                    var fileLenBfrExtnsn = FullOutputPath.Length - 4;
                                    var sequence = counterVar.ToString();
                                    var withoutExtension = FullOutputPath.Substring(0, fileLenBfrExtnsn);
                                    var extension = FullOutputPath.Substring(fileLenBfrExtnsn);
                                    FullOutputPath = withoutExtension + "_" + sequence + extension;
                                    while (Exists(FullOutputPath))
                                    {
                                        counterVar++;
                                        sequence = counterVar.ToString();
                                        FullOutputPath = withoutExtension + "_" + sequence + extension;
                                    }
                                }
                                //The File Reader reads the template document; then I get the template Document's dimensions.
                                using (FileReader = new PdfReader(InputFile))
                                {
                                    FileReader.GetPageSize(1);
                                    MyDoc = new Document(FileReader.GetPageSize(1));


                                    using (NewFileStream = new FileStream(FullOutputPath, FileMode.Create, FileAccess.Write))
                                    using (stamper = new PdfStamper(FileReader, NewFileStream))
                                    {
                                        Copy = new PdfSmartCopy(MyDoc, NewFileStream);
                                        MyDoc.Open();


                                        //Each field in my pdf is called "form" [they're called form fields]
                                        form = stamper.AcroFields;


                                        BSuccess = form.SetField("Name", field[3]);
                                        BSuccess = form.SetField("AccountNumber", field[2]);


                                        //Address Block
                                        BSuccess = form.SetField("AddressLine1", field[5]);
                                        BSuccess = form.SetField("AddressLine2", field[6]);
                                        BSuccess = form.SetField("City", field[7]);
                                        BSuccess = form.SetField("State", field[8]);
                                        BSuccess = form.SetField("Zip", field[9]);
                                        //End Address Block


                                        BSuccess = form.SetField("StudentSSN", field[1]);
                                        BSuccess = form.SetField("AmountofInterestPaid", field[4]);


                                        //"Flatten" the form so it wont be editable/usable anymore
                                        stamper.FormFlattening = true;

                                        //Copy 1 page of the template to the output
                                        Copy.AddPage(Copy.GetImportedPage(FileReader, 1));

                                        //Add meta data
                                        MyDoc.AddAuthor("Educational Data Systems, Inc");
                                        MyDoc.AddCreator("EDSI");
                                        MyDoc.AddSubject(FormName);
                                        MyDoc.AddTitle(Title);
                                    }
                                }
                            }

                            catch (Exception err)
                            {
                                Console.WriteLine(err);
                            }

                            finally
                            {
                                OutputFileName = null;
                            }

                            counter++;
                        }

                        //counter++;
                    }

                    #endregion

                    //Currently unused.
                    #region Prepare Merge

                    //IEnumerable<string> filenames = OutputFileList;
                    //var MergedPDFName = OutputPath + "MonthlyStatements\\Statements_" + TodaysDate + ".pdf";

                    ////Clever way to check if a file exists. If it does, it automatically adds a _n where n in the number sequence.
                    //if (System.IO.File.Exists(MergedPDFName))
                    //{
                    //    var counterVar = 2;
                    //    var fileLenBfrExtnsn = MergedPDFName.Length - 4;
                    //    var sequence = counterVar.ToString();
                    //    var withoutExtension = MergedPDFName.Substring(0, fileLenBfrExtnsn);
                    //    var Extension = MergedPDFName.Substring(fileLenBfrExtnsn);
                    //    MergedPDFName = withoutExtension + "_" + sequence + Extension;
                    //    while (System.IO.File.Exists(MergedPDFName))
                    //    {
                    //        counterVar++;
                    //        sequence = counterVar.ToString();
                    //        MergedPDFName = withoutExtension + "_" + sequence + Extension;
                    //    }
                    //}

                    //bSuccess = MergePDFs(filenames, MergedPDFName);
                }
            }

            #endregion
        }

        #endregion

        //Currently unused.
        #region Merge PDFs

        //public bool MergePDFs(IEnumerable<string> FileNames, string MergedPDFName)
        //{


        //    return true;
        //}


        #endregion
    }

    #endregion

    #region Perkins Reassignment

    public class PerkinsReassignment
    {
        #region Variables

        public string _outputPath { get; set; }
        public string FormName { get; set; }   // form name
        public string Title { get; set; }      // title
        public PdfSmartCopy Copy { get; set; }
        public Document MyDoc { get; set; }
        public FileStream NewFileStream { get; set; }
        public string Text { get; set; }
        public PdfReader FileReader { get; set; }
        string SchoolId { get; set; }
        public List<string> OutputFileList { get; set; }
        public string InputFile { get; set; }
        public string CsvFile { get; set; }
        public string OutputPath { get; set; }
        public string OutputFileName { get; set; }
        public string TodaysDate { get; set; }
        public string FullOutputPath { get; set; }
        public Boolean BSuccess;

        #endregion

        #region PrintInterestPaid

        //Using a CSV File
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2202:Do not dispose objects multiple times")]
        public void PrintPerkinsAssignment()
        {
            //Currently unused
            #region Radio Options

            //const string yes = "YES";
            //const string no = "NO";
            //const string perkins = "PERKINS";
            //const string direct = "DIRECT";
            //const string defense = "DEFENSE";
            //const string inSchool = "IN SCHOOL";
            //const string gracePeriod = "GRACE PERIOD";
            //const string deferment = "DEFERMENT";
            //const string repayment = "REPAYMENT";
            //const string hardship = "HARDSHIP";
            //const string incarceration = "INCARCERATION";
            //const string unemployment = "UNEMPLOYMENT";
            //const string liquidation = "LIQUIDATION";
            //const string refusalToPay = "REFUSAL TO PAY";
            //const string addressUnknown = "ADDRESS UNKNOWN";
            //const string disability = "TOTAL AND PERMANENT DISABILITY";


            #endregion

            #region Manage Directory

            if (DirectoryInstance.ToUserDirectory == null)
            {
                DirectoryInstance.ToUserDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            }


            /**********************************************************************
            *                Delete All Statements Before Starting               *
            **********************************************************************/

            TodaysDate = DateTime.Now.ToString("yyyy");

            var csvList = new List<string>(10);

            Array.ForEach(Directory.GetFiles(@"C:\AS400Report\Print Queue\CSV Files", "assign*.*"), csvList.Add);

            foreach (var t in csvList)
            {
                CsvFile = t;
                InputFile = DirectoryInstance.ToUserDirectory + @"\Google Drive\PrintingProgram\CompletedTemplates\PerkinsReassignmentForm.pdf";
                OutputPath = @"C:\AS400Report\Reports\PerkinsReassignment\";


                var bytes = ReadAllBytes(CsvFile);

                TemplateSingleton.Lock(CsvFile,
                    (f) =>
                    {
                        {
                            f.Write(bytes, 0, bytes.Length);
                        }
                    });

                using (var sr = new StreamReader(CsvFile))
                {
                    //2d Array. Left dimension is the page number, Right dimension is the text field number.
                    var fields = new List<List<string>>();
                    while (!sr.EndOfStream)
                    {
                        var field = sr.ReadLine().Split(',').ToList();
                        fields.Add(field);
                    }

                    foreach (var field in fields)
                    {
                        for (var quotekiller = 0; quotekiller < field.Count; quotekiller++)
                        {
                            field[quotekiller] = field[quotekiller].Replace('"'.ToString(), "");
                        }
                    }


                    OutputFileList = new List<string>(fields.Count - 1);
                    for (var i = 1; i < fields.Count; i++)
                    {
                        SchoolId = fields[i][2];
                        OutputFileName = SchoolId + "_" + "PerkinsReassignment_" + TodaysDate + ".pdf";
                        FullOutputPath = OutputPath + OutputFileName;
                        OutputFileList.Add(FullOutputPath);
                    }

                    #endregion

                    #region Print Reports


                    MyDoc = null;
                    FileReader = null;

                    var counter = 0;

                    FullOutputPath = OutputFileList[0];




                    //Clever way to check if a file exists. If it does, it automatically adds a _n where n in the number sequence.
                    if (Exists(FullOutputPath))
                    {
                        var counterVar = 2;
                        var fileLenBfrExtnsn = FullOutputPath.Length - 4;
                        var sequence = counterVar.ToString();
                        var withoutExtension = FullOutputPath.Substring(0, fileLenBfrExtnsn);
                        var extension = FullOutputPath.Substring(fileLenBfrExtnsn);
                        FullOutputPath = withoutExtension + "_" + sequence + extension;
                        while (Exists(FullOutputPath))
                        {
                            counterVar++;
                            sequence = counterVar.ToString();
                            FullOutputPath = withoutExtension + "_" + sequence + extension;
                        }
                    }

                    foreach (var field in fields.Skip(1))
                    {
                        //Meta Data
                        var scid = field[2].Substring(0, 4);
                        SchoolId = scid;
                        FormName = SchoolId + "'s Perkins Reassignment Form";
                        Title = SchoolId + "'s Perkins Reassignment Form";
                        if (counter >= fields.Count)
                        {
                            Console.Write("What's happening here?");
                        }

                        try
                        {
                            /****************************************************************
                             * Print the Perkins Reassignment Data                          *
                             *      Contains details on student such as Account Number, SSN,*
                             *      Amount Paid, and Address.                               *
                             ****************************************************************/
                            if (field == fields[1])
                            {
                                MyDoc = null;
                                FileReader = null;
                            }


                            FullOutputPath = OutputFileList[counter];

                            if (Exists(FullOutputPath))
                            {
                                var counterVar = 2;
                                var fileLenBfrExtnsn = FullOutputPath.Length - 4;
                                var sequence = counterVar.ToString();
                                var withoutExtension = FullOutputPath.Substring(0, fileLenBfrExtnsn);
                                var extension = FullOutputPath.Substring(fileLenBfrExtnsn);
                                FullOutputPath = withoutExtension + "_" + sequence + extension;
                                while (Exists(FullOutputPath))
                                {
                                    counterVar++;
                                    sequence = counterVar.ToString();
                                    FullOutputPath = withoutExtension + "_" + sequence + extension;
                                }
                            }
                            //The File Reader reads the template document; then I get the template Document's dimensions.
                            using (FileReader = new PdfReader(InputFile))
                            {
                                MyDoc = new Document(FileReader.GetPageSize(1));


                                using (NewFileStream = new FileStream(FullOutputPath, FileMode.Create, FileAccess.Write))
                                {
                                    PdfStamper stamper;
                                    using (stamper = new PdfStamper(FileReader, NewFileStream))
                                    {
                                        Copy = new PdfSmartCopy(MyDoc, NewFileStream);
                                        MyDoc.Open();


                                        //Each field in my pdf is called "form" [they're called form fields]
                                        var form = stamper.AcroFields;

                                        #region Borrower Information

                                        //Block 1 Current Name
                                        BSuccess = form.SetField("CurrentName", field[4]);

                                        //Block 2 Previous Name
                                        BSuccess = form.SetField("PreviousName", "N/A");

                                        //Block 3 Social Security Number
                                        BSuccess = form.SetField("SSN", field[0]);

                                        //Block 4 Date of Birth
                                        BSuccess = form.SetField("DoB", field[15]);

                                        //Block 5 Departure Date (MM/DD/YYY)
                                        BSuccess = form.SetField("DepartureDate", field[29]);

                                        //Block 6 Current Address (Number and Street)
                                        BSuccess = form.SetField("CurrentAddress", field[9]);

                                        //Block 7 Telephone Number
                                        BSuccess = form.SetField("PhoneNumber", field[16]);

                                        //Block 8 City
                                        BSuccess = form.SetField("City", field[11]);

                                        //Block 9 State
                                        BSuccess = form.SetField("State", field[12]);

                                        //Block 10 Zip Code
                                        BSuccess = form.SetField("Zip", field[13]);

                                        #endregion

                                        #region Cosigner Information

                                        //Block 11 Name of Cosigner
                                        BSuccess = form.SetField("NameCosigner", "");

                                        //Block 12 Social Security Number of Cosigner
                                        BSuccess = form.SetField("SSNCosigner", "");

                                        //Block 13 Current Address Cosigner (Number and Street)
                                        BSuccess = form.SetField("CurrentAddressCosigner", "");

                                        //Block 14 Telephone Number Cosigner
                                        BSuccess = form.SetField("PhoneNumberCosigner", "");

                                        //Block 15 City Cosigner
                                        BSuccess = form.SetField("CityCosigner", "");

                                        //Block 16 State Cosigner
                                        BSuccess = form.SetField("StateCosigner", "");

                                        //Block 17 Zip Code Cosigner
                                        BSuccess = form.SetField("ZipCosigner", "");

                                        #endregion

                                        #region Loan Information: Historical

                                        //Block 18 Type of Loan - Radio
                                        BSuccess = form.SetField("18.0", "•");

                                        //Block 19 Applicable Interest Rate on the Loan
                                        BSuccess = form.SetField("InterestRate", "5.00");

                                        //Block 20 Date of last Disbursement (MM/DD/YYYY)
                                        BSuccess = form.SetField("DateLastDisbursement", field[22]);

                                        //Block 21 Date Last Grace Period Ended (MM/DD/YYYY)
                                        BSuccess = form.SetField("DateLastGracePeriod", field[31]);

                                        //Block 22 Date of Default, if defaulted (MM/DD/YYYY)
                                        BSuccess = form.SetField("DateDefault", field[30]);

                                        //Block 23 Was this Loan Accelerated? - Radio + Text
                                        BSuccess = form.SetField("23.1", "•");
                                        if (field[5] == "Yes".ToUpper())
                                        {
                                            BSuccess = form.SetField("23.0", "•");
                                        }
                                        else
                                        {
                                            BSuccess = form.SetField("23.1", "•");
                                            BSuccess = form.SetField("DateAccelerated", "N/A");
                                        }

                                        //Block 24 Was this Loan Litigated? - Radio + Text
                                        BSuccess = form.SetField("24.1", "•");
                                        BSuccess = form.SetField("DateLitigated", "N/A");

                                        //Block 25 Borrower Repayment Status, if loan is not in default - Radio
                                        if (field[71] != "DEFAULT")
                                        {
                                            if (field[71].ToUpper() == "Repayment".ToUpper())
                                            {
                                                BSuccess = form.SetField("25.3", "•");
                                            }

                                        }



                                        //Block 26 Reason the loan has been reassigned - Radio + Text
                                        BSuccess = form.SetField("26.3", "•");

                                        #endregion

                                        Copy.AddPage(Copy.GetImportedPage(FileReader, 1));
                                        //myDoc.NewPage();

                                        #region Loan Information: Financial

                                        //Block 27 Disbursement Amount
                                        BSuccess = form.SetField("DisbursementAmount", field[19]);

                                        //Block 28 Principal Amount Adjusted
                                        BSuccess = form.SetField("PrincipalAmtAdjusted", field[20]);

                                        //Block 29 Principal Amount Repaid
                                        BSuccess = form.SetField("PrinAmtRepaid", field[32]);

                                        //Block 30 Principal Amount Canceled
                                        BSuccess = form.SetField("PrinAmtCancelled", field[36]);

                                        //Block 31 Principal Amount Outstanding
                                        BSuccess = form.SetField("PrinAmtOutstanding", field[23]);

                                        //Block 32 Collection Costs Repaid
                                        BSuccess = form.SetField("CollectionsCostRepaid", field[35]);

                                        //Block 33 Interest Repaid
                                        BSuccess = form.SetField("InterestRepaid", field[33]);

                                        //Block 34 Interest Canceled
                                        BSuccess = form.SetField("InterestCancelled", field[37]);

                                        //Block 35 Interest Due
                                        BSuccess = form.SetField("InterestDue", field[25]);

                                        //Block 36 Collection Costs, Penalty, Late Charges
                                        BSuccess = form.SetField("CollCosts_Penalties_LateCharges", field[26]);

                                        //Block 37 Total Amount Outstanding (Sum of items 31, 35, 36)
                                        var x = Convert.ToDouble(field[23]);
                                        var y = Convert.ToDouble(field[25]);
                                        var z = Convert.ToDouble(field[26]);
                                        var sum = x + y + z;
                                        BSuccess = form.SetField("TotalAmtOutstanding", sum.ToString());

                                        #endregion

                                        #region Cancellation Information

                                        //Each row contains Type of Cancellation, Percentage Rate, Principal Canceled, Interest Canceled,
                                        //Cancellation Service Start Date (MM/DD/YYY), and Cancellation Service End Date (MM/DD/YYYY)
                                        var counter2 = 0;
                                        for (var counterVar = 0; counterVar < 30; counterVar++, counter2++)
                                        {
                                            if (field[38 + counterVar] == "0")
                                                break;

                                            BSuccess = form.SetField("TypeCancellation" + counter2.ToString(), field[38 + counterVar] /*+ "/" + Field[39 + counterVar]*/);
                                            counterVar++;   //counterVar++;
                                            BSuccess = form.SetField("PercentageRate." + counter2.ToString(), field[38 + counterVar]);
                                            counterVar++;
                                            BSuccess = form.SetField("PrinCancelled." + counter2.ToString(), field[38 + counterVar]);
                                            counterVar++;
                                            BSuccess = form.SetField("IntrCancelled." + counter2.ToString(), field[38 + counterVar]);
                                            counterVar++;
                                            BSuccess = form.SetField("CancellationStartDate." + counter2.ToString(), field[38 + counterVar]);
                                            counterVar++;
                                            BSuccess = form.SetField("CancellationEndDate." + counter2.ToString(), field[38 + counterVar]);
                                        }

                                        #endregion


                                        //"Flatten" the form so it wont be editable/usable anymore
                                        stamper.FormFlattening = true;

                                        //Copy 1 page of the template to the output

                                        Copy.AddPage(Copy.GetImportedPage(FileReader, 2));


                                        stamper.Close();

                                        //Add meta data
                                        MyDoc.AddAuthor("Educational Data Systems, Inc");
                                        MyDoc.AddCreator("EDSI");
                                        MyDoc.AddSubject(FormName);
                                        MyDoc.AddTitle(Title);
                                    }
                                }
                            }
                        }

                        catch (Exception err)
                        {
                            Console.WriteLine(err);
                        }

                        finally
                        {
                            OutputFileName = null;
                        }

                        counter++;
                    }

                    //counter++;


                    #endregion

                    //Currently unused
                    #region Prepare Merge

                    //        IEnumerable<string> filenames = OutputFileList;
                    //        var MergedPDFName = OutputPath + "MonthlyStatements\\Statements_" + TodaysDate + ".pdf";

                    //        //Clever way to check if a file exists. If it does, it automatically adds a _n where n in the number sequence.
                    //        if (System.IO.File.Exists(MergedPDFName))
                    //        {
                    //            var counterVar = 2;
                    //            var fileLenBfrExtnsn = MergedPDFName.Length - 4;
                    //            var sequence = counterVar.ToString();
                    //            var withoutExtension = MergedPDFName.Substring(0, fileLenBfrExtnsn);
                    //            var Extension = MergedPDFName.Substring(fileLenBfrExtnsn);
                    //            MergedPDFName = withoutExtension + "_" + sequence + Extension;
                    //            while (System.IO.File.Exists(MergedPDFName))
                    //            {
                    //                counterVar++;
                    //                sequence = counterVar.ToString();
                    //                MergedPDFName = withoutExtension + "_" + sequence + Extension;
                    //            }
                    //        }

                    //        bSuccess = MergePDFs(filenames, MergedPDFName);
                }
            }

            #endregion
        }

        #endregion

        //Currently unused.
        #region Merge PDFs

        //public bool MergePDFs(IEnumerable<string> FileNames, string MergedPDFName)
        //{


        //    return true;
        //}


        #endregion
    }

    #endregion

    #region Default

    // ------------------------------------------------------------------------------
    // Creates a Default PDF document
    // ------------------------------------------------------------------------------
    public class Default : PdfCreator
    {
        #region Variables

        public string FullOutputPath { get; set; }
        public string OutputPath { get; set; }
        public string FormName { get; set; }   // form name
        public string Title { get; set; }      // title
        public PdfSmartCopy Copy { get; set; }
        public Document Doc { get; set; }
        public FileStream NewFileStream { get; set; }
        public string Text { get; set; }
        public PdfReader FileReader { get; set; }
        public string Hex;
        public string Hex2;
        public Color Color;
        public Color Color2;
        public byte A;
        public byte A2;
        public byte R;
        public byte R2;
        public byte G;
        public byte G2;
        public byte B;
        public byte B2;
        public int Left;
        public int Right;
        public int Top;
        public int Bottom;
        public BaseFont FoundationFont;
        public BaseFont BookFont;
        public BaseFont AgMediumBookFont;
        PdfWriter Writer { get; set; }

        #endregion

        #region PrintDefaultPDF

        public void PrintDefaultPdf(List<string> filesToPrint)
        {
            //Setting The Font
            const string fontpath = @"C:\Windows\Fonts\";
            var monoFont = BaseFont.CreateFont(fontpath + "Consola.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);

            //Foundation Roman(Used in Letterhead)
            BaseFont.CreateFont(fontpath + "FoundationRoman.otf", BaseFont.CP1252, BaseFont.EMBEDDED);

            //AG Book, used everywhere else. (Medium is used for certain bolded portions).
            var bookFont = BaseFont.CreateFont(fontpath + "AGRoundedBook Book.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            var agMediumBookFont = BaseFont.CreateFont(fontpath + "AGSchoolbookMediumABook Book.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);

            var batchPath = PrintQueue.DirectoryInstance.ToUserDirectory + @"\Google Drive\Judy's Folder\Batchsheets\";
            



            // generate one page per statement
            foreach (var file in filesToPrint)
            {
                var cfg = TemplateSingleton.Instance.FileCfgReportPairs
                    .FirstOrDefault(x => x.Value.Contains(file))
                    .Key;
                FormName = cfg.FormName;
                Title = cfg.Title;
                var pdfName = cfg.PdfPath.Substring(cfg.PdfPath.LastIndexOf(@"\", StringComparison.Ordinal));
                FullOutputPath = (FormName == "batch_sheet")? PrintQueue.DirectoryInstance.ToUserDirectory + @"\Google Drive\Judy's Folder\Batchsheets" + pdfName : cfg.PdfPath;


                //Clever way to check if a file exists. If it does, it automatically adds a _n where n in the number sequence.
                if (Exists(FullOutputPath))
                {
                    var counter = 2;
                    var fileLenBfrExtnsn = FullOutputPath.Length - 4;
                    var sequence = counter.ToString();
                    var withoutExtension = FullOutputPath.Substring(0, fileLenBfrExtnsn);
                    var extension = FullOutputPath.Substring(fileLenBfrExtnsn);
                    FullOutputPath = withoutExtension + "_" + sequence + extension;
                    while (Exists(FullOutputPath))
                    {
                        counter++;
                        sequence = counter.ToString();
                        FullOutputPath = withoutExtension + "_" + sequence + extension;
                    }
                }

                //_fullOutputPath = _fullOutputPath.Substring(0, _fullOutputPath.Length - 3);
                FontSize = cfg.FontSize;

                ObjectLibrary.ObjLib.FullOutputPath = FullOutputPath;

                var letterhead = DirectoryInstance.ToUserDirectory + @"\Google Drive\PrintingProgram\TestDocs\Watermark\EDSILogo.PNG";

                //Creates the Document from the template's dimensions in a using block to save memory.
                using (Doc = (cfg.LandScape) ? new Document
                    (
                            PageSize.A4.Rotate(),
                            Left,
                            Right,
                            Top,
                            Bottom
                    ) : new Document
                    (
                            PageSize.A4,
                            Left,
                            Right,
                            Top,
                            Bottom
                            ))
                {

                    //NewFileStream is the output PDF, PDFSmartCopy copies the template and pastes it to the PDFDoc. [later in code]
                    using (NewFileStream = new FileStream(FullOutputPath, FileMode.Create, FileAccess.ReadWrite))
                    {
                        using (Writer = PdfWriter.GetInstance(Doc, NewFileStream))
                        {
                            Writer.CloseStream = false;

                            //Open the Document.
                            Doc.Open();



                            //Use a StreamReader to read the file into a string.
                            using (var sreader = new StreamReader(file))
                            {
                                //The string is equal to the stream reader read to end of file.
                                Text = sreader.ReadToEnd();
                            }

                            //FCFC Scanner Object; Store it in my ObjectLibrary.
                            ScannerText = new FcfcScanner(Text);
                            ObjectLibrary.ObjLib.Fcfc = ScannerText;

                            //Send FCFC Scanner the text from the file, PageArray is filled with the return value.
                            PageArray = new List<string>(100);

                            //Fill PageArray with the returned List<string> from FCFC Scanner.
                            PageArray = ScannerText.NextPage(Text);
                            PageArray.TrimExcess();

                            //PDF Content Byte (lets me add text or images to the page).
                            var cb = Writer.DirectContent;

                            while (ScannerText.PageLength > 0)
                            {

                                if (cfg.LandScape)
                                {
                                    //Applying the EDSI Letterhead
                                    var letterHead = iTextSharp.text.Image.GetInstance(letterhead);
                                    letterHead.SetAbsolutePosition(650, 530);
                                    letterHead.ScalePercent(30);

                                    //Allignment
                                    letterHead.Alignment = iTextSharp.text.Image.UNDERLYING;
                                    //letterHead.RotationDegrees = 270;

                                    //Add the Letterhead
                                    cb.AddImage(letterHead);
                                }
                                else
                                {
                                    //Applying the EDSI Letterhead
                                    var letterHead = iTextSharp.text.Image.GetInstance(letterhead);
                                    letterHead.SetAbsolutePosition(400, 750);
                                    letterHead.ScalePercent(30);

                                    //Allignment
                                    letterHead.Alignment = iTextSharp.text.Image.UNDERLYING;
                                    //letterHead.RotationDegrees = 270;

                                    //Add the Letterhead
                                    cb.AddImage(letterHead);
                                }



                                //Setting up a content byte for absolute positioned text
                                const float xPos = 42f;
                                var offset = FontSize * 1.26f;
                                var yPos = (cfg.TopPrintOffset > 0) ? cfg.TopPrintOffset : (cfg.LandScape) ? (offset * 55.0f) : ((offset * 80.0f) + 100);

                                cb = Writer.DirectContent;
                                yPos = yPos - (3 * offset);

                                if (!cfg.LandScape)
                                    yPos = 700f;

                                for (var i = 0; i < ScannerText.NumLines; i++)
                                {
                                    if (cfg.Title != "Disbursement Reports")
                                    {
                                        if (PageArray[i] == "\r\n" || PageArray[i] == "\n" || PageArray[i] == "\r")
                                        {
                                            continue;
                                        }
                                    }
                                    cb.BeginText();
                                    cb.SetFontAndSize(monoFont, FontSize);
                                    cb.SetTextMatrix(xPos, yPos);  //(xPos, yPos)
                                    cb.SetColorFill(BaseColor.BLACK);
                                    //cb.SaveState();
                                    cb.ShowTextAligned(Element.ALIGN_JUSTIFIED, PageArray[i], xPos, yPos, 0);
                                    cb.EndText();
                                    yPos -= offset;
                                    
                                }
                                if (ScannerText.MorePages)
                                    Doc.NewPage();
                                PageArray = ScannerText.NextPage(Text);
                            }

                            Doc.Close();
                            Writer.Close();
                        }

                        NewFileStream.Close();
                    }
                }
                Writer = null;
                Doc = null;
                NewFileStream = null;

                DirectoryInstance.NumberCompleted++;

            }
        }

        #endregion

        #region IsFileLocked

        public static bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                return true;
            }
            finally
            {
                stream?.Close();
            }
            return false;
        }

        #endregion
    }


    #endregion



    #endregion

    #region TemplateSingleton

    /************************************************************
    / Class:   TemplateSingleton                   			    *
    / Purpose:  Stores useful Template Values. Also has a method	*
    /			    to Pair File lists and store the Pairings.  *
    / Methods:  Choose Template, Pair File Lists                 *
    *************************************************************/
    public sealed class TemplateSingleton
    {

        #region Singleton Variables

        private static readonly TemplateSingleton instance = new TemplateSingleton();
        public List<List<string>> StreamList { get; set; }   //Automatically implemented property.
        public List<List<List<string>>> ListOfStringLists { get; set; }   //Automatically implemented property.
        public Dictionary<ReportConfig, string> FileCfgReportPairs { get; set; }
        public static TemplateSingleton Instance
        {
            get
            {
                return instance;
            }
        }

        public Dictionary<ReportConfig, List<List<string>>> CfgReportPairs { get; set; }

        private Templates ReportTemplate { get; set; }

        private int NumTemplates { get; set; }

        public List<ReportConfig> CfgList { get; set; }

        public int TotalReports { get; set; }

        public int CurrentReportNum { get; set; }

        public List<Templates> ReportsList { get; set; }

        #endregion

        #region Singleton Constructor

        private TemplateSingleton()
        {

        }

        #endregion

        #region ChooseTemplate

        /************************************************************
        / Method:   ChooseTemplate  								*
        / Purpose:  Determines which template to use, similar to 	*
        /			    main(). This method should be called when 	*
        /			    you're ready to print your Stream data.		*
        / Returns:  Template Type.                                  *
        *************************************************************/
        public Templates ChooseTemplate(ReportConfig cfg)
        {
            var reportType = cfg.Title.Substring(0, 4);
            reportType = reportType.ToUpper();

            ReportTemplate = (Templates)reportType;
            reportType = ReportTemplate.ToString();

            //Check the Enum type
            if (reportType == Templates.Statement.ToString())
            {
                ObjectLibrary.ObjLib.InputFileName = Templates.Statement.ToString();
            }

            else if (reportType == Templates.Blank.ToString())
            {
                ObjectLibrary.ObjLib.InputFileName = Templates.Blank.ToString();
            }

            else if (reportType == Templates.Invoice.ToString())
            {
                ObjectLibrary.ObjLib.InputFileName = Templates.Invoice.ToString();
            }

            else if (reportType == Templates.StudentJournal.ToString())
            {
                //Send to StudentJournal Method
                ObjectLibrary.ObjLib.InputFileName = Templates.StudentJournal.ToString();
            }

            else
            {
                ObjectLibrary.ObjLib.InputFileName = Templates.Default.ToString();
            }

            return ReportTemplate;
        }

        #endregion

        #region Pair File Lists

        public void PairFileLists()
        {

            Instance.ReportsList = new List<Templates>(DirectoryInstance.FileEntries.Length);
            foreach (var file in DirectoryInstance.FileEntries)
            {
                var cfg = DirectoryInstance.GetConfigFile(DirectoryInstance.NumPaired);
                Instance.CfgList.Add(cfg);
                Instance.ReportsList.Add(Instance.ChooseTemplate(cfg));
                DirectoryInstance.NumPaired++;
            }

            //So I tried something weird here. I have a struct with a constructor. I made a list of structs here.
            var pairStructs = new List<FileCfgPair>(DirectoryInstance.FileEntries.Length);
            FileCfgPair pairStruct;

            var i = 0;

            //Then, in a loop, I filled the list with structs that build my key value pairs.
            foreach (var file in DirectoryInstance.FileEntries)
            {
                pairStruct = new FileCfgPair(file, Instance.CfgList[i]);
                pairStructs.Add(pairStruct);
                i++;
            }

            //I'm bad at working with Enumerables, so I broke them into dictionary definitions that are easier to use.
            FileCfgReportPairs = new Dictionary<ReportConfig, string>();

            //foreach (KeyValuePair<List<List<string>>, ReportConfig> pair in myEnumerable)
            foreach (var pair in pairStructs)

            {
                try
                {
                    //For each different Document, Pair the Document and its CFG file.
                    FileCfgReportPairs.Add(pair.DocumentConfigType, pair.DocumentReportType);
                }
                catch
                {
                    Console.Write("Error assigning Key Value Pairs.");
                    break;
                }
            }
        }

        #endregion

        #region Lock

        public static void Lock(string path, Action<FileStream> action)
        {
            using (var autoResetEvent = new AutoResetEvent(false))
            {

                while (true)
                {
                    try
                    {
                        using (var file = Open(path,
                                                    FileMode.OpenOrCreate,
                                                    FileAccess.ReadWrite,
                                                    FileShare.Write))
                        {
                            action?.Invoke(file);
                            break;
                        }
                    }
                    catch (IOException)
                    {
                        using (var fileSystemWatcher =
                            new FileSystemWatcher(Path.GetDirectoryName(path))
                            {
                                EnableRaisingEvents = true
                            })
                        {

                            fileSystemWatcher.Changed +=
                                (o, e) =>
                                {
                                    if (Path.GetFullPath(e.FullPath) == Path.GetFullPath(path))
                                    {
                                        if (autoResetEvent != null) autoResetEvent.Set();
                                    }
                                };

                            autoResetEvent.WaitOne();
                        }
                    }
                }
            }
        }

        #endregion

    };

    #endregion

    #region Document Pairing Strategies

    #region Template Type Safe Enum Class

    /************************************************************
    / Template Type Safe Enum Class:                			*
    / Purpose:  While it sounds like it makes more sense in my 	*
    /			    TemplateHandler class, it actually doesn't.	*
    /			    My Parent class needs to know when it should*
    /               add text or leave it blank.                 *
    / Returns:  The format of the Report                        *
    *************************************************************/
    public sealed class Templates
    {
        #region Dictionary

        private static readonly Dictionary<string, Templates> Instance = new Dictionary<string, Templates>();

        #endregion

        #region Conversion Operator

        //Convert the Prefix into the Template
        public static explicit operator Templates(string str)
        {
            Templates result;
            return Instance.TryGetValue(str, out result) ? result : Default;
        }

        #endregion

        #region Template Variables

        //Variables, can add more later if needed.
        private readonly String _name;

        public static readonly Templates Error = new Templates("EROR");
        public static readonly Templates Statement = new Templates("STMT");
        public static readonly Templates Invoice = new Templates("INVO");
        public static readonly Templates StudentJournal = new Templates("STUD");
        public static readonly Templates Blank = new Templates("BLNK");
        public static readonly Templates Default = new Templates("0");

        #endregion

        #region Private Constructor

        private Templates(String prefix)
        {
            var prefix1 = prefix;
            Instance[prefix1] = this;
            switch (prefix1)
            {
                case "STMT":
                    {
                        _name = "PDFStatement";
                        break;
                    }
                case "EROR":
                    {
                        _name = "Error";
                        break;
                    }
                case "JOUR":
                case "STUD":
                    {
                        _name = "StudentJournal";
                        break;
                    }
                case "REMI":
                case "INVO":
                    {
                        _name = "Invoice";
                        break;
                    }
                case "BLNK":
                    {
                        prefix1 = "Blank";
                        break;
                    }
                default:
                    {
                        _name = "Default";
                        break;
                    }
            }
        }
        #endregion

        #region ToString() Override

        public override String ToString()
        {
            return _name;
        }

        #endregion

    }

    #endregion

    #region Zip

    /************************************************************
    / Class:   Zip + its IEnumerable Pair          			    *
    / Purpose:  Allows me to pair to lists together so I can 	*
    /			    iterate through both of them simultaneously *
    / Returns:  IEnumerable Key Value Pair.                     *
    *************************************************************/
    public static class Zipper
    {
        public static IEnumerable<KeyValuePair<TLeft, TRight>> Zip<TLeft, TRight>(
            this IEnumerable<TLeft> left, IEnumerable<TRight> right)
        {
            return Zip(left, right, (x, y) => new KeyValuePair<TLeft, TRight>(x, y));
        }

        // accepts a projection from the caller for each pair
        public static IEnumerable<TResult> Zip<TLeft, TRight, TResult>(
            this IEnumerable<TLeft> left, IEnumerable<TRight> right,
            Func<TLeft, TRight, TResult> selector)
        {
            using (var leftE = left.GetEnumerator())
            using (var rightE = right.GetEnumerator())
            {
                while (leftE.MoveNext() && rightE.MoveNext())
                {
                    var handler = selector;
                    if (handler != null)
                        yield return handler(leftE.Current, rightE.Current);
                }
            }
        }
    }

    #endregion

    #region FilePairing Methodology

    /************************************************************
    / Struct:   FileCFGPair          			                *
    / Purpose:  Allows me to pair to a file name to a cfg type. *
    / Returns:  No Return Value                                 *
    *************************************************************/
    struct FileCfgPair
    {
        public string DocumentReportType;
        public ReportConfig DocumentConfigType;

        public FileCfgPair(string fileToPrint, ReportConfig cfg)
        {
            DocumentConfigType = cfg;
            DocumentReportType = fileToPrint;
        }
    }

    #endregion

    #endregion
}






