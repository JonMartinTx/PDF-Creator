//using Report;
//using System;
//using System.Diagnostics;
//using System.IO;
//using System.Linq;
//using System.Text.RegularExpressions;
//using System.Xml.Serialization;

//namespace ConfigBuilder
//{
//    #region De-serialized XML

//    /****************************************************************************
//     * The following classes are designed to build an xml file for every foreign*
//     *      file type. If a file is not recognized by the approved templates or *
//     *      CFG files, then it will generate the necessary XML to allow the file*
//     *      to print normally.                                                  *
//     ****************************************************************************/

//    class Report
//    {
//        public string Title { get; set; }
//        public string Path { get; set; }
//        public string SearchPattern { get; set; }
//        public string FormName { get; set; }
//        public Printer _printer;
//        public Pdf _pdf { get; set; }
//        public Layout _layout { get; set; }

//        public Report(string t, string p, string f, Printer printer, Pdf pdf, Layout l)
//        {
//            Title = t;
//            Path = p;
//            SearchPattern = "*.*";
//            FormName = f;
//            _printer = printer;
//            _pdf = pdf;
//            _layout = l;
//        }
//    }

//    class Printer
//    {
//        public bool HardCopy { get; set; }
//        public readonly string PrinterName = "HP OfficeJet Pro 8630 (Network)";

//        public Printer()
//        {
//            HardCopy = false;
//        }
//    }

//    class Pdf
//    {
//        public bool CreatePdf { get; set; }
//        public string PdfName { get; set; }
//        public string PdfPath { get; set; }
//        public PdfMargins _pdfMargins { get; set; }

//        public Pdf(bool b, string n, string p, PdfMargins M)
//        {
//            CreatePdf = b;
//            PdfName = n;
//            PdfPath = p;
//            _pdfMargins = M;
//        }
//    }

//    class PdfMargins
//    {
//        public int Top { get; set; }
//        public int Bottom { get; set; }
//        public int Left { get; set; }
//        public int Right { get; set; }

//        public PdfMargins(int t, int b, int l, int r)
//        {
//            Top = t;
//            Bottom = b;
//            Left = l;
//            Right = r;
//        }
//    }

//    class Layout
//    {
//        public Font _font { get; set; }
//        public Margins _margins { get; set; }
//        public string Orientation { get; set; }

//        public Layout(Font f)
//        {
//            _font = f;
//        }
//    }

//    class Margins
//    {
//        public int Top { get; set; }
//        public int Bottom { get; set; }
//        public int Left { get; set; }
//        public int Right { get; set; }

//        public Margins(int t, int b, int l, int r)
//        {
//            Top = t;
//            Bottom = b;
//            Left = l;
//            Right = r;
//        }
//    }

//    class Font
//    {
//        public string FontName { get; set; }
//        public float FontSize { get; set; }

//        public Font(string name, int size)
//        {
//            FontName = name;
//            FontSize = size;
//        }
//    }

//    #endregion

//    #region CFGBuilder

//    class CFGBuilder
//    {

//        #region Properties

//        [XmlArray]
//        [XmlArrayItem(typeof(Report))]
//        [XmlArrayItem(typeof(Printer))]
//        [XmlArrayItem(typeof(Pdf))]
//        [XmlArrayItem(typeof(PdfMargins))]
//        [XmlArrayItem(typeof(Layout))]
//        [XmlArrayItem(typeof(Margins))]
//        public Report _report { get; set; }
//        public string _fileName;
//        public string FileName
//        {
//            get {   return _fileName;   }
//            set {   _fileName = value;  }
//        }


//        #endregion

//        #region Constructor

//        public CFGBuilder(string n)
//        {
//            FileName = n;
//        }

//        #endregion

//        #region Methods

//        #region GenerateXML

//        public void GenerateXML()
//        {
//            var todaysDate = DateTime.Today;

//            var date = todaysDate.ToString();
//            //Need to pull file data from inside the open file. I'm going to include the FileStream here, more work needs to be done.
//            using (var reader = new StreamReader(FileName))
//            {
//                string Text = reader.ReadToEnd();

//                var ScannerText = new FCFCScanner(Text);

//                var ModText = (ScannerText.NextPage(Text)).AsEnumerable();

//                //This is currently hard coded to NOT print to paper.
//                var myPrinter = new ConfigBuilder.Printer();

//                //Default Values will change font size based on the amount printed. It will be a range between 7 and 12.
//                var myFont = new Font("Consolas", 8);

//                //Hard Coded. To be honest, the margins don't matter whatsoever anymore.
//                var myMargins = new Margins(75, 75, 40, 40);
//                var myLayout = new Layout(myFont);



//                var Minimum = "";
//                var Maximum = "";

//                Maximum = ModText.OrderByDescending(a => a.Length).First().ToString();
//                Minimum = ModText.OrderBy(a => a.Length).First().ToString();

//                myLayout.Orientation = (Maximum.Length >= 120) ? "LandScape" : "Portrait";

//                var myPdfMargins = new PdfMargins(75, 75, 40, 40);



//                var hsNum = new Regex("[1-2][0-2][0-1][0-9][0-5]");
//                var colNum = new Regex("[1-7][0-9][0-7][0-9]");

//                var match = (Regex.IsMatch(Text, "[1-2][0-2][0-1][0-9][0-5]")) ? hsNum.Match(Text) : (Regex.IsMatch(Text, "[1-7][0-9][0-7][0-9]")) ? colNum.Match(Text) : null;

//                string prefix;

//                var SchoolNumber = (match != null) ? match.ToString() : null;

//                switch (SchoolNumber)
//                {
//                    case "1006": prefix = "BOY"; break;
//                    case "1007": prefix = "ACI"; break;
//                    case "1008": prefix = "ACC"; break;
//                    case "1087": prefix = "AAI"; break;
//                    case "1255": prefix = "ALC"; break;
//                    case "1321": prefix = "LAC"; break;
//                    case "1533": prefix = "MCC"; break;
//                    case "1908": prefix = "CEN"; break;
//                    case "1911": prefix = "COL"; break;
//                    case "1929": prefix = "KWU"; break;
//                    case "1946": prefix = "TAB"; break;
//                    case "2447": prefix = "WCC"; break;
//                    case "2567": prefix = "YOR"; break;
//                    case "3167": prefix = "USA"; break;
//                    case "3176": prefix = "CAL"; break;
//                    case "3415": prefix = "EAS"; break;
//                    case "5041": prefix = "MHB"; break;
//                    case "6755": prefix = "BMC"; break;
//                    case "7437": prefix = "PTI"; break;
//                    case "12020": prefix = "NAT"; break;
//                    case "12035": prefix = "WNH"; break;
//                    case "12040": prefix = "STA"; break;
//                    case "12045": prefix = "WES"; break;
//                    case "12050": prefix = "STM"; break;
//                    case "12060": prefix = "EPI"; break;
//                    case "12065": prefix = "HOT"; break;
//                    case "12070": prefix = "KEN"; break;
//                    case "12075": prefix = "SUF"; break;
//                    case "12080": prefix = "GOU"; break;
//                    case "12085": prefix = "WIL"; break;
//                    case "12090": prefix = "CHO"; break;
//                    case "12095": prefix = "DFA"; break;
//                    case "12100": prefix = "TAF"; break;
//                    case "12105": prefix = "GUN"; break;
//                    case "12110": prefix = "NMH"; break;
//                    case "12115": prefix = "CAN"; break;
//                    case "20100": prefix = "WAD"; break;
//                    case "20110": prefix = "CHD"; break;
//                    case null: prefix = "NUL"; break;
//                    default: prefix = "ERR"; break;
//                }

//                var myPdf = (prefix != "NUL" && prefix != "ERR") ? new Pdf(
//                        true,
//                        prefix.ToLower() + "_" + date + ".pdf",
//                        @"\SharePoint\EDSIReports - Reports\Reports\" + SchoolNumber + @"\",
//                        myPdfMargins
//                        )

//                        : new Pdf(
//                        true,
//                        date + ".pdf",
//                        @"\SharePoint\EDSIReports - Reports\Reports\Default\",
//                        myPdfMargins
//                        );




//                var myReport = new Report("NewReport", @"\SharePoint\EDSIReports - Documents\prtq\", "NewForm", myPrinter, myPdf, myLayout);

//                var serializer = new XmlSerializer(typeof(Report));
//                serializer.Serialize(File.Create(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\SharePoint\EDSIReports - AS400Report\AS400Report\PrintSetUp\" + "NewReport.xml"), myReport);

//                CreateInstructions();
//                Process.Start("notepad.exe", Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\SharePoint\EDSIReports - Documents\TestDocs\NewXmlInstructions");
//            }


//        }

//        #endregion

//        #region CreateInstructions

//        public void CreateInstructions()
//        {

//            var lines = new string[20];

//            lines[0] = "You've created a new XML Configuration File." + Environment.NewLine + "Before it can be used, you must follow the instructions below to update the code." + Environment.NewLine;
//            lines[1] = "/t1) Go to C:\\Users\\Owner\\SharePoint\\EDSIReports - AS400Report\\AS400Report\\PrintSetUp and locate the xml file called \"NewReport.xml\". Open it.";
//            lines[2] = "\t2) In a second window, Go to " + @"C:\Users\Owner\SharePoint\EDSIReports - Documents\prtq" + "and right click the file named " + FileName + ". Open with NotePad++.";
//            lines[3] = "\t3) Compare the Report to your xml file: ";
//            lines[4] = "\t\tA) Change the Title of the xml Document (located under <Report> <Title> TEXT TO CHANGE </Title>. The title of the config file should match the title of the Report you opened in step 2.";
//            lines[5] = "\t\tB) Change the FormName (located at <Report> <FormName> TEXT TO CHANGE </FormName>) to match the new Title you just updated.";
//            lines[6] = "Verify that the Orientation, near the bottom of the xml, correctly shows Landscape or Portrait. If the orientation of wrong, type the correct orientation, capitalizing the first letter" + Environment.NewLine
//                + "\t\t\t Example: Landscape or Portrait";
//            lines[7] = "\t4) Save the xml file and close it. Rename the xml file to the first 4 letters of the Report's Title, in all caps." + Environment.NewLine + "\t\t Example: BATC for Batchsheet or TRAN for Transaction Journal";
//            lines[8] = "\t5) Close the file to print.";
//            lines[9] = "\t6) Go to " + @"C:\Users\Owner\SharePoint\EDSIReports - AS400Report\AS400Report\AS400Report" + " and open AS400Report.csproj.";
//            lines[10] = "\t\tA) On the left hand side of the page, there is a pane called 'Solution Explorer'. Inside that explorer, search for a file called ReportConfig.cs and open it.";
//            lines[11] = "\t\t\t• You may also locate this by going to the top of the program (where you see File, Edit, etc) and selecting View -> Solution Explorer or by typing CTRL + ALT + 0"
//                + Environment.NewLine + "\t\t\t• Alternatively, you can navigate to " + @"C: \Users\Owner\SharePoint\EDSIReports - AS400Report\AS400Report\AS400Report" + " and right clicking ReportConfig.cs -> open with NotePad ++.";
//            lines[12] = "\t\tB) Go to line number 184. Insert this text at the beginning of the line: 'prefix == \"[Enter your four letter xml name from step 4 in ALL CAPS]\" ||'. It's important that you DO NOT INCLUDE the single quotes."
//                + Environment.NewLine + "\t\t\t• Your text should match the format of the text above. Ensure that it does.";

//            lines[13] = Environment.NewLine + "When you have verified that the above steps have been taken, Save the Program. You may now attempt to run the program. Should the program fail, call Jon at (928) 208-5308 or email at jonmartin1990@gmail.com.";


//            System.IO.File.WriteAllLines(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\SharePoint\EDSIReports - Documents\TestDocs\NewXmlInstructions", lines);

//        }

//        #endregion

//        #endregion
//    }

//    #endregion
//}
