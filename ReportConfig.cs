using System;
using System.IO;
using System.Xml;

namespace Report
{
    public class ReportConfig
    {
        #region Variables

        public string Title;
        public string InputPath;
        public string SearchPattern;
        public string Printer;
        public string FormName;
        public int TopMargin;
        public int BottomMargin;
        public int LeftMargin;
        public int RightMargin;
        public bool LandScape;
        public string FontName;
        public int FontSize;
        public bool BPdf;
        public bool BPrint;
        public string PdfPath;
        public int PdfTop;
        public int PdfBottom;
        public int PdfLeft;
        public int PdfRight;
        public string ErrMsg;
        private string _pdfFile;
        private string _pdfPath;
        public float TopPrintOffset;
        public bool SuppressBlankLines;

        #endregion

        #region Constructor

        public ReportConfig(string configFileName, string inputFile)
        {
            Title = "Generic Report";
            InputPath = string.Empty;
            SearchPattern = "*.*";
            Printer = "HP OfficeJet Pro 8630 (Network)";
            FormName = "Generic";
            TopMargin = 75;
            BottomMargin = 75;
            LeftMargin = 40;
            RightMargin = 40;
            LandScape = false;
            FontName = "Consolas";
            FontSize = 8;
            BPdf = false;
            BPrint = true;
            PdfPath = String.Empty;
            PdfTop = 10;
            PdfBottom = 10;
            PdfLeft = 10;
            PdfRight = 10;
            _pdfFile = String.Empty;
            _pdfPath = String.Empty;
            TopPrintOffset = 0.0f;
            SuppressBlankLines = false;

            LoadConfig(configFileName);

            if (BPdf)
            {
                _pdfFile = ScanReplace(inputFile);
                if (BPdf && _pdfPath.Length > 0 && _pdfFile.Length > 0)
                    PdfPath = _pdfPath.TrimEnd() + _pdfFile;
            }
        }

        #endregion

        #region Methods

        #region Load Config

        // Load the XML configuration
        private void LoadConfig(string fileName)
        {
            XmlNode node;

            var doc = new XmlDocument();

            try
            {
                doc.Load(fileName);

                if ((node = doc.SelectSingleNode("/Report/Title")) != null)
                    Title = node.InnerText;

                if ((node = doc.SelectSingleNode("/Report/Path")) != null)
                    InputPath = node.InnerText;

                if ((node = doc.SelectSingleNode("/Report/Printer/HardCopy")) != null)
                    BPrint = (node.InnerText == "true");

                if ((node = doc.SelectSingleNode("/Report/Printer/PrinterName")) != null)
                    Printer = node.InnerText;

                if ((node = doc.SelectSingleNode("/Report/Pdf/CreatePdf")) != null)
                    BPdf = (node.InnerText == "true");

                if ((node = doc.SelectSingleNode("/Report/Pdf/PdfName")) != null)
                    _pdfFile = node.InnerText;

                if ((node = doc.SelectSingleNode("/Report/Pdf/PdfPath")) != null)
                    _pdfPath = node.InnerText;

                if ((node = doc.SelectSingleNode("/Report/Layout/Pdf/PdfMargins/Top")) != null)
                    TopMargin = Convert.ToInt16(node.InnerText);

                if ((node = doc.SelectSingleNode("/Report/Layout/Pdf/PdfMargins/Bottom")) != null)
                    BottomMargin = Convert.ToInt16(node.InnerText);

                if ((node = doc.SelectSingleNode("/Report/Layout/Pdf/PdfMargins/Left")) != null)
                    LeftMargin = Convert.ToInt16(node.InnerText);

                if ((node = doc.SelectSingleNode("/Report/Layout/Pdf/PdfMargins/Right")) != null)
                    RightMargin = Convert.ToInt16(node.InnerText);

                if ((node = doc.SelectSingleNode("/Report/Pdf/TopPrintOffset")) != null)
                    TopPrintOffset = (float) Convert.ToDouble(node.InnerText);

                if ((node = doc.SelectSingleNode("/Report/Pdf/SuppressBlankLines")) != null)
                    SuppressBlankLines = (node.InnerText == "true");

                if ((node = doc.SelectSingleNode("/Report/FormName")) != null)
                    FormName = node.InnerText;

                if ((node = doc.SelectSingleNode("/Report/Layout/Margins/Top")) != null)
                    TopMargin = Convert.ToInt16(node.InnerText);

                if ((node = doc.SelectSingleNode("/Report/Layout/Margins/Bottom")) != null)
                    BottomMargin = Convert.ToInt16(node.InnerText);

                if ((node = doc.SelectSingleNode("/Report/Layout/Margins/Left")) != null)
                    LeftMargin = Convert.ToInt16(node.InnerText);

                if ((node = doc.SelectSingleNode("/Report/Layout/Margins/Right")) != null)
                    RightMargin = Convert.ToInt16(node.InnerText);

                if ((node = doc.SelectSingleNode("/Report/Layout/Orientation")) != null)
                {
                    LandScape = node.InnerText == "Landscape";
                }

                if ((node = doc.SelectSingleNode("/Report/Layout/Font/FontName")) != null)
                    FontName = node.InnerText;

                if ((node = doc.SelectSingleNode("/Report/Layout/Font/FontSize")) != null)
                    FontSize = Convert.ToInt16(node.InnerText);
            }
            catch (Exception e)
            {
                ErrMsg = e.Message;
            }
        }

        #endregion

        #region ScanReplace

        // Function to scan for school id, fund id and date substitution values and replace them with the real values
        // ----------------------------------------------------------------------------------------------------------
        private string ScanReplace(string origPath)
        {
            var origFileName = String.Empty;
            var newFileName = String.Empty;
            string prefix, strScid, strCons, strDate;
            int scid, cons;

            origFileName = Path.GetFileName(origPath);
            scid = cons = 0;
            prefix = origFileName.Substring(0, 4);
            if (prefix == "JOUR" || prefix == "TRAN" || prefix == "TPPJ" || prefix == "BILL" || prefix == "CRED" ||
                prefix == "LOWB" || prefix == "PAST" || prefix == "PAID" || prefix == "COHO" || prefix == "SEPA" ||
                prefix == "BADA" || prefix == "FISC" || prefix == "NEWL" || prefix == "ENRO" || prefix == "BATC" ||
                prefix == "DELE" || prefix == "INVO" || prefix == "EPSR" || prefix == "COLL" || prefix == "HIST" ||
                prefix == "DISB" || prefix == "EPSR" || prefix == "SLP0")
            {
                strScid = origFileName.Substring(4, 3);
                strCons = origFileName.Substring(7, 3);


                if (prefix != "EPSR" && prefix != "INVO" && prefix != "HIST" && prefix != "SLP0")
                {
                    try
                    {
                        scid = AbbrevToNumber(strScid);
                        cons = Convert.ToInt32(strCons);
                        strScid = scid.ToString();
                        strCons = cons.ToString();
                    }
                    catch
                    {
                        scid = cons = 0;
                        strScid = scid.ToString();
                        strCons = cons.ToString();
                    }
                }
                var thisDay = DateTime.Today;
                strDate = thisDay.Year.ToString();
                strDate = strDate.TrimEnd() + thisDay.Month.ToString("D2");
                strDate = strDate.TrimEnd() + thisDay.Day.ToString("D2");

                // uncomment next line for month-end reports to force the month-ending date
                //strDate = "20160930";

                newFileName = _pdfFile.Replace("%scid%", strScid);
                newFileName = newFileName.Replace("%cons%", strCons);
                newFileName = newFileName.Replace("%date%", strDate);

                _pdfPath = _pdfPath.Replace("%scid%", strScid);
            }
            else
            {
                newFileName = _pdfFile;

                // Does file name already exist?
                if (File.Exists(newFileName))
                    newFileName.Replace(".pdf", "_copy.pdf");
            }

            return (newFileName);
        }

        #endregion

        #region AbbrevToNumber

        // change 3-char school abbreviation to school number
        // --------------------------------------------------
        private static int AbbrevToNumber(string strScid)
        {
            var scid = 0;

            switch(strScid)
            {
                case "BOY": scid = 1006; break;
                case "ACI": scid = 1007; break;
                case "ACC": scid = 1008; break;
                case "AAI": scid = 1087; break;
                case "ALC": scid = 1255; break;
                case "LAC": scid = 1321; break;
                case "MCC": scid = 1533; break;
                case "CEN": scid = 1908; break;
                case "COL": scid = 1911; break;
                case "KWU": scid = 1929; break;
                case "TAB": scid = 1946; break;
                case "WCC": scid = 2447; break;
                case "YOR": scid = 2567; break;
                case "USA": scid = 3167; break;
                case "CAL": scid = 3176; break;
                case "EAS": scid = 3415; break;
                case "MHB": scid = 5041; break;
                case "BMC": scid = 6755; break;
                case "PTI": scid = 7437; break;
                case "NAT": scid = 12020; break;
                case "WNH": scid = 12035; break;
                case "STA": scid = 12040; break;
                case "WES": scid = 12045; break;
                case "STM": scid = 12050; break;
                case "EPI": scid = 12060; break;
                case "HOT": scid = 12065; break;
                case "KEN": scid = 12070; break;
                case "SUF": scid = 12075; break;
                case "GOU": scid = 12080; break;
                case "WIL": scid = 12085; break;
                case "CHO": scid = 12090; break;
                case "DFA": scid = 12095; break;
                case "TAF": scid = 12100; break;
                case "GUN": scid = 12105; break;
                case "NMH": scid = 12110; break;
                case "CAN": scid = 12115; break;
                case "WAD": scid = 20100; break;
                case "CHD": scid = 20110; break;
                default   : scid = 0;     break;
            }

            return (scid);
        }

        #endregion

        #endregion
    }
}
