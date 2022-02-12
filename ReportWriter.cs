#define __WINDOWS__
#define __MONO__

using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Text;

namespace Report
{
    public class ReportWriter : PrintDocument
    {
        #region Property_Variables
        private readonly string _fmtopt;
        private string _text;      // text to print
        static FcfcScanner _page;   // object to handle FCFC formatting
        #endregion

        #region Class_Properties
        // set the text to print
        public string TextToPrint
        {
            get { return _text; }
            set
            {
                _text = value;
                _page = new FcfcScanner(_text);
            }
        }

        // override the font
        public Font PrinterFont { get; set; }

        // set landscape true or false
        public bool LandScape { get; set; }

        // Margins
        public int LeftMargin { get; set; }
        public int RightMargin { get; set; }
        public int TopMargin { get; set; }
        public int BottomMargin { get; set; }
        public string PrinterName { get; set; }
        #endregion

        #region Class_Contsructors
        public ReportWriter()
        {
            _fmtopt = "NONE";
            _text = string.Empty;
            PrinterName = string.Empty;
            LeftMargin = RightMargin = TopMargin = BottomMargin = 0;
        }

        public ReportWriter(string fmtOption)
        {
            _fmtopt = fmtOption;
            _text = string.Empty;
            PrinterName = string.Empty;
            LeftMargin = RightMargin = TopMargin = BottomMargin = 0;
        }

        public ReportWriter(string fmtOption, string str)
        {
            _fmtopt = fmtOption;
            _text = str;
            _page = new FcfcScanner(_text);
            PrinterName = string.Empty;
            LeftMargin = RightMargin = TopMargin = BottomMargin = 0;
        }
        #endregion

        #region onbeginPrint
#if !(__MONO__)
		// Override the default onbeginPrint method of the PrintObject object
		protected override void OnBeginPrint(System.Drawing.Printing.PrintEventArgs e)
		{
			// Run base code
			base.OnBeginPrint(e);

			// check to see if the user provided a font
			// if they didn't, then default to Times New Roman
			if (_font == null) {
				// create the font we need
				_font = new Font("Times New Roman", 9);
			}
			base.DefaultPageSettings.Landscape = _landscape;

			// override the printer name if an alternate printer name has been provided
			if (_printerName != string.Empty)
				base.PrinterSettings.PrinterName = _printerName;
		}
#endif
        #endregion

        #region OnPrintPage
        // Override the default OnPrintPage method of the PrintDocument. This provides the
        // print logic for our document
        protected override void OnPrintPage(PrintPageEventArgs e)
        {
            // Run base code
            base.OnPrintPage(e);

            // declare local variables
            int printHeight;
            int printWidth;
            int leftMargin;
            int rightMargin;
            int topMargin, bottomMargin;
            int lines;
            int chars;
            int formLen;
            string strToPrint;

            // Set print area size and margins
            {
                leftMargin = (LeftMargin > 0) ? LeftMargin : DefaultPageSettings.Margins.Left;  // X
                rightMargin = (RightMargin > 0) ? RightMargin : DefaultPageSettings.Margins.Right;  // Y
                topMargin = (TopMargin > 0) ? TopMargin : DefaultPageSettings.Margins.Top;
                bottomMargin = (BottomMargin > 0) ? BottomMargin : DefaultPageSettings.Margins.Bottom;

                printHeight = DefaultPageSettings.PaperSize.Height - topMargin - bottomMargin;
                printWidth = DefaultPageSettings.PaperSize.Width - leftMargin - rightMargin;
            }

            // check if user selected to print in landscape mode. If so, then swap height/width parameters
            if (DefaultPageSettings.Landscape)
            {
                int tmp;
                tmp = printHeight;
                printHeight = printWidth;
                printWidth = tmp;
            }

            // Now we need to determine the total number of lines
            var numLines = (int)printHeight / PrinterFont.Height;

            // create a rectangle printing area for our document
            var printArea = new RectangleF(leftMargin, rightMargin, printWidth, printHeight);

            // use the StringFormat class for the text layout of our document
            var format = new StringFormat(StringFormatFlags.LineLimit);

            // Scan for form feed character
            if (_fmtopt == "FCFC")
            {
                _page.NextPage();
                strToPrint = _page.Page;
                formLen = _page.PageLength;
            }
            else
            {
                formLen = 0;
                //strToPrint = _text.Substring(curChar, formLen);
                strToPrint = _text.Substring(0, formLen);
            }

            // fit as many characters as we can into the print area
            if (formLen > 0)
                e.Graphics.MeasureString(strToPrint, PrinterFont, new SizeF(printWidth, printHeight), format, out chars, out lines);
            else
                e.Graphics.MeasureString(strToPrint, PrinterFont, new SizeF(printWidth, printHeight), format, out chars, out lines);

            // Print the page
            if (formLen > 0)
                e.Graphics.DrawString(strToPrint, PrinterFont, Brushes.Black, printArea, format);
            else
                e.Graphics.DrawString(strToPrint, PrinterFont, Brushes.Black, printArea, format);

            // Determine if there is more text to print. If so, tell the printer there is more coming
            e.HasMorePages = _page.MorePages;
        }
        #endregion

        #region RemoveZeros
        // scan for form feed
#pragma warning disable CC0068 // Unused Method
        private int ScanFormFeed(string line, ref string prtStr, ref int offset)
        {
            var sb = new StringBuilder();
            const string STR_CR_LF = "\r\n";
            var pageLen = 0;
            int lineLen;
            var fromPos = 0;
            var bDone = false;
            var firstTime = true;
            char[] firstChar;

            byte[] searchChars = { 0x0D, 0x0A };
            var searchStr = Encoding.Default.GetString(searchChars);

            while (!bDone)
            {
                try
                {
                    firstChar = line.ToCharArray(fromPos, 1);
                }
                catch
                {
                    break;
                }
                fromPos++;
                offset++;
                switch (firstChar[0])
                {

                    case '0':
                        {
                            sb.Append(STR_CR_LF);
                            pageLen += 2;
                            break;
                        }
                    case '-':
                        {
                            sb.Append(STR_CR_LF);
                            sb.Append(STR_CR_LF);
                            pageLen += 4;
                            break;
                        }
                    case '1':
                        {
                            if (!firstTime)
                                bDone = true;
                            break;
                        }
                    case '\0':
                        {
                            bDone = true;
                            break;
                        }
                }
                firstTime = false;
                lineLen = line.Substring(fromPos).IndexOf(searchStr);
                if (lineLen < 0 || bDone)
                {
                    pageLen = sb.Length;
                    if (lineLen < 0)
                        offset += line.Length;
                    break;
                }
                lineLen += 2;
                offset += lineLen;
                pageLen += lineLen;
                sb.Append(line.Substring(fromPos, lineLen));
                fromPos += lineLen;
            }

            prtStr = sb.ToString();
            return (pageLen);
        }
#pragma warning restore CC0068 // Unused Method

        // Replace any zeros in the size to a 1.  Zeros will mess up printing area
        public static int RemoveZeros(int value)
        {
            // check the value passed into the function. If the value is a zero (0), then return a 1.
            // otherwise return the value passed in
            switch (value)
            {
                case 0:
                    //return 1;
                    return 0;
                default:
                    return value;
            }
        }
        #endregion

        #region Miscellaneous
        private void PrintPage2(object sender, PrintPageEventArgs ev)
        {
            float linesPerPage = 0;
            float yPos = 0;
            var count = 0;
            var skipBefore = 0;
            var spaceAfter = 0;
            float leftMargin = ev.MarginBounds.Left;
            float topMargin = ev.MarginBounds.Top;
            string line = null;
            const string BLANK_LINE = "    ";
            char leadChar;

            // calculate the lines per page
            linesPerPage = ev.MarginBounds.Height / PrinterFont.GetHeight(ev.Graphics);
            using (var sr = new StreamReader("/home/tracym/workspace/Edsi/PrintQueue/EPSR14378.txt"))
            {
                // Print each line of the file
                while (count < linesPerPage &&
                       ((line = sr.ReadLine()) != null))
                {
                    switch (_fmtopt)
                    {
                        case "PRTCTL":
                        {
                            skipBefore = GetSkipBefore(line);
                            if (skipBefore > 0)
                                skipBefore--;
                            spaceAfter = GetSpaceAfter(line);
                            if (spaceAfter > 0)
                                spaceAfter--;
                            break;
                        }
                        case "FCFC":
                            leadChar = line[0];
                            switch (leadChar)
                            {
                                case '1':
                                {
                                    ev.HasMorePages = true;
                                    break;
                                }
                            }

                            break;
                    }
                }
                // skip before to leave appropriate number of blank lines
                while (skipBefore > 0 && count < skipBefore && count < linesPerPage)
                {
                    // print blank line
                    yPos = topMargin + (count * PrinterFont.GetHeight(ev.Graphics));
                    using (var stringFormat = new StringFormat())
                    {
                        ev.Graphics.DrawString(BLANK_LINE, PrinterFont, Brushes.Black, leftMargin, yPos, stringFormat);
                    }
                    count++;
                }

                // space after to leave appropriate number of blank lines
                while (spaceAfter-- > 0 && count < linesPerPage)
                {
                    // print blank line
                    yPos = topMargin + (count * PrinterFont.GetHeight(ev.Graphics));
                    using (var stringFormat = new StringFormat())
                    {
                        ev.Graphics.DrawString(BLANK_LINE, PrinterFont, Brushes.Black, leftMargin, yPos, stringFormat);
                    }
                    count++;
                }

                // print actual text line
                yPos = topMargin + (count * PrinterFont.GetHeight(ev.Graphics));
                if (_fmtopt == "PRTCTL")
                    using (var stringFormat = new StringFormat())
                    {
                        ev.Graphics.DrawString(line.Substring(4), PrinterFont, Brushes.Black, leftMargin, yPos, stringFormat);
                    }
                else
                    using (var stringFormat = new StringFormat())
                    {
                        ev.Graphics.DrawString(line.Substring(1), PrinterFont, Brushes.Black, leftMargin, yPos, stringFormat);
                    }

                count++;
            }


            // if more lines exist, print another page.
            if (line != null)
                ev.HasMorePages = true;
            else
                ev.HasMorePages = false;
        }

        private int GetSkipBefore(string line)
        {
            int lineNbr;

            try
            {
                lineNbr = Convert.ToInt32(line.Substring(0, 3));
            }
            catch
            {
                lineNbr = 0;
            }
            return (lineNbr);
        }

        private static int GetSpaceAfter(string line)
        {
            int lineNbr;

            try
            {
                lineNbr = Convert.ToInt32(line.Substring(3, 1));
            }
            catch
            {
                lineNbr = 0;
            }
            return (lineNbr);
        }
        #endregion
    }
}

