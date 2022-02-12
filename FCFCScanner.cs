using AS400Report;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Report
{
    public class FcfcScanner
    {
        #region Class_Properties
        private string _page;      // page to print
        private int _curPageOffset;      // current byte count (offset) for original _text. Object user must set once initially via constructor
        private int _pageLen;      // length of _page after NextPage method
        private bool _morePages;   // true if _curPageOffset < _textLen
        private int _numLines;     // number of lines in the current page
        private List<int> _wordsPerEachPage;
        private readonly List<string> _allTextToPrint;
        // set the text to print
        public string Text { get; set; }
        // get the text length
        public int TextLength { get; set; }
        // get the formatted page to print
        public string Page
        {
            get { return _page; }
        }
        public int PageLength
        {
            get { return _pageLen; }
        }
        // get boolean for whether more pages are available
        public bool MorePages
        {
            get { return _morePages; }
        }
        // get boolean for whether more pages are available
        public int NumLines
        {
            get { return _numLines; }
        }
        //Holds an array of ints: Words per page.
        public List<int> WordsPerEachPage
        {
            get { return _wordsPerEachPage; }
        }
        // gets the current page offset.
        public int CurrentPageOffset
        {
            get { return _curPageOffset; }
        }
        #endregion

        #region Class_Contsructors
        public FcfcScanner(string fcfcText)
        {
            Text = fcfcText;
            TextLength = Text.Length;
            _page = string.Empty;
            _curPageOffset = 0;
            _pageLen = 0;
            _morePages = (TextLength > 0);
            _numLines = 0;
        }

        public FcfcScanner(List<string> fcfcText)
        {

            _page = string.Empty;
            _curPageOffset = 0;
            _pageLen = 0;
            _morePages = (TextLength > 0);
            _numLines = 0;

            _allTextToPrint = new List<string>(PrintQueue.DirectoryInstance.FileEntries.Length);
            for (var i = 1; i < PrintQueue.DirectoryInstance.FileEntries.Length; i++)
            {
                _allTextToPrint.Add(fcfcText[i]);
            }
            //for (int i = 0; i < PrintQueue.DirectoryInstance.FileEntries.Length; i++)
            //{
            //    this._allTextToPrint[i] = fcfcText[i];
            //}


        }
        #endregion

        #region Methods

        #region HasMorePages()

        public bool HasMorePages()
        {
            // Position to start of current page
            if (_curPageOffset >= TextLength)
            {
                _morePages = false;
                _page = "";
                return false;
            }
            else
                return true;
        }

        #endregion

        #region NextPage()

        // ------------------------------------------------------------------------------
        // Scan for form feed/spacing codes to determine the next page to print
        // Returns 0 if successful
        // ------------------------------------------------------------------------------
        public int NextPage()
        {
            var sb = new StringBuilder();
            string line;
            const string STR_CR_LF = "\r\n";
            var pageLen = 0;
            int lineLen;
            var fromPos = 0;
            var bDone = false;
            var firstTime = true;
            bool bSkip;
            char[] firstChar;

            _morePages = HasMorePages();
            // Position to start of current page
            if (!_morePages)
            {
                _page = "";
                return (-1);
            }
            line = Text.Substring(_curPageOffset);

            // evaluate lines until either EOF of a form-feed is encountered
            char[] searchChars = { '\r', '\n' };
            while (!bDone)
            {
                try
                {
                    // firstChar on a line is the FCFC spacing character
                    firstChar = line.ToCharArray(fromPos, 3);
                }
                catch
                {
                    break;
                }
                fromPos++;
                bSkip = false;
                _curPageOffset++;
                switch (firstChar[0])
                {

                    case '0':
                        {
                            if (firstChar[1] == 0x0D || firstChar[1] == 0x0A)    // If CR or LF
                            {
                                bSkip = true;
                                _curPageOffset++;
                                fromPos++;
                            }
                            if (firstChar[2] == 0x0A && bSkip)                   // If CR AND LF
                            {
                                _curPageOffset++;
                                fromPos++;
                            }
                            sb.Append(STR_CR_LF);
                            pageLen += 2;
                            break;
                        }
                    case '-':
                        {
                            if (firstChar[1] == 0x0D || firstChar[1] == 0x0A)    // If CR or LF
                            {
                                bSkip = true;
                                _curPageOffset++;
                                fromPos++;
                            }
                            if (firstChar[2] == 0x0A && bSkip)                   // If CR AND LF
                            {
                                _curPageOffset++;
                                fromPos++;
                            }
                            sb.Append(STR_CR_LF);
                            sb.Append(STR_CR_LF);
                            pageLen += 4;
                            break;
                        }
                    case '1':
                        {
                            if (!firstTime)
                                bDone = true;
                            else
                            {
                                if (firstChar[1] == 0x0D || firstChar[1] == 0x0A)    // If CR or LF
                                {
                                    bSkip = true;
                                    _curPageOffset++;
                                    fromPos++;
                                }
                                if (firstChar[2] == 0x0A && bSkip)                   // If CR AND LF
                                {
                                    _curPageOffset++;
                                    fromPos++;
                                }
                            }
                            break;
                        }
                    case '\0':
                        {
                            bDone = true;
                            break;
                        }
                }
                firstTime = false;
                if (!bSkip)
                {
                    lineLen = line.Substring(fromPos).IndexOfAny(searchChars);
                    //lineLen2 = line.Substring(fromPos).IndexOf(searchStr);
                    if (lineLen < 0 || bDone)
                    {
                        pageLen = sb.Length;
                        if (lineLen < 0)
                            _curPageOffset += line.Length;
                        break;
                    }
                    char[] eorChars = { '\0', '\0' };
                    line.CopyTo((fromPos + lineLen), eorChars, 0, 2);
                    if (eorChars[0] == '\r' && eorChars[1] == '\n')
                        lineLen += 2;
                    else
                        lineLen += 1;

                    _curPageOffset += lineLen;
                    pageLen += lineLen;
                    sb.Append(line.Substring(fromPos, lineLen));
                    fromPos += lineLen;
                }
            }

            _page = sb.ToString();
            pageLen = (pageLen > 0) ? pageLen + 2 : pageLen;        // not sure why I added this?
            _pageLen = pageLen;
            _wordsPerEachPage.Add(_pageLen);
            _morePages = (_curPageOffset < TextLength - 1);
            return (0);
        }

        #endregion

        //Currently Unused
        #region NextPage(arg)

        //------------------------------------------------------------------------------
        //Scan for form feed/spacing codes to determine the next page to print
        //Returns 0 if successful
        //------------------------------------------------------------------------------
        public List<string> NextPage(string textToPrint)
        {
            var pageLines = new string[110];
            string line;
            const string STR_CR_LF = "\r\n";
            var pageLen = 0;
            int lineLen;
            var fromPos = 0;
            var lineNbr = 0;
            var bDone = false;
            var firstTime = true;
            bool bSkip;
            char[] firstChar;

            if (_curPageOffset > 0)
            {

                pageLines.ToList().Clear();
                pageLines.ToArray();
            }

            // Position to start of current page
            if (_curPageOffset > TextLength)
            {
                _morePages = false;
                _page = "";
                return pageLines.Where(x => !string.IsNullOrEmpty(x)).ToList();
            }
            line = Text.Substring(_curPageOffset);

            // evaluate lines until either EOF of a form-feed is encountered
            char[] searchChars = { '\r', '\n' };

            while (!bDone)
            {
                try
                {
                    // firstChar on a line is the FCFC spacing character
                    firstChar = line.ToCharArray(fromPos, 3);
                }
                catch
                {
                    break;
                }
                fromPos++;
                bSkip = false;
                _curPageOffset++;
                switch (firstChar[0])
                {

                    case '0':
                        {
                            if (firstChar[1] == 0x0D || firstChar[1] == 0x0A)    // If CR or LF
                            {
                                bSkip = true;
                                _curPageOffset++;
                                fromPos++;
                            }
                            if (firstChar[2] == 0x0A && bSkip)                   // If CR AND LF
                            {
                                _curPageOffset++;
                                fromPos++;
                            }
                            pageLines[lineNbr++] = STR_CR_LF;
                            pageLen += 2;
                            break;
                        }
                    case '-':
                        {
                            if (firstChar[1] == 0x0D || firstChar[1] == 0x0A)    // If CR or LF
                            {
                                bSkip = true;
                                _curPageOffset++;
                                fromPos++;
                            }
                            if (firstChar[2] == 0x0A && bSkip)                   // If CR AND LF
                            {
                                _curPageOffset++;
                                fromPos++;
                            }
                            pageLines[lineNbr++] = STR_CR_LF;
                            pageLines[lineNbr++] = STR_CR_LF;
                            pageLen += 4;
                            break;
                        }
                    case '1':
                        {
                            if (!firstTime)
                            {
                                bDone = true;
                                fromPos--;       // needs to start with "1\r\n" on next call rather than "\r\n"
                                _curPageOffset--;
                                bSkip = true;
                            }
                            else
                            {
                                if (firstChar[1] == 0x0D || firstChar[1] == 0x0A)    // If CR or LF
                                {
                                    bSkip = true;
                                    _curPageOffset++;
                                    fromPos++;
                                }
                                if (firstChar[2] == 0x0A && bSkip)                   // If CR AND LF
                                {
                                    _curPageOffset++;
                                    fromPos++;
                                }
                            }
                            break;
                        }
                    case '\0':
                        {
                            bDone = true;
                            break;
                        }
                }
                firstTime = false;
                if (!bSkip)
                {
                    lineLen = line.Substring(fromPos).IndexOfAny(searchChars);
                    //lineLen = line.Substring(fromPos).IndexOf(searchStr);
                    if (lineLen < 0 || bDone)
                    {
                        //pageLen = sb.Length;
                        if (lineLen < 0)
                            _curPageOffset += line.Length;
                        break;
                    }
                    char[] eorChars = { '\0', '\0' };
                    line.CopyTo((fromPos + lineLen), eorChars, 0, 2);
                    if (eorChars[0] == '\r' && eorChars[1] == '\n')
                        lineLen += 2;
                    else
                        lineLen += 1;
                    _curPageOffset += lineLen;
                    pageLen += lineLen;
                    pageLines[lineNbr++] = line.Substring(fromPos, lineLen - 2);
                    fromPos += lineLen;
                }
            }

            // this._page = sb.ToString();
            //pageLen = (pageLen > 0) ? pageLen + 2 : pageLen;        // not sure why I added this?
            _pageLen = pageLen;
            _morePages = (_curPageOffset < TextLength - 1);
            _numLines = lineNbr;

            return pageLines.ToArray().Where(x => !string.IsNullOrEmpty(x)).ToList();
        }

        #endregion

        #region Sync From Print Queue

        public void SyncFromPrintQueue()
        {
            Text = PrintQueue.DirectoryInstance.Text;
            TextLength = PrintQueue.DirectoryInstance.TextLen;
            _page = PrintQueue.DirectoryInstance.Page;
            _curPageOffset = PrintQueue.DirectoryInstance.CurPageOffset;
            _pageLen = PrintQueue.DirectoryInstance.PageLen;
            _morePages = PrintQueue.DirectoryInstance.MorePages;
            _numLines = PrintQueue.DirectoryInstance.NumLines;
            _wordsPerEachPage = PrintQueue.DirectoryInstance.WordsPerEachPage;
        }

        #endregion

        #endregion
    }
}

