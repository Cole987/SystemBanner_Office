//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using SystemBannerOffice.Markings;

namespace SystemBannerOffice
{
    //the text of the document split in portions: those that start with a portion mark, and how many have none
    public class PortionScan
    {
        public List<Marking> Marks = new List<Marking>();
        public List<string> Errors = new List<string>();
        public int Unmarked;
    }

    //what the add-in does in a document, the same way in Word, Excel and PowerPoint; every call is late bound
    public abstract class DocumentHost
    {
        //document properties, the classification metadata other tools (mail guards, file servers) can read
        public const string ClassificationProperty = "Classification";
        public const string BannerProperty = "SystemBanner.Banner";
        public const string ClassifiedByProperty = "SystemBanner.ClassifiedBy";
        public const string AgencyProperty = "SystemBanner.Agency";
        public const string ReasonProperty = "SystemBanner.Reason";
        public const string DerivedFromProperty = "SystemBanner.DerivedFrom";
        public const string DeclassifyOnProperty = "SystemBanner.DeclassifyOn";
        public const string DateOfOriginProperty = "SystemBanner.DateOfOrigin";
        public const string OriginalProperty = "SystemBanner.Original";
        public const string SourcesProperty = "SystemBanner.Sources";

        //a portion mark at the start of a paragraph: "(S//NF) text"
        private static readonly Regex PortionStart = new Regex(@"^\s*\(([^()\r\n]{1,300})\)");

        public object Application;
        public object Document;

        //the host of the active document, null when no document is open
        public static DocumentHost Active(object application)
        {
            DocumentHost host = Create(application);
            if (host == null)
            {
                return null;
            }
            host.Document = Late.TryGet(application, host is WordHost ? "ActiveDocument" : host is ExcelHost ? "ActiveWorkbook" : "ActivePresentation");
            return host.Document == null ? null : host;
        }

        //the host of a given document, from the events that pass it
        public static DocumentHost For(object application, object document)
        {
            DocumentHost host = Create(application);
            if (host != null)
            {
                host.Document = document;
            }
            return host;
        }

        private static DocumentHost Create(object application)
        {
            string name = Convert.ToString(Late.Get(application, "Name"));
            DocumentHost host;
            if (name.IndexOf("Word") >= 0) host = new WordHost();
            else if (name.IndexOf("Excel") >= 0) host = new ExcelHost();
            else if (name.IndexOf("PowerPoint") >= 0) host = new PowerPointHost();
            else return null;
            host.Application = application;
            return host;
        }

        //major version of Office: 9 for Office 2000, 12 for 2007, 14 for 2010
        public int Version
        {
            get
            {
                string version = Convert.ToString(Late.Get(Application, "Version"));
                int dot = version.IndexOf('.');
                int major;
                return int.TryParse(dot > 0 ? version.Substring(0, dot) : version, out major) ? major : 0;
            }
        }

        public string GetProperty(string name)
        {
            object properties = Late.Get(Document, "CustomDocumentProperties");
            try
            {
                object property = Late.Item(properties, name);
                return Convert.ToString(Late.Get(property, "Value"));
            }
            catch (Exception)
            {
                return null; //no such property
            }
        }

        public void SetProperty(string name, string value)
        {
            object properties = Late.Get(Document, "CustomDocumentProperties");
            try
            {
                object property = Late.Item(properties, name);
                if (string.IsNullOrEmpty(value))
                {
                    Late.Call(property, "Delete");
                }
                else
                {
                    Late.Set(property, "Value", value);
                }
            }
            catch (Exception)
            {
                if (!string.IsNullOrEmpty(value))
                {
                    //msoPropertyTypeString = 4, not linked to content
                    Late.Call(properties, "Add", name, false, 4, value);
                }
            }
        }

        //the banner stored in the document, null when it is not classified yet
        public Marking StoredBanner()
        {
            string text = GetProperty(BannerProperty);
            if (string.IsNullOrEmpty(text))
            {
                return null;
            }
            ParseResult result = MarkingParser.Parse(text);
            return result.Success ? result.Marking : null;
        }

        public AuthorityBlock StoredAuthorityBlock()
        {
            AuthorityBlock block = new AuthorityBlock();
            block.ClassifiedBy = GetProperty(ClassifiedByProperty);
            block.AgencyAndOffice = GetProperty(AgencyProperty);
            block.Reason = GetProperty(ReasonProperty);
            block.DerivedFrom = GetProperty(DerivedFromProperty);
            block.DeclassifyOn = GetProperty(DeclassifyOnProperty);
            block.Original = GetProperty(OriginalProperty) == "1";
            DateTime origin;
            if (DateTime.TryParseExact(GetProperty(DateOfOriginProperty) ?? "", "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out origin))
            {
                block.DateOfOrigin = origin;
            }
            string sources = GetProperty(SourcesProperty);
            if (!string.IsNullOrEmpty(sources))
            {
                block.Sources.AddRange(sources.Split('\n'));
            }
            return block;
        }

        public void StoreAuthorityBlock(AuthorityBlock block)
        {
            SetProperty(ClassifiedByProperty, block.ClassifiedBy);
            SetProperty(AgencyProperty, block.AgencyAndOffice);
            SetProperty(ReasonProperty, block.Original ? block.Reason : null);
            SetProperty(DerivedFromProperty, block.Original ? null : block.DerivedFrom);
            SetProperty(DeclassifyOnProperty, block.DeclassifyOn);
            SetProperty(OriginalProperty, block.Original ? "1" : "0");
            SetProperty(DateOfOriginProperty, block.DateOfOrigin.ToString("yyyyMMdd"));
            SetProperty(SourcesProperty, block.Sources.Count > 0 ? string.Join("\n", block.Sources.ToArray()) : null);
        }

        //writes the banner line at the top and bottom of every page and stores it in the document properties
        public void Mark(Marking banner)
        {
            string text = MarkingFormatter.Banner(banner);
            SetProperty(BannerProperty, text);
            SetProperty(ClassificationProperty, MarkingFormatter.LevelTitle(banner.Level == Level.Restricted ? Level.Confidential : banner.Level));
            int[] colors = Presets.ColorsFor(banner);
            WriteBanner(text, Late.Rgb(colors[0], colors[1], colors[2]), Late.Rgb(colors[3], colors[4], colors[5]));
        }

        //a national banner of SystemBanner, on computers set to the French, German, British, NATO or UN systems
        public void Mark(Preset preset)
        {
            SetProperty(BannerProperty, null);
            SetProperty(ClassificationProperty, preset.Text);
            WriteBanner(preset.Text, Late.Rgb(preset.Colors[0], preset.Colors[1], preset.Colors[2]), Late.Rgb(preset.Colors[3], preset.Colors[4], preset.Colors[5]));
        }

        public string StoredClassification()
        {
            return GetProperty(ClassificationProperty);
        }

        protected abstract void WriteBanner(string text, int background, int foreground);

        public virtual bool SupportsPortions
        {
            get { return false; }
        }

        //the paragraphs of the document, for the roll-up
        protected virtual IEnumerable<string> Paragraphs()
        {
            return new string[0];
        }

        public PortionScan ScanPortions()
        {
            PortionScan scan = new PortionScan();
            foreach (string paragraph in Paragraphs())
            {
                string text = paragraph.Trim('\a', '\v', '\f', ' ', '\t', '\r', '\n', '\u0007');
                if (text.Length == 0)
                {
                    continue;
                }
                Match match = PortionStart.Match(text);
                if (!match.Success)
                {
                    scan.Unmarked++;
                    continue;
                }
                ParseResult result = MarkingParser.Parse(match.Groups[1].Value);
                if (result.Success)
                {
                    scan.Marks.Add(result.Marking);
                }
                else if (LooksLikeMarking(match.Groups[1].Value))
                {
                    scan.Errors.Add("(" + match.Groups[1].Value + "): " + string.Join(" ", result.Errors.ToArray()));
                }
                else
                {
                    scan.Unmarked++; //a parenthesis that is not a portion mark, "(see annex)"
                }
            }
            return scan;
        }

        private static bool LooksLikeMarking(string text)
        {
            return text.StartsWith("//") || Regex.IsMatch(text, @"^(TS|S|C|U|TOP SECRET|SECRET|CONFIDENTIAL|UNCLASSIFIED)(//|$)");
        }

        //puts the portion mark at the start of the paragraph of the cursor, replacing the one already there
        public virtual void InsertPortionMark(string mark)
        {
            throw new NotSupportedException();
        }

        public virtual void WriteAuthorityBlock(string[] lines)
        {
            throw new NotSupportedException();
        }

        //replaces a portion mark at the start of a paragraph's text, or adds one
        protected static int ExistingMarkLength(string paragraphText)
        {
            Match match = PortionStart.Match(paragraphText);
            if (!match.Success || !MarkingParser.Parse(match.Groups[1].Value).Success)
            {
                return 0;
            }
            int length = match.Length;
            if (length < paragraphText.Length && paragraphText[length] == ' ')
            {
                length++;
            }
            return length;
        }
    }

    public class WordHost : DocumentHost
    {
        private const int HeaderFooterPrimary = 1, HeaderFooterFirstPage = 2, HeaderFooterEvenPages = 3;
        private const string AuthorityBookmark = "SystemBannerAuthority";

        public override bool SupportsPortions
        {
            get { return true; }
        }

        protected override void WriteBanner(string text, int background, int foreground)
        {
            object sections = Late.Get(Document, "Sections");
            int count = Late.Count(sections);
            for (int s = 1; s <= count; s++)
            {
                object section = Late.Item(sections, s);
                foreach (int kind in new int[] { HeaderFooterPrimary, HeaderFooterFirstPage, HeaderFooterEvenPages })
                {
                    WriteHeaderFooter(Late.Item(Late.Get(section, "Headers"), kind), "SystemBannerTop" + s + "_" + kind, true, text, background, foreground);
                    WriteHeaderFooter(Late.Item(Late.Get(section, "Footers"), kind), "SystemBannerBottom" + s + "_" + kind, false, text, background, foreground);
                }
            }
        }

        //the banner is its own paragraph, first in the header and last in the footer, kept in a bookmark so the user's header content stays
        private void WriteHeaderFooter(object headerFooter, string bookmark, bool top, string text, int background, int foreground)
        {
            if (Convert.ToBoolean(Late.TryGet(headerFooter, "LinkToPrevious") ?? false))
            {
                return; //the previous section's header already shows it
            }
            object bookmarks = Late.Get(Document, "Bookmarks");
            object range;
            if (Convert.ToBoolean(Late.Call(bookmarks, "Exists", bookmark)))
            {
                range = Late.Get(Late.Item(bookmarks, bookmark), "Range");
            }
            else
            {
                object whole = Late.Get(headerFooter, "Range");
                string existing = Convert.ToString(Late.Get(whole, "Text")).Trim('\r', '\n', ' ');
                if (existing.Length > 0)
                {
                    Late.Call(whole, top ? "InsertParagraphBefore" : "InsertParagraphAfter");
                }
                object paragraphs = Late.Get(Late.Get(headerFooter, "Range"), "Paragraphs");
                object paragraph = top ? Late.Get(paragraphs, "First") : Late.Get(paragraphs, "Last");
                range = Late.Get(paragraph, "Range");
                int start = Convert.ToInt32(Late.Get(range, "Start"));
                int end = Convert.ToInt32(Late.Get(range, "End"));
                string paragraphText = Convert.ToString(Late.Get(range, "Text"));
                if (paragraphText.EndsWith("\r"))
                {
                    end--; //keep the paragraph mark out of the bookmark
                }
                range = Late.Get(Late.Get(headerFooter, "Range"), "Duplicate");
                Late.Call(range, "SetRange", start, end);
            }
            Late.Set(range, "Text", text);
            object font = Late.Get(range, "Font");
            Late.TrySet(font, "Bold", true);
            if (!Late.TrySet(font, "Color", foreground))
            {
                Late.TrySet(font, "ColorIndex", foreground == 0 ? 1 : 8); //Word 2000 has no RGB font color: black or white
            }
            object format = Late.Get(range, "ParagraphFormat");
            Late.TrySet(format, "Alignment", 1); //wdAlignParagraphCenter
            Late.TrySet(Late.Get(format, "Shading"), "BackgroundPatternColor", background);
            Late.Call(bookmarks, "Add", bookmark, range);
        }

        protected override IEnumerable<string> Paragraphs()
        {
            //one string for the whole body is much faster than a COM call per paragraph; table cells end with "\r\a"
            string text = Convert.ToString(Late.Get(Late.Get(Document, "Content"), "Text"));
            return text.Split('\r');
        }

        public override void InsertPortionMark(string mark)
        {
            object selection = Late.Get(Application, "Selection");
            object paragraph = Late.Get(Late.Item(Late.Get(selection, "Paragraphs"), 1), "Range");
            int start = Convert.ToInt32(Late.Get(paragraph, "Start"));
            string text = Convert.ToString(Late.Get(paragraph, "Text"));
            int leading = text.Length - text.TrimStart(' ', '\t').Length;
            int existing = ExistingMarkLength(text.Substring(leading));
            object range = Late.Call(Document, "Range", start + leading, start + leading + existing);
            Late.Set(range, "Text", "(" + mark + ") ");
        }

        public override void WriteAuthorityBlock(string[] lines)
        {
            object bookmarks = Late.Get(Document, "Bookmarks");
            object range;
            if (Convert.ToBoolean(Late.Call(bookmarks, "Exists", AuthorityBookmark)))
            {
                range = Late.Get(Late.Item(bookmarks, AuthorityBookmark), "Range");
            }
            else
            {
                range = Late.Get(Late.Get(Application, "Selection"), "Range");
            }
            Late.Set(range, "Text", string.Join("\v", lines)); //line breaks keep the block in one paragraph
            Late.Call(bookmarks, "Add", AuthorityBookmark, range);
        }
    }

    public class ExcelHost : DocumentHost
    {
        protected override void WriteBanner(string text, int background, int foreground)
        {
            //Excel headers and footers only take text; colors exist from Excel 2007 (&K), an escaped "&" is "&&"
            string escaped = text.Replace("&", "&&");
            string color = "";
            if (Version >= 12)
            {
                int red = background & 0xFF, green = (background >> 8) & 0xFF, blue = (background >> 16) & 0xFF;
                color = "&K" + red.ToString("X2") + green.ToString("X2") + blue.ToString("X2");
            }
            string header = "&\"Arial,Bold\"&12" + color + escaped;
            object sheets = Late.Get(Document, "Worksheets");
            int count = Late.Count(sheets);
            for (int i = 1; i <= count; i++)
            {
                object setup = Late.Get(Late.Item(sheets, i), "PageSetup");
                Late.Set(setup, "CenterHeader", header);
                Late.Set(setup, "CenterFooter", header);
            }
        }
    }

    public class PowerPointHost : DocumentHost
    {
        private const string TopShape = "SystemBannerTop", BottomShape = "SystemBannerBottom";
        private const float BannerHeight = 20;

        public override bool SupportsPortions
        {
            get { return true; }
        }

        protected override void WriteBanner(string text, int background, int foreground)
        {
            List<object> masters = new List<object>();
            masters.Add(Late.Get(Document, "SlideMaster"));
            if (Convert.ToInt32(Late.TryGet(Document, "HasTitleMaster") ?? 0) != 0)
            {
                masters.Add(Late.Get(Document, "TitleMaster")); //PowerPoint 2000 to 2003 title slides use their own master
            }
            object setup = Late.Get(Document, "PageSetup");
            float width = Convert.ToSingle(Late.Get(setup, "SlideWidth"));
            float height = Convert.ToSingle(Late.Get(setup, "SlideHeight"));
            foreach (object master in masters)
            {
                object shapes = Late.Get(master, "Shapes");
                RemoveShape(shapes, TopShape);
                RemoveShape(shapes, BottomShape);
                AddBanner(shapes, TopShape, 0, width, text, background, foreground);
                AddBanner(shapes, BottomShape, height - BannerHeight, width, text, background, foreground);
            }
        }

        private static void RemoveShape(object shapes, string name)
        {
            for (int i = Late.Count(shapes); i >= 1; i--)
            {
                object shape = Late.Item(shapes, i);
                if (Convert.ToString(Late.Get(shape, "Name")) == name)
                {
                    Late.Call(shape, "Delete");
                }
            }
        }

        private static void AddBanner(object shapes, string name, float top, float width, string text, int background, int foreground)
        {
            object shape = Late.Call(shapes, "AddTextbox", 1, 0f, top, width, BannerHeight); //msoTextOrientationHorizontal
            Late.Set(shape, "Name", name);
            object fill = Late.Get(shape, "Fill");
            Late.Set(fill, "Visible", -1);
            Late.Set(Late.Get(fill, "ForeColor"), "RGB", background);
            object frame = Late.Get(shape, "TextFrame");
            Late.TrySet(frame, "WordWrap", -1);
            Late.TrySet(frame, "AutoSize", 0);
            object range = Late.Get(frame, "TextRange");
            Late.Set(range, "Text", text);
            object font = Late.Get(range, "Font");
            Late.Set(font, "Bold", -1);
            Late.Set(font, "Size", 12);
            Late.Set(Late.Get(font, "Color"), "RGB", foreground);
            Late.Set(Late.Get(range, "ParagraphFormat"), "Alignment", 2); //ppAlignCenter
        }

        protected override IEnumerable<string> Paragraphs()
        {
            List<string> paragraphs = new List<string>();
            object slides = Late.Get(Document, "Slides");
            int count = Late.Count(slides);
            for (int i = 1; i <= count; i++)
            {
                object shapes = Late.Get(Late.Item(slides, i), "Shapes");
                int shapeCount = Late.Count(shapes);
                for (int j = 1; j <= shapeCount; j++)
                {
                    object shape = Late.Item(shapes, j);
                    if (Convert.ToInt32(Late.TryGet(shape, "HasTextFrame") ?? 0) == 0)
                    {
                        continue;
                    }
                    object frame = Late.Get(shape, "TextFrame");
                    if (Convert.ToInt32(Late.TryGet(frame, "HasText") ?? 0) == 0)
                    {
                        continue;
                    }
                    paragraphs.AddRange(Convert.ToString(Late.Get(Late.Get(frame, "TextRange"), "Text")).Split('\r'));
                }
            }
            return paragraphs;
        }

        public override void InsertPortionMark(string mark)
        {
            object selection = Late.Get(Late.Get(Application, "ActiveWindow"), "Selection");
            if (Convert.ToInt32(Late.Get(selection, "Type")) != 3) //ppSelectionText
            {
                throw new InvalidOperationException("Put the cursor in the text that the portion mark applies to.");
            }
            object paragraph = Late.Call(Late.Get(selection, "TextRange"), "Paragraphs", 1, 1);
            string text = Convert.ToString(Late.Get(paragraph, "Text"));
            int existing = ExistingMarkLength(text);
            if (existing > 0)
            {
                Late.Call(Late.Call(paragraph, "Characters", 1, existing), "Delete");
            }
            Late.Call(paragraph, "InsertBefore", "(" + mark + ") ");
        }
    }
}
