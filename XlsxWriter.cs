using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace SsmsExcelExport
{
    /// <summary>Minimal streaming .xlsx writer: one worksheet per grid, bold header row, frozen header, auto filter.</summary>
    sealed class XlsxWriter : IDisposable
    {
        readonly FileStream _file;
        readonly ZipArchive _zip;
        readonly List<string> _sheetNames = new List<string>();
        bool _closed;

        public XlsxWriter(string path)
        {
            _file = new FileStream(path, FileMode.Create, FileAccess.Write);
            _zip = new ZipArchive(_file, ZipArchiveMode.Create);
        }

        public Sheet AddSheet(string desiredName, IList<string> headers)
        {
            string name = UniqueName(desiredName);
            _sheetNames.Add(name);
            var entry = _zip.CreateEntry("xl/worksheets/sheet" + _sheetNames.Count + ".xml", CompressionLevel.Optimal);
            return new Sheet(entry.Open(), headers);
        }

        static readonly Regex BadChars = new Regex(@"[\\/?*\[\]:]", RegexOptions.Compiled);

        string UniqueName(string n)
        {
            n = BadChars.Replace(n ?? "Sheet", "_").Trim('\'');
            if (n.Length == 0) n = "Sheet";
            if (n.Length > 31) n = n.Substring(0, 31);
            string candidate = n; int i = 2;
            while (_sheetNames.Any(s => string.Equals(s, candidate, StringComparison.OrdinalIgnoreCase)))
            {
                string suffix = " (" + i++ + ")";
                candidate = n.Substring(0, Math.Min(n.Length, 31 - suffix.Length)) + suffix;
            }
            return candidate;
        }

        public void Close()
        {
            if (_closed) return;
            _closed = true;

            var ct = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            for (int i = 1; i <= _sheetNames.Count; i++)
                ct.Append("<Override PartName=\"/xl/worksheets/sheet" + i + ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            ct.Append("</Types>");
            Put("[Content_Types].xml", ct.ToString());

            Put("_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");

            var wb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>");
            var rels = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (int i = 0; i < _sheetNames.Count; i++)
            {
                wb.Append("<sheet name=\"" + Esc(_sheetNames[i]) + "\" sheetId=\"" + (i + 1) + "\" r:id=\"rId" + (i + 1) + "\"/>");
                rels.Append("<Relationship Id=\"rId" + (i + 1) + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet" + (i + 1) + ".xml\"/>");
            }
            wb.Append("</sheets></workbook>");
            rels.Append("<Relationship Id=\"rIdStyles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>");
            Put("xl/workbook.xml", wb.ToString());
            Put("xl/_rels/workbook.xml.rels", rels.ToString());

            Put("xl/styles.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font><font><b/><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/></cellXfs></styleSheet>");
            _zip.Dispose();
            _file.Dispose();
        }

        void Put(string name, string content)
        {
            using (var w = new StreamWriter(_zip.CreateEntry(name, CompressionLevel.Optimal).Open(), new UTF8Encoding(false)))
                w.Write(content);
        }

        public void Dispose()
        {
            if (!_closed) { try { Close(); } catch { } }
        }

        internal static string Esc(string s)
        {
            var sb = new StringBuilder(s.Length + 8);
            foreach (char c in s)
            {
                if (c == '&') sb.Append("&amp;");
                else if (c == '<') sb.Append("&lt;");
                else if (c == '>') sb.Append("&gt;");
                else if (c == '"') sb.Append("&quot;");
                else if (c == '\t' || c == '\n' || c == '\r' || (c >= 0x20 && c != 0xFFFE && c != 0xFFFF && !(c >= 0xD800 && c <= 0xDFFF))) sb.Append(c);
                else if (c >= 0xD800 && c <= 0xDFFF) sb.Append(c); // keep surrogate halves; pairs are valid
            }
            return sb.ToString();
        }

        internal static string ColName(int idx)
        {
            var s = "";
            for (idx++; idx > 0; idx = (idx - 1) / 26) s = (char)('A' + (idx - 1) % 26) + s;
            return s;
        }

        public sealed class Sheet
        {
            readonly StreamWriter _w;
            readonly int _cols;
            int _row;
            static readonly Regex Num = new Regex(@"^-?(0|[1-9]\d{0,14})(\.\d{1,15})?$", RegexOptions.Compiled);

            internal Sheet(Stream s, IList<string> headers)
            {
                _w = new StreamWriter(s, new UTF8Encoding(false), 65536);
                _cols = headers.Count;
                _w.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/><sheetData>");
                WriteRow(headers, true);
            }

            public void AddRow(IList<string> values) { WriteRow(values, false); }

            void WriteRow(IList<string> values, bool header)
            {
                _row++;
                _w.Write("<row r=\"" + _row + "\">");
                for (int c = 0; c < values.Count; c++)
                {
                    string v = values[c];
                    if (v == null || (!header && v == "NULL")) continue;
                    string r = ColName(c) + _row;
                    if (!header && Num.IsMatch(v))
                        _w.Write("<c r=\"" + r + "\"><v>" + v + "</v></c>");
                    else
                    {
                        if (v.Length > 32767) v = v.Substring(0, 32767);
                        _w.Write("<c r=\"" + r + "\" t=\"inlineStr\"" + (header ? " s=\"1\"" : "") + "><is><t xml:space=\"preserve\">" + Esc(v) + "</t></is></c>");
                    }
                }
                _w.Write("</row>");
            }

            public void Finish()
            {
                _w.Write("</sheetData>");
                if (_cols > 0) _w.Write("<autoFilter ref=\"A1:" + ColName(_cols - 1) + Math.Max(_row, 1) + "\"/>");
                _w.Write("</worksheet>");
                _w.Dispose();
            }
        }
    }
}
