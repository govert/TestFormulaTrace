using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;

public static class CopyMxlCommand
{
    // Entry points (same shape as yours)
    [ExcelCommand(MenuName = "CopyMXL", MenuText = "Copy MXL Formula Values", ShortCut = "%c")]
    public static void CopyMXLFormulaValues() => CopyMXLFormulas(includeValues: false);
    [ExcelCommand(MenuName = "CopyMXL", MenuText = "Copy MXL Formulas and Values", ShortCut = "%C")]
    public static void CopyMXLFormulaAndValues() => CopyMXLFormulas(includeValues: true);

    private static void CopyMXLFormulas(bool includeValues)
    {
        var app = ExcelDnaUtil.Application as Excel.Application;
        if (app == null) return;

        var selection = app.Selection as Excel.Range;
        if (selection == null) return;

        var results = new List<string>();

        // Worst bug in original: de-dup by formula string (can collide across cells).
        // Fix: de-dup by anchor cell identity (sheet + absolute address).
        var seenAnchors = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (Excel.Range cell in selection.Cells)
        {
            // Handles:
            // - legacy CSE arrays (HasArray/CurrentArray)
            // - dynamic spills (SpillParent/SpillRange)
            // - normal cells
            var spillInfo = SpillHelpers.GetSpillInfo(cell);
            var anchor = spillInfo.Anchor;
            if (anchor == null) continue;

            var anchorKey = GetAnchorKey(anchor);
            if (!seenAnchors.Add(anchorKey))
                continue;

            // Only work on formulas that contain MXL*
            if (!SafeBool(() => anchor.HasFormula))
                continue;

            // Prefer Formula2 if available (dynamic arrays, modern syntax), else Formula
            var formula = ExcelDynamicArrayHelpers_GetFormulaTextDynamic(anchor);
            if (string.IsNullOrEmpty(formula))
                continue;

            if (formula.IndexOf("MXL", StringComparison.OrdinalIgnoreCase) < 0)
                continue;

            // "Worst" correctness problem in original was regex over formula text.
            // Here we do a minimal-but-safe improvement:
            // - evaluate nested MXL calls (innermost first)
            // - replace A1-style refs outside of string literals
            // - handle A1# spill refs (resolve to spill range values)
            var evaluatedFormula = RewriteFormula_Minimal(app, anchor.Worksheet, formula);

            if (!includeValues)
            {
                results.Add(evaluatedFormula);
                continue;
            }

            // Values:
            // - For dynamic spills/legacy arrays, include the whole range values
            // - Otherwise just the anchor cell value
            object val2 = SafeGet(() => spillInfo.Range?.Value2);
            var valueText = Value2ToTsv(val2);

            results.Add($"{evaluatedFormula} => {valueText}");
        }

        if (results.Count == 0)
        {
            MessageBox.Show("No MXL formulas found in selection", "CopyMXL",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var resultString = string.Join(Environment.NewLine, results);
        Clipboard.SetText(resultString);
        MessageBox.Show($"Copied {results.Count} MXL formula(s) to clipboard\r\n{resultString}", "CopyMXL",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    // --------------------------
    // Minimal "fixed worst problems" rewrite pipeline
    // --------------------------

    private static string RewriteFormula_Minimal(Excel.Application app, Excel.Worksheet sheet, string formula)
    {
        // 1) Reduce nested MXL calls (innermost first), leaving the outermost MXL intact.
        // This is intentionally conservative: it only targets nested MXL occurrences.
        var afterNested = ReduceNestedMxl(app, sheet, formula);

        // 2) Replace A1-style refs outside quotes (safe-ish), including A1# spill refs.
        // Still not a full Excel parser, but avoids the biggest footgun: touching inside strings.
        var afterRefs = ReplaceA1RefsOutsideStrings(app, sheet, afterNested);

        return afterRefs;
    }

    // Evaluate innermost MXL*(...) segments and replace them with literals,
    // but keep the outermost MXL*(...) intact.
    private static string ReduceNestedMxl(Excel.Application app, Excel.Worksheet sheet, string formula)
    {
        // Find all MXL-like function call spans using a simple stack, skipping text in quotes.
        var spans = FindMxlCallSpans(formula);
        if (spans.Count <= 1) return formula; // no nesting

        // Keep the outermost span intact; reduce all inner ones (process from end to not break indices)
        var outermost = spans.OrderByDescending(s => s.Length).First();

        var innerSpans = spans
            .Where(s => !(s.Start == outermost.Start && s.Length == outermost.Length))
            .OrderByDescending(s => s.Start)
            .ToList();

        var result = formula;

        foreach (var span in innerSpans)
        {
            try
            {
                // Re-evaluate span boundaries against current result:
                // since we're replacing from right-to-left and spans are ordered by Start desc,
                // original indices remain valid as long as we use original string.
                // For simplicity in "minimal" version, we assume spans refer to the original formula.
                // If you want bulletproof, re-scan each iteration.
                var text = formula.Substring(span.Start, span.Length);
                var expr = "=" + text;

                // Worksheet.Evaluate gives correct context (names, tables, etc.)
                object val = sheet.Evaluate(expr);

                var literal = ToFormulaLiteral(val);
                result = result.Remove(span.Start, span.Length).Insert(span.Start, literal);
            }
            catch
            {
                // If an inner MXL call can't be evaluated, leave it as-is.
            }
        }

        return result;
    }

    // Replace A1, $A$1, A1:B3, Sheet!A1, 'Sheet 1'!A1#, etc. outside of quoted strings.
    // Includes spill refs (#): resolves to SpillRange and serializes as TSV inside quotes.
    private static string ReplaceA1RefsOutsideStrings(Excel.Application app, Excel.Worksheet sheet, string formula)
    {
        var sb = new StringBuilder(formula.Length);
        int i = 0;
        bool inString = false;

        while (i < formula.Length)
        {
            char ch = formula[i];

            // Handle string literals: "..." with "" escape.
            if (ch == '"')
            {
                sb.Append(ch);
                i++;

                if (inString)
                {
                    // If next is also quote, it's an escape: keep inString true.
                    if (i < formula.Length && formula[i] == '"')
                    {
                        sb.Append('"');
                        i++;
                    }
                    else
                    {
                        inString = false;
                    }
                }
                else
                {
                    inString = true;
                }
                continue;
            }

            if (inString)
            {
                sb.Append(ch);
                i++;
                continue;
            }

            // Try to parse a reference token starting at i.
            if (TryReadReferenceToken(formula, i, out var token, out var tokenLen))
            {
                var rng = ReferenceResolver.ResolveReference(app, sheet, token);
                if (rng != null)
                {
                    // If single cell: use display text as quoted string (matches your example).
                    // If multi-cell (spill/range): TSV as a single quoted string.
                    string replacement;
                    if (IsSingleCell(rng))
                    {
                        replacement = QuoteForFormula(SafeString(() => rng.Text?.ToString()) ?? SafeString(() => rng.Value2?.ToString()) ?? "");
                    }
                    else
                    {
                        object v2 = SafeGet(() => rng.Value2);
                        string tsv = Value2ToTsv(v2);
                        replacement = QuoteForFormula(tsv);
                    }

                    sb.Append(replacement);
                    i += tokenLen;
                    continue;
                }
            }

            sb.Append(ch);
            i++;
        }

        return sb.ToString();
    }

    // --------------------------
    // Reference token reader (minimal)
    // --------------------------

    // Reads:
    //  - A1, $A$1, A$1, $A1
    //  - A1:B3
    //  - Sheet1!A1, 'Sheet 1'!A1
    //  - ... plus optional trailing #
    private static bool TryReadReferenceToken(string s, int start, out string token, out int length)
    {
        token = null;
        length = 0;

        int i = start;

        // Optional sheet qualifier: Sheet! or 'Sheet Name'!
        int sheetStart = i;
        if (i < s.Length && s[i] == '\'')
        {
            // quoted sheet: '...'
            i++;
            while (i < s.Length)
            {
                if (s[i] == '\'')
                {
                    i++;
                    break;
                }
                i++;
            }
            if (i < s.Length && s[i] == '!')
            {
                i++; // include !
            }
            else
            {
                return false;
            }
        }
        else
        {
            // unquoted sheet: letters/numbers/_ until !
            int j = i;
            while (j < s.Length && (char.IsLetterOrDigit(s[j]) || s[j] == '_' || s[j] == '.'))
                j++;

            if (j < s.Length && s[j] == '!')
            {
                // accept as sheet qualifier
                i = j + 1;
            }
        }

        // Now parse first cell address: [$]COL[$]ROW
        int addrStart = i;

        // optional $
        if (i < s.Length && s[i] == '$') i++;

        // column letters (A..XFD etc.)
        int colStart = i;
        while (i < s.Length && char.IsLetter(s[i]))
            i++;
        if (i == colStart) return false;

        // optional $
        if (i < s.Length && s[i] == '$') i++;

        // row digits
        int rowStart = i;
        while (i < s.Length && char.IsDigit(s[i]))
            i++;
        if (i == rowStart) return false;

        // Optional range part : second address
        int afterFirst = i;
        if (i < s.Length && s[i] == ':')
        {
            i++; // :
            // second addr
            if (i < s.Length && s[i] == '$') i++;

            int col2Start = i;
            while (i < s.Length && char.IsLetter(s[i]))
                i++;
            if (i == col2Start) return false;

            if (i < s.Length && s[i] == '$') i++;

            int row2Start = i;
            while (i < s.Length && char.IsDigit(s[i]))
                i++;
            if (i == row2Start) return false;
        }

        // Optional spill marker '#'
        if (i < s.Length && s[i] == '#')
            i++;

        // Guard against grabbing function names like AVERAGE(...) etc.
        // We require a "non-identifier" boundary before start.
        if (start > 0 && (char.IsLetterOrDigit(s[start - 1]) || s[start - 1] == '_'))
            return false;

        token = s.Substring(start, i - start);
        length = i - start;
        return true;
    }

    // --------------------------
    // MXL call span finder (minimal)
    // --------------------------

    private sealed class Span { public int Start; public int Length; }

    private static List<Span> FindMxlCallSpans(string formula)
    {
        var spans = new List<Span>();
        var stack = new Stack<(int openParenIndex, int nameStartIndex)>();

        bool inString = false;

        for (int i = 0; i < formula.Length; i++)
        {
            char ch = formula[i];

            if (ch == '"')
            {
                if (inString)
                {
                    // handle escaped quote
                    if (i + 1 < formula.Length && formula[i + 1] == '"')
                    {
                        i++;
                        continue;
                    }
                    inString = false;
                }
                else inString = true;

                continue;
            }

            if (inString) continue;

            // detect MXL*(
            if (IsStartOfMxlCall(formula, i, out int nameStart, out int openParen))
            {
                stack.Push((openParen, nameStart));
                i = openParen; // continue after '('
                continue;
            }

            if (ch == ')' && stack.Count > 0)
            {
                (openParen, nameStart) = stack.Pop();
                spans.Add(new Span { Start = nameStart, Length = (i - nameStart) + 1 });
            }
        }

        return spans;
    }

    private static bool IsStartOfMxlCall(string s, int i, out int nameStart, out int openParen)
    {
        nameStart = -1;
        openParen = -1;

        // Must be at an identifier boundary
        if (i > 0 && (char.IsLetterOrDigit(s[i - 1]) || s[i - 1] == '_')) return false;

        // Must start with MXL (case-insensitive)
        if (i + 3 > s.Length) return false;
        if (!s.Substring(i, 3).Equals("MXL", StringComparison.OrdinalIgnoreCase)) return false;

        int j = i + 3;
        while (j < s.Length && (char.IsLetterOrDigit(s[j]) || s[j] == '_'))
            j++;

        if (j < s.Length && s[j] == '(')
        {
            nameStart = i;
            openParen = j;
            return true;
        }

        return false;
    }

    // --------------------------
    // Formatting helpers
    // --------------------------

    private static string ExcelDynamicArrayHelpers_GetFormulaTextDynamic(Excel.Range anchor)
    {
        try
        {
            dynamic d = anchor;
            var f2 = d.Formula2 as string;
            if (!string.IsNullOrEmpty(f2)) return f2;
        }
        catch { /* ignore */ }

        try
        {
            return anchor.Formula?.ToString();
        }
        catch
        {
            return null;
        }
    }

    private static string Value2ToTsv(object value2)
    {
        if (value2 == null) return "";

        if (value2 is object[,] grid)
        {
            int rows = grid.GetLength(0);
            int cols = grid.GetLength(1);
            var sb = new StringBuilder();

            for (int r = 1; r <= rows; r++)
            {
                if (r > 1) sb.AppendLine();
                for (int c = 1; c <= cols; c++)
                {
                    if (c > 1) sb.Append('\t');
                    sb.Append(ScalarToText(grid[r, c]));
                }
            }
            return sb.ToString();
        }

        return ScalarToText(value2);
    }

    private static string ScalarToText(object v)
    {
        if (v == null) return "";

        if (v is double d) return d.ToString("G17", CultureInfo.InvariantCulture);
        if (v is bool b) return b ? "TRUE" : "FALSE";

        return Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
    }

    // For this simplified version, match your example: everything becomes a quoted string.
    // (You can later switch to typed literals if desired.)
    private static string ToFormulaLiteral(object value)
        => QuoteForFormula(value == null ? "" : Convert.ToString(value, CultureInfo.InvariantCulture) ?? "");

    private static string QuoteForFormula(string text)
    {
        // Excel string literal uses "" to escape "
        var safe = (text ?? "").Replace("\"", "\"\"");
        return $"\"{safe}\"";
    }

    private static bool IsSingleCell(Excel.Range rng)
    {
        try { return rng.Rows.Count == 1 && rng.Columns.Count == 1; }
        catch { return true; }
    }

    private static string GetAnchorKey(Excel.Range anchor)
    {
        // Sheet + absolute address. Include workbook name if you want cross-workbook safety.
        var sheetName = SafeString(() => anchor.Worksheet?.Name) ?? "";
        var addr = SafeString(() => anchor.Address[RowAbsolute: true, ColumnAbsolute: true]) ?? anchor.Address;
        return $"{sheetName}!{addr}";
    }

    private static bool SafeBool(Func<bool> f) { try { return f(); } catch { return false; } }
    private static string SafeString(Func<string> f) { try { return f(); } catch { return null; } }
    private static object SafeGet(Func<object> f) { try { return f(); } catch { return null; } }
}
