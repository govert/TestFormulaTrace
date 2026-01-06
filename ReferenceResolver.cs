using System;
using Excel = Microsoft.Office.Interop.Excel;

public static class ReferenceResolver
{
    public static Excel.Range ResolveReference(
        Excel.Application app,
        Excel.Worksheet currentSheet,
        string referenceToken)
    {
        if (app == null) throw new ArgumentNullException(nameof(app));
        if (currentSheet == null) throw new ArgumentNullException(nameof(currentSheet));
        if (string.IsNullOrWhiteSpace(referenceToken)) return null;

        bool isSpillRef = referenceToken.EndsWith("#", StringComparison.Ordinal);
        string token = isSpillRef ? referenceToken.Substring(0, referenceToken.Length - 1) : referenceToken;

        // token may be:
        //  - A1
        //  - $A$1
        //  - A1:B3
        //  - Sheet1!A1
        //  - 'Sheet 1'!A1:B3
        //
        // We'll let Excel resolve the address string into a Range.
        Excel.Range rng = null;

        try
        {
            // If token has "!" treat as fully qualified; app.Range works with active sheet context too,
            // but it's safer to use currentSheet.Parent (workbook) if you want.
            if (token.Contains("!"))
                rng = app.Range[token];
            else
                rng = currentSheet.Range[token];
        }
        catch
        {
            return null;
        }

        if (!isSpillRef)
            return rng;

        // Spill reference: resolve spill range based on the *top-left cell* of the referenced range.
        // Excel defines A1# from a single cell anchor; if user writes A1:B3#, it's unusual.
        // We'll implement as: take top-left of rng and return its spill range if it has one.
        Excel.Range anchor = null;
        try { anchor = (Excel.Range)rng.Cells[1, 1]; } catch { anchor = rng; }

        var spill = SpillHelpers.GetSpillInfo(anchor);
        return spill.IsDynamicSpill ? spill.Range : anchor; // if no spill, treat as anchor
    }
}
