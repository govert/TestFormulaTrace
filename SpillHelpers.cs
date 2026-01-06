using System;
using Excel = Microsoft.Office.Interop.Excel;

public static class SpillHelpers
{
    public sealed class SpillInfo
    {
        public Excel.Range Anchor { get; set; }      // formula cell (top-left)
        public Excel.Range Range { get; set; }       // spilled result range (or CurrentArray for CSE)
        public bool IsLegacyArray { get; set; }
        public bool IsDynamicSpill { get; set; }
    }

    public static SpillInfo GetSpillInfo(Excel.Range anyCell)
    {
        // This is about the best information: https://www.reddit.com/r/excel/comments/9v7w16/dynamic_array_formulas_new_properties_in_vba/

        if (anyCell == null) throw new ArgumentNullException(nameof(anyCell));

        // Legacy CSE array
        try
        {
            if (anyCell.HasArray)
            {
                var arr = anyCell.CurrentArray;
                var anchor = (Excel.Range)arr.Cells[1, 1];
                return new SpillInfo
                {
                    Anchor = anchor,
                    Range = arr,
                    IsLegacyArray = true,
                    IsDynamicSpill = false
                };
            }
        }
        catch { /* ignore */ }

        // Dynamic spill (late-bound)
        Excel.Range spillParent = null;
        Excel.Range spillingToRange = null;

        try
        {
            dynamic d = anyCell;
            // These properties exist in newer Excel; in older Excel this throws.
            spillParent = (Excel.Range)d.SpillParent;
        }
        catch { /* ignore */ }

        try
        {
            dynamic d = anyCell;
            spillingToRange = (Excel.Range)d.SpillingToRange;
        }
        catch { /* ignore */ }

        // If we have a parent, prefer its spill range (more stable)
        if (spillParent != null)
        {
            try
            {
                dynamic p = spillParent;
                spillingToRange = (Excel.Range)p.SpillingToRange;
            }
            catch { /* keep whatever we already have */ }

            return new SpillInfo
            {
                Anchor = spillParent,
                Range = spillingToRange ?? spillParent,
                IsLegacyArray = false,
                IsDynamicSpill = true
            };
        }

        // If no parent but SpillRange exists and includes the cell, treat top-left as anchor
        if (spillingToRange != null)
        {
            var topLeft = (Excel.Range)spillingToRange.Cells[1, 1];
            return new SpillInfo
            {
                Anchor = topLeft,
                Range = spillingToRange,
                IsLegacyArray = false,
                IsDynamicSpill = true
            };
        }

        // Normal single cell
        return new SpillInfo
        {
            Anchor = anyCell,
            Range = anyCell,
            IsLegacyArray = false,
            IsDynamicSpill = false
        };
    }
}
