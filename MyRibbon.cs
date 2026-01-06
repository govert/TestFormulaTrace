using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

namespace TestFormulaTrace
{
    [ComVisible(true)]
    public class MyRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return RibbonResources.Ribbon;
        }

        public override object? LoadImage(string imageId)
        {
            // This will return the image resource with the name specified in the image='xxxx' tag
            return RibbonResources.ResourceManager.GetObject(imageId);
        }

        public void CopyMXLValuesPressed(IRibbonControl control)
        {
            CopyMxlCommand.CopyMXLFormulaAndValues();
        }

        public void CopyMxlFormulaPressed(IRibbonControl control)
        {
            CopyMxlCommand.CopyMXLFormulaValues();
        }
    }
}
