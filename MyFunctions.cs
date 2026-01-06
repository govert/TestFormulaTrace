using ExcelDna.Integration;

namespace TestFormulaTrace
{
    public static class MyFunctions
    {
        [ExcelFunction(Description = "My first .NET function")]
        public static string SayHello(string name)
        {
            return "Hello " + name;
        }

        // An ExcelFunction that takes an object[] input, and writes out a concatenation of the inputs in {,,} format
        [ExcelFunction(Description = "Concatenate a range of inputs into a string")]
        public static string MXL(params object[] inputs)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.Append("{");
            for (int i = 0; i < inputs.Length; i++)
            {
                if (i > 0)
                    sb.Append(",");
                sb.Append(inputs[i]?.ToString() ?? "null");
            }
            sb.Append("}");
            return sb.ToString();

        }
    }
}
