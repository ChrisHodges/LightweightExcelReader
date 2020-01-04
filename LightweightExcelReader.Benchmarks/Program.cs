using BenchmarkDotNet.Running;

namespace LightweightExcelReader.Benchmarks
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            /*var benchmarks = new Benchmarks();
            //benchmarks.NPoi();
            benchmarks.OpenXml();
            benchmarks.ExcelDataReader();
            benchmarks.LightweightExcelReader();
            //benchmarks.GemBox();*/

            BenchmarkRunner.Run<Benchmarks>();
        }
    }
}