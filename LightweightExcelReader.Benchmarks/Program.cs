using BenchmarkDotNet.Running;

namespace LightweightExcelReader.Benchmarks
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            //BenchmarkRunner.Run<ReadNextPerformance>();
            BenchmarkRunner.Run<Benchmarks>();
        }
    }
}