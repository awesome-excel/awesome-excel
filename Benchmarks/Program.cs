using BenchmarkDotNet.Running;

namespace Benchmarks;

public class Program
{
    private static void Main(string[] args)
    {
        var summary = BenchmarkRunner.Run<Generate_Excel_BigDataSets>();
    }
}