using System;
using System.Collections.Generic;
using System.IO;
using BenchmarkDotNet.Attributes;
using ExcelDataReader;

namespace LightweightExcelReader.Benchmarks
{
    public class ReadNextPerformance
    {
        [Benchmark]
        public void SkipNulls()
        {
            var fileName = TestHelper.TestsheetPath("MassiveFile.xlsx");
            var reader = new ExcelReader(fileName)
            {
                ReadNextBehaviour = ReadNextBehaviour.SkipNulls
            };
            var sheet = reader["sheet1"];
            var list = new List<object>();
            while (sheet.ReadNext())
            {
                list.Add(sheet.Value);
            }
            if (list.Count != 199998)
            {
                throw new Exception($"Expected 199998 items, but got {list.Count}");
            }
        }
        
        [Benchmark]
        public void ReadAllNulls()
        {
            var fileName = TestHelper.TestsheetPath("MassiveFile.xlsx");
            var reader = new ExcelReader(fileName)
            {
                ReadNextBehaviour = ReadNextBehaviour.ReadAllNulls
            };
            var sheet = reader["sheet1"];
            var list = new List<object>();
            while (sheet.ReadNext())
            {
                list.Add(sheet.Value);
            }
            if (list.Count != 399991)
            {
                throw new Exception($"Expected 399991 items but got {list.Count}");
            }
        }
        
        [Benchmark]
        public void ExcelDataReader()
        {
            var list = new List<object>();
            var fileName = TestHelper.TestsheetPath("MassiveFile.xlsx");
            using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read())
                        {
                            for (var i = 0; i < reader.FieldCount; i++)
                            {
                                list.Add(reader.GetValue(i));
                            }
                        }
                    } while (reader.NextResult());
                }
            }

            if (list.Count != 399992)
            {
                throw new Exception($"Expected 399991 items but got {list.Count}");
            }
        }
        
        [Benchmark]
        public void SkipNullsCompactFile()
        {
            var fileName = TestHelper.TestsheetPath("MassiveFileFull.xlsx");
            var reader = new ExcelReader(fileName)
            {
                ReadNextBehaviour = ReadNextBehaviour.SkipNulls
            };
            var sheet = reader["sheet1"];
            var list = new List<object>();
            while (sheet.ReadNext())
            {
                list.Add(sheet.Value);
            }
            if (list.Count != 199998)
            {
                throw new Exception($"Expected 199998 items, but got {list.Count}");
            }
        }
        
        [Benchmark]
        public void ReadAllNullsCompactFile()
        {
            var fileName = TestHelper.TestsheetPath("MassiveFileFull.xlsx");
            var reader = new ExcelReader(fileName)
            {
                ReadNextBehaviour = ReadNextBehaviour.ReadAllNulls
            };
            var sheet = reader["sheet1"];
            var list = new List<object>();
            while (sheet.ReadNext())
            {
                list.Add(sheet.Value);
            }
            if (list.Count != 199998)
            {
                throw new Exception($"Expected 399991 items but got {list.Count}");
            }
        }
        
        [Benchmark]
        public void ExcelDataReaderCompactFile()
        {
            var list = new List<object>();
            var fileName = TestHelper.TestsheetPath("MassiveFileFull.xlsx");
            using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read())
                        {
                            for (var i = 0; i < reader.FieldCount; i++)
                            {
                                list.Add(reader.GetValue(i));
                            }
                        }
                    } while (reader.NextResult());
                }
            }

            if (list.Count != 199998)
            {
                throw new Exception($"Expected 399991 items but got {list.Count}");
            }
        }
    }
}