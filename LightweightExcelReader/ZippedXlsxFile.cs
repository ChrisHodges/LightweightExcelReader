using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace LightweightExcelReader
{
    internal class ZippedXlsxFile : IZippedXlsxFile
    {
        private ZipArchive _archive;
        private readonly Stream _fileStream;
        private readonly Dictionary<int, Stream> _openWorksheetStreams = new Dictionary<int, Stream>();
        private Stream _workbookXml;
        private ZipArchiveEntry[] _worksheetEntries;

        public ZippedXlsxFile(string filePath)
        {
            _fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            Initialize();
        }

        public ZippedXlsxFile(Stream excelStream)
        {
            _fileStream = excelStream;
            Initialize();
        }

        public Stream WorkbookXml
        {
            get
            {
                if (_workbookXml == null)
                {
                    _workbookXml = _archive.Entries.First(x => x.FullName.EndsWith("workbook.xml")).Open();
                }

                return _workbookXml;
            }
        }

        public Stream GetWorksheetStream(int i)
        {
            if (_openWorksheetStreams.ContainsKey(i))
            {
                return _openWorksheetStreams[i];
            }

            if (_worksheetEntries.Length <= i)
            {
                throw new ArgumentOutOfRangeException(nameof(i),$"Sheet with zero-based index {i} was not found in the workbook. Workbook contains {_worksheetEntries.Length} sheets.");
            }

            _openWorksheetStreams.Add(i, _worksheetEntries[i].Open());
            return _openWorksheetStreams[i];
        }

        public XslxSharedStringsStream SharedStringsStream { get; private set; }
        public XslxIsDateTimeStream IsDateTimeStream { get; private set; }

        public void Dispose()
        {
            _archive.Dispose();
            _fileStream.Dispose();
        }

        private void Initialize()
        {
            _archive = new ZipArchive(_fileStream, ZipArchiveMode.Read);
            //CSH 28012020 - We order by the length of the name string then by the name itself, so that Sheet10 appears immediately
            //after Sheet9 rather than after Sheet1
            _worksheetEntries = _archive.Entries.Where(x => x.FullName.StartsWith("xl/worksheets/sheet"))
                .OrderBy(x => x.Name.Length)
                .ThenBy(x => x.Name)
                .ToArray();
            var sharedStringsEntry = _archive.Entries.FirstOrDefault(x => x.FullName.EndsWith("sharedStrings.xml"));
            if (sharedStringsEntry != null) {
            SharedStringsStream = new XslxSharedStringsStream(sharedStringsEntry.Open());
            }
            IsDateTimeStream =
                new XslxIsDateTimeStream(_archive.Entries.First(x => x.FullName.EndsWith("styles.xml")).Open());
        }
    }
}