using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using ExcelNumberFormat;

namespace LightweightExcelReader
{

    internal class XslxIsDateTimeStream : IDictionary<int, bool>, IEnumerator<KeyValuePair<int, bool>>
    {
        public void Reset()
        {
            throw new NotImplementedException();
        }

        object IEnumerator.Current { get; }
        public KeyValuePair<int, bool> Current { get; private set; }
        
        private readonly Dictionary<string, bool> _formatDictionary = new Dictionary<string, bool>
        {
            {"0", false},
            {"1", false},
            {"2", false},
            {"3", false},
            {"4", false},
            {"9", false},
            {"10", false},
            {"11", false},
            {"12", false},
            {"13", false},
            {"14", true},
            {"15", true},
            {"16", true},
            {"17", true},
            {"18", true},
            {"19", true},
            {"20", true},
            {"21", true},
            {"22", true},
            {"37", false},
            {"38", false},
            {"39", false},
            {"40", false},
            {"45", true},
            {"46", true},
            {"47", true},
            {"48", false},
            {"49", false},
        };
        private readonly Dictionary<int, bool> _storedKeys = new Dictionary<int, bool>();
        private readonly XmlReader _xmlReader;
        private int _readIndex = -1;
        private bool _readingCellXfs;

        public XslxIsDateTimeStream(Stream xmlStream)
        {
            _xmlReader = XmlReader.Create(xmlStream,
                new XmlReaderSettings {ConformanceLevel = ConformanceLevel.Fragment});
        }

        public IEnumerator<KeyValuePair<int, bool>> GetEnumerator()
        {
            return this;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(KeyValuePair<int, bool> item)
        {
            throw new NotImplementedException();
        }

        public void Clear()
        {
        }

        public bool Contains(KeyValuePair<int, bool> item)
        {
            throw new NotImplementedException();
        }

        public void CopyTo(KeyValuePair<int, bool>[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public bool Remove(KeyValuePair<int, bool> item)
        {
            throw new NotImplementedException();
        }

        public int Count { get; }
        public bool IsReadOnly { get; }

        public void Add(int key, bool value)
        {
            throw new NotImplementedException();
        }

        public bool ContainsKey(int key)
        {
            throw new NotImplementedException();
        }

        public bool Remove(int key)
        {
            throw new NotImplementedException();
        }

        public bool TryGetValue(int key, out bool value)
        {
            throw new NotImplementedException();
        }

        private void HandleReadingXfElement()
        {
            _readIndex++;
            var fmtId = _xmlReader.GetAttribute("numFmtId");
            _storedKeys[_readIndex] = _formatDictionary[fmtId];
            Current = new KeyValuePair<int, bool>(_readIndex, _storedKeys[_readIndex]);
        }

        public bool MoveNext()
        {
            if (!_readingCellXfs)
            {
                AdvanceXmlReaderToCellXfs();
            }
            
            while (_xmlReader.Read())
            {
                if (_readingCellXfs && _xmlReader.IsStartOfElement("xf"))
                {
                    HandleReadingXfElement();
                    return true;
                }
            }

            return false;
        }

        public bool this[int key]
        {
            get
            {
                if (key > _readIndex)
                {
                    return AdvanceToIndex(key);
                }

                return _storedKeys[key];
            }
            set => throw new NotImplementedException();
        }

        public ICollection<int> Keys { get; }
        public ICollection<bool> Values { get; }

        private void AdvanceXmlReaderToCellXfs()
        {
            while (_xmlReader.Read())
            {
                if (_xmlReader.IsStartOfElement("numFmt"))
                {
                    var attribute = _xmlReader.GetAttribute("numFmtId");
                    if (!_formatDictionary.ContainsKey(attribute)) {
                        _formatDictionary.Add(attribute,
                            new NumberFormat(_xmlReader.GetAttribute("formatCode")).IsDateTimeFormat);
                    }
                }
                
                if (_xmlReader.IsStartOfElement("cellXfs"))
                {
                    _readingCellXfs = true;
                    break;
                }
            }
        }

        private bool AdvanceToIndex(int key)
        {
            if (!_readingCellXfs)
            {
                AdvanceXmlReaderToCellXfs();
            }
            while (_xmlReader.Read())
            {
                if (_readingCellXfs && _xmlReader.IsStartOfElement("xf"))
                {
                    HandleReadingXfElement();
                    if (_readIndex == key)
                    {
                        return _storedKeys[_readIndex];
                    }
                }
            }

            throw new KeyNotFoundException($"The key '{key}' was not found in the dictionary");
        }

        public int? GetFirstDateTimeStyle()
        {
            foreach (var item in _storedKeys)
            {
                if (item.Value)
                {
                    return item.Key;
                }
            }

            while (MoveNext())
            {
                if (Current.Value)
                {
                    return Current.Key;
                }
            }

            return default;
        }

        public void Dispose()
        {
        }
    }
}