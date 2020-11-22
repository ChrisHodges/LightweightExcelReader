using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace LightweightExcelReader
{
    internal class XslxSharedStringsStream : IDictionary<int, string>
    {
        private readonly XmlReader _xmlReader;
        private int _readIndex = -1;

        internal XslxSharedStringsStream(Stream xmlStream)
        {
            _xmlReader = XmlReader.Create(xmlStream,
                new XmlReaderSettings {ConformanceLevel = ConformanceLevel.Fragment});
            _xmlReader.MoveToContent();
        }

        private Dictionary<int, string> _storedKeys { get; } = new Dictionary<int, string>();

        public IEnumerator<KeyValuePair<int, string>> GetEnumerator()
        {
            return _storedKeys.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(KeyValuePair<int, string> item)
        {
            throw new NotImplementedException();
        }

        public void Clear()
        {
            throw new NotImplementedException();
        }

        public bool Contains(KeyValuePair<int, string> item)
        {
            throw new NotImplementedException();
        }

        public void CopyTo(KeyValuePair<int, string>[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public bool Remove(KeyValuePair<int, string> item)
        {
            throw new NotImplementedException();
        }

        public int Count { get; }
        public bool IsReadOnly => true;

        public void Add(int key, string value)
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

        public bool TryGetValue(int key, out string value)
        {
            throw new NotImplementedException();
        }

        public string this[int key]
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
        public ICollection<string> Values { get; }

        private string GetFormattedValue()
        {
            var returnString = "";
            while (!_xmlReader.IsEndOfElement("si"))
            {
                if (_xmlReader.IsStartOfElement("t"))
                {
                    _xmlReader.Read();
                    returnString = returnString + _xmlReader.Value;
                }

                _xmlReader.Read();
            }

            return returnString;
        }

        private string AdvanceToIndex(int key)
        {
            while (_xmlReader.Read())
            {
                if (_xmlReader.IsStartOfElement("si"))
                {
                    _readIndex++;
                }

                if (_xmlReader.IsStartOfElement("r"))
                {
                    var compoundVal = GetFormattedValue();
                    _storedKeys[_readIndex] = compoundVal;
                    if (_readIndex == key)
                    {
                        return compoundVal;
                    }
                }

                if (_xmlReader.IsStartOfElement("t"))
                {
                    _xmlReader.Read();
                    var val = _xmlReader.Value;
                    _storedKeys[_readIndex] = val;
                    if (_readIndex == key)
                    {
                        return val;
                    }
                }
            }

            throw new KeyNotFoundException($"The key '{key}' was not found in the dictionary");
        }
    }
}