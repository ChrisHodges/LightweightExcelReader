using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using FluentAssertions;
using Xunit;

namespace LightweightExcelReader.Tests.TestHelpers
{
    public class CultureGenerator : IEnumerable<object[]>
    {
        private readonly IEnumerable<object[]> _cultures;

        public CultureGenerator()
        { 
            _cultures = CultureInfo.GetCultures(CultureTypes.AllCultures).Select(x => new []{x}).AsEnumerable();
        }

        public IEnumerator<object[]> GetEnumerator()
        {
            return _cultures.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}