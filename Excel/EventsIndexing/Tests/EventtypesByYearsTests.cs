#if DEBUG

using System;
using System.Linq;
using System.Collections.Generic;
using NUnit.Framework;

namespace Logy.Api.Mw.Excel.Tests
{
    public class EventtypesByYearsTests
    {
        [Test]
        public void Symmetric3(){
			var byYears = new Dictionary<short, string>();
			byYears.Add(1, "a,b,c");
			byYears.Add(2, "b,a,c");
			var rows = EventtypesByYears.Do(byYears);
			Assert.AreEqual(2, rows["a"]["b"]);        
			Assert.AreEqual(2, rows["a"]["c"]);
			Assert.AreEqual(2, rows["b"]["c"]);
        }

        [Test]
        public void Count2(){
			var byYears = new Dictionary<short, string>();
			byYears.Add(1, "b,a,b");
			var rows = EventtypesByYears.Do(byYears);
			Assert.AreEqual(1, rows["b"]["a"]);
			Assert.AreEqual(1, rows["a"]["b"]);
        }
    }
}
#endif
