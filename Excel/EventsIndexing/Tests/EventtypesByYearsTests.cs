#if DEBUG

using System.Collections.Generic;
using NUnit.Framework;

namespace Routines.Excel.EventsIndexing.Tests
{
    public class EventtypesByYearsTests
    {
        [Test]
        public void Do_Symmetric()
        {
            var byYears = new Dictionary<int, string> {{1, "a,b,c"}, {2, "b,a,c"}};
            var rows = EventtypesByYears.Do(byYears);
            Assert.AreEqual(2, rows["a"]["b"]);
            Assert.AreEqual(2, rows["a"]["c"]);
            Assert.AreEqual(2, rows["b"]["c"]);
        }

        [Test]
        public void Do_Count()
        {
            var byYears = new Dictionary<int, string> {{1, "b,a,b"}};
            var rows = EventtypesByYears.Do(byYears);
            Assert.AreEqual(1, rows["b"]["a"]);
            Assert.AreEqual(1, rows["a"]["b"]);
        }

        [Test]
        public void Do_Shift()
        {
            var byYears = new Dictionary<int, string> { { 1, "a" }, { 2, "b" } };
            var rows = EventtypesByYears.Do(byYears, null, 1);
            Assert.AreEqual(1, rows["b"]["a"]);
            Assert.AreEqual(0, rows["a"].Count);
        }
    }
}
#endif
