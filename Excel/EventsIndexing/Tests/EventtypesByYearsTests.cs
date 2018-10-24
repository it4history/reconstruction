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
            var byYears = new Graphes(new Dictionary<int, string> { { 1, "a,b,c" }, { 2, "b,a,c" } });
            var rows = EventtypesByYears.Do(byYears);
            Assert.AreEqual(2, rows["a"]["b"]);
            Assert.AreEqual(2, rows["a"]["c"]);
            Assert.AreEqual(2, rows["b"]["c"]);
        }

        [Test]
        public void Do_Count()
        {
            var byYears = new Graphes(new Dictionary<int, string> {{1, "b,a,b"}});
            var rows = EventtypesByYears.Do(byYears);
            Assert.AreEqual(1, rows["b"]["a"]);
            Assert.AreEqual(1, rows["a"]["b"]);
        }

        [Test]
        public void Do_Shift()
        {
            var byYears = new Graphes(new Dictionary<int, string>
            {
                { 1, "a,b" },
                { 2, "b" },
                { 3, "c,a" }
            });
            var rows = EventtypesByYears.Do(byYears, 1);
            Assert.AreEqual(@"$, a, b, c, 
a, 0, 1, 0, 
b, 1, 1, 1, 
c, 0, 0, 0, 
", LauncherBase.OutputConsole(rows));

            rows = EventtypesByYears.Do(byYears, 2);
            Assert.AreEqual(@"$, a, b, c, 
a, 1, 0, 1, 
b, 1, 0, 1, 
c, 0, 0, 0, 
", LauncherBase.OutputConsole(rows));
        }
    }
}
#endif
