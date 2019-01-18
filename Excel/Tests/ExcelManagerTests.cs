#if DEBUG

using System.IO;
using Logy.Entities.Localization;
using NUnit.Framework;

namespace Logy.Api.Mw.Excel.Tests
{
    public class ExcelManagerTests
    {
        [Test]
        public void Read()
        {
            var man = new ExcelManager(Path.Combine("Mw/Excel/Tests", "Крым реконструкция.xls"));
            man.Read();
            Assert.AreEqual(377, man.Sheet.PhysicalNumberOfRows);
        }

        [Test]
        public void TrimTitle()
        {
            Assert.AreEqual("1  test", ExcelManager.TrimTitle("1 [sdf] test"));
        }

        [Test]
        public void GetYears()
        {
            Assert.IsNull(ExcelManager.GetYears(" test"));

            // numbers till 10 are not treated as year
            Assert.IsNull(ExcelManager.GetYears("1 test"));
            Assert.AreEqual(new[] { "195" }, ExcelManager.GetYears("195 test test"));
            Assert.AreEqual(new[] { "1380", "1387" }, ExcelManager.GetYears("1380 - 1387 asdfasdf"));
            Assert.AreEqual(new[] { "1511", "1512" }, ExcelManager.GetYears("1511—1512 годах"));
            Assert.AreEqual(new[] { "1428" }, ExcelManager.GetYears("1428 году test"));
            Assert.AreEqual(new[] { "1695" }, ExcelManager.GetYears("1695 25 test"));
        }

        [Test]
        public void GetDescription()
        {
            Assert.AreEqual("test", ExcelManager.GetDescription("12 test"));
            Assert.AreEqual("test test", ExcelManager.GetDescription("1905 test test."));
            Assert.AreEqual("asdfasdf", ExcelManager.GetDescription("1380 - 1387 asdfasdf ."));
            Assert.AreEqual("a", ExcelManager.GetDescription(" a"));
            Assert.AreEqual("asdfasdf", ExcelManager.GetDescription("1380 - 1387гг. asdfasdf "));
            Assert.AreEqual("asdfasdf", ExcelManager.GetDescription("1380 - 1387 гг. asdfasdf "));
            Assert.AreEqual("asdfasdf", ExcelManager.GetDescription("1380 - 1387г. asdfasdf"));
            Assert.AreEqual("asdfasdf", ExcelManager.GetDescription("1380 - 1387г asdfasdf "));
            Assert.AreEqual("asdfasdf", ExcelManager.GetDescription("1380 - 1387 гг asdfasdf "));
            Assert.AreEqual("test", ExcelManager.GetDescription("1428 году test"));
            Assert.AreEqual("test", ExcelManager.GetDescription("1428 года test"));
            Assert.AreEqual("генуэзские колонии", ExcelManager.GetDescription("1396 генуэзские колонии"));
            Assert.AreEqual("генуэзцы", ExcelManager.GetDescription("1396 годах генуэзцы"));
            Assert.AreEqual("генуэзцы", ExcelManager.GetDescription("1381 г., генуэзцы"));
        }

        [Test]
        public void GetLanguage()
        {
            Assert.AreEqual("uk", ExcelManager.GetLanguage("Вікіпедія.Хронологія історії України"));
            Assert.AreEqual("ru", ExcelManager.GetLanguage("Википедия.Генуэзские колонии в Северном Причерноморье"));
        }

        [Test]
        public void GetUrl()
        {
            Assert.AreEqual("[[wikiuk:Хронологія історії України]]", ExcelManager.GetUrl("Вікіпедія.Хронологія історії України"));
            Assert.AreEqual(
                "[[wikiru:Генуэзские колонии в Северном Причерноморье]]",
                ExcelManager.GetUrl("Википедия.Генуэзские колонии в Северном Причерноморье"));
        }

        [Test]
        public void GetRowNumFromText()
        {
            Assert.AreEqual("190", ExcelManager.GetRowNumFromText("excelParameters=[[1-row::190]]", "1"));
        }

        [Test]
        public void GetJsonTranslation()
        {
            TranslationManager.Language = "ru";
            Assert.AreEqual("Описание", JsonManager.GetJsonTranslation(ExcelFileColumns.Description));
        }
    }
}
#endif
