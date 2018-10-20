#if DEBUG

using System.Collections.Generic;
using Logy.Api.Mw.Excel;
using NPOI.HSSF.UserModel;
/*
 * ячейки G3059, A6664, D8564, P19684 и т.д. и ниже пустые. Ёто правильно?
<img src="https://ic.pics.livejournal.com/it4history/77674790/10303/10303_original.png" alt="" title="">

также вижу новые индексы, возможно, это индексы места, но обрабатываю их как раньше индексы типа событий (может ошибаюсь)

считаю, что в каждой €чейке по одному индексу, то есть символ $ не есть разделителем

использую тот же <a href="https://github.com/it4history/reconstruction/blob/master/Excel/EventsIndexing/EventtypesByYears.cs">самый алгоритм</a>

полученные таблицы размером 7516 на 7516 в формате CSV, потому что генерирование XLSX не дождалс€, наход€тс€ в папке https://github.com/it4history/reconstruction/tree/master/Excel/EventsIndexing/Tests/outTech10

 */
namespace Routines.Excel.EventsIndexing.Tests
{
    public class Tech10Launcher : LauncherBase
    {
        public override string FileNameIn
        {
            get { return "00 “ехфайл 10.xls"; }
        }

        public override string FolderOut
        {
            get { return "outTech10"; }
        }

        public override OutputType FileOutputType
        {
            get { return OutputType.Csv; }
        }

        protected override Dictionary<int, string> GroupEventsByYear(ExcelManager man, Dictionary<string, string> legends)
        {
            var byYears = new Dictionary<int, string>();
            foreach (HSSFRow row in man.Records)
            {
                for (int colGroup = 0; colGroup < 12; colGroup++)
                {
                    var i = colGroup * 2;
                    if (row.RowNum >= 3058 && colGroup >= 11
                        || row.RowNum >= 6663 && colGroup >= 10
                        || row.RowNum >= 8563 && colGroup >= 9
                        || row.RowNum >= 19683 && colGroup >= 8
                        || row.RowNum >= 25229 && colGroup >= 7
                        || row.RowNum >= 37180 && colGroup >= 6
                        || row.RowNum >= 45344 && colGroup >= 5
                        || row.RowNum >= 54304 && colGroup >= 4
                        )
                    {
                    }
                    else
                    {
                        var year = (int) row.Cells[i].NumericCellValue;
                        var index = row.Cells[i + 1].ToString();
                        byYears[year] = (byYears.ContainsKey(year) ? byYears[year] : null) + index + ",";
                    }
                }
            }
            return byYears;
        }
    }
}
#endif