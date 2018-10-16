#if DEBUG

using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NUnit.Framework;

namespace Logy.Api.Mw.Excel.Tests.EventsIndexing
{
    public class T1Tests
    {
    	const string path = "Excel/EventsIndexing/Tests";
    	
        [Test]
        public void DoFull()
        {
            var byYears = Read();
            var rows = EventtypesByYears.Do(byYears);
            Output(rows, "сводная");
        }

		[Test]
        public void Do20()
        {
            var byYears = Read();
            var filter = "яш кя ля пс ую щб рю бж фг фю сг цх ак цз ск фм аю ип нв";
            var rows = EventtypesByYears.Do(byYears, filter);
            Output(rows, "сводная20", filter);
        }
      
        public Dictionary<short, string> Read()
        {
            var man = new ExcelManager(Path.Combine(path, "00_База_2018_10_01.xls"));
            man.Read();
            
            // group by year
            var byYears = new Dictionary<short, string>();
            foreach (HSSFRow row in man.Records)
            {
                var indices = man.GetValue(row, "Индекс");
                if (!string.IsNullOrEmpty(indices)) {
                	var year = short.Parse(man.GetValue(row, "-99000"));
                	byYears[year] = (byYears.ContainsKey(year) ? byYears[year] : null) + indices;
                }
            }
            return byYears;
        }

        public void Output(Dictionary<string, Dictionary<string, int>> rows, 
                           string name,
                           string outOrder = null) {
        	var file = string.Format("../../{0}/out/{1}.xlsx", path, name);
       		var sorted = outOrder == null 
       			? (IEnumerable<string>)rows.Keys.Where(s => rows[s].Count > 0).OrderBy(s => s)
       			: outOrder.Split(' ');
        	/*var console = File.CreateText(file);
       		Console.WriteLine();
       		console.Write(",");
  			foreach (var col in sorted) {
          		console.Write(col + ", ");
       		}
       		console.WriteLine();
       		foreach (var row in sorted) {
          		console.Write(row + ", ");
          		var cols = rows[row];
	  			foreach (var col in sorted) {
          			if (cols.ContainsKey(col)) {
          				console.Write("{0}", Math.Abs(cols[col]));
          			}
	          		console.Write(", ");
          		}
         		console.WriteLine();
            }
       		console.Close();
       		*/
       		
	      	using (var stream = new FileStream(file, FileMode.Create, FileAccess.Write))
	        {
	            var wb = new XSSFWorkbook();
	            var sheet = wb.CreateSheet(name);
	            var cH = wb.GetCreationHelper();
	            
	            var erowN = 0;
	            var ecellN = 0;
                var erow = sheet.CreateRow(erowN++);
                var cell = erow.CreateCell(ecellN++);
				cell.SetCellValue("");
				foreach (var col in sorted) {
	                cell = erow.CreateCell(ecellN++);
					cell.SetCellValue(col);
	       		}

				foreach (var row in sorted) {
                	erow = sheet.CreateRow(erowN++);
                	ecellN = 0;
	                cell = erow.CreateCell(ecellN++);
					cell.SetCellValue(row);

					var cols = rows[row];
	  				foreach (var col in sorted) {
	          			if (cols.ContainsKey(col)) {
			                cell = erow.CreateCell(ecellN++);
			                cell.SetCellValue(cols[col]);
    	      			}
    	      		}
    	        }
	            wb.Write(stream);
	        }
        }
    }
}
#endif
