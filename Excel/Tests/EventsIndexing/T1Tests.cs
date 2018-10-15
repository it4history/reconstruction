#if DEBUG

using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using NPOI.HSSF.UserModel;
using NUnit.Framework;

namespace Logy.Api.Mw.Excel.Tests.EventsIndexing
{
    public class T1Tests
    {
		void Link(Dictionary<string, Dictionary<string, int>> rows, string row, KeyValuePair<string, int> column)
		{
			if (column.Value > 0) 
			{ // no linking for already linked when column.Value < 0
		   		var symmetricColumns = rows[column.Key];
				int symmetricValue;
				if (symmetricColumns.TryGetValue(row, out symmetricValue))
				    {
						if (symmetricValue < 0)
						{
							//symmetricColumns[row] -= column.Value;
							//throw new ApplicationException("impossible to write symmetric value twise");
						}
				    }
				else
					symmetricColumns[row] = column.Value;
			}
		}

        [Test]
        public void Symmetric3(){
			var byYears = new Dictionary<short, string>();
			byYears.Add(1, "a,b,c");
			byYears.Add(2, "b,a,c");
			var rows = Do(byYears);
			Assert.AreEqual(1, rows["a"]["b"]);        
			Assert.AreEqual(1, rows["a"]["c"]);
			Assert.AreEqual(1, rows["b"]["c"]);
        }

        [Test]
        public void Count2(){
			var byYears = new Dictionary<short, string>();
			byYears.Add(1, "b,a,b");
			var rows = Do(byYears);
			Assert.AreEqual(2, rows["b"]["a"]);
			Assert.AreEqual(2, rows["a"]["b"]);
        }
        
        [Test]
        public void Read()
        {
            var man = new ExcelManager(Path.Combine("Excel/Tests/EventsIndexing", "00_База_2018_10_01.xls"));
            man.Read();
            
            // group by year
            var byYears = new Dictionary<short, string>();
            foreach (HSSFRow row in man.Records)
            {
                var indices = man.GetValue(row, "Индекс");
                if (!string.IsNullOrEmpty(indices)) {
                	var year = short.Parse(man.GetValue(row, "-99000"));
                	byYears[year] = indices 
                		+ (byYears.ContainsKey(year) ? byYears[year] : null);
                }
            }
            
            var rows = Do(byYears);
            
			// output
            var console = File.CreateText("out.csv");
       		Console.WriteLine();
       		var sorted = from s in rows.Keys
						where rows[s].Count > 0
       					orderby s
						select s;
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
        }
        
        public Dictionary<string, Dictionary<string, int>> Do(Dictionary<short, string> byYears) {
            // fill rows and find columns names
            // rowName, columnName, links count
            var rows = new Dictionary<string, Dictionary<string, int>>();
            var rowsCoef = new Dictionary<string, int>(); 
            foreach (var indices in byYears.Values) {
            	string rowName = null;
	            Dictionary<string, int> columns = null;
	            foreach(var index in indices.Split(',')) {
            		if (!string.IsNullOrEmpty(index)){
	            		Dictionary<string, int> list;
            			if (rows.ContainsKey(index))
            			{
            				list=rows[index];
            			}
            			else{
            				list=new Dictionary<string, int>();
            				rows.Add(index, list);
            				rowsCoef.Add(index, 1);
            			}

	            		if (columns == null){
	            			// first index
	            			columns = list;
	            			rowName = index;
	            		}
	            		else{
	            			if (index!=rowName) // ignore self2self links
	            			{
	            				if(columns.ContainsKey(index))
	            				{
	            					columns[index]++;
	            				}else
	            					columns[index]=1;
	            			}
	            			else{
	            				rowsCoef[rowName]++;
	            			}
	            		}
            		}
	            }
            }
            
            /* increase coef
  			foreach (var pair in rowsCoef) {
            	if (pair.Value > 1) {
            		var cols = new List<string>(rows[pair.Key].Keys);
            		foreach (var col in cols) {
            			rows[pair.Key][col] = rows[pair.Key][col] * pair.Value;
            		}
            	}
            }*/
            
            // make complete graph
  			foreach (var pair in rows) {
          		var row = pair.Key;
	  			foreach (var colPair in pair.Value) {
            		var column = colPair.Key;
       				Link(rows, row, colPair);
		  			foreach (var p in pair.Value) {
            			if (p.Key != column) {
            				Link(rows, p.Key, colPair);
            			}
            		}
	            }
            }

            return rows;
        }
    }
}
#endif
