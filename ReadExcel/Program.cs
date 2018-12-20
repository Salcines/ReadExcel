using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadExcel
{
	class Program
	{
		static void Main(string[] args)
		{
			Excel.Application application = new Excel.Application();
			Excel.Workbook workbook = application.Workbooks.Open(@"prueba_VS.xlsx");
			Excel.Worksheet worksheet = workbook.Sheets[1];
			Excel.Range range = worksheet.UsedRange;

			int rowCount = range.Rows.Count;
			int colCount = range.Columns.Count;

			//write the data by screen. In Excel the row and column begins in 1
			for (int i = 1; i < rowCount; i++)
			{
				for (int j = 1; j < colCount; j++)
				{
					if (j == 1)
					{
						Console.Write("\r\n");
					}

					if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
					{
						Console.Write($"{ range[i, j].Value2.ToString()} \t");
					}
				}
			}

			//clean
			GC.Collect();
			GC.WaitForPendingFinalizers();

			/*
			rule of thumb for release COM objects
			never use two dots. The resources must be referenced and release individually
			ex: [something].[something].[something] is bad
			*/

			//release COM objects to fullfy kill the excel process from running in background
			Marshal.ReleaseComObject(range);
			Marshal.ReleaseComObject(worksheet);

			Console.WriteLine("\n\nRelease objects");

			//close and release
			workbook.Close();
			Marshal.ReleaseComObject(workbook);

			Console.WriteLine("Close File");

			//quit and release
			application.Quit();
			Marshal.ReleaseComObject(application);

			Console.WriteLine("Quit Excel");

			Console.ReadKey();


		
		}
	}
}
