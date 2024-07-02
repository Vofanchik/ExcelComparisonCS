using Business.Models;
using Business.Services.Excel;
using Common.Models;
using Common.Transactions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tests.Services
{
	public class ExcelTestService
	{
		public ExcelService ExcelService { get; private set; }
		public string OutputDirectory { get; private set; }

		private string _ignoreString = "N/A";

		public ExcelTestService(string outputDirectory)
		{
			this.ExcelService = new ExcelService();
			this.OutputDirectory = outputDirectory;
		}

		public MethodResult Compare(string expectedPath, string actualPath, string testName, string testOutputDirectory, string errorFileName)
		{
			var expected = ExcelService.ReadFile(expectedPath);
			var actual = ExcelService.ReadFile(actualPath);
			//var end = expected.ExcelPackage.Workbook.Worksheets[1].Dimension.End;
			var info = new List<ExcelTest_CoordinateInfo>();
            for (int i = 1; i <= expected.ExcelPackage.Workbook.Worksheets.Count; i++)
			{
                if (expected.ExcelPackage.Workbook.Worksheets[i].Hidden==eWorkSheetHidden.Visible)
                {
					var start = expected.ExcelPackage.Workbook.Worksheets[i].Dimension.End;
					var end = actual.ExcelPackage.Workbook.Worksheets[i].Dimension.End;
					info.Add(new ExcelTest_CoordinateInfo
					{
						expectedColumns = start.Column,
						expectedRows = start.Row,
						actualColumns = end.Column,
						actualRows = end.Row,
						ColumnOffset = this.ExcelService.ColumnNumber_FromColumnLetter("A"),
						RowOffset = 1,
						Name = expected.ExcelPackage.Workbook.Worksheets[i].Name
					});
				}
            }
			return Compare(info, expected, actual, testName, testOutputDirectory, errorFileName);
		}

		public MethodResult Compare(List<ExcelTest_CoordinateInfo> info, ExcelResult expected, ExcelResult actual, string testName, string testOutputDirectory, string errorFileName)
		{
			var mainDataComparisonResults = new List<ExcelComparison>();
			int errorCountSum = 0;
			int fullCountSum = 0;
			double percentFailed = 0;
			bool fileCopied = false;
			ExcelResult copyResult = new ExcelResult();
			foreach (var i in info) {
				mainDataComparisonResults.Clear();
				var expectedWS = expected.ExcelPackage.Workbook.Worksheets[i.Name];
				var actualWS = actual.ExcelPackage.Workbook.Worksheets[i.Name];
				var xs = this.Compare(i, expectedWS, actualWS);
				mainDataComparisonResults.AddRange(xs);
				var flag = mainDataComparisonResults.All(a => a.ExpectedAndActual_AreEqual);

				if (!flag)
				{
                    if (!fileCopied)
                    {
						copyResult = this.ExcelService.CopyExcelFile(actual, errorFileName);
						fileCopied = true;
					}
					var message = new StringBuilder();
					var errorCount = mainDataComparisonResults.Where(a => !a.ExpectedAndActual_AreEqual).Count();
					var fullCount = mainDataComparisonResults.Count();
					errorCountSum += errorCount;
					fullCountSum += fullCount;
					percentFailed = Math.Round(((double)errorCount / fullCount) * 100);
					message.AppendLine($"All numeric cells are expected to be identical. { percentFailed }% ({ errorCount } of { fullCount }) error rate.");
					foreach (var excelComparisonResult in mainDataComparisonResults.Where(a => !a.ExpectedAndActual_AreEqual))
					{
						excelComparisonResult.Message = $"{ this.ExcelService.ColumnLetter_FromColumnNumber(excelComparisonResult.Column) }{ excelComparisonResult.Row } 旧: { excelComparisonResult.Expected } 新: { excelComparisonResult.Actual }";
						message.AppendLine(excelComparisonResult.Message);
					}
					var _message = message.ToString();
					this.CreateOutputWorksheet(mainDataComparisonResults, actual, testName, testOutputDirectory, errorFileName, _message,i.Name, copyResult);
				}
			}
			if (errorCountSum==0) { return MethodResult.Ok(Status.Success, "Files are the same."); }
			percentFailed = Math.Round(((double)errorCountSum / fullCountSum) * 100);
			return MethodResult.Fail($"Files are not the same. Failure: { percentFailed }%.\r\nDetails here: { testOutputDirectory }/{ errorFileName }");
		}

		public List<ExcelComparison> Compare(ExcelTest_CoordinateInfo info, ExcelWorksheet expected, ExcelWorksheet actual)
		{
			var result = new List<ExcelComparison>();
			int actualRow = info.RowOffset;
			for (int expectedRow = info.RowOffset; expectedRow < info.expectedRows; expectedRow++) {
				if (expected.Row(expectedRow).Hidden | expected.Cells[expectedRow, info.RowOffset, expectedRow, info.expectedRows].Count() == 0)
				{
					continue;
				}
				for (; actualRow < info.actualRows; actualRow++)
				{
					if (actual.Cells[actualRow, info.ColumnOffset, actualRow, info.actualColumns].Count(a=>a.Value != null) == 0)
					{
						continue;
					}
					int actualCol = info.ColumnOffset;
					for (int expectedCol = info.ColumnOffset; expectedCol < info.expectedColumns; expectedCol++)
					{
						if (expected.Column(expectedCol).Hidden | expected.Cells[info.ColumnOffset, expectedCol, info.actualColumns, expectedCol].Count() == 0)
						{
							continue;
						}
						var accepted = expected.Cells[expectedRow, expectedCol]?.Value?.ToString();
						for (; actualCol < info.actualColumns; actualCol++)
						{
							if (actual.Cells[info.ColumnOffset, actualCol, info.actualColumns, actualCol].Count() == 0)
							{
								continue;
							}
							var working = actual.Cells[actualRow, actualCol]?.Value?.ToString();
							result.Add(new ExcelComparison
							{
								Row = actualRow,
								Column = actualCol,
								Expected = accepted,
								Actual = working
							});
							actualCol++;
							break;
						}
					}
					actualRow++;
					break;
				}
			}
			return result;
		}

		public async Task Compare(string acceptanceDirectory, string testName, string testOutputDirectory, List<ExcelTest_CoordinateInfo> info, Task<ExcelResult> func)
		{
			var acceptanceFilename = $"{ testName }.xlsx";
			var errorFileName = $"{ testName }_Errors.xlsx";
			//read data, write to excel, save file for comparison if something doesn't compare correctly
			using (var actual = await func) {
				actual.Filename = $"{ testName }_{ actual.Filename }";
				actual.Directory = testOutputDirectory;
				//saving will close the stream
				this.ExcelService.Save(actual);
				//read acceptance file
				using (var expected = this.ExcelService.ReadFile(acceptanceFilename, acceptanceDirectory)) {
					this.Compare(info, expected, actual, testName, testOutputDirectory, errorFileName);
				}
			}
		}

		public async Task Compare(string acceptanceDirectory, string testName, string testOutputDirectory, List<ExcelTest_CoordinateInfo> info, Task<Envelope<ExcelResult>> func)
		{
			var acceptanceFilename = $"{ testName }.xlsx";
			var errorFileName = $"{ testName }_Errors.xlsx";
			//read data, write to excel, save file for comparison if something doesn't compare correctly
			var actual = await func;
			actual.Result.Filename = $"{ testName }_{ actual.Result.Filename }";
			actual.Result.Directory = testOutputDirectory;
			//saving will close the stream
			this.ExcelService.Save(actual.Result);
			//read acceptance file
			using (var expected = this.ExcelService.ReadFile(acceptanceFilename, acceptanceDirectory)) {
				this.Compare(info, expected, actual.Result, testName, testOutputDirectory, errorFileName);
			}
			actual.Result.Dispose();
		}

		public void CreateOutputWorksheet(List<ExcelComparison> excelComparisonResult, ExcelResult actual, string testName, string testOutputDirectory, string errorFilename, string message, string sheetName, ExcelResult copyResult)
		{
			//var copyResult = this.ExcelService.CopyExcelFile(actual, errorFilename);
			//gotta read it again, because the stream will be closed
			using (var errorFile = this.ExcelService.ReadFile(copyResult.Filename, copyResult.Directory)) {
				var errorWS = errorFile.ExcelPackage.Workbook.Worksheets[sheetName];
				foreach (var item in excelComparisonResult.Where(a => !a.ExpectedAndActual_AreEqual)) {
					this.ExcelService.Format_RedBackground_BlackText(errorWS.Cells[item.Row, item.Column], item.Message);
				}
				//errorWS.Cells[1, 1].Value = message;
				this.ExcelService.AutoFit_All_Columns(errorWS);
				this.ExcelService.Save(errorFile);
			}
		}
	}
}