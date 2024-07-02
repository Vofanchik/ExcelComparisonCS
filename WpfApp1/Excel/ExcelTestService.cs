﻿using Business.Models;
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

		public MethodResult Compare(string expectedPath, string actualPath, string testName, string testOutputDirectory, string errorFileName, Int32 ws)
		{
			var expected = ExcelService.ReadFile(expectedPath);
			var actual = ExcelService.ReadFile(actualPath);
			var end = expected.ExcelPackage.Workbook.Worksheets[ws].Dimension.End;
			var info = new List<ExcelTest_CoordinateInfo> {
				//main data
				new ExcelTest_CoordinateInfo {
					Columns = end.Column,
					Rows = end.Row,
					ColumnOffset = this.ExcelService.ColumnNumber_FromColumnLetter("A"),
					RowOffset = 1
				}
			};
			return Compare(info, expected, actual, testName, testOutputDirectory, errorFileName, ws);
		}

		public MethodResult Compare(List<ExcelTest_CoordinateInfo> info, ExcelResult expected, ExcelResult actual, string testName, string testOutputDirectory, string errorFileName, Int32 ws)
		{
			var expectedWS = expected.ExcelPackage.Workbook.Worksheets[ws];
			var actualWS = actual.ExcelPackage.Workbook.Worksheets[ws];
			var mainDataComparisonResults = new List<ExcelComparison>();
			foreach (var i in info) {
				var xs = this.Compare(i, expectedWS, actualWS);
				mainDataComparisonResults.AddRange(xs);
			}
			var isOk = mainDataComparisonResults.All(a => a.ExpectedAndActual_AreEqual);
			if (isOk) { return MethodResult.Ok(Status.Success, "Листы одинаковые."); }
			var message = new StringBuilder();
			var errorCount = mainDataComparisonResults.Where(a => !a.ExpectedAndActual_AreEqual).Count();
			var fullCount = mainDataComparisonResults.Count();
			var percentFailed = Math.Round(((double)errorCount / fullCount) * 100);
			message.AppendLine($"All numeric cells are expected to be identical. { percentFailed }% ({ errorCount } of { fullCount }) error rate.");
			foreach (var excelComparisonResult in mainDataComparisonResults.Where(a => !a.ExpectedAndActual_AreEqual)) {
				excelComparisonResult.Message = $"{ this.ExcelService.ColumnLetter_FromColumnNumber(excelComparisonResult.Column) }{ excelComparisonResult.Row } До: { excelComparisonResult.Expected } После: { excelComparisonResult.Actual }";
				message.AppendLine(excelComparisonResult.Message);
			}
			var _message = message.ToString();
			this.CreateOutputWorksheet(mainDataComparisonResults, actual, testName, testOutputDirectory, errorFileName, _message, ws);
			return MethodResult.Fail($"Листы разные. отличие на: { percentFailed }%.\r\nДетали здесь: { testOutputDirectory }/{ errorFileName }");
		}

		public List<ExcelComparison> Compare(ExcelTest_CoordinateInfo info, ExcelWorksheet expected, ExcelWorksheet actual)
		{
			var result = new List<ExcelComparison>();
			for (int i = 0; i < info.Rows; i++) {
				var row = info.RowOffset + i;
				for (int j = 0; j < info.Columns; j++) {
					var col = info.ColumnOffset + j;
					var accepted = expected.Cells[row, col]?.Value?.ToString();
					if (accepted == _ignoreString) {
						continue;
					}
					var working = actual.Cells[row, col]?.Value?.ToString();
					result.Add(new ExcelComparison
					{
						Row = row,
						Column = col,
						Expected = accepted,
						Actual = working
					});
				}
			}
			return result;
		}

		public async Task Compare(string acceptanceDirectory, string testName, string testOutputDirectory, List<ExcelTest_CoordinateInfo> info, Task<ExcelResult> func, Int32 ws)
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
					this.Compare(info, expected, actual, testName, testOutputDirectory, errorFileName, ws);
				}
			}
		}

		public async Task Compare(string acceptanceDirectory, string testName, string testOutputDirectory, List<ExcelTest_CoordinateInfo> info, Task<Envelope<ExcelResult>> func , Int32 ws)
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
				this.Compare(info, expected, actual.Result, testName, testOutputDirectory, errorFileName, ws);
			}
			actual.Result.Dispose();
		}

		public void CreateOutputWorksheet(List<ExcelComparison> excelComparisonResult, ExcelResult actual, string testName, string testOutputDirectory, string errorFilename, string message, Int32 ws)
		{
			var copyResult = this.ExcelService.CopyExcelFile(actual, errorFilename);
			//gotta read it again, because the stream will be closed
			using (var errorFile = this.ExcelService.ReadFile(copyResult.Filename, copyResult.Directory)) {
				var errorWS = errorFile.ExcelPackage.Workbook.Worksheets[ws];
				foreach (var item in excelComparisonResult.Where(a => !a.ExpectedAndActual_AreEqual)) {
					this.ExcelService.Format_RedBackground_BlackText(errorWS.Cells[item.Row, item.Column], item.Message);
				}
				//errorWS.Cells[1, 1].Value = message;
				//this.ExcelService.AutoFit_All_Columns(errorWS);
				this.ExcelService.Save(errorFile);
			}
		}
	}
}