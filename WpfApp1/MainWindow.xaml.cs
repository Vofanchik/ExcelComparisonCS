using Business.Common;
using System;
using System.IO;
using System.Windows;
using Tests.Services;

namespace WpfApp1
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();
		}

		private void btnGetFile1(object sender, RoutedEventArgs e)
		{
			var filename = getFilePath();
			txtFileName1.Text = filename; 
		}

		private void btnGetFile2(object sender, RoutedEventArgs e)
		{
			var filename = getFilePath();
			txtFileName2.Text = filename;
		}

		private string getFilePath()
		{
			// Configure open file dialog box
			var dlg = new Microsoft.Win32.OpenFileDialog();
			dlg.FileName = "Document"; // Default file name
			dlg.DefaultExt = ".xlsx"; // Default file extension
			dlg.Filter = "Excel documents (.xlsx)|*.xlsx"; // Filter files by extension

			// Show open file dialog box
			var result = dlg.ShowDialog();

			// Process open file dialog box results
			if (result == true) {
				// Open document
				return dlg.FileName;
			}
			return string.Empty;
		}

		private void btnCompare(object sender, RoutedEventArgs e)
		{
			if (string.IsNullOrEmpty(txtFileName1.Text)) {
				MessageBox.Show("Первое поле пустое.");
			}
			if (string.IsNullOrEmpty(txtFileName2.Text)) {
				MessageBox.Show("Второе поле пустое.");
			}
			compare(txtFileName1.Text, txtFileName2.Text, workSheetNumber.SelectedItem.ToString());
		}

		private void compare(string file1, string file2, string workSheetNumber)
		{
			var workSheet = Int32.Parse(workSheetNumber);
            var dirName = Path.GetDirectoryName(file1);
			var excelTestService = new ExcelTestService(dirName);
			var result = excelTestService.Compare(file1, file2, "Compare", dirName, "compare_errors.xlsx", workSheet);
			MessageBox.Show(result.Message);
		}


        private void ComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {

        }

        private void txtFileName1_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {

        }
    }
}