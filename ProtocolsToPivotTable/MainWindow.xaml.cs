using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Diagnostics;

namespace ProtocolsToPivotTable {
	/// <summary>
	/// Логика взаимодействия для MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window {
		public MainWindow() {
			InitializeComponent();
		}

		private void Button_Click(object sender, RoutedEventArgs e) {
			OpenFileDialog openFileDialog = new OpenFileDialog();
			openFileDialog.Filter = "Книга Excel (*.xls*)|*.xls*";
			openFileDialog.CheckFileExists = true;
			openFileDialog.CheckPathExists = true;
			openFileDialog.Multiselect = false;
			openFileDialog.RestoreDirectory = true;

			if (openFileDialog.ShowDialog() == true)
				ParseExcelFile(openFileDialog.FileName);
		}

		private async void ParseExcelFile(string fileName) {
			Console.WriteLine("ParseExcelFile: " + fileName);

			Cursor = Cursors.Wait;
			buttonDoIt.IsEnabled = false;

			await Task.Run(() => {
				ItemProtocol itemProtocol = ExcelReader.ReadProtocol(fileName);
				XmlWriter.WriteToXml(itemProtocol);
			});

			buttonDoIt.IsEnabled = true;
			Cursor = Cursors.Arrow;

			MessageBox.Show("Завершено");
		}
	}
}
