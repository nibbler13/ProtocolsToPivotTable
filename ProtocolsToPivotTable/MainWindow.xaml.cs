﻿using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
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

			string path = Directory.GetCurrentDirectory() + "\\Results_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "\\";
			if (!Directory.Exists(path))
				Directory.CreateDirectory(path);

			await Task.Run(() => {
				List<ItemProtocol> protocols = ExcelReader.ReadProtocols(fileName);

				foreach (ItemProtocol itemProtocol in protocols)
					XmlWriter.WriteToXml(itemProtocol, path);
			});

			buttonDoIt.IsEnabled = true;
			Cursor = Cursors.Arrow;

			MessageBox.Show("Завершено");

			Process.Start(path);
		}
	}
}
