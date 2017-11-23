using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ProtocolsToPivotTable {
	class ExcelReader {
		public static List<ItemProtocol> ReadProtocols(string fileName) {
			Console.WriteLine("ReadProtocol: " + fileName);
			Dictionary<string, DataTable> dataTables = ReadExcelFile(fileName);
			IdGenerator idGenerator = new IdGenerator();
			List<ItemProtocol> protocols = new List<ItemProtocol>();

			foreach (KeyValuePair<string, DataTable> item in dataTables)
				protocols.Add(ReadProtocol(item, idGenerator));

			return protocols;
		}

		private static ItemProtocol ReadProtocol(KeyValuePair<string, DataTable> keyValuePair, IdGenerator idGenerator) {
			DataRowCollection rows = keyValuePair.Value.Rows;
			bool isMainBlockStarted = false;
			ItemProtocol itemProtocol = new ItemProtocol();

			for (int rowGroup = 0; rowGroup < rows.Count; rowGroup++) {
				string groupName = rows[rowGroup][0].ToString();
				if (string.IsNullOrEmpty(groupName))
					continue;

				if (!isMainBlockStarted) {
					if (groupName.ToLower().StartsWith("жалобы (со слов пациента):")) {
						isMainBlockStarted = true;
						itemProtocol.ProtocolName = rows[rowGroup][2].ToString();
						itemProtocol.PlaceId = idGenerator.Id;
						itemProtocol.SheetName = keyValuePair.Key;
						itemProtocol.Postfix = new String(keyValuePair.Key.Where(Char.IsDigit).ToArray()); 
					}

					continue;
				}

				ItemGroup itemGroup = new ItemGroup() {
					GroupName = groupName,
					GroupId = idGenerator.Id
				};

				for (int rowParameter = rowGroup; rowParameter < rows.Count; rowParameter++) {
					string groupNameInner = rows[rowParameter][0].ToString();

					if (!string.IsNullOrEmpty(groupNameInner) &&
						!groupNameInner.Equals(groupName)) {
						rowGroup = rowParameter - 1;
						break;
					}

					string parameterName = rows[rowParameter][1].ToString();
					if (string.IsNullOrEmpty(parameterName))
						parameterName = itemGroup.GroupName;

					ItemParameter itemParameter = new ItemParameter() {
						ParameterName = parameterName,
						CodeParams = idGenerator.Id,
						ReftToDict = idGenerator.Id
					};

					for (int rowDictionary = rowParameter; rowDictionary < rows.Count; rowDictionary++) {
						string parameterNameInner = rows[rowDictionary][1].ToString();
						string groupNameInner2 = rows[rowDictionary][0].ToString();

						if (!string.IsNullOrEmpty(groupNameInner2) &&
							!groupNameInner2.Equals(groupName)) {
							rowParameter = rowDictionary - 1;
							break;
						}

						if (!string.IsNullOrEmpty(parameterNameInner) &&
							!parameterNameInner.Equals(parameterName)) {
							rowParameter = rowDictionary - 1;
							break;
						}

						string dictionaryValue = rows[rowDictionary][2].ToString();
						if (string.IsNullOrEmpty(dictionaryValue))
							continue;
						
						Regex regex = new Regex("[ ]{2,}", RegexOptions.None);
						dictionaryValue = dictionaryValue.TrimStart(' ').TrimEnd(' ');
						dictionaryValue = regex.Replace(dictionaryValue, " ");
						dictionaryValue = dictionaryValue.Replace("\r\n", " ");
						dictionaryValue = dictionaryValue.Replace(Environment.NewLine, " ");
						dictionaryValue = dictionaryValue.Replace("\n", " ");

						itemParameter.Dictionary.Add(new ItemDictionary() {
							DictionaryValue = dictionaryValue,
							DicId = idGenerator.Id
						});

						if (rowDictionary == rows.Count - 1) {
							rowParameter = rowDictionary;
							break;
						}
					}

					if (itemParameter.Dictionary.Count > 0)
						itemGroup.ListOfItemParameter.Add(itemParameter);
				}

				itemProtocol.ListOfItemGroup.Add(itemGroup);
			}

			return itemProtocol;
		}

		private static Dictionary<string, DataTable> ReadExcelFile(string fileFullPath) {
			Dictionary<string, DataTable> dataTables = new Dictionary<string, DataTable>();

			if (!File.Exists(fileFullPath))
				return dataTables;

			try {
				using (OleDbConnection conn = new OleDbConnection()) {
					conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileFullPath + ";" +
						"Extended Properties='Excel 12.0 Xml;HDR=NO;'";

					using (OleDbCommand comm = new OleDbCommand()) {
						conn.Open();
						DataTable dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, 
							new object[] { null, null, null, "TABLE" });
						foreach (DataRow row in dtSchema.Rows) {
							string sheetName = row.Field<string>("TABLE_NAME");

							comm.CommandText = "Select * from [" + sheetName + "]";
							comm.Connection = conn;

							using (OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter()) {
								oleDbDataAdapter.SelectCommand = comm;
								DataTable dataTable = new DataTable();
								oleDbDataAdapter.Fill(dataTable);
								sheetName = sheetName.Replace("#", ".").Replace("$", "");
								dataTables.Add(sheetName, dataTable);
							}
						}

						conn.Close();
					}
				}
			} catch (Exception e) {
				Console.WriteLine(e.Message + " " + e.StackTrace);
			}

			return dataTables;
		}

	}
}
