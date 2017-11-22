using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProtocolsToPivotTable {
	class ExcelReader {
		public static ItemProtocol ReadProtocol(string fileName) {
			Console.WriteLine("ReadProtocol: " + fileName);
			DataTable dataTable = ReadExcelFile(fileName, "");
			DataRowCollection rows = dataTable.Rows;
			IdGenerator idGenerator = new IdGenerator();

			bool isMainBlockStarted = false;
			ItemProtocol itemProtocol = new ItemProtocol();
			for (int rowGroup = 0; rowGroup < rows.Count; rowGroup++) {
				Console.WriteLine("rowGroup: " + rowGroup);
				string groupName = rows[rowGroup][0].ToString();
				if (string.IsNullOrEmpty(groupName))
					continue;

				if (!isMainBlockStarted) {
					if (groupName.ToLower().StartsWith("жалобы (со слов пациента):")) {
						isMainBlockStarted = true;
						itemProtocol.ProtocolName = rows[rowGroup][2].ToString();
						itemProtocol.PlaceId = idGenerator.Id;
					}

					continue;
				}

				ItemGroup itemGroup = new ItemGroup() {
					GroupName = groupName,
					GroupId = idGenerator.Id
				};
				for (int rowParameter = rowGroup; rowParameter < rows.Count; rowParameter++) {
					Console.WriteLine("rowParameter: " + rowParameter);
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
						Console.WriteLine("rowDictionary: " + rowDictionary);
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

		private static DataTable ReadExcelFile(string fileFullPath, string sheetName) {
			DataTable dataTable = new DataTable();

			if (!File.Exists(fileFullPath))
				return dataTable;

			try {
				using (OleDbConnection conn = new OleDbConnection()) {
					conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileFullPath + ";" +
						"Extended Properties='Excel 12.0 Xml;HDR=NO;'";

					using (OleDbCommand comm = new OleDbCommand()) {
						if (string.IsNullOrEmpty(sheetName)) {
							conn.Open();
							DataTable dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, 
								new object[] { null, null, null, "TABLE" });
							sheetName = dtSchema.Rows[0].Field<string>("TABLE_NAME");
							conn.Close();
						} else
							sheetName += "$";

						comm.CommandText = "Select * from [" + sheetName + "]";
						comm.Connection = conn;

						using (OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter()) {
							oleDbDataAdapter.SelectCommand = comm;
							oleDbDataAdapter.Fill(dataTable);
						}
					}
				}
			} catch (Exception e) {
				Console.WriteLine(e.Message + " " + e.StackTrace);
			}

			return dataTable;
		}

	}
}
