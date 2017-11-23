using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProtocolsToPivotTable {
	class XmlWriter {
		public static void WriteToXml(ItemProtocol itemProtocol) {
			System.Xml.XmlWriterSettings xmlWriterSettings = new System.Xml.XmlWriterSettings() {
				Indent = true,
				IndentChars = "  "
			};

			using(System.Xml.XmlWriter writer = System.Xml.XmlWriter.Create(itemProtocol.SheetName + ".xml", xmlWriterSettings)) {
				writer.WriteStartDocument();
				writer.WriteStartElement("WorkPlace");

				writer.WriteAttributeString("PLACEID", itemProtocol.PlaceId.ToString());
				writer.WriteAttributeString("NAMEPLACE", itemProtocol.ProtocolName);
				writer.WriteAttributeString("F25_NAMEPLACE", "Протокол телемедицинской консультации");
				writer.WriteAttributeString("PARENTPLACEID", "-1");
				writer.WriteAttributeString("PLACETYPE", "0");
				writer.WriteAttributeString("NAPRAVPRINTTYPE", "0");
				writer.WriteAttributeString("BOLPRINTTYPE", "0");
				writer.WriteAttributeString("DIAGPRINTTYPE", "0");
				writer.WriteAttributeString("RECIPEPRINTTYPE", "0");

				foreach (ItemGroup itemGroup in itemProtocol.ListOfItemGroup) {
					writer.WriteStartElement("Grp" + itemGroup.GroupId.ToString());

					writer.WriteAttributeString("GROUPID", itemGroup.GroupId.ToString());
					writer.WriteAttributeString("GROUPNAME", itemGroup.GroupName);
					writer.WriteAttributeString("F25_GROUPNAME", "");
					writer.WriteAttributeString("VIEWINWEB", "1");
					
					foreach (ItemParameter itemParameter in itemGroup.ListOfItemParameter) {
						writer.WriteStartElement("Prm" + itemParameter.CodeParams.ToString());

						writer.WriteAttributeString("CODEPARAMS", itemParameter.CodeParams.ToString());
						writer.WriteAttributeString("NAMEPARAMS", itemParameter.ParameterName + itemProtocol.Postfix);
						writer.WriteAttributeString("TYPEPARAMS", "6");
						writer.WriteAttributeString("REFTODICT", itemParameter.ReftToDict.ToString());
						writer.WriteAttributeString("COMMENT", "");
						writer.WriteAttributeString("DEFAULTVALUE", "");
						writer.WriteAttributeString("MTYPEID", "990000021");
						writer.WriteAttributeString("F25_STYLE_N", "1");
						writer.WriteAttributeString("F25_NAMEPARAMS", itemParameter.ParameterName);
						writer.WriteAttributeString("PARAMWRAP", "1");

						writer.WriteEndElement();
					}

					writer.WriteEndElement();
				}

				writer.WriteStartElement("Reference");

				foreach (ItemGroup itemGroup in itemProtocol.ListOfItemGroup) {
					foreach (ItemParameter itemParameter in itemGroup.ListOfItemParameter) {
						writer.WriteStartElement("Ref" + itemParameter.ReftToDict.ToString());

						writer.WriteAttributeString("REFID", itemParameter.ReftToDict.ToString());
						writer.WriteAttributeString("REFNAME", itemParameter.ParameterName + itemProtocol.Postfix);

						int order = 0;
						foreach (ItemDictionary itemDictionary in itemParameter.Dictionary) {
							writer.WriteStartElement("Dic" + itemDictionary.DicId.ToString());

							writer.WriteAttributeString("DICID", itemDictionary.DicId.ToString());
							writer.WriteAttributeString("PARENTDICID", "0");
							writer.WriteAttributeString("REFID", itemParameter.ReftToDict.ToString());
							writer.WriteAttributeString("DICNAME", itemDictionary.DictionaryValue);
							writer.WriteAttributeString("DICORDER", order.ToString());

							writer.WriteEndElement();

							order++;
						}

						writer.WriteEndElement();
					}
				}

				writer.WriteEndElement();

				writer.WriteStartElement("ParamNorm");
				writer.WriteEndElement();
				
				writer.WriteStartElement("Templates");
				writer.WriteEndElement();

				writer.WriteStartElement("ExportInfo");
				writer.WriteAttributeString("ComputerName", "MSCS-TSERVER5");
				writer.WriteAttributeString("WindowsUserName", "vyatkin");
				writer.WriteEndElement();

				writer.WriteEndElement();
				writer.WriteEndDocument();
			}
		}
	}
}
