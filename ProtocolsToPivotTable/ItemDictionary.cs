using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProtocolsToPivotTable {
	class ItemDictionary {
		public string DictionaryValue { get; set; }
		public int DicId { get; set; }

		public ItemDictionary() {
			DictionaryValue = "";
			DicId = 0;
		}
	}
}
