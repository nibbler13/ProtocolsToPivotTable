using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProtocolsToPivotTable {
	class ItemParameter {
		public string ParameterName { get; set; }
		public List<ItemDictionary> Dictionary { get; set; }
		public int CodeParams { get; set; }
		public int ReftToDict { get; set; }

		public ItemParameter() {
			ParameterName = "";
			Dictionary = new List<ItemDictionary>();
			CodeParams = 0;
			ReftToDict = 0;
		}
	}
}
