using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProtocolsToPivotTable {
	class ItemProtocol {
		public string ProtocolName { get; set; }
		public List<ItemGroup> ListOfItemGroup { get; set; }
		public int PlaceId { get; set; }

		public ItemProtocol() {
			ProtocolName = "";
			ListOfItemGroup = new List<ItemGroup>();
			PlaceId = 0;
		}
	}
}
