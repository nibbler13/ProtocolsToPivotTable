using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProtocolsToPivotTable {
	class ItemGroup {
		public string GroupName { get; set; }
		public List<ItemParameter> ListOfItemParameter { get; set; }
		public int GroupId { get; set; }

		public ItemGroup() {
			GroupName = "";
			ListOfItemParameter = new List<ItemParameter>();
			GroupId = 0;
		}
	}
}
