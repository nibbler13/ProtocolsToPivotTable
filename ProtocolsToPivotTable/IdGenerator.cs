using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProtocolsToPivotTable {
	class IdGenerator {
		private int id;
		public int Id {
			get {
				return id++;
			}
			private set {
				id = value;
			}
		}

		public IdGenerator() {
			id = 660000000;
		}
	}
}
