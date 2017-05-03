using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace FdxLeadAssignmentPlugin
{
    [DataContractAttribute]    
    public class Lead
    {
        [DataMember]
        public string goldMineId { get; set; }

        [DataMember]
        public bool goNoGo { get; set; }
    }
}
