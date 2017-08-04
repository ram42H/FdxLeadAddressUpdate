using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace FdxLeadAddressUpdate
{
    [DataContractAttribute]
    public class API_PutResponse
    {
        [DataMember]
        public bool goNoGo { get; set; }

        [DataMember]
        public bool status { get; set; }
    }
}
