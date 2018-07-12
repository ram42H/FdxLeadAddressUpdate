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

        [DataMember]
        public bool status { get; set; }

        [DataMember]
        public decimal prspectScore { get; set; }

        [DataMember]
        public decimal prospectPercentile { get; set; }

        [DataMember]
        public decimal prospectPriority { get; set; }

        [DataMember]
        public string prospectGroup { get; set; }

        [DataMember]
        public string rateSource { get; set; }

        [DataMember]
        public decimal pprRate { get; set; }

        [DataMember]
        public decimal subRate { get; set; }

        [DataMember]
        public int prospectRadius { get; set; }

        [DataMember]
        public string priceListName { get; set; }

        [DataMember]
        public string prospectScoreBlankMessage { get; set; }
    }
}
