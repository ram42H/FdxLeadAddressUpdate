using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FdxLeadAddressUpdate
{
    class ProspectData
    {
        public Guid? ProspectGroupId { get; set; }

        public string ProspectGroupName { get; set; }

        public Guid? PriceListId { get; set; }

        public string PriceListName { get; set; }

        public decimal? Priority { get; set; }

        public decimal? Score { get; set; }

        public decimal? Percentile { get; set; }

        public string RateSource { get; set; }

        public decimal? PPRRate { get; set; }

        public decimal? SubRate { get; set; }

        public int? Radius { get; set; }

        public DateTime? LastUpdated { get; set; }
    }
}
