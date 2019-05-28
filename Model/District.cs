using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Volby.Model
{
    public class District
    {
        public string Code { get; set; }
        public int TotalVoters { get; set; }
        public int Voted { get; set; }

        public List<PartyResult> PartyList { get; set; }
    }
}
