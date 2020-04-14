using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MetanitAngular.Models
{
    public class Phones
    {
        Dictionary<string, Phone> Calls;
        public Dictionary<string, int> Stages { get; set; }

        public Phones()
        {
            Calls = new Dictionary<string, Phone>();
            Stages = new Dictionary<string, int>();
        }
        public void AddCall(FullCall call)
        {
            if (!Calls.ContainsKey(call.phoneNumber))
            {
                Calls[call.phoneNumber] = new Phone();
            }
            Calls[call.phoneNumber].AddCall(call);
        }
        public Dictionary<string, Phone> getPhones()
        {
            return Calls;
        }
        
    }
}
