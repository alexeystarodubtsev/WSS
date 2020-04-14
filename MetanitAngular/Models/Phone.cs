using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MetanitAngular.Models
{
    public class Phone
    {
        public Dictionary<string, List<OneCall>> stages;
        public Phone()
        {
            stages = new Dictionary<string, List<OneCall>>();
        }
        public void AddCall(FullCall call)
        {
            if (!stages.ContainsKey(call.stage.ToUpper()))
            {
                stages[call.stage.ToUpper()] = new List<OneCall>();
            }
            stages[call.stage.ToUpper()].Add(new OneCall(call));
        } 
    }
}
