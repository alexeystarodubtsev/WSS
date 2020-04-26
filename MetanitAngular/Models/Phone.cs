using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MetanitAngular.Models
{
    public class Phone
    {
        public Dictionary<string, List<OneCall>> stages;
        public string link { get; }
        public string phoneNumber { get; }
        public List<string> Managers = new List<string>();
        public Phone(string link, string PhoneNumber)
        {
            this.link = link;
            this.phoneNumber = PhoneNumber;
            stages = new Dictionary<string, List<OneCall>>();
        }
        public string GetManager()
        {
            string managers = "";
            foreach (var man in Managers)
            {
                managers += man + ", ";

            }
            managers = managers.Trim(' ').Trim(',');
            return managers;
        }
        public void AddCall(FullCall call)
        {
            if (!stages.ContainsKey(call.stage.ToUpper()))
            {
                stages[call.stage.ToUpper()] = new List<OneCall>();
            }
            if (!Managers.Contains(call.Manager))
                Managers.Add(call.Manager);
            stages[call.stage.ToUpper()].Add(new OneCall(call));
        } 
    }
}
