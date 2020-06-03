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
        public string DealState= "";
        public DateTime DateDeal = DateTime.MinValue;
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
            var NameStage = call.stage.ToUpper().Trim();
            Dictionary<string, string> OldNewStage = new Dictionary<string, string>();
            string oldNameStage = "Консультация по коллекции/Макет коллаж 17";
            string newName = "Консультация по коллекции/Макет коллаж Максимально баллов без учета расширения чека (19)";
            OldNewStage[oldNameStage.ToUpper()] = newName.ToUpper();
            oldNameStage = "КОНСУЛЬТАЦИЯ ПО КОЛЛЕКЦИИ / МАКЕТ КОЛЛАЖ МАКСИМАЛЬНО БАЛЛОВ БЕЗ УЧЕТА РАСШИРЕНИЯ ЧЕКА(19)";
            OldNewStage[oldNameStage.ToUpper()] = newName.ToUpper();
            oldNameStage = "Заказ 11";
            newName = "Заказ 13";
            OldNewStage[oldNameStage.ToUpper()] = newName.ToUpper();
            oldNameStage = "Предварительный просчет 17";
            newName = "Предварительный просчет 19";
            OldNewStage[oldNameStage.ToUpper()] = newName.ToUpper();
            oldNameStage = "Если нет оплаты 16";
            newName = "Если нет оплаты 18";
            OldNewStage[oldNameStage.ToUpper()] = newName.ToUpper();

            if (OldNewStage.ContainsKey(NameStage.Trim()))
               NameStage = OldNewStage[NameStage.Trim()];
            if (!stages.ContainsKey(NameStage))
            {
                stages[NameStage] = new List<OneCall>();
            }
            if (!Managers.Contains(call.Manager))
                Managers.Add(call.Manager);
            stages[NameStage].Add(new OneCall(call));
            DateTime DateAnalyze;
            if (call.phoneNumber == "8 (903) 279-23-46")
            {

            }
            if (call.StateDeal.ToUpper().Trim() == "В РАБОТЕ")
            {
                DateAnalyze = call.DateNextContact;
            }
            else
            {
                DateAnalyze = call.date;
            }
               
            if (DateAnalyze > DateDeal || (DealState.ToUpper().Trim() == "В РАБОТЕ" && DateAnalyze == DateDeal) || DealState == "")
            {
                DateDeal = DateAnalyze;
                DealState = call.StateDeal.Trim();
            }
        } 
    }
}
