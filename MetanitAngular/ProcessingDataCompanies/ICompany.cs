using ClosedXML.Excel;
using MetanitAngular.Models;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using static MetanitAngular.Excel.DataStructsForPrintCalls;

namespace MetanitAngular.ProcessingDataCompanies
{
    interface ICompany
    {
        void AddCall(FullCall call);
        List<CallIncoming> getIncomeWithoutOutGoing();
        List<CallPerWeek> getCallsPerWeek();
        List<CallOneStage> getCallsOneStage();
        List<CallPreAgreement> getCallsPreAgreement();
        List<firstCallsToClient> getfirstCallForBelfan();
        void FillStageDictionary(XLWorkbook wb);
        void ParserCheckLists(IEnumerable<IFormFile> files);
    }
}
