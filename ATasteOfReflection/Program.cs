using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ATasteOfReflection
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var reportModel = new ReportModel(); 
            reportModel.GenerateReportWithReflection(1234); //set JobId = 1234 here as an example
            
            
        }
          
    }
}


