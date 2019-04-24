using System;
using System.Collections.Generic;

namespace CustomDemoUsingClosedXML
{
    public class ResidentInfomation
    {
        public string Name { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string ZipCode { get; set; }

    }
    class Program
    {

        static void Main(string[] args)
        {
            var demo = new DemoClosedXML();

            demo.GetExcelFile();
        }
    }
}
