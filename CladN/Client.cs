using ClosedXML.Excel;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CladN
{
    internal class Client
    {
        public int RowNumber;
        public int Code { get; set; }
        public string Name { get; set; }
        public string Adress { get; set; }
        public string PersonName { get; set; }

        public readonly int Id;

        public Client(int id, int code, string name, string adress, string personName)
        {
            Id = id;
            Code = code;
            Name = name;
            Adress = adress;
            PersonName = personName;
        }

        public string GetInfo() => $" {Name},\n {Adress},\n {PersonName}";

        public void SetPersonName(string personName, IXLWorksheet workSheet)
        {
            workSheet.Cell(Id, "D").Value = personName;
            //workSheet.Row(Id).Cell(5).InsertData(personName);
            workSheet.Columns().AdjustToContents();
            PersonName = personName;
        }
    }
}
