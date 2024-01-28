using ClosedXML.Excel;

namespace CladN
{
    public class Client
    {
        public int RowNumber;
        public int Code {get;}
        public string Name {get;}
        public string Adress {get;}
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
            workSheet.Columns().AdjustToContents();
            PersonName = personName;
        }
    }
}
