using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Bibliography;
using System.Globalization;

namespace CladN
{
    class DataProcessor
    {
        private string _filePath;
        private XLWorkbook _workbook;
        public List<Product> productList = new List<Product>();
        public List<Client> clientList = new List<Client>();
        public List<Request> requestList = new List<Request>();
        dynamic Number;


        public DataProcessor(string filePath)
        {
            _filePath = filePath;
            _workbook = new XLWorkbook(filePath);
        }
        public void SaveFile()
        {
            _workbook.Save();
        }
        public void GetProducts(IXLWorksheet productWorksheet)
        {
            productList = new List<Product>();
            var rows = productWorksheet.RangeUsed().RowsUsed();
            foreach (var row in rows)
            {
                productList.Add(new Product(
                    id: row.RowNumber(),
                    code: int.TryParse(row.Cell(1).Value.ToString(), out int c) ? c : 0,
                    name: row.Cell(2).Value.ToString(),
                    unit: row.Cell(3).Value.ToString(),
                    price: row.Cell(4).Value.IsNumber ? row.Cell(4).Value.GetNumber() : 0));
            }
        }
        public void GetClients(IXLWorksheet clientWorksheet)
        {
            clientList = new List<Client>();
            var rows = clientWorksheet.RangeUsed().RowsUsed();
            foreach (var row in rows)
            {
                clientList.Add(new Client(
                    id: row.RowNumber(),
                    code: int.TryParse(row.Cell(1).Value.ToString(), out int c) ? c : 0,
                    name: row.Cell(2).Value.ToString(),
                    adress: row.Cell(3).Value.ToString(),
                    personName: row.Cell(4).Value.ToString()));
            }
        }
        public void GetRequests(IXLWorksheet requestWorksheet)
        {
            requestList = new List<Request>();
            var rows = requestWorksheet.RangeUsed().RowsUsed();
            foreach (var row in rows)
            {
                requestList.Add(new Request(
                    id: row.RowNumber(),
                    code: int.TryParse(row.Cell(1).Value.ToString(), out int c) ? c : 0,
                    codeProduct: int.TryParse(row.Cell(2).Value.ToString(), out int cP) ? cP : 0,
                    codeClient: int.TryParse(row.Cell(3).Value.ToString(), out int cC) ? cC : 0,              //row.Cell(6).Value.IsNumber ? row.Cell(6).Value.GetNumber(): 0
                    requestNubmer: int.TryParse(row.Cell(4).Value.ToString(), out int rN) ? rN : 0,
                    countNeeded: int.TryParse(row.Cell(5).Value.ToString(), out int cN) ? cN : 0,
                    datePost: row.Cell(6).Value.IsDateTime ? row.Cell(6).Value.GetDateTime() : new DateTime()));
            }
        }
        public Client GetClientByName(string clientName)
        {
            GetClients(_workbook.Worksheet(2));
            return clientList.FirstOrDefault(p => p.Name.ToUpper() == clientName.ToUpper());
        }
        public Product GeProductByName(string productName)
        {
            GetProducts(_workbook.Worksheet(1));
            return productList.FirstOrDefault(p => p.Name.ToUpper() == productName.ToUpper());
        }
        public void DisplayCustomersByProduct(string productName)
        {
            GetClients(_workbook.Worksheet(2));
            GetRequests(_workbook.Worksheet(3));


            var product = GeProductByName(productName);
            var re = requestList.Where(r => r.CodeProduct == product?.Code);
            clientList.ForEach(c =>
            {
                foreach (var r in re)
                {
                    if (r?.CodeClient == c.Code)
                    {
                        Console.WriteLine($"Информация о заказчике: {c.GetInfo()}\n Дата заказа: {r.DatePost} Кол-во: {r.CountNeeded} Цена за шт.: {product.Price}");
                    }
                }

            });
        }

        public void ChangeContactPerson(string organizationName, string newContactPerson)
        {
            GetClientByName(organizationName)?.SetPersonName(personName: newContactPerson, _workbook.Worksheet(2));
            SaveFile();
            Console.WriteLine("Успешно изменено");
        }

        public void DisplayGoldenCustomer(int year, int month)
        {
            GetClients(_workbook.Worksheet(2));
            GetRequests(_workbook.Worksheet(3));

            var requestsByDate = requestList.Where(p => p.DatePost.Month == month && p.DatePost.Year == year);

            //clientList.Where(c =>
            //{
            //    foreach (var r in requestsByDate)
            //    {
            //        if (r.CodeClient == c.Code)
            //        {
            //            return true;
            //        }
            //    }
            //    return false;
            //
            //});
            var orderCounts = requestsByDate.GroupBy(order => order.CodeClient).ToDictionary(group => group.Key, group => group.Count());

            int maxOrderClientCode = orderCounts.FirstOrDefault(x => x.Value == orderCounts.Values.Max()).Key;

            Client maxOrderClient = clientList.FirstOrDefault(cliente => cliente.Code == maxOrderClientCode);
            Console.WriteLine(maxOrderClient.Name);

        }
    }
}
