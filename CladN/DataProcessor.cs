using ClosedXML.Excel;


namespace CladN
{
    class DataProcessor
    {
        private XLWorkbook _workbook;
        public List<Product> productList;
        public List<Client> clientList;
        public List<Request> requestList;
        dynamic Number;


        public DataProcessor(string filePath)
        {
            _workbook = new XLWorkbook(filePath);
            InitialiseOrUpdate();
        }

        public void SaveFile()
        {
            _workbook.Save();
        }

        public void InitialiseOrUpdate()
        {
            GetProducts(_workbook.Worksheet(1));
            GetClients(_workbook.Worksheet(2));
            GetRequests(_workbook.Worksheet(3));
        }

        #region получение и запись данных для Товаров, Клиентов, Запросов в соответствующие List<objectName>
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
        #endregion

        /// <summary>
        /// Получение клиента по его имени
        /// </summary>
        /// <param name="clientName"></param>
        /// <returns></returns>
        public Client GetClientByName(string clientName)
        {
            return clientList.FirstOrDefault(p => p.Name.ToUpper() == clientName.ToUpper());
        }
        /// <summary>
        /// Получение продукта по его имени
        /// </summary>
        /// <param name="productName"></param>
        /// <returns></returns>
        public Product GeProductByName(string productName)
        {
            return productList.FirstOrDefault(p => p.Name.ToUpper() == productName.ToUpper());
        }

        /// <summary>
        /// Отображение данных Клиентов, Заявок и Продукта по имени продукта
        /// </summary>
        /// <param name="productName"></param>
        public void DisplayCustomersByProduct(string productName)
        {
            var product = GeProductByName(productName);
            var reuests = requestList.Where(r => r.CodeProduct == product?.Code);
            clientList.ForEach(c =>
            {
                foreach (var r in reuests)
                {
                    if (r?.CodeClient == c.Code)
                    {
                        Console.WriteLine($"Информация о заказчике: {c.GetInfo()}\n Дата заказа: {r.DatePost} Кол-во: {r.CountNeeded} Цена за шт.: {product.Price}");
                    }
                }

            });
        }
        /// <summary>
        /// Изменение контактного лица клиента по имени клиента
        /// </summary>
        /// <param name="organizationName"></param>
        /// <param name="newContactPerson"></param>
        public void ChangeContactPerson(string organizationName, string newContactPerson)
        {
            GetClientByName(organizationName)?.SetPersonName(personName: newContactPerson, _workbook.Worksheet(2));
            SaveFile();
            Console.WriteLine("Успешно изменено");
        }
        /// <summary>
        /// Отображение золотого клиента(больше всего заказов в указанный год, месяц, среди прочих клиентов)
        /// </summary>
        /// <param name="year"></param>
        /// <param name="month"></param>
        public void DisplayGoldenCustomer(int year, int month)
        {
            var requestsByDate = requestList.Where(p => p.DatePost.Month == month && p.DatePost.Year == year);

            var orderCounts = requestsByDate.GroupBy(order => order.CodeClient).ToDictionary(group => group.Key, group => group.Count());

            int maxOrderClientCode = orderCounts.FirstOrDefault(x => x.Value == orderCounts.Values.Max()).Key;

            Client maxOrderClient = clientList.FirstOrDefault(cliente => cliente.Code == maxOrderClientCode);
            Console.WriteLine(maxOrderClient.Name);

        }
    }
}
