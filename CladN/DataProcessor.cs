using ClosedXML.Excel;


namespace CladN
{
    class DataProcessor
    {
        private XLWorkbook _workbook;
        public HashSet<Product> productList;
        public List<Client> clientList;
        public List<Request> requestList;


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

        #region получение и запись данных для Товаров, Клиентов, Запросов в соответствующие коллекции
        public void GetProducts(IXLWorksheet productWorksheet)
        {
            productList = new HashSet<Product>();
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
                var product = productList?.FirstOrDefault(p => p.Code == (row.Cell(2).Value.IsNumber ? row.Cell(2).Value.GetNumber() : 0));
                var client = clientList?.FirstOrDefault(p => p.Code == (row.Cell(3).Value.IsNumber ? row.Cell(3).Value.GetNumber() : 0));
                requestList.Add(new Request(
                    id: row.RowNumber(),
                    code: int.TryParse(row.Cell(1).Value.ToString(), out int c) ? c : 0,
                    product: product,
                    client: client,
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
            var requests = requestList.Where(r => r.Product == product).ToList();
            requests.ForEach(r => Console.WriteLine($"Информация о заказчике: {r.Client.GetInfo()}\n Дата заказа: {r.DatePost} Кол-во: {r.CountNeeded} Цена за шт.: {product.Price}"));
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
            var requestsByDate = requestList
                .Where(p => p.DatePost.Month == month && p.DatePost.Year == year);

            var goldenClient = requestsByDate.GroupBy(r => r.Client)
                .Select(g => new { Client = g.Key, RequestCount = g.Count() })
                .OrderByDescending(t => t.RequestCount)
                .FirstOrDefault()?.Client;
            if (goldenClient != null)
                Console.WriteLine(goldenClient.Name);
            else
                Console.WriteLine("Золотой клиент не найден");

        }
    }
}
