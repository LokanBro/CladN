using System;
using CladN;
using System.Collections.Generic;
using System.Linq;
class Program
{
    static void Main()
    {
        Console.WriteLine("Добро пожаловать в приложение управления данными!");

        string filePath = RequestFilePath();

        // Создаем экземпляр класса для работы с данными
        DataProcessor dataProcessor = new DataProcessor(filePath);

        while (true)
        {
            Console.WriteLine("\nВыберите команду:");
            Console.WriteLine("1. Вывести информацию о клиентах по товару");
            Console.WriteLine("2. Изменить контактное лицо клиента");
            Console.WriteLine("3. Определить золотого клиента");
            Console.WriteLine("4. Завершить программу");

            int choice = GetChoice(1, 4);

            switch (choice)
            {
                case 1:
                    Console.Write("Введите наименование товара: ");
                    string productName = Console.ReadLine();
                    dataProcessor.DisplayCustomersByProduct(productName);
                    break;

                case 2:
                    Console.Write("Введите название организации: ");
                    string organizationName = Console.ReadLine();
                    Console.Write("Введите ФИО нового контактного лица: ");
                    string newContactPerson = Console.ReadLine();
                    dataProcessor.ChangeContactPerson(organizationName, newContactPerson);
                    break;

                case 3:
                    Console.Write("Введите год: ");
                    int year = GetIntInput();
                    Console.Write("Введите месяц: ");
                    int month = GetChoice(1, 12);
                    dataProcessor.DisplayGoldenCustomer(year, month);
                    break;

                case 4:
                    Console.WriteLine("Программа завершена.");
                    return;
            }
        }
    }

    static string RequestFilePath()
    {
        Console.Write("Введите путь до файла с данными: ");
        return "I:\\Windows\\Downloads\\1.xlsx"; //Console.ReadLine();
    }

    static int GetIntInput()
    {
        int result;
        while (!int.TryParse(Console.ReadLine(), out result))
        {
            Console.WriteLine("Введите корректное целочисленное значение.");
        }
        return result;
    }

    static int GetChoice(int minValue, int maxValue)
    {
        int choice;
        do
        {
            Console.Write($"Введите значение от {minValue} до {maxValue}: ");
        } 
        while (!int.TryParse(Console.ReadLine(), out choice) || choice < minValue || choice > maxValue);
        return choice;
    }
}
