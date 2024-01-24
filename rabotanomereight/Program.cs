using System;

class Program
{
    static void Main()
    {
        int n;
        bool isValidInput = false;
        double[] array = new double[n];
        double sum = 0;

        // Вводим значение n до тех пор, пока оно не будет корректным
        do
        {
            Console.Write("Введите количество элементов одномерного массива: ");
            string input = Console.ReadLine();

            isValidInput = int.TryParse(input, out n);

            if (!isValidInput)
            {
                Console.WriteLine("Ошибка: введите корректное целое числовое значение.");
            }
        } 
        while (!isValidInput);


        // Заполняем массив значениями и находим сумму положительных элементов
        for (int i = 0; i < n; i++)
        {
            bool isValidNumber = false;

            do
            {
                Console.Write($"Введите {i + 1}-й элемент одномерного массива: ");
                string input = Console.ReadLine();

                isValidNumber = double.TryParse(input, out array[i]);

                if (!isValidNumber)
                {
                    Console.WriteLine("Введите корректное числовое значение.");
                }
            } 
            while (!isValidNumber);

            if (array[i] > 0)
            {
                sum += array[i];
            }
        }

        Console.WriteLine($"Сумма положительных элементов: {sum}");
        Console.ReadKey(true);
    }
}
