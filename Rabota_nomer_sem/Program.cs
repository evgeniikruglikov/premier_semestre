using System;

class Program
{
    static void Main()
    {
        int n;
        bool validInput = false;

        do
        {
            Console.Write("Введите число n (n > 1): ");
            string input = Console.ReadLine();

            validInput = int.TryParse(input, out n) && n > 1;

            if (!validInput)
            {
                Console.WriteLine("Некорректный ввод. Пожалуйста, повторите попытку.");
            }
        } while (!validInput);

        long sum = 0;
        for (int i = 1; i <= n; i++)
        {
            sum += Factorial(i);
        }

        Console.WriteLine($"Сумма факториалов чисел от 1 до {n} равна {sum}.");
        Console.ReadKey(true);

    }

    static long Factorial(int number)
    {
        long factorial = 1;
        for (int i = 2; i <= number; i++)
        {
            factorial *= i;
        }
        return factorial;
    }
}