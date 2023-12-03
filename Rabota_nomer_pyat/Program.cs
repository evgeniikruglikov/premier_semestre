using System;

namespace Rabota_nomer_pyat
{
    class Program
    {
        static void Main()
        {
            Console.WriteLine("Введите символ арифметического оператора:");
            string input = Console.ReadLine();

            
            char oper = input[0];
            switch (oper)
            {
                case '+':
                    Console.WriteLine("Символ '+' соответствует оператору сложения.");
                    break;
                case '-':
                    Console.WriteLine("Символ '-' соответствует оператору вычитания.");
                    break;
                case '*':
                    Console.WriteLine("Символ '*' соответствует оператору умножения.");
                    break;
                case '/':
                    Console.WriteLine("Символ '/' соответствует оператору деления.");
                    break;
                case '%':
                    Console.WriteLine("Символ '%' соответствует оператору остатка от деления.");
                    break;
                default:
                    Console.WriteLine("Введен некорректный символ оператора.");
                    break;
            }
            

            Console.ReadLine();
        }
    }
}