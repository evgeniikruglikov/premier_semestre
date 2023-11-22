using System;

namespace rabotanomertri
{
    class Program
    {
        static void Main(string[] args)
        {
            double a, b, m, x, y, z;

            bool flag1, flag2, flag3;

            Console.Write("Введите значение переменной m: ");
            flag1 = Double.TryParse(Console.ReadLine(), out m);

            Console.Write("Введите значение переменной b: ");
            flag2 = Double.TryParse(Console.ReadLine(), out b);

            Console.Write("Введите значение переменной x: ");
            flag3 = Double.TryParse(Console.ReadLine(), out x);

            if (flag1 && flag2 && flag3)
            {
                if (m == b)
                {
                    a = m;
                    Console.WriteLine("Первое условие а" + a);
                }

                else
                {
                    a = Math.Exp(m + b);
                    Console.WriteLine("Второе условие а" + a);
                }

                if (a == 5 * b)
                {
                    y = Math.Sin(a) + Math.Tan(b);
                    Console.WriteLine("Второе условие у" + y);
                }

                if (a > 5 * b)
                {
                    y = a - 5;
                    Console.WriteLine("Первое условие у" + y);
                }

                else
                {
                    y = a * Math.Cos(a);
                    Console.WriteLine("Третье условие у" + y);
                }

                z = y * Math.Cos(y) + x * Math.Sin(y) + Math.Sqrt(Math.Pow(x, 2) - b);
                Console.WriteLine(y * Math.Cos(y));
                Console.WriteLine(x * Math.Sin(y));
                Console.WriteLine(Math.Sqrt(Math.Pow(x, 2) - b));

                if (flag1 && flag2 && flag3)
                {
                    Console.WriteLine("Значение арифметического выражения: " + z);
                }

                else
                {
                    Console.WriteLine("Невозможно решить.");
                }
            }

            else
            {
                Console.WriteLine("Введены некорректные данные.");
            }

            Console.ReadKey(true);
        }
    }
}