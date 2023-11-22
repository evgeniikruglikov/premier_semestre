using System;

namespace rabotanomerchetire
{
    class Programm
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Введите первое число:");
            int a = Convert.ToInt32(Console.ReadLine());

            Console.WriteLine("Введите второе число:");
            int b = Convert.ToInt32(Console.ReadLine());

            Console.WriteLine("Введите третье число:");
            int c = Convert.ToInt32(Console.ReadLine());

            if (a < b)
            {
                if (c > b)
                    Console.WriteLine("{0}, {1}, {2}", c, b, a);

                else
                {
                    if (a < c)
                        Console.WriteLine("{0}, {1}, {2}", b, c, a);

                    else
                        Console.WriteLine("{0}, {1}, {2}", b, a, c);
                }
            }

            else
            {
                if (c < b)
                    Console.WriteLine("{0}, {1}, {2}", a, b, c);
                else
                {
                    if (a < c)
                        Console.WriteLine("{0}, {1}, {2}", c, a, b);

                    else
                        Console.WriteLine("{0}, {1}, {2}", a, c, b);

                }
            }

            Console.Read();
        }
    }
}