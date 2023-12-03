using System;

namespace Rabota_nomer_shest
{
    class Program
    {
        static void Main()
        {
            double k, k0, ke, dk, p, r, g;

            Console.WriteLine("Введите значение k0:");
            k0 = Convert.ToDouble(Console.ReadLine());

            Console.WriteLine("Введите значение ke:");
            ke = Convert.ToDouble(Console.ReadLine());

            Console.WriteLine("Введите значение dk:");
            dk = Convert.ToDouble(Console.ReadLine());

            Console.WriteLine("Введите значение p:");
            p = Convert.ToDouble(Console.ReadLine());

            Console.WriteLine("Введите значение r:");
            r = Convert.ToDouble(Console.ReadLine());

            while (k0 < ke)
            {
                for (k = k0; k <= ke; k += dk)
                {
                    g = Math.Pow(p, 3) - ((Math.Sqrt(p + 1.0 / Math.Cos(k) - r)) / (Math.Cos(k) / Math.Sin(k)));
                    Console.WriteLine("При k = " + k + ", g  равняется " + g);
                    k0++;
                }
                break;
            } 

                Console.ReadLine();
        }
    }
}