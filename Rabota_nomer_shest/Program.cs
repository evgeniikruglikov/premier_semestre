using System;

class Program
{
    static void Main(string[] args)
    {
        double k0, ke, dk, p, r;

        Console.WriteLine("Введите значения k0, ke, шаг dk и p, r:");

        // Ввод и проверка данных
        while (true)
        {
            if (double.TryParse(Console.ReadLine(), out k0) &&
                double.TryParse(Console.ReadLine(), out ke) &&
                double.TryParse(Console.ReadLine(), out dk) &&
                double.TryParse(Console.ReadLine(), out p) &&
                double.TryParse(Console.ReadLine(), out r))
            {
                break; // Если данные введены корректно, выходим из цикла ввода
            }

            Console.WriteLine("Некорректный ввод. Повторите попытку.");
            Console.WriteLine("Введите значения k0, ke, шаг dk и p, r:");
        }

        Calculate(k0, ke, dk, p, r);
    }

    static void Calculate(double k0, double ke, double dk, double p, double r)
    {
        for (double k = k0; k <= ke; k += dk)
        {
            if ((Math.Sqrt(p + (1 / Math.Cos(k)) - r)) >= 0)  
            {   
                if (1.0 / Math.Tan(k) != 0)
                {
                    double g = Math.Pow(p, 3) - ((Math.Sqrt(p + (1 / Math.Cos(k)) - r)) / (1.0 / Math.Tan(k)));  
                    Console.WriteLine($"При k = {k}, g = {g}");
                    
                }
            }
            else
            {
                Console.WriteLine($"При k = {k}, Выколотая точка");
            }
        }
        Console.ReadKey(true);
    }
}