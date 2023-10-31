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
                } 
                else
                { 
                    a = Math.Pow(Math.E, (m + b)); 
                } 
                if (a > 5 * b) 
                { 
                    y = a - 5; 
                } 
                if (a < 5 * b)
                {
                    y = Math.Sin(a) + Math.Tan(b);
                }
                else 
                { 
                    y = a * Math.Cos(a); 
                } 
                z = y * Math.Cos(y) + x * Math.Sin(y) + Math.Sqrt(Math.Pow(x, 2) - b); 
                if(flag1 && flag2 && flag3) 
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
