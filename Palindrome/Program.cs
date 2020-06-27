using System;
using System.Collections.Generic;

namespace Palindrome
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> UserInputs = new List<string>();
            Console.WriteLine("Enter list items and type exit");
            while (true)
            {
                var input = Console.ReadLine();
                if (input.Equals("exit", StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }
                UserInputs.Add(input);
            }

            foreach (string item in UserInputs)
            {
                if (IsPalindrome(item))
                    Console.WriteLine(item+ " is Palindromes");
            }
            Console.ReadKey();
        }
        public static bool IsPalindrome(string value)
        {
            int min = 0;
            int max = value.Length - 1;
            while (true)
            {
                if (min > max)
                {
                    return true;
                }
                char a = value[min];
                char b = value[max];
                if (char.ToLower(a) != char.ToLower(b))
                {
                    return false;
                }
                min++;
                max--;
            }
        }
    }
}
