using System;

namespace ExcelTransformation
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter file url: ");
            string fileUrl = Console.ReadLine();

            XlsTransformator transformator = new XlsTransformator();
            transformator.Transform(fileUrl);
        }
    }
}