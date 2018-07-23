using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemoCode
{
    class Program
    {
        static void Main(string[] args)
        {
            Rectangle rectangle = new Rectangle();
            rectangle.SetHeight(5);
            rectangle.SetWidth(3);
            Console.WriteLine(rectangle.GetWidth()*rectangle.GetHeight());

            Sqaure sqaure = new Sqaure();
            sqaure.SetHeight(5);
            sqaure.SetWidth(3);
            Console.WriteLine(sqaure.GetWidth() * sqaure.GetHeight());

            Console.ReadLine();

        }
    }
}
