using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordToPdfConverter
{
    public class Utilities
    {
        public static void ConsoleWriteLine(string message)
        {
            Console.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + ": " + message);
        }
    }
}
