using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListTableTOExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args == null || args.Length == 0)
                Console.WriteLine("enter command correctly");

            string connectionString = "";
    

            foreach (string arg in args)
            {
                if (arg.Contains("connectionString:"))
                    connectionString = Split(arg);
              

            }
          
            //string Conection = @"  ";
            var result= new ExecuteQuery(connectionString).ListTable();
            new ListTableINExcel().CreateExcel(result);

        }
        private static string Split(string str)
        {
            var index = str.IndexOf(":");

            if (index < 0)
                return null;

            return str.Substring(index + 1);
        }
    }
}
