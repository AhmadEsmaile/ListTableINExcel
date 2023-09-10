using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListTableTOExcel.DAL
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class CanDbInsertAttribute : Attribute
    {
        public string Name { get; set; }
        public string Argument { get; set; }
    }
}
