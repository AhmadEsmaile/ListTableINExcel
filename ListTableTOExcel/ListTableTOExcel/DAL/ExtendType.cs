using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListTableTOExcel.DAL
{
    public static class ExtendType
    {
        public static object GetDefaultValue(this Type t)
        {
            return (t.IsValueType && Nullable.GetUnderlyingType(t) == null) ? Activator.CreateInstance(t) : null;
        }

    }
}
