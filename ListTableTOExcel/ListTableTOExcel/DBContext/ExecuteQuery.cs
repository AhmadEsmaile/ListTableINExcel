using ListTableTOExcel.DAL;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListTableTOExcel
{
    public class ExecuteQuery : AdoRepository
    {
        public ExecuteQuery(string connStr) : base(connStr)
        {

        }
        public List<TableModel> ListTable()
        {
            var result = ListTableData();


            return result;
        }

        public List<TableModel> ListTableData()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("SELECT");
            sb.AppendLine("\tsch.TABLE_CATALOG [DBName],");
            sb.AppendLine("\tsch.TABLE_SCHEMA [Schema],");
            sb.AppendLine("\tCAST(st.name AS NVARCHAR(200)) [Table],");
            sb.AppendLine("\tUPPER(sch.DATA_TYPE) [DataType],");
            sb.AppendLine("\tsch.IS_NULLABLE [Nullable],");
            sb.AppendLine("\tCAST(sc.name AS NVARCHAR(200)) [Column],");
            sb.AppendLine("\tCAST(sep.value AS NVARCHAR(200)) [Description]");
            sb.AppendLine("FROM sys.tables st");
            sb.AppendLine("INNER JOIN sys.columns sc ON st.object_id = sc.object_id");
            sb.AppendLine("LEFT JOIN INFORMATION_SCHEMA.COLUMNS sch ON st.name = sch.TABLE_NAME ");
            sb.AppendLine("\t\t\t\t\tAND sc.name = sch.COLUMN_NAME");
            sb.AppendLine("\t\t\t\t\tAND schema_name(st.schema_id) = sch.TABLE_SCHEMA");
            sb.AppendLine("LEFT JOIN sys.extended_properties sep ON st.object_id = sep.major_id");
            sb.AppendLine("\t\t\t\t\tAND sc.column_id = sep.minor_id");
            sb.AppendLine("\t\t\t\t\tAND sep.name = 'MS_Description'");
            sb.AppendLine("WHERE sch.TABLE_SCHEMA <> 'APP' AND sch.TABLE_SCHEMA <> 'CMN' AND sch.TABLE_SCHEMA <> 'org'");
            sb.AppendLine("ORDER BY sch.TABLE_SCHEMA, st.name, sch.ORDINAL_POSITION ");
            var query = sb.ToString();
            var result = this.ExecuteQuery<TableModel>(query);
            return result;
        }
    }
}
