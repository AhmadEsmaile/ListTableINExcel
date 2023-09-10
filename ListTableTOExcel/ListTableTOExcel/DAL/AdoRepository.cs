using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ListTableTOExcel.DAL
{
    public class AdoRepository
    {
        /// <summary>
        /// Get/set flag indicating whether the List&lt;T&gt; ExecuteXXX&lt;T&gt;() methods should 
        /// throw an exception if the DataTable retrieved by the query does not match the model 
        /// being created (it compares the number of datatable columns with the number of assigned 
        /// values in the model). The default falue is false.
        /// </summary>
        public bool FailOnMismatch { get; set; }
        /// <summary>
        /// Get/set value indicating the timeout value (in seconds)
        /// </summary>
        public int TimeoutSecs { get; set; }
        /// <summary>
        /// Get/(protected)set the connection string.
        /// </summary>
        public string ConnectionString { get; protected set; }
        /// <summary>
        /// Get/set a flag indicating whether the a return value parameter is added to the sql 
        /// parameter list if it's missing. This only applies to the SetData method 
        /// (insert/update/delete functionality). In order for this to work, you MUST return 
        /// @@ROWCOUNT from your stored proc.
        /// </summary>
        public bool AddReturnParamIfMissing { get; set; }
        /// <summary>
        /// Get/set the bulk insert batch size
        /// </summary>
        public int BulkInsertBatchSize { get; set; }
        /// <summary>
        /// Get/set the number of seconds before the bulk copy times out
        /// </summary>
        public int BulkCopyTimeout { get; set; }

        /// <summary>
        /// Create instance of DBObject, and set default values for properties.
        /// </summary>
        /// <param name="connStr"></param>
        public AdoRepository(string connStr)
        {
            if (string.IsNullOrEmpty(connStr))
            {
                throw new ArgumentNullException("connection string");
            }
            this.ConnectionString = connStr;
            // five minutes should be enough, right?
            this.TimeoutSecs = 300;
            this.FailOnMismatch = false;
            this.AddReturnParamIfMissing = true;
            this.BulkInsertBatchSize = 100;
            this.BulkCopyTimeout = 600;
        }

        /// <summary>
        /// Calls SqlCommand.ExecuteDataReader() to retrieve a dataset from the database.
        /// </summary>
        /// <param name="cmdText">The storedproc or query to execute</param>
        /// <param name="parameters">The parameters to use in the storedproc/query</param>
        /// <returns></returns>
        protected DataTable GetData(string cmdText, SqlParameter[] parameters = null, CommandType cmdType = CommandType.StoredProcedure)
        {
            // by defining these variables OUTSIDE the using statements, we can evaluate them in 
            // the debugger even when the using's go out of scope.
            SqlConnection conn = null;
            SqlCommand cmd = null;
            SqlDataReader reader = null;
            DataTable data = null;

            // create the connection
            using (conn = new SqlConnection(this.ConnectionString))
            {
                // open it
                conn.Open();
                // create the SqlCommand object
                using (cmd = new SqlCommand(cmdText, conn) { CommandTimeout = this.TimeoutSecs, CommandType = cmdType })
                {
                    // give the SqlCommand object the parameters required for the stored proc/query
                    if (parameters != null)
                    {
                        cmd.Parameters.AddRange(parameters);
                    }
                    //create the SqlDataReader
                    using (reader = cmd.ExecuteReader())
                    {
                        // move the data to a DataTable
                        data = new DataTable();
                        data.Load(reader);
                    }
                }
            }
            // return the DataTable object to the calling method
            return data;
        }

        /// <summary>
        /// Calls SqlCommand.ExecuteNonQuery to save data to the database.
        /// </summary>
        /// <param name="connStr"></param>
        /// <param name="cmdText"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        protected int SetData(string cmdText, SqlParameter[] parameters, CommandType cmdType = CommandType.StoredProcedure)
        {
            int result = 0;
            SqlConnection conn = null;
            SqlCommand cmd = null;
            using (conn = new SqlConnection(this.ConnectionString.Base64Decode()))
            {
                conn.Open();
                using (cmd = new SqlCommand(cmdText, conn) { CommandTimeout = this.TimeoutSecs, CommandType = cmdType })
                {
                    SqlParameter rowsAffected = null;
                    if (parameters != null)
                    {
                        cmd.Parameters.AddRange(parameters);
                        // if this is a stored proc and we want to add a return param
                        if (cmdType == CommandType.StoredProcedure && this.AddReturnParamIfMissing)
                        {
                            // see if we already have a return parameter
                            rowsAffected = parameters.FirstOrDefault(x => x.Direction == ParameterDirection.ReturnValue);
                            // if we don't, add one.
                            if (rowsAffected == null)
                            {
                                rowsAffected = cmd.Parameters.Add(new SqlParameter("@rowsAffected", SqlDbType.Int) { Direction = ParameterDirection.ReturnValue });
                            }
                        }
                    }
                    result = cmd.ExecuteNonQuery();
                    result = (rowsAffected != null) ? (int)rowsAffected.Value : result;
                }
            }
            return result;
        }

        /// <summary>
        /// Converts a value from its database value to something we can use (we need this because 
        /// we're using reflection to populate our entities)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        protected static T ConvertFromDBValue<T>(object obj, T defaultValue)
        {
            T result = (obj == null || obj == DBNull.Value) ? default(T) : (T)obj;
            return result;
        }

        /// <summary>
        /// Creates the list of entities from the specified DataTable object. We do this because we 
        /// have two methods that both need to do the same thing.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        protected List<T> MakeEntityFromDataTable<T>(DataTable data)
        {
            Type objType = typeof(T);
            List<T> collection = new List<T>();
            // if we got back data
            if (data != null && data.Rows.Count > 0)
            {
                // we're going to count how many properties in the model were assigned from the 
                // datatable.
                int matched = 0;

                foreach (DataRow row in data.Rows)
                {
                    // create an instance of our object
                    T item = (T)Activator.CreateInstance(objType);

                    // get our object type's properties
                    PropertyInfo[] properties = objType.GetProperties();

                    // set the object's properties as they are found.
                    foreach (PropertyInfo property in properties)
                    {
                        if (data.Columns.Contains(property.Name))
                        {
                            Type pType = property.PropertyType;
                            var defaultValue = pType.GetDefaultValue();
                            var value = row[property.Name];
                            value = ConvertFromDBValue(value, defaultValue);
                            property.SetValue(item, value);
                            matched++;
                        }
                    }
                    if (matched != data.Columns.Count && this.FailOnMismatch)
                    {
                        throw new Exception("Data retrieved does not match specified model.");
                    }
                    collection.Add(item);
                }
            }
            return collection;
        }

        /// <summary>
        /// Executes the named stored proc (using ExecuteReader) that gets data from the database. 
        /// It uses reflection to set property values in the specified type. If nothing was returned 
        /// from the stored proc, the returned list will be empty.
        /// </summary>
        /// <typeparam name="T">The type of the list item</typeparam>
        /// <param name="storedProc"></param>
        /// <param name="parameters"></param>
        /// <returns>A list of the specified type.</returns>
        /// <remarks>Useage: List&lt;MyObject&gt; list = this.ExecuteStoredProc&lt;MyObject&gt;(...)</remarks>
        public List<T> ExecuteStoredProc<T>(string storedProc, params SqlParameter[] parameters)
        {
            if (string.IsNullOrEmpty(storedProc))
            {
                throw new ArgumentNullException("stored procedure");
            }
            // get the data from the database
            DataTable data = this.GetData(storedProc, parameters, CommandType.StoredProcedure);
            List<T> collection = this.MakeEntityFromDataTable<T>(data);
            return collection;
        }

        /// <summary>
        /// Executes the named stored proc (using ExecuteNonQuery) that stores data in the database. 
        /// </summary>
        /// <param name="storedProc"></param>
        /// <param name="parameters"></param>
        /// <returns>The number of records affected</returns>
        public int ExecuteStoredProc(string storedProc, params SqlParameter[] parameters)
        {
            if (string.IsNullOrEmpty(storedProc))
            {
                throw new ArgumentNullException("stored procedure");
            }

            // Save the data to the database. If you don't explicitly return @@ROWCOUNT from your 
            // stored proc, the return value will always be -1, regardless of how many rows are 
            // actually affected.
            int result = this.SetData(storedProc, parameters, CommandType.StoredProcedure);

            return result;
        }

        /// <summary>
        /// Executes the specifid query (using ExecuteReader) that gets data from the database. 
        /// It uses reflection to set property values in the specified type. If nothing was returned 
        /// from the stored proc, the returned list will be empty.
        /// </summary>
        /// <typeparam name="T">The type of the list item</typeparam>
        /// <param name="storedProc"></param>
        /// <param name="parameters"></param>
        /// <returns>A list of the specified type.</returns>
        /// <remarks>Useage: List&lt;MyObject&gt; list = this.ExecuteStoredProc&lt;MyObject&gt;(...)</remarks>
        public List<T> ExecuteQuery<T>(string query, params SqlParameter[] parameters)
        {
            if (string.IsNullOrEmpty(query))
            {
                throw new ArgumentNullException("query");
            }
            DataTable data = this.GetData(query, parameters, CommandType.Text);
            List<T> collection = this.MakeEntityFromDataTable<T>(data);
            return collection;
        }

        /// <summary>
        /// Executes the specified query text (using ExecuteNonQuery) that stores data in the 
        /// database. 
        /// </summary>
        /// <param name="storedProc"></param>
        /// <param name="parameters"></param>
        /// <returns>The number of records affected (if you didn't use SET NOCOUNT ON in 
        /// your batch)</returns>
        public int ExecuteQuery(string query, params SqlParameter[] parameters)
        {
            if (string.IsNullOrEmpty(query))
            {
                throw new ArgumentNullException("query");
            }

            // Save the data to the database. If you use SET NOCOUNT ON in your query, the return 
            // value will always be -1, regardless of how many rows are actually affected.
            int result = this.SetData(query, parameters, CommandType.Text);
            return result;
        }

        /// <summary>
        /// Performs a simply bulk insert into a table in the database. The schema MUST be part of 
        /// the table name if the target schema isn't "dbo".
        /// </summary>
        /// <param name="dataTable"></param>
        /// <returns></returns>
        public int DoBulkInsert(DataTable dataTable)
        {
            // If you have an auto-incrementing identity column, make sure you decorate the column 
            // with DbCanInsert attribute. If you don't, it will try to put the first available 
            // property into that db table column, and will throw an exception if the types don't 
            // match.
            int recordsAffected = 0;
            SqlConnection conn = null;
            SqlBulkCopy bulk = null;
            using (conn = new SqlConnection(this.ConnectionString))
            {
                conn.Open();
                using (bulk = new SqlBulkCopy(conn)
                {
                    BatchSize = this.BulkInsertBatchSize
                    ,
                    BulkCopyTimeout = this.BulkCopyTimeout
                    ,
                    DestinationTableName = dataTable.TableName
                })
                {
                    bulk.WriteToServer(dataTable);
                }
            }
            return recordsAffected;
        }

        /// <summary>
        /// Performs a simple bulk insert into a table in the database. The schema MUST be part of 
        /// the table name if the target schema isn't "dbo".
        /// </summary>
        /// <param name="dataTable"></param>
        /// <returns></returns>
        public int DoBulkInsert<T>(IEnumerable<T> data, string tableName, bool byDBInsertAttribute = false)
        {
            int result = 0;
            DataTable dataTable = null;

            if (data.Count() > 0)
            {
                using (dataTable = new DataTable() { TableName = tableName })
                {
                    Type type = typeof(T);
                    MethodInfo method = type.GetMethod("GetEntityProperties");

                    // get the properties regardless of whether or not the object is using EntityBase
                    PropertyInfo[] properties = (method == null) ? type.GetProperties().Where(prop => Attribute.IsDefined(prop, typeof(CanDbInsertAttribute))).ToArray()
                                                                 : (PropertyInfo[])method.Invoke(this, null);

                    foreach (PropertyInfo property in properties)
                    {
                        dataTable.Columns.Add(new DataColumn(property.Name, property.PropertyType));
                    }
                    foreach (T entity in data)
                    {
                        DataRow row = dataTable.NewRow();
                        foreach (PropertyInfo property in properties)
                        {
                            row[property.Name] = property.GetValue(entity);
                        }
                        dataTable.Rows.Add(row);
                    }
                }
                result = this.DoBulkInsert(dataTable);
            }
            return result;
        }

    }
}
