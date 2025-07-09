using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Web;
using System.Threading.Tasks;
using System.Web.Http.Results;
using System.Reflection;

namespace AddOnGs.Extenstion
{
    public class DbStoreProcContext
    {
        private readonly string _connection;

        public DbStoreProcContext(string _connectionString)
        {
            _connection = _connectionString;
        }

        public IDbConnection CreateConnectionRead() => new SqlConnection(_connection);
        public IDbConnection CreateConnectionWrite() => new SqlConnection(_connection);
        public async Task<bool> ExecuteStoreNonQueryAsync(string store, SqlParameter[] sqlParameter = null)
        {
            SqlConnection sqlConnection = new SqlConnection(_connection);
            try
            {
                sqlConnection.Open();

                SqlCommand command = new SqlCommand(store, sqlConnection);
                command.CommandType = CommandType.StoredProcedure;
                if (sqlParameter != null && sqlParameter.Length != 0)
                {
                    command.Parameters.AddRange(sqlParameter);
                }

                SqlDataReader dr = await command.ExecuteReaderAsync();
                if (dr.Read())
                {
                    var isSuccess = (bool)dr["result"];
                    var msg = (string)dr["Message"];

                    return isSuccess;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                sqlConnection.Close();
                throw ex;
            }
            finally { sqlConnection.Close(); }
        }
        public async Task<List<T>> ExcuteStoreQueryListAsync<T>(string store, SqlParameter[] sqlParameter = null)
        {
            SqlConnection sqlConnection = new SqlConnection(_connection);
            try
            {
                sqlConnection.Open();

                SqlCommand command = new SqlCommand(store, sqlConnection);
                command.CommandType = CommandType.StoredProcedure;
                if (sqlParameter != null && sqlParameter.Length != 0)
                {
                    command.Parameters.AddRange(sqlParameter);
                }

                SqlDataReader dr = await command.ExecuteReaderAsync();
                List<T> list = new List<T>();
                T obj = default;
                while (dr.Read())
                {
                    obj = Activator.CreateInstance<T>();
                    foreach (PropertyInfo prop in obj.GetType().GetProperties())
                    {
                        if (!object.Equals(dr[prop.Name], DBNull.Value))
                        {
                            prop.SetValue(obj, Convert.ChangeType(dr[prop.Name], prop.PropertyType), null);
                        }
                    }
                    list.Add(obj);
                }
                return list;
            }
            catch (Exception ex)
            {
                sqlConnection.Close();
                throw ex;
            }
            finally { sqlConnection.Close(); }
        }
        public async Task<T> ExcuteStoreQuerySingleAsync<T>(string store, SqlParameter[] sqlParameter = null)
        {
            SqlConnection sqlConnection = new SqlConnection(_connection);
            try
            {
                sqlConnection.Open();

                SqlCommand command = new SqlCommand(store, sqlConnection);
                command.CommandType = CommandType.StoredProcedure;
                if (sqlParameter != null && sqlParameter.Length != 0)
                {
                    command.Parameters.AddRange(sqlParameter);
                }

                SqlDataReader dr = await command.ExecuteReaderAsync();
                T obj = default;
                while (dr.Read())
                {
                    obj = Activator.CreateInstance<T>();
                    foreach (PropertyInfo prop in obj.GetType().GetProperties())
                    {
                        if (!object.Equals(dr[prop.Name], DBNull.Value))
                        {
                            prop.SetValue(obj, Convert.ChangeType(dr[prop.Name], prop.PropertyType), null);
                        }
                    }
                }
                return obj;
            }
            catch (Exception ex)
            {
                sqlConnection.Close();
                throw ex;
            }
            finally { sqlConnection.Close(); }
        }
    }
}