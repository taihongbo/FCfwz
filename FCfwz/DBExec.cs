using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace FCfwz
{

    public class DBExec
    {
        private SqlConnection connection = null;
        private SqlTransaction transaction = null;
        public SqlConnection Connection
        {
            get
            {
                return connection;
            }
            set
            {
                connection = value;
            }
        }
        public SqlTransaction Transaction
        {
            get
            {
                return transaction;
            }
            set
            {
                transaction = value;
            }
        }
        public DBExec()
        {

        }
        public DBExec(SqlConnection connection, SqlTransaction transaction)
        {
            this.connection = connection;
            this.transaction = transaction;
        }

        public DataRow GetSingleRow(string strSQL, SqlTransaction transaction, params SqlParameter[] parameters)
        {
            try
            {

                SqlDataAdapter adapter = new SqlDataAdapter(strSQL, connection);
                if (parameters != null)
                    adapter.SelectCommand.Parameters.AddRange(parameters);

                if (transaction != null)
                {
                    adapter.SelectCommand.Transaction = transaction;
                }
                else
                {
                    if (this.transaction != null)
                    {
                        adapter.SelectCommand.Transaction = this.transaction;
                    }
                }
                DataSet dataSet = new DataSet();
                adapter.Fill(dataSet);
                if (dataSet.Tables.Count > 0)
                {
                    if (dataSet.Tables[0].Rows.Count > 0)
                    {
                        return dataSet.Tables[0].Rows[0];
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataTable GetTableWithParameter(string strSQL, SqlTransaction transaction, params SqlParameter[] parameters)
        {
            try
            {

                SqlDataAdapter adapter = new SqlDataAdapter(strSQL, connection);
                adapter.SelectCommand.Parameters.AddRange(parameters);

                if (transaction != null)
                {
                    adapter.SelectCommand.Transaction = transaction;
                }
                else
                {
                    if (this.transaction != null)
                    {
                        adapter.SelectCommand.Transaction = this.transaction;
                    }
                }
                //DataSet dataSet = new DataSet();
                DataTable table = new DataTable();
                adapter.Fill(table);
                if (table.Rows.Count > 0)
                {
                    return table;
                }
                return null;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool Exists(string strSQL, SqlTransaction transaction, params SqlParameter[] parameters)
        {
            using (SqlCommand command = new SqlCommand())
            {
                if (transaction != null)
                {
                    command.Transaction = transaction;
                }
                else
                {
                    if (this.transaction != null)
                    {
                        command.Transaction = this.transaction;
                    }
                }
                command.Connection = connection;
                command.CommandText = strSQL;
                try
                {
                    command.Parameters.AddRange(parameters);
                    if (Convert.ToInt32(command.ExecuteScalar()) > 0)
                    {
                        return true;
                    }
                    return false;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public bool Exists(string strSQL, SqlTransaction transaction)
        {
            using (SqlCommand command = new SqlCommand())
            {
                if (transaction != null)
                {
                    command.Transaction = transaction;
                }
                else
                {
                    if (this.transaction != null)
                    {
                        command.Transaction = this.transaction;
                    }
                }
                command.Connection = connection;
                command.CommandText = strSQL;
                try
                {
                    if (Convert.ToInt32(command.ExecuteScalar()) > 0)
                    {
                        return true;
                    }
                    return false;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public int ExecuteSql(string strSQL, SqlTransaction transaction, params SqlParameter[] parameters)
        {
            using (SqlCommand command = new SqlCommand())
            {
                if (transaction != null)
                {
                    command.Transaction = transaction;
                }
                else
                {
                    if (this.transaction != null)
                    {
                        command.Transaction = this.transaction;
                    }
                }
                command.Connection = connection;
                command.CommandText = strSQL;

                int result = -1;

                try
                {
                    if (parameters != null)
                    {
                        command.Parameters.AddRange(parameters);
                    }

                    result = command.ExecuteNonQuery();

                }
                catch (Exception ex)
                {
                    throw ex;
                }

                return result;
            }
        }

        public int ExecuteCountSql(string strSQL, SqlTransaction transaction, params SqlParameter[] parameters)
        {
            using (SqlCommand command = new SqlCommand())
            {
                if (transaction != null)
                {
                    command.Transaction = transaction;
                }
                else
                {
                    if (this.transaction != null)
                    {
                        command.Transaction = this.transaction;
                    }
                }
                command.Connection = connection;
                command.CommandText = strSQL;
                try
                {
                    command.Parameters.AddRange(parameters);
                    return Convert.ToInt32(command.ExecuteScalar());
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }
        public int ExecuteAddSql(string strSQL, SqlTransaction transaction, params SqlParameter[] parameters)
        {
            using (SqlCommand command = new SqlCommand())
            {
                if (transaction != null)
                {
                    command.Transaction = transaction;
                }
                else
                {
                    if (this.transaction != null)
                    {
                        command.Transaction = this.transaction;
                    }
                }
                command.Connection = connection;
                command.CommandText = strSQL;
                try
                {
                    command.Parameters.AddRange(parameters);
                    return Convert.ToInt32(command.ExecuteScalar());
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public static void ExecuteSqlTran()
        {

        }

        public int GetCount(string strSQL, SqlTransaction transaction, params SqlParameter[] parameters)
        {
            try
            {

                SqlDataAdapter adapter = new SqlDataAdapter(strSQL, connection);
                if (transaction != null)
                {
                    adapter.SelectCommand.Transaction = transaction;
                }
                else
                {
                    if (this.transaction != null)
                    {
                        adapter.SelectCommand.Transaction = this.transaction;
                    }
                }
                DataSet dataSet = new DataSet();
                adapter.Fill(dataSet);
                if (dataSet.Tables.Count > 0)
                {
                    if (dataSet.Tables[0].Rows.Count > 0)
                    {
                        return dataSet.Tables[0].Rows.Count;
                    }
                }
                return 0;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataTable Query(string strSQL, SqlTransaction transaction, string strTab)
        {
            try
            {

                SqlDataAdapter adapter = new SqlDataAdapter(strSQL, connection);
                if (transaction != null)
                {
                    adapter.SelectCommand.Transaction = transaction;
                }
                else
                {
                    if (this.transaction != null)
                    {
                        adapter.SelectCommand.Transaction = this.transaction;
                    }
                }
                DataSet dataSet = new DataSet();
                adapter.Fill(dataSet, strTab);
                return dataSet.Tables[strTab];
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public object GetScalar(string strSQL, SqlTransaction transaction, params SqlParameter[] parameters)
        {
            using (SqlCommand command = new SqlCommand())
            {
                if (transaction != null)
                {
                    command.Transaction = transaction;
                }
                else
                {
                    if (this.transaction != null)
                    {
                        command.Transaction = this.transaction;
                    }
                }
                command.Connection = connection;
                command.CommandText = strSQL;
                try
                {
                    command.Parameters.AddRange(parameters);
                    return command.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public DataSet ExecuteProcedure(string procName, int id, int topNum, out int count, params string[] strWhere)
        {
            throw new Exception("ExecuteStoredProcedure");
        }
    }

}
