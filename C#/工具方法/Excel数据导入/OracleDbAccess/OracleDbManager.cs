using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel数据导入.OracleDbAccess
{
    class OracleDbManager : IDisposable
    {
        #region 变量

        /// <summary>
        /// 数据库连接字符串串
        /// </summary>
        private string m_connectionString = string.Empty;

        /// <summary>
        /// 数据库连接对象
        /// </summary>
        private OracleConnection m_connection;

        /// <summary>
        /// 数据库命令
        /// </summary>
        private OracleCommand m_command;

        #endregion

        #region 属性

        /// <summary>
        /// 获取数据库连接
        /// </summary>
        private OracleConnection m_OracleConnection
        {
            get
            {
                if (m_connection == null)
                {
                    m_connection = new OracleConnection(m_connectionString);
                }

                lock (m_connection)
                {
                    if (m_connection.State != System.Data.ConnectionState.Open)
                    {
                        m_connection.Close();
                        m_connection.Open();
                    }
                }

                return m_connection;
            }
        }

        #endregion

        #region 构造函数

        /// <summary>
        /// 有参构造
        /// </summary>
        /// <param name="connStr">数据库连接串</param>
        public OracleDbManager(string connStr)
        {
            m_connectionString = connStr;
        }

        #endregion

        #region 资源释放


        /*
         * 微软官方示例：https://docs.microsoft.com/en-us/visualstudio/code-quality/ca1063?view=vs-2019
         * **/

        /// <summary>
        /// 资源释放
        /// </summary>
        public void Dispose()
        {
            Dispose(true); //释放托管资源
            GC.SuppressFinalize(this); //不调用析构函数
        }

        ~OracleDbManager()
        {
            //释放非托管资源
            Dispose(false);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {//释放托管资源 
                if (m_command != null)
                {
                    m_command.Dispose();
                    m_command = null;
                }

                if (m_connection != null)
                {
                    m_connection.Dispose();
                    m_connection = null;
                }

            }

            //释放非托管资源
            //当前无非托管资源
        }

        #endregion

        #region 方法

        /// <summary>
        /// 执行非查询类SQL
        /// </summary>
        /// <param name="sql"></param>
        /// <returns>SQL执行影响的行数。 -1：执行失败</returns>
        public int ExecuteNonQuerySql(string sql)
        {
            lock (m_OracleConnection)
            {
                int res = -1;
                OracleCommand oracleCommand = new OracleCommand();
                oracleCommand.Connection = m_OracleConnection;
                oracleCommand.CommandType = CommandType.Text;
                oracleCommand.CommandText = sql;
                res = oracleCommand.ExecuteNonQuery();
                return res;
            }
        }

        /// <summary>
        /// 执行非查询类SQL 动态参数绑定(顺序无关，按照参数名绑定)
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="oracleParameters"></param>
        /// <returns></returns>
        public int ExecuteNonQuerySql(string sql, OracleParameter[] oracleParameters)
        {
            lock (m_OracleConnection)
            {
                int res = -1;
                OracleCommand oracleCommand = new OracleCommand();
                oracleCommand.Connection = m_OracleConnection;
                oracleCommand.CommandType = CommandType.Text;
                oracleCommand.CommandText = sql;
                oracleCommand.Parameters.AddRange(oracleParameters);
                oracleCommand.BindByName = true;
                res = oracleCommand.ExecuteNonQuery();
                return res;
            }
        }

        /// <summary>
        /// 执行查询SQL语句 获取DataTable类型数据 动态参数按名称绑定
        /// </summary>
        /// <param name="sql">查询sql语句</param>
        /// <param name="oracleParameters">参数</param>
        /// <returns></returns>
        public DataTable Query(string sql, OracleParameter[] oracleParameters)
        {
            lock (m_OracleConnection)
            {
                DataSet ds = new DataSet();
                OracleCommand oracleCommand = new OracleCommand();
                oracleCommand.Connection = m_OracleConnection;
                oracleCommand.CommandType = CommandType.Text;
                oracleCommand.CommandText = sql;
                oracleCommand.Parameters.AddRange(oracleParameters);
                oracleCommand.BindByName = true;
                using (OracleDataAdapter oracleDataAdapter = new OracleDataAdapter(oracleCommand))
                {
                    oracleDataAdapter.Fill(ds, "DS");
                }
                return ds == null ? null : ds.Tables.Count < 1 ? null : ds.Tables[0];
            }
        }

        /// <summary>
        /// 执行查询SQL语句 获取DataTable类型数据
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public DataTable Query(string sql)
        {
            lock (m_OracleConnection)
            {
                DataSet ds = new DataSet();
                using (OracleDataAdapter oracleDataAdapter = new OracleDataAdapter(sql, m_OracleConnection))
                {
                    oracleDataAdapter.Fill(ds, "DS");
                }
                return ds == null ? null : ds.Tables.Count < 1 ? null : ds.Tables[0];
            }
        }

        /// <summary>
        /// 插入数据
        /// </summary>
        /// <param name="srcDt"></param>
        /// <param name="insertSql"></param>
        /// <param name="oracleParameters"></param>
        /// <returns></returns>
        public int InsertData(DataTable srcDt, string insertSql, OracleParameter[] oracleParameters)
        {
            lock (m_OracleConnection)
            {
                int resRow = -1;
                OracleCommand oracleCommand = new OracleCommand();
                oracleCommand.Connection = m_OracleConnection;
                OracleTransaction m_OraTrans = m_OracleConnection.BeginTransaction(IsolationLevel.ReadCommitted);//创建事务对象
                try
                {
                    oracleCommand.CommandType = CommandType.Text;
                    oracleCommand.Transaction = m_OraTrans; //绑定事务
                    oracleCommand.CommandText = insertSql;  //绑定插入语句
                    oracleCommand.Parameters.AddRange(oracleParameters); //绑定参数
                    oracleCommand.BindByName = true;
                    //DataTable标记数据为新增数据
                    srcDt.AcceptChanges();
                    foreach (DataRow dr in srcDt.Rows)
                    {
                        dr.SetAdded();
                    }

                    using (OracleDataAdapter oracleDataAdapter = new OracleDataAdapter())
                    {
                        oracleDataAdapter.InsertCommand = oracleCommand;
                        resRow = oracleDataAdapter.Update(srcDt);
                        srcDt.AcceptChanges();
                        m_OraTrans.Commit(); //提交事物
                    }
                }
                catch (Exception)
                {
                    m_OraTrans.Rollback();
                    throw;
                }

                return resRow;
            }
        }

        #endregion
    }
}
