using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel数据导入.OracleDbAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.ManagedDataAccess.Client;
using System.Data;

namespace Excel数据导入.OracleDbAccess.Tests
{
    [TestClass()]
    public class OracleDbManagerTests
    {
        /// <summary>
        /// 测试用数据库连接串
        /// </summary>
        private static readonly string m_connStr = "User Id=NURSE_MANAGE_BASE;Password=mandala;Data Source=(DESCRIPTION =(ADDRESS_LIST =(ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.73.128)(PORT = 1521)))(CONNECT_DATA =(SERVICE_NAME = orcl)))";

        private OracleDbManager oracleDbManager = new OracleDbManager(m_connStr);

        /// <summary>
        /// 测试 查询 sql文查询
        /// </summary>
        [TestMethod()]
        public void QueryTest()
        {
            try
            {
                string selectSql = "SELECT table_name FROM User_Tables";
                DataTable dtRes = oracleDbManager.Query(selectSql);
 
                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message + ex.StackTrace);
            }
        }

        /// <summary>
        /// 测试 查询功能
        /// 1、动态参数绑定
        /// 2、参数按名称绑定，与顺序无关
        /// </summary>
        [TestMethod()]
        public void Query1Test()
        {
            try
            {
                string selectSql = "SELECT * FROM User_Tables ut WHERE ut.STATUS = :STATUS AND ut.LOGGING = :LOGGING";
                OracleParameter[] oracleParameters = new OracleParameter[2];
                oracleParameters[0] = new OracleParameter("LOGGING", "YES");
                oracleParameters[1] = new OracleParameter("STATUS", "VALID");
                DataTable dtRes = oracleDbManager.Query(selectSql, oracleParameters);

                Assert.IsTrue(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message + ex.StackTrace);
            }
        }

        /// <summary>
        /// 测试 执行非查询sql语句。create、drop、delete、insert、update
        /// </summary>
        [TestMethod()]
        public void ExecuteNonQuerySqlTest()
        {
            try
            {
                string testTableName = string.Format("Test_{0}{1}", Guid.NewGuid().ToString().Substring(0, 3), DateTime.Now.ToString("yyyyMMddHHmmss"));
                //Create Table
                string createTableSql = string.Format("CREATE TABLE {0}(ID VARCHAR2(32),NAME VARCHAR2(30),BIRTHDAY DATE,SEX VARCHAR2(6),CREATE_TIME DATE)", testTableName);
                oracleDbManager.ExecuteNonQuerySql(createTableSql);

                //Insert Data
                //sql文本插入
                string insertSql1 = string.Format("INSERT INTO {0}(ID,NAME,BIRTHDAY,SEX,CREATE_TIME) VALUES('001','Tom',SYSDATE-6000-13/24,'男',SYSDATE)", testTableName);
                string insertSql2 = string.Format("INSERT INTO {0}(ID,NAME,BIRTHDAY,SEX,CREATE_TIME) VALUES('002','Tom',SYSDATE-6000-13/24,'男',SYSDATE)", testTableName);
                oracleDbManager.ExecuteNonQuerySql(insertSql1);
                oracleDbManager.ExecuteNonQuerySql(insertSql2);
                //动态参数sql插入
                string insertSqlParams = string.Format("INSERT INTO {0}(ID,NAME,BIRTHDAY,SEX,CREATE_TIME) VALUES(:ID,:NAME,:BIRTHDAY,:SEX,:CREATE_TIME)", testTableName);
                OracleParameter[] oracleParametersInsert = new OracleParameter[5];
                oracleParametersInsert[0] = new OracleParameter("ID", "002");
                oracleParametersInsert[1] = new OracleParameter("NAME", "Jack");
                oracleParametersInsert[2] = new OracleParameter("BIRTHDAY", DateTime.Now.AddDays(-6000).AddHours(-13));
                oracleParametersInsert[3] = new OracleParameter("SEX", "女");
                oracleParametersInsert[4] = new OracleParameter("CREATE_TIME", DateTime.Now);
                oracleDbManager.ExecuteNonQuerySql(insertSqlParams, oracleParametersInsert);

                //Update Data
                //sql文本更新
                string updateSql = string.Format("UPDATE {0} SET BIRTHDAY = BIRTHDAY+1,SEX='女',CREATE_TIME = SYSDATE WHERE ID='001'", testTableName);
                oracleDbManager.ExecuteNonQuerySql(updateSql);
                //动态参数sql更新
                string updateSqlParams = string.Format("UPDATE {0} SET BIRTHDAY = BIRTHDAY+2,SEX=:SEX,CREATE_TIME = :CREATE_TIME WHERE ID=:ID", testTableName);
                OracleParameter[] oracleParametersUpdate = new OracleParameter[3];
                oracleParametersUpdate[0] = new OracleParameter("ID", "002");
                oracleParametersUpdate[1] = new OracleParameter("SEX", "女");
                oracleParametersUpdate[2] = new OracleParameter("CREATE_TIME", DateTime.Now);
                oracleDbManager.ExecuteNonQuerySql(updateSqlParams, oracleParametersUpdate);

                //Delete Data
                //sql文本删除
                string deleteSql = string.Format("DELETE {0} WHERE ID='001'", testTableName);
                oracleDbManager.ExecuteNonQuerySql(deleteSql);
                //动态参数 sql删除
                string deleteSqlParams = string.Format("DELETE {0} WHERE ID=:ID", testTableName);
                OracleParameter[] oracleParametersDelete = new OracleParameter[1];
                oracleParametersDelete[0] = new OracleParameter("ID", "002");
                oracleDbManager.ExecuteNonQuerySql(deleteSqlParams, oracleParametersDelete);

                //Drop Table
                string dropSql = string.Format("DROP TABLE {0}", testTableName);
                oracleDbManager.ExecuteNonQuerySql(dropSql);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message + ex.StackTrace);
            }

            Assert.IsTrue(true);
        }

        /// <summary>
        /// 测试 批量插入
        /// </summary>
        [TestMethod()]
        public void InsertDataTest()
        {
            DataTable srcDt = null;
            try
            {
                int dataSize = 1000 * 10;
                //创建待插入的大量数据
                srcDt = new DataTable();
                srcDt.Columns.Add(new DataColumn("ID", typeof(string)));
                srcDt.Columns.Add(new DataColumn("NAME", typeof(string)));
                srcDt.Columns.Add(new DataColumn("BIRTHDAY", typeof(DateTime)));
                srcDt.Columns.Add(new DataColumn("SEX", typeof(string)));
                srcDt.Columns.Add(new DataColumn("CREATE_TIME", typeof(DateTime)));

                int idNum = 1000;
                string nameStr = "Man";
                DateTime baseDate = DateTime.Now;
                string sexStr = idNum % 2 == 0 ? "男" : "女";
                //创建表数据
                DataRow dr;
                for (int i = 0; i < dataSize; i++)
                {
                    dr = srcDt.NewRow();
                    dr["ID"] = idNum++;
                    dr["NAME"] = nameStr + idNum;
                    dr["BIRTHDAY"] = baseDate.AddMinutes(idNum);
                    dr["SEX"] = idNum % 2 == 0 ? "男" : "女";
                    dr["CREATE_TIME"] = baseDate.AddSeconds(idNum).ToString("yyyy/MM/dd HH:mm:ss");
                    srcDt.Rows.Add(dr);
                }

                //创建对应的数据表
                string testTableName = string.Format("Test_{0}{1}", Guid.NewGuid().ToString().Substring(0, 3), DateTime.Now.ToString("yyyyMMddHHmmss"));
                string createTableSql = string.Format("CREATE TABLE {0}(ID VARCHAR2(32),NAME VARCHAR2(30),BIRTHDAY DATE,SEX VARCHAR2(6),CREATE_TIME DATE)", testTableName);
                oracleDbManager.ExecuteNonQuerySql(createTableSql);

                //构建插入SQL 参数 插入数据
                string insertSqlParams = string.Format("INSERT INTO {0}(ID,NAME,BIRTHDAY,SEX,CREATE_TIME) VALUES(:ID,:NAME,:BIRTHDAY,:SEX,:CREATE_TIME)", testTableName);
                OracleParameter[] oracleParametersInsert = new OracleParameter[5];
                oracleParametersInsert[0] = new OracleParameter("ID", OracleDbType.Varchar2, 32, "ID");
                oracleParametersInsert[1] = new OracleParameter("NAME", OracleDbType.Varchar2, 30, "NAME");
                oracleParametersInsert[2] = new OracleParameter("BIRTHDAY", OracleDbType.Date, 22, "BIRTHDAY");
                oracleParametersInsert[3] = new OracleParameter("SEX", OracleDbType.Varchar2, 6, "SEX");
                oracleParametersInsert[4] = new OracleParameter("CREATE_TIME", OracleDbType.Date, 22, "CREATE_TIME");
                int resRow = oracleDbManager.InsertData(srcDt, insertSqlParams, oracleParametersInsert);

                //删除表
                string dropSql = string.Format("DROP TABLE {0} PURGE", testTableName);
                oracleDbManager.ExecuteNonQuerySql(dropSql);
                srcDt.Dispose();
                srcDt = null;
                if (resRow != dataSize)
                {
                    Assert.Fail(string.Format("漏插数据，目标插入数据：{0}行，实际插入{1}行。", dataSize, resRow));
                }
            }
            catch (Exception ex)
            {
                if (srcDt != null)
                {
                    srcDt.Dispose();
                    srcDt = null;
                }
                Assert.Fail(ex.Message + ex.StackTrace);
            }

            Assert.IsTrue(true);

        }
    }
}