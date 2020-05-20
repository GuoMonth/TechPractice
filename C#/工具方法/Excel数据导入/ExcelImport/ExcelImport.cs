using Aspose.Cells;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace Excel数据导入.ExcelImport
{
    class ExcelImport
    {
        /// <summary>
        /// 导入配置文件路径
        /// </summary>
        private string m_excelImportConfigXmlPath; //= Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelImportConfig.xml");

        /// <summary>
        /// 全局数据访问帮助类
        /// </summary>
        private OracleDbAccess.OracleDbManager m_oracleDbManager;

        /// <summary>
        /// 匹配默认值字符串内容 懒惰匹配
        /// </summary>
        private Regex m_GetDefalutValueStr = new Regex(@"\$\{.*?\}");

        /// <summary>
        /// 缓存查询回来的excel映射数据。每次查询回来n个映射值，存入List集合中。
        /// key:excel中的列名
        /// value：映射的数据库对应字段的数据。item1:excle中数据 item2:映射的数据库中数据
        /// </summary>
        private Dictionary<string, List<Tuple<string, string>>> m_MapColData;

        /// <summary>
        /// 构造函数
        /// </summary>
        public ExcelImport(string connStr, string xmlImportConfig)
        {
            //获取excel导入数据库的连接串
            m_oracleDbManager = new OracleDbAccess.OracleDbManager(connStr);
            m_excelImportConfigXmlPath = xmlImportConfig;
        }

        #region 方法

        /// <summary>
        /// 导入excel数据到数据库中
        /// 1、excel文件
        /// 2、对应的xml配置
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <param name="menuName"></param>
        /// <returns></returns>
        public bool ImportExcelData(string excelFilePath, string menuName)
        {
            DataTable excelData = null, importData = null;
            try
            {
                int numRows = -1;
                int allRows = 0;
                excelData = GetExcelData(excelFilePath); //获取excel中数据
                EntityExcelImportConfig excelImportConfigEntity = GetExcelImportXmlConfig(m_excelImportConfigXmlPath, menuName); //读取excel导入配置
                importData = GetImportData(excelImportConfigEntity, excelData); //根据配置和excel数据，生成待导入数据，此时会做一个校验xml导入配置的列必须为excelData列集合的子集
                allRows = importData.Rows.Count;
                excelData.Dispose();
                excelData = null;
                if (importData != null && importData.Rows.Count > 0 && importData.Columns.Count > 0)
                {
                    //生成sql
                    string insertSql = string.Format("INSERT INTO {0}({1}) VALUES({2})"
                        , excelImportConfigEntity.m_TableName
                        , string.Join(",", excelImportConfigEntity.m_ColsDefine.Select(c => c.m_DbTableColName))
                        , string.Join(",", excelImportConfigEntity.m_ColsDefine.Select(c => ":" + c.m_DbTableColName)));

                    //生成参数
                    var paramsOracle = excelImportConfigEntity.m_ColsDefine.Select(c => new OracleParameter(c.m_DbTableColName, c.m_OracleDbType, c.m_DataMaxLength, c.m_DbTableColName));

                    numRows = m_oracleDbManager.InsertData(importData, insertSql, paramsOracle.ToArray());
                }
                importData.Dispose();
                importData = null;
                return numRows == allRows;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (excelData != null)
                {
                    excelData.Dispose();
                    excelData = null;
                }
                if (importData != null)
                {
                    importData.Dispose();
                    importData = null;
                }
            }

        }


        /// <summary>
        /// 获取excel中数据内容
        /// </summary>
        /// <param name="filePath">excel文件路径</param>
        /// <returns></returns>
        private DataTable GetExcelData(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException("错误：excel文件路径为空！");
            }

            if (!File.Exists(filePath))
            {
                throw new ArgumentException(string.Format("错误：无法从给定的路径【{0}】获取到excel文件", filePath));
            }

            DataTable dtRes = null;
            //获取excel文件内容
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open))
            {
                //Aspose.Cells.LoadFormat 参考文档：https://apireference.aspose.com/cells/net/aspose.cells/loadformat
                LoadFormat loadFormat = filePath.EndsWith("xlsx") ? LoadFormat.Xlsx
                    : filePath.EndsWith("xls") ? LoadFormat.Excel97To2003
                    : LoadFormat.Auto; //判定excel文件格式，进行内容数据读取
                Workbook workbook = new Workbook(fileStream);
                Worksheet worksheet = workbook.Worksheets[0]; //只读取第一个sheet
                int rowsCount = worksheet.Cells.MaxDataRow + 1;//Rows.Count;
                int colmnsCount = worksheet.Cells.MaxDataColumn + 1;//.Columns.Count;
                dtRes = worksheet.Cells.ExportDataTable(0, 0, rowsCount, colmnsCount, true);
            }

            return dtRes;
        }

        /// <summary>
        /// 读取xml，获取excel导入配置
        /// </summary>
        /// <param name="xmlFilePath">xml配置文件路径</param>
        /// <param name="menuName">菜单名称</param>
        private EntityExcelImportConfig GetExcelImportXmlConfig(string xmlFilePath, string menuName)
        {
            if (string.IsNullOrEmpty(xmlFilePath))
            {
                throw new ArgumentNullException("错误：xml文件路径为空！");
            }

            if (!File.Exists(xmlFilePath))
            {
                throw new ArgumentException(string.Format("错误：无法从给定的路径【{0}】获取到xml文件", xmlFilePath));
            }

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(xmlFilePath);
            XmlNode xn = xmlDocument.DocumentElement.SelectSingleNode(string.Format("ECXEL_IMPORT[@MENU_NAME='{0}']", menuName));
            if (xn == null)
            {
                throw new ArgumentException(string.Format("无法读取到{0}中的ECXEL_IMPORT[@MENU_NAME='{1}']节点", xmlFilePath, menuName));
            }

            EntityExcelImportConfig excelImportConfigEntity = new EntityExcelImportConfig(xn as XmlElement);
            return excelImportConfigEntity;
        }

        /// <summary>
        /// 获取导入数据填充至指定导入DataTable结构中
        /// </summary>
        /// <param name="excelImportConfigEntity"></param>
        /// <param name="excelData"></param>
        private DataTable GetImportData(EntityExcelImportConfig excelImportConfigEntity, DataTable excelData)
        {

            //1、生成数据源结构
            DataTable dbData = GetImortDataTableStructure(excelImportConfigEntity); //存入数据库中的数据结构
            DataRow dr;

            //根据excel中的内容，填充数据
            foreach (DataRow excelRow in excelData.Rows)
            {
                //此处获取map映射数据？
                //获取1000个内容的缓存。字典存储 Dic<string:colName,List<string:valvue>>
                dr = dbData.NewRow();
                foreach (var col in excelImportConfigEntity.m_ColsDefine)
                {//业务规则：1、如果excel没有对应列，则直接使用定义的默认值填充。2、如果excel有对应列，但是为空，那么使用定义的默认值，非空，则使用excel中的值。
                    dr[col.m_DbTableColName] = !excelData.Columns.Contains(col.m_ExcelColName) || string.IsNullOrEmpty(excelRow[col.m_ExcelColName].ToString()) ?
                        col.m_DefalutValue :
                        col.m_HasMapValue ? GetMapData(col.m_MapToTableName, col.m_ExcelColName, col.m_MapToColName, excelRow[col.m_ExcelColName].ToString(), excelData)
                        : excelRow[col.m_ExcelColName];
                }
                dbData.Rows.Add(dr);
            }

            return dbData;
        }

        /// <summary>
        /// 获取excel此列值对应的映射值。从数据库中查询获取。但是有个问题，频繁查询会有性能问题，于是做了缓存，每次按顺序查询一千条记录。缓存每次会覆盖。缓存存在m_MapColData中。
        /// </summary>
        /// <param name="mapTableName"></param>
        /// <param name="srcColName"></param>
        /// <param name="tragetColName"></param>
        /// <param name="excelValue"></param>
        /// <param name="excelAllData"></param>
        /// <returns></returns>
        private string GetMapData(string mapTableName, string excelColName, string tragetColName, string excelValue, DataTable excelAllData)
        {
            if (m_MapColData == null)
            {//全局缓存变量
                m_MapColData = new Dictionary<string, List<Tuple<string, string>>>();
            }

            if (m_MapColData.ContainsKey(excelColName) && m_MapColData[excelColName].Any(c => c.Item1.Equals(excelValue)))
            {
                return m_MapColData[excelColName].First(c => c.Item1.Equals(excelValue)).Item2;
            }
            else
            {
                //查询数据
                int rowIndex = excelAllData.Rows.IndexOf(excelAllData.Rows.Cast<DataRow>().First(c => c[excelColName].ToString().Equals(excelValue)));
                string selectSql = string.Format(@"SELECT {0},{1} FROM {2} WHERE {3} IN ({4})", excelColName, tragetColName, mapTableName, excelColName
                    , string.Join(",", excelAllData.Rows.Cast<DataRow>().Skip(rowIndex).Take(1000).Select(c => "'" + c[excelColName].ToString() + "'")) //查询时每次缓存以后的1000条数据
                    );
                DataTable dt = m_oracleDbManager.Query(selectSql);

                //填充查询数据到缓存
                if (!m_MapColData.ContainsKey(excelColName))
                {
                    m_MapColData.Add(excelColName, new List<Tuple<string, string>>());
                }

                foreach (DataRow dr in dt.Rows)
                {
                    m_MapColData[excelColName].Add(new Tuple<string, string>(dr[excelColName.ToUpper()].ToString(), dr[tragetColName.ToUpper()].ToString()));
                }

                return m_MapColData[excelColName].First(c => c.Item1.Equals(excelValue)).Item2;
            }
        }

        /// <summary>
        /// 获取指定表的列配置信息 从Oracle系统表：USER_TAB_COLS 中获取。
        /// 包含的列：COLUMN_NAME, DATA_TYPE, CHAR_LENGTH,DATA_PRECISION,DATA_SCALE
        /// </summary>
        /// <param name="tableName">数据表名称</param>
        /// <returns></returns>
        private DataTable GetDbColumnDefineInfo(string tableName)
        {
            string selectSql = "SELECT COLUMN_NAME, DATA_TYPE, CHAR_LENGTH, DATA_PRECISION, DATA_SCALE FROM USER_TAB_COLS WHERE TABLE_NAME = :TABLE_NAME";
            OracleParameter[] oracleParameters = new OracleParameter[1];
            oracleParameters[0] = new OracleParameter("TABLE_NAME", tableName);
            DataTable dt = m_oracleDbManager.Query(selectSql, oracleParameters);
            return dt;
        }

        /// <summary>
        /// 根据配置信息和数据库中表的定义，生成导入源数据的DataTable结构
        /// </summary>
        /// <param name="excelImportConfigEntity"></param>
        /// <returns></returns>
        private DataTable GetImortDataTableStructure(EntityExcelImportConfig excelImportConfigEntity)
        {
            DataTable srcDataTable = new DataTable();
            DataTable dbDefineColInfo = GetDbColumnDefineInfo(excelImportConfigEntity.m_TableName);
            DataRow[] dataRows;
            DataColumn dataColumn;
            foreach (var xmlColDefine in excelImportConfigEntity.m_ColsDefine)
            {
                dataRows = dbDefineColInfo.Select(string.Format("COLUMN_NAME = '{0}'", xmlColDefine.m_DbTableColName));
                if (dataRows == null || dataRows.Length != 1)
                {
                    throw new Exception(string.Format("导入excel数据xml【{0}】配置错误。数据库表【{1}】不存在此列【{2}】", excelImportConfigEntity.m_ID, excelImportConfigEntity.m_TableName, xmlColDefine.m_DbTableColName));
                }
                xmlColDefine.m_DataType = GetColType(dataRows[0]["DATA_TYPE"].ToString(), dataRows[0]["DATA_SCALE"].ToString()); //获取列的数据类型（C#代码中的类型）
                xmlColDefine.m_OracleDbType = GetColOracleType(dataRows[0]["DATA_TYPE"].ToString(), dataRows[0]["DATA_SCALE"].ToString()); //获取列的数据类型（Oracle中的类型）
                xmlColDefine.m_DataMaxLength = GetColLength(dataRows[0]["CHAR_LENGTH"].ToString(), dataRows[0]["DATA_PRECISION"].ToString(), dataRows[0]["DATA_SCALE"].ToString()); //获取列最大长度
                dataColumn = new DataColumn(xmlColDefine.m_DbTableColName, xmlColDefine.m_DataType); //创建列 设置列名称、数据类型
                if (xmlColDefine.m_OracleDbType == OracleDbType.Char || xmlColDefine.m_OracleDbType == OracleDbType.NChar
                    || xmlColDefine.m_OracleDbType == OracleDbType.Varchar2 || xmlColDefine.m_OracleDbType == OracleDbType.NVarchar2)
                {//只有字符串类型设置最大长度。进行长度校验
                    dataColumn.MaxLength = xmlColDefine.m_DataMaxLength;
                }
                srcDataTable.Columns.Add(dataColumn);
            }

            return srcDataTable;
        }

        /// <summary>
        /// 获取数据列的数据类型
        /// </summary>
        /// <param name="typeName">Oracle中类型名称</param>
        /// <param name="scale">Oracle数字类型时，数字小数点位数</param>
        /// <returns>C#内置数据类型</returns>
        private Type GetColType(string typeName, string scale = "")
        {
            typeName = typeName.ToUpper();
            switch (typeName)
            {
                case "CLOB":
                case "VARCHAR":
                case "VARCHAR2":
                case "NVARCHAR2":
                case "CHAR":
                    return typeof(string);
                case "DATE":
                    return typeof(DateTime);
                case "NUMBER":
                    return string.IsNullOrEmpty(scale) ? typeof(int) : typeof(double);
                case "BLOB":
                    return typeof(byte[]);
                default:
                    return typeof(string);
            }
        }

        /// <summary>
        /// 获取列对应的Oracle数据库类型
        /// </summary>
        /// <param name="typeName"></param>
        /// <param name="scale"></param>
        /// <returns></returns>
        private OracleDbType GetColOracleType(string typeName, string scale = "")
        {
            typeName = typeName.ToUpper();
            switch (typeName)
            {
                case "CLOB":
                    return OracleDbType.Clob;
                case "VARCHAR":
                    return OracleDbType.Varchar2;
                case "VARCHAR2":
                    return OracleDbType.Varchar2;
                case "NVARCHAR2":
                    return OracleDbType.NVarchar2;
                case "CHAR":
                    return OracleDbType.Char;
                case "DATE":
                    return OracleDbType.Date;
                case "NUMBER":
                    return string.IsNullOrEmpty(scale) ? OracleDbType.Int32 : OracleDbType.Double;
                case "BLOB":
                    return OracleDbType.Blob;
                default:
                    return OracleDbType.Varchar2;
            }
        }

        /// <summary>
        /// 获取列的最大数据长度
        /// </summary>
        /// <param name="dataLength">数据长度</param>
        /// <param name="dataPrecision">数字类型时，数字总位数</param>
        /// <param name="dataScale">数字类型时，数字小数点位数</param>
        /// <returns></returns>
        private int GetColLength(string dataLength, string dataPrecision, string dataScale)
        {
            int dataLengthNum = Convert.ToInt32(dataLength);
            if (dataLengthNum == 0)
            {//非字符串类型：CHAR、VARCHAR2、NCHAR、NVARCHAR2
                if (!string.IsNullOrEmpty(dataPrecision))
                {//数字类型
                    int dataPrecisionNum = Convert.ToInt32(dataPrecision);
                    int dataScaleNum = string.IsNullOrEmpty(dataScale) ? 0 : Convert.ToInt32(dataScale);
                    return dataPrecisionNum - dataScaleNum;
                }
                else
                {//非字符串、非数字类型 Date、CLOB、BOLB等
                    return Int32.MaxValue;
                }
            }
            else
            {//字符串类型：CHAR、VARCHAR2、NCHAR、NVARCHAR2
                return dataLengthNum;
            }
        }

        /// <summary>
        /// 将数据保存到Oracle数据库
        /// 1万条数据直接提交到数据库，超过一万条数据，分批提交，每次最多提交1万条。
        /// </summary>
        private void SavaDataToOracleDb(DataTable dataTable, string sql, OracleParameter[] oracleParameters)
        {
            //TODO:excel导入 大量数据分批提交
        }
        #endregion
    }
}
