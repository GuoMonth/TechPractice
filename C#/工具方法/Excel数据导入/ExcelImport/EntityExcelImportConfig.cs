using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace Excel数据导入.ExcelImport
{
    /// <summary>
    /// 郭升
    /// 2020年5月19日
    /// </summary>
    internal class EntityExcelImportConfig
    {
        /*
         XML配置示例：
         <ECXEL_IMPORT ID="测试1" TABLE_NAME="Test_A">
            <COL_DEFINE>
                <COL EXCEL_COL_NAME="ID"  DBTABLE_COL_NAME="ID" DEFAULT_VALUE="${ID}"></COL> <!--匹配${ID} \$\{.*?\} -->
                <COL EXCEL_COL_NAME="工号"  DBTABLE_COL_NAME="DOC_ID" MAPTO="${USERINFO}.{ID}"></COL> <!--匹配.{ID} \.\{.*?\} -->
                <COL EXCEL_COL_NAME=""  DBTABLE_COL_NAME="CONTAINER_ID" DEFAULT_VALUE="divInfo"></COL> <!--添加的默认值列-->
                <COL EXCEL_COL_NAME=""  DBTABLE_COL_NAME="DELETE_FLAG" DEFAULT_VALUE="0"></COL>
			    <COL EXCEL_COL_NAME=""  DBTABLE_COL_NAME="UPDATE_FLAG" DEFAULT_VALUE="i"></COL>
			    <COL EXCEL_COL_NAME=""  DBTABLE_COL_NAME="AUTHOR" DEFAULT_VALUE="admin(12345)"></COL>
			    <COL EXCEL_COL_NAME=""  DBTABLE_COL_NAME="AUDITING" DEFAULT_VALUE="0"></COL>
			    <COL EXCEL_COL_NAME="创建时间"  DBTABLE_COL_NAME="CREATE_TIME" DEFAULT_VALUE="${CURRENT_TIME}"></COL>
			    <COL EXCEL_COL_NAME=""  DBTABLE_COL_NAME="MENU_ID" DEFAULT_VALUE="20160802152932914-efe53dadc65a"></COL>
			    <COL EXCEL_COL_NAME="入库时间"  DBTABLE_COL_NAME="入库时间"></COL>
			    <COL EXCEL_COL_NAME="出库时间"  DBTABLE_COL_NAME="出库时间"></COL>
			    <COL EXCEL_COL_NAME="备注"  DBTABLE_COL_NAME="备注"></COL>
            </COL_DEFINE>
            <AFTER_IMPORT_CALL_BACK>FUNCTION_NAME</AFTER_IMPORT_CALL_BACK>
        </ECXEL_IMPORT>
        
        1、<ECXEL_IMPORT>节点说明
            一个excel的导入配置单元
            属性： 
                ID：导入配置唯一识别号，不能重复，据此调用对应的导入配置。
                TABLE_NAME：对应的数据库表名
        2、<COL_DEFINE>节点说明
            定义excel中列和数据库表字段的映射关系，默认进行数据类型检查
            属性：
            EXCEL_COL_NAME：（可缺省）excel中列名。如果此属性为空。说明此列在excel中没有对应的列，使用默认值进行数据导入。如果没有默认值，则导入空数据。
            DBTABLE_COL_NAME：对应数据库表的列名。
            DEFAULT_VALUE：（可缺省）定义列的默认值。
                1)变量默认值。如果其值为${ID}、${CURRENT_TIME} 这种形式，那么就用系统内部运算出来的值进行替换。（目前只支持ID、CURRENT_TIME。其中ID为GUID随机生成，CURRENT_TIME为当前系统时间）
                2)固定默认值。直接定义为DEFAULT_VALUE="divInfo"，那么此列的空值都将用“divInfo”进行填充。
            MAPTO：（可缺省）映射值定义。${USERINFO}.{ID} 意思为此列值使用对应的userinfo表的ID字段内容进行填充。
        3、<AFTER_IMPORT_CALL_BACK>节点 
            回调函数定义。导入完成后回调。此节点可缺省，节点值为导入完成后回调函数的名称。
         */
        private XmlElement m_xmlElement;

        #region 属性        

        /// <summary>
        /// excel导入配置的xml节点
        /// </summary>
        private XmlElement m_XmlElement
        {
            get
            {
                return m_xmlElement;
            }
        }

        /// <summary>
        /// 获取菜单名称
        /// </summary>
        public string m_ID
        {
            get
            {
                return m_XmlElement.GetAttribute("ID");
            }
        }

        /// <summary>
        /// 获取导入对应的数据库表名称
        /// </summary>
        public string m_TableName
        {
            get
            {
                return m_XmlElement.GetAttribute("TABLE_NAME");
            }
        }

        /// <summary>
        /// 列集合定义 私有变量缓存
        /// </summary>
        private IEnumerable<ExcelImportColDefineEntity> m_colsDefine;

        /// <summary>
        /// 列定义集合
        /// </summary>
        public IEnumerable<ExcelImportColDefineEntity> m_ColsDefine
        {
            get
            {
                if (m_colsDefine == null)
                {
                    m_colsDefine = m_XmlElement.SelectNodes("COL_DEFINE/COL").Cast<XmlElement>().Select(c => new ExcelImportColDefineEntity(c)).ToList();
                }
                return m_colsDefine;
            }
        }

        /// <summary>
        /// 导入完成后调用
        /// 节点值：AFTER_IMPORT_CALL_BACK
        /// </summary>
        public string m_AfterImportCallBack
        {
            get
            {
                return m_XmlElement.SelectSingleNode("AFTER_IMPORT_CALL_BACK").Value;
            }
        }

        #endregion

        #region 构造函数

        /// <summary>
        /// 有参构造函数 【ECXEL_IMPORT】 节点
        /// </summary>
        /// <param name="xe"></param>
        public EntityExcelImportConfig(XmlElement xe)
        {
            m_xmlElement = xe;
        }

        #endregion

        #region 方法

        /// <summary>
        /// 重写 xml化
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {

            return string.Format("<ECXEL_IMPORT ID=\"{0}\" TABLE_NAME=\"{1}\"><COL_DEFINE>{2}</COL_DEFINE><AFTER_IMPORT_CALL_BACK>{3}</AFTER_IMPORT_CALL_BACK></ECXEL_IMPORT>"
                , m_ID, m_TableName
                , string.Join("", m_ColsDefine.Select(c=>c.ToString()))
                , m_AfterImportCallBack
                );
        }

        #endregion

    }

    #region 一个列的定义
    /// <summary>
    /// excel导入一个列的定义
    /// </summary>
    internal class ExcelImportColDefineEntity
    {
        private XmlElement m_xmlElement;

        /// <summary>
        /// 匹配默认值字符串内容 懒惰匹配
        /// </summary>
        private Regex m_GetDefalutValueStr = new Regex(@"\$\{.*?\}");

        /// <summary>
        /// 匹配映射字段MAPTO中定义的表名
        /// </summary>
        private Regex m_GetMapToTableNameStr = new Regex(@"\$\{.*?\}");

        /// <summary>
        /// 匹配映射字段MAPTO中定义的列名
        /// </summary>
        private Regex m_GetMapToColNameStr = new Regex(@"\.\{.*?\}");

        #region 属性

        /// <summary>
        /// 一个列定义的节点
        /// </summary>
        public XmlElement m_XmlElement
        {
            get
            {
                return m_xmlElement;
            }
        }

        /// <summary>
        /// excel中的列名
        /// </summary>
        public string m_ExcelColName
        {
            get
            {
                return m_XmlElement.GetAttribute("EXCEL_COL_NAME");
            }
        }

        /// <summary>
        /// 数据库中的列名
        /// </summary>
        public string m_DbTableColName
        {
            get
            {
                return m_XmlElement.GetAttribute("DBTABLE_COL_NAME");
            }
        }

        /// <summary>
        /// 代码中列的数据类型
        /// </summary>
        public Type m_DataType
        {
            get; set;
        }

        /// <summary>
        /// Oracle数据库中列的数据类型
        /// </summary>
        public OracleDbType m_OracleDbType
        {
            get; set;
        }

        /// <summary>
        /// 列数据最大长度
        /// </summary>
        public int m_DataMaxLength
        {
            get; set;
        }

        /// <summary>
        /// 此列是否定义了默认值
        /// </summary>
        public bool m_HasDefalutValue
        {
            get
            {
                return m_XmlElement.HasAttribute("DEFAULT_VALUE");
            }
        }

        /// <summary>
        /// 获取此列定义的默认值
        /// </summary>
        public object m_DefalutValue
        {
            get
            {
                string str = m_XmlElement.GetAttribute("DEFAULT_VALUE");
                return GetDefaultValue(str);
            }
        }

        /// <summary>
        /// 此列是否定义了映射值
        /// </summary>
        public bool m_HasMapValue
        {
            get
            {
                return m_XmlElement.HasAttribute("MAPTO");
            }
        }

        /// <summary>
        /// 获取定义映射值的表名
        /// </summary>
        public string m_MapToTableName
        {
            get
            {
                return m_HasMapValue ? m_GetMapToTableNameStr.Match(m_XmlElement.GetAttribute("MAPTO")).Value.Trim('$', '{', '}') : string.Empty;
            }
        }

        /// <summary>
        /// 获取定义映射值对应的列名
        /// </summary>
        public string m_MapToColName
        {
            get
            {
                return m_HasMapValue ? m_GetMapToColNameStr.Match(m_XmlElement.GetAttribute("MAPTO")).Value.Trim('.', '{', '}') : string.Empty;
            }
        }

        #endregion

        #region 构造函数

        /// <summary>
        /// 有参构造函数 【ECXEL_IMPORT/COL_DEFINE/COL】节点
        /// </summary>
        /// <param name="xe"></param>
        public ExcelImportColDefineEntity(XmlElement xe)
        {
            m_xmlElement = xe;
        }

        #endregion

        #region 方法

        /// <summary>
        /// 获取默认值
        /// 1、变量默认值，其格式为 ${ID}、${CURRENT_TIME} 这种形式，使用正则表达式进行判断。
        /// 2、固定默认值。举例：直接定义DEFAULT_VALUE="divInfo"，那么此列的空值都将是“divInfo”
        /// </summary>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        private object GetDefaultValue(string defaultValue)
        {
            if (m_GetDefalutValueStr.IsMatch(defaultValue))
            {//变量默认值
                string defaultValueStr = m_GetDefalutValueStr.Match(defaultValue).Value.Trim('$', '{', '}');
                EnumDefaultValue excelImportDefaultValue;
                if (Enum.TryParse<EnumDefaultValue>(defaultValueStr, out excelImportDefaultValue))
                {//转换为枚举类型进行匹配
                    switch (excelImportDefaultValue)
                    {
                        case EnumDefaultValue.NONE:
                            return DBNull.Value;
                        case EnumDefaultValue.ID:
                            return Guid.NewGuid().ToString();
                        case EnumDefaultValue.CURRENT_TIME:
                            return DateTime.Now;
                        default:
                            return DBNull.Value;
                    }
                }
                else
                {//异常 未定义的变量默认值
                    throw new ArgumentException(string.Format("【{0}】解析为【{1}】，系统无对应的枚举EnumDefaultValue对象", defaultValue, defaultValueStr));
                }
            }
            else
            {//固定默认值
                return defaultValue;
            }
        }

        /// <summary>
        /// 返回xml格式字符串
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return string.Format("<COL EXCEL_COL_NAME = \"{0}\"  DBTABLE_COL_NAME=\"{1}\"  DATA_TYPE=\"{2}\" DEFAULT_VALUE=\"{3}\"><COL/>"
                , m_ExcelColName, m_DbTableColName, m_DataType.ToString(), m_DefalutValue);
        }

        #endregion

    }
    #endregion
}
