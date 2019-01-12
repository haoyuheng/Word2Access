using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ADOX;
using System.IO;

namespace Word2Access
{
    public static class AccessDbHelper
    {
        /// <summary>
        /// 创建access数据库
        /// </summary>
        /// <param name="filePath">数据库文件的全路径，如 D:\\NewDb.mdb</param>
        public static bool CreateAccessDb(string filePath)
        {
            ADOX.Catalog catalog = new Catalog();
            if (!File.Exists(filePath))
            {
                try
                {
                    catalog.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Jet OLEDB:Engine Type=5");
                }
                catch (System.Exception ex)
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// 在access数据库中创建表
        /// </summary>
        /// <param name="filePath">数据库表文件全路径如D:\\NewDb.mdb 没有则创建 </param> 
        /// <param name="tableName">表名</param>
        /// <param name="colums">ADOX.Column对象数组</param>
        public static bool CreateAccessTable(string filePath, string tableName, params ADOX.Column[] colums)
        {
            ADOX.Catalog catalog = new Catalog();
            //数据库文件不存在则创建
            if (!File.Exists(filePath))
            {
                try
                {
                    catalog.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Jet OLEDB:Engine Type=5");
                }
                catch (System.Exception ex)
                {
                    return false;
                }
            }
            ADODB.Connection cn = new ADODB.Connection();
            cn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath, null, null, -1);
            catalog.ActiveConnection = cn;
            ADOX.Table table = new ADOX.Table();
            table.Name = tableName;
            foreach (var column in colums)
            {
                table.Columns.Append(column);
            }
            colums[0].ParentCatalog = catalog;
            colums[0].Properties["AutoIncrement"].Value = true; //设置自动增长
            table.Keys.Append("FirstTablePrimaryKey", KeyTypeEnum.adKeyPrimary, colums[0], null, null); //定义主键
            catalog.Tables.Append(table);
            cn.Close();
            return true;
        }
        //========================================================================================调用
        //ADOX.Column[] columns = {
        //                     new ADOX.Column(){Name="id",Type=DataTypeEnum.adInteger,DefinedSize=9},
        //                     new ADOX.Column(){Name="col1",Type=DataTypeEnum.adWChar,DefinedSize=50},
        //                     new ADOX.Column(){Name="col2",Type=DataTypeEnum.adLongVarChar,DefinedSize=50}
        //                 };
        // AccessDbHelper.CreateAccessTable("d:\\111.mdb", "testTable", columns);
    }
}