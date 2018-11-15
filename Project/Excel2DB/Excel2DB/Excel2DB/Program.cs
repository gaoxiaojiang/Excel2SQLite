using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.SQLite;
using System.IO;
using System.Text;

namespace Excel2DB
{
    class Program
    {
        static string AppPath = System.Environment.CurrentDirectory;
        static string DBName = "GameDB";
        static string CsDirectoryName = "CS";
        static string CsPath = "";
        static SQLiteDBHelper dbHelper;

        static void Log(string _mes)
        {
            Console.WriteLine(_mes);
        }



        static void Main(string[] args)
        {
            ExistsCSDirectory();

            string dbName = CreatDBFile();

            dbHelper = new SQLiteDBHelper(dbName);

            var files = Directory.GetFiles(Path.Combine(Path.GetDirectoryName(AppPath), "Excel"), "*.xlsx");
            foreach (var file in files)
            {
                if (file.Contains("~"))
                    continue;
                Log($"开始转换{Path.GetFileNameWithoutExtension(file)}...");
                Change2DB(dbHelper, file);
            }

            Log("DB生成完成...");
            Console.ReadKey();
        }



        /// <summary>
        /// 操作cs生成目录
        /// </summary>
        static void ExistsCSDirectory()
        {
            CsPath = Path.Combine(Path.GetDirectoryName(AppPath), CsDirectoryName);
            if (Directory.Exists(CsPath))
            {
                Directory.Delete(CsPath, true);
            }
            Directory.CreateDirectory(CsPath);
        }




        ///1.创建db文件
        static string CreatDBFile()
        {
            string db = Path.Combine(Path.GetDirectoryName(AppPath), DBName, DBName + ".db3");
            try
            {
                if (Directory.Exists(Path.GetDirectoryName(db)))
                {
                    Directory.Delete(Path.GetDirectoryName(db), true);
                }

                Directory.CreateDirectory(Path.GetDirectoryName(db));

                SQLiteConnection.CreateFile(db);
            }
            catch (Exception _e)
            {
                Log($"删除旧的DB或创建新的异常：{_e }");
            }
            return db;
        }



        ///2转DB
        static void Change2DB(SQLiteDBHelper sqliteDBHelper, string ExcelPath)
        {
            IWorkbook workbook = null;
            FileStream fileStream = new FileStream(ExcelPath, FileMode.Open, FileAccess.Read);
            string tableName = Path.GetFileNameWithoutExtension(ExcelPath);
            if (ExcelPath.IndexOf(".xlsx") > 0) // 2007版本  
            {
                workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook  
            }
            else if (ExcelPath.IndexOf(".xls") > 0) // 2003版本  
            {
                workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook  
            }
            else
            {
                return;
            }
            ISheet sheet = workbook.GetSheetAt(0);
            if (sheet != null)
            {
                //创建标投
                IRow rowItem1 = sheet.GetRow(2);
                IRow rowItem2 = sheet.GetRow(3);

                List<string> ColumnsName = new List<string>();
                List<string> ColumnsType = new List<string>();


                List<string> CSType = new List<string>();
                int cellCount = rowItem1.LastCellNum;
                for (int k = 0; k < cellCount; k++)
                {
                    ICell cell = rowItem1.GetCell(k);
                    ICell cel2 = rowItem2.GetCell(k);
                    if (cell != null)
                    {
                        string cellValue = cell.StringCellValue;
                        ColumnsName.Add(cel2.StringCellValue);
                        ColumnsType.Add(CS2DbType(cell.StringCellValue));
                        CSType.Add(GetCSharpType(cell.StringCellValue));
                    }            
                }
                sqliteDBHelper.CreateTable(tableName, ColumnsName.ToArray(), ColumnsType.ToArray());

                using (DbTransaction transaction = sqliteDBHelper.connection.BeginTransaction())
                {
                    using (DbCommand cmd = sqliteDBHelper.connection.CreateCommand())
                    {
                        int startRow = 4;
                        int rowCount = sheet.LastRowNum;
                        for (int i = startRow; i <= rowCount; i++)
                        {
                            IRow rowItem = sheet.GetRow(i);
                            List<string> ContentStr = new List<string>();
                            bool caInsert = true;
                            for (int k = 0; k < rowItem.LastCellNum; k++)
                            {
                                ICell cellItem = rowItem.GetCell(k);
                                if (cellItem != null && !string.IsNullOrEmpty( cellItem.ToString()))
                                {
                                    ContentStr.Add(cellItem.ToString());
                                }
                                else
                                {
                                    caInsert = false;
                                    break;
                                }
                            }
                            if (caInsert)
                            {
                                //解决SQL注入漏洞  cmd是SqlCommand对象
                                cmd.CommandText = sqliteDBHelper.GetInsertValueStr(tableName, ColumnsName.ToArray());// @"select count(*) from UserInfo where UserName=@UserName and UserPwd=@UserPwd";
                  
                                for (int k = 0; k < ColumnsName.Count; k++)
                                {
                                    DbParameter para = cmd.CreateParameter();
                                    para.ParameterName="@"+ColumnsName[k]+"";  //SQL参数化
                                    para.Value = ContentStr[k];  //SQL参数化
                                    cmd.Parameters.Add(para);
                                }
                                cmd.ExecuteScalar();
                            }
                        }
                    }
                    transaction.Commit();
                }
                CreatCs(Path.GetFileNameWithoutExtension(ExcelPath), Path.Combine(CsPath, "CS_" + Path.GetFileNameWithoutExtension(ExcelPath) + ".cs"), ColumnsName.ToArray(), CSType);
            }
        }

        /// <summary>
        /// 创建C#代码文件
        /// </summary>
        /// <param name="TabelName"></param>
        /// <param name="classPath"></param>
        /// <param name="ColumnName"></param>
        /// <param name="_type"></param>
        static void CreatCs(string TabelName, string classPath, string[] ColumnName, List<string> _type)
        {
            string className = Path.GetFileNameWithoutExtension(classPath);
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("using UnityEngine;");
            sb.AppendLine("using System.Collections;");
            sb.AppendLine("using System.Collections.Generic;");
            sb.AppendLine("using Mono.Data.Sqlite;");

            sb.AppendLine($"public class {className}");
            sb.AppendLine("{");
            sb.AppendLine("    public class DataEntry");
            sb.AppendLine("    {");
            List<string> ColumnsName = new List<string>();
            for (int i = 0; i < _type.Count; i++)
            {
                string _t = _type[i];

                sb.AppendLine($"        public { _t} {ColumnName[i]} = {GetTypeDefault(_t)};");
            }
            sb.AppendLine("    }");

            sb.AppendLine($"    public Dictionary<int, DataEntry> m_kDataEntryTable = new Dictionary<int, DataEntry>();");


            sb.AppendLine("    public void Init()");
            sb.AppendLine("    {");
            sb.AppendLine("        m_kDataEntryTable.Clear();");
            sb.AppendLine($"        System.String kSqlCMD = \"SELECT * FROM {TabelName}\";");
            sb.AppendLine("        m_kDataEntryTable.Clear();");
            sb.AppendLine("        SqliteDataReader kDataReader = DBManager.Instance.Query(kSqlCMD);");

            if (_type.Contains("Vector3"))
            {
                sb.AppendLine($"        string[] v3int = null;");
            }

            sb.AppendLine("        while (kDataReader.HasRows && kDataReader.Read())");
            sb.AppendLine("        {");
            sb.AppendLine($"            DataEntry kNewEntry = new DataEntry();");
            for (int i = 0; i < _type.Count; i++)
            {
                string qian = "kNewEntry." + ColumnName[i];
                if (_type[i] == "System.Int32")
                {
                    sb.AppendLine($"            {qian} = kDataReader.GetInt32({i});");
                }
                else if (_type[i] == "System.String")
                {
                    sb.AppendLine($"            {qian} = kDataReader.GetString({i});");
                }
                else if (_type[i] == "System.Boolean")
                {
                    sb.AppendLine($"            {qian} = kDataReader.GetBoolean({i});");
                }
                else if (_type[i] == "System.Single")
                {
                    sb.AppendLine($"            {qian} = kDataReader.GetFloat({i});");
                }
                else if (_type[i] == "System.Int64")
                {
                    sb.AppendLine($"            {qian} = kDataReader.GetInt64({i});");
                }
                else if (_type[i] == "System.Decimal")
                {
                    sb.AppendLine($"            {qian} = kDataReader.GetDecimal({i});");
                }
                else if (_type[i] == "Vector3")
                {
                    sb.AppendLine($"            v3int = kDataReader.GetString({i}).Split(' ');");
                    sb.AppendLine($"            {qian} = new Vector3(float.Parse(v3int[0]), float.Parse(v3int[1]), float.Parse(v3int[2]));");
                }
            }

            sb.AppendLine("            m_kDataEntryTable[kNewEntry._ID] = kNewEntry;");
            sb.AppendLine("        }");
            sb.AppendLine("        kDataReader.Close();");

            sb.AppendLine("    }");

            sb.AppendLine($"    public DataEntry GetEntryPtr(int iID)");
            sb.AppendLine("    {");
            sb.AppendLine("        if (m_kDataEntryTable.ContainsKey(iID))");
            sb.AppendLine("        {");
            sb.AppendLine("            return m_kDataEntryTable[iID];");
            sb.AppendLine("        }");
            sb.AppendLine("        return null;");
            sb.AppendLine("    }");

            sb.AppendLine("    public bool ContainsID(int iID)");
            sb.AppendLine("    {");
            sb.AppendLine("        return m_kDataEntryTable.ContainsKey(iID);");
            sb.AppendLine("    }");
            sb.AppendLine("}");
            File.WriteAllText(classPath, sb.ToString());
        }

        /// <summary>
        /// 每种类型的默认值
        /// </summary>
        /// <param name="_type"></param>
        /// <returns></returns>
        private static string GetTypeDefault(string _type)
        {
            string DbType = "";

            if (_type == "System.Int32" || _type == "System.Single" || _type == "System.Decimal")
            {
                DbType = "0";
            }
            else if (_type == "System.String")
            {
                DbType = "\"\"";
            }
            else if (_type == "System.Boolean")
            {
                DbType = "false";
            }
            else if (_type == "Vector3")
            {
                DbType = "Vector3.zero";
            }
            else
            {
                Log($"没有找到合适的类型：{DbType}");
            }
            return DbType;
        }

        /// <summary>
        /// 转化为为C#对应的类型
        /// </summary>
        /// <param name="_type"></param>
        /// <returns></returns>
        private static string GetCSharpType(string _type)
        {
            string DbType = "";
            if (_type == "Int")
            {
                DbType = "System.Int32";
            }
            else if (_type == "String")
            {
                DbType = "System.String";
            }
            else if ( _type == "Vector3")
            {
                DbType = "Vector3";
            }
            else if (_type == "Decimal")
            {
                DbType = "System.Decimal";
            }
            else if (_type == "Boolean")
            {
                DbType = "System.Boolean";
            }
            else if (_type == "Single")
            {
                DbType = "System.Single";
            }
            else
            {
                Log($"没有找到合适的类型：{_type}");
            }
            return DbType;
        }

        /// <summary>
        /// 把类型转化为数据库类型
        /// </summary>
        /// <param name="_type"></param>
        /// <returns></returns>
        private static string CS2DbType(string _type)
        {
            string DbType = "";
            if (_type == "Int")
            {
                DbType = "INT";
            }
            else if (_type == "String" || _type == "Vector3")
            {
                DbType = "Text";
            }
            else if (_type == "Decimal")
            {
                DbType = "DECIMAL";
            }
            else if (_type == "Boolean")
            {
                DbType = "BOOL";
            }
            else if (_type == "Single")
            {
                DbType = "REAL";
            }
            else
            {
                Log($"没有找到合适的类型：{_type}");
            }
            return DbType;
        }
    }
}
