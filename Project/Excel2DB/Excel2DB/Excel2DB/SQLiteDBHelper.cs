using System.Data.SQLite;

/// <summary> 
/// 生成DB的一个类
/// </summary> 
public class SQLiteDBHelper
{
    private string connectionString = string.Empty;
    public SQLiteConnection connection;
    /// <summary> 
    /// 构造函数 
    /// </summary> 
    /// <param name="dbPath">SQLite数据库文件路径</param> 
    public SQLiteDBHelper(string dbPath)
    {
        this.connectionString = "Data Source=" + dbPath;
        connection = new SQLiteConnection(connectionString);
            connection.Open();
    }

    public int CreateTable(string tableName, string[] colNames, string[] colTypes)
    {
        string queryString = "CREATE TABLE IF NOT EXISTS " + tableName + "( " + colNames[0] + " " + colTypes[0];
        for (int i = 1; i < colNames.Length; i++)
        {
            queryString += ", " + colNames[i] + " " + colTypes[i];
        }
        queryString += "  ) ";
        return ExecuteNonQuery(queryString,null);
    }

    /// <summary>
    /// 向指定数据表中插入数据
    /// </summary>
    /// <returns>The values.</returns>
    /// <param name="tableName">数据表名称</param>
    /// <param name="values">插入的数值</param>
    public string GetInsertValueStr(string tableName, string[] values)
    {
        string queryString = $"INSERT INTO {tableName} VALUES ( {"@"+ values[0]} ";
        for (int i = 1; i < values.Length; i++)
        {
            queryString += ", " + "@" + values[i];
        }
        queryString += " )";
        return queryString;
    }

    /// <summary> 
    /// 对SQLite数据库执行增删改操作，返回受影响的行数。 
    /// </summary> 
    /// <param name="sql">要执行的增删改的SQL语句</param> 
    /// <param name="parameters">执行增删改语句所需要的参数，参数必须以它们在SQL语句中的顺序为准</param> 
    /// <returns></returns> 
    public int ExecuteNonQuery(string sql, SQLiteParameter[] parameters)
    {
        int affectedRows = 0;

   
            using (SQLiteCommand command = new SQLiteCommand(connection))
            {
                command.CommandText = sql;
                if (parameters != null)
                {
                    command.Parameters.AddRange(parameters);
                }
                affectedRows = command.ExecuteNonQuery();
            }

        return affectedRows;

    }

}
