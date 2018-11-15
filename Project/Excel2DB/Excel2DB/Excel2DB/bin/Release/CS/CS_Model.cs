using UnityEngine;
using System.Collections;
using System.Collections.Generic;
using Mono.Data.Sqlite;
public class CS_Model
{
    public class DataEntry
    {
        public System.Int32 _ID = 0;
        public System.String _Name = "";
        public System.String _Icon = "";
        public System.Single _Move = 0;
        public Vector3 _Scale = Vector3.zero;
        public System.Int32 _Volume = 0;
        public System.Boolean _Enable = false;
    }
    public Dictionary<int, DataEntry> m_kDataEntryTable = new Dictionary<int, DataEntry>();
    public void Init()
    {
        m_kDataEntryTable.Clear();
        System.String kSqlCMD = "SELECT * FROM Model";
        m_kDataEntryTable.Clear();
        SqliteDataReader kDataReader = DBManager.Instance.Query(kSqlCMD);
        string[] v3int = null;
        while (kDataReader.HasRows && kDataReader.Read())
        {
            DataEntry kNewEntry = new DataEntry();
            kNewEntry._ID = kDataReader.GetInt32(0);
            kNewEntry._Name = kDataReader.GetString(1);
            kNewEntry._Icon = kDataReader.GetString(2);
            kNewEntry._Move = kDataReader.GetFloat(3);
            v3int = kDataReader.GetString(4).Split(' ');
            kNewEntry._Scale = new Vector3(float.Parse(v3int[0]), float.Parse(v3int[1]), float.Parse(v3int[2]));
            kNewEntry._Volume = kDataReader.GetInt32(5);
            kNewEntry._Enable = kDataReader.GetBoolean(6);
            m_kDataEntryTable[kNewEntry._ID] = kNewEntry;
        }
        kDataReader.Close();
    }
    public DataEntry GetEntryPtr(int iID)
    {
        if (m_kDataEntryTable.ContainsKey(iID))
        {
            return m_kDataEntryTable[iID];
        }
        return null;
    }
    public bool ContainsID(int iID)
    {
        return m_kDataEntryTable.ContainsKey(iID);
    }
}
