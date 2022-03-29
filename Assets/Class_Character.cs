using System.Collections;
using System.Collections.Generic;
using UnityEngine;

/// <summary>
/// Class_Character: ScriptableObject
/// </summary>
public class Class_Character : ScriptableObject
{
    ///<summary>
    /// 性別
    ///</summary>
    public enum Sex
    {
        Female,
        Male,
        Unknown,
    }


    /// <summary>
    /// Wrapper for JsonUtility
    /// </summary>
    [System.Serializable]
    public class Wrapper
    {
        public List<Row> Rows = new List<Row>();
    }

    ///<summary>
    /// Table Row
    ///</summary>
    [System.Serializable]
    public class Row
    {
        ///<summary>
        /// [ID:必須]
        ///</summary>
        public int          ID;
        public int          CharacterNo;
        public string       CharacterName;
        public float        Size;
        public List<string> FriendName = new List<string>();
        public Sex          Sex;
    }

    /// <summary>
    /// データカウント
    /// </summary>
    public int Count
    {
        get
        {
            return Rows == null ? 0 : Rows.Count;
        }
    }

    /// <summary>
    /// テーブルデータ
    /// </summary>
    public List<Row> Rows;
    Dictionary<int, Row> IDRows;

    /// <summary>
    /// factory: create instance
    /// </summary>
    /// <returns></returns>
    public static Class_Character CreateInstance()
    {
        return Object.Instantiate(ScriptableObject.CreateInstance<Class_Character>());
    }

    /// <summary>
    /// .ctor
    /// </summary>
    public Class_Character()
    {
        Rows = new List<Row>();
    }
    
    /// <summary>
    /// set json data
    /// </summary>
    public void SetRowsByJson(string json)
    {
        Rows = JsonUtility.FromJson<Wrapper>(json).Rows;
        IDRows = null;
    }

    /// <summary>
    /// 行を配列順で取得する
    /// </summary>
    /// <param name="index">配列番号</param>
    public Row GetRow(int index)
    {
        if (index < 0 || index >= Rows.Count)
        {
            Debug.LogError($"XlsToJson_ClassTemplate: out of range. index = {index}");
            return null;
        }
        return Rows[index];
    }

    /// <summary>
    /// ID
    /// </summary>
    public Row FindRowByID(int val, bool errorLog = true)
    {
        if (IDRows == null)
        {
            IDRows = new Dictionary<int, Row>();
            Rows.ForEach( (row) => IDRows.Add(row.ID, row) );
        }
        if (IDRows.ContainsKey(val) == false)
        {
            if (errorLog == true)
            {
                Debug.LogError($"cannot find: {val}");
            }
            return null;
        }
        return IDRows[val];
    }

}
