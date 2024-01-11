using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public enum StageTime
{
    Morning,
    Noon,
    Night,
}

/// <summary>
/// Class_Stage: ScriptableObject
/// </summary>
public class Class_Stage : ScriptableObject
{
    public enum AMPM
    {
        AM,
        PM,
    }

    public const float  ALPHA_DEFAULT = 1f;
    public const float  ALPHA_ZERO    = 0f;
    ///<summary>
    /// BG_MORNING = "bg01A"
    ///</summary>
    public const string BG_MORNING    = "bg01A";
    public const string BG_NOON       = "bg01B";
    public const string BG_NIGHT      = "bg01C";

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
        /// [ID:required]
        ///</summary>
        public int       ID;
        ///<summary>
        /// Name
        ///</summary>
        public string    BgName;
        ///<summary>
        /// PPS Profile
        ///</summary>
        public string    ProfileName;
        ///<summary>
        /// Light
        ///</summary>
        public bool      Brightness;
        ///<summary>
        /// AM or PM
        ///</summary>
        public AMPM      AMPM;
        public StageTime StageTime;
        public SubClass  SubClass = new SubClass();
        ///<summary>
        /// α
        ///</summary>
        public float     Alpha;
    }

    [System.Serializable]
    public class SubClass
    {
        public int       Member0;
        public int       Member1;
        public Child     Child = new Child();
    }

    [System.Serializable]
    public class Child
    {
        public int       Member0;
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
    public static Class_Stage CreateInstance()
    {
        return Object.Instantiate(ScriptableObject.CreateInstance<Class_Stage>());
    }

    /// <summary>
    /// .ctor
    /// </summary>
    public Class_Stage()
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
            Rows.ForEach( (row) => { if (IDRows.ContainsKey(row.ID) == false) IDRows.Add(row.ID, row); } );
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
