using System.Collections;
using System.Collections.Generic;
using UnityEngine;
#if UNITY_EDITOR
using System.Linq;
using UnityEditor;
#endif

/// <summary>
/// Stage: ScriptableObject singleton accessor
/// </summary>
public partial class Stage : MonoBehaviour, iXlsToJsonAccessor
{
    [SerializeField]
    Class_Stage        Table;

    /// <summary>
    /// データカウント
    /// </summary>
    public static int Count
    {
        get
        {
            return table.Count;
        }
    }
    /// <summary>
    /// 全行を返す
    /// </summary>
    public static List<Class_Stage.Row> Rows
    {
        get
        {
            return table.Rows;
        }
    }

    static Class_Stage table;
#if UNITY_EDITOR
    static int lastId;
#endif

    /// <summary>
    /// awake
    /// </summary>
    void Awake()
    {
        Initialize();
    }

    /// <summary>
    /// Initialize
    /// </summary>
    public void Initialize()
    {
        table = Table;
#if UNITY_EDITOR
        setLastId();
#endif
    }

    /// <summary>
    /// Check Serialized table Exists
    /// </summary>
    public bool CheckSerializedTableExists()
    {
        return Table != null;
    }

    /// <summary>
    /// テーブルを設定
    /// </summary>
    public static void SetTable(Object obj)
    {
        if (obj is TextAsset)
        {
            // Json
            table = Class_Stage.CreateInstance();
            table.SetRowsByJson(obj.ToString());
        }
        else
        {
            // Scriptable Object
            table = obj as Class_Stage;
        }
#if UNITY_EDITOR
        setLastId();
#endif
    }

    /// <summary>
    /// テーブルを取得
    /// </summary>
    public static Class_Stage GetTable()
    {
        return table;
    }

    /// <summary>
    /// 行を配列順で取得する
    /// </summary>
    /// <param name="index">配列番号</param>
    public static Class_Stage.Row GetRow(int index)
    {
        return table?.GetRow(index);
    }

#if UNITY_EDITOR
    /// <summary>
    /// 新規行を追加. ID 自動生成
    /// </summary>
    public static void Add(Class_Stage.Row row)
    {
        lastId = lastId + 1;
        row.ID = lastId;
        
        table.Rows.Add(row);

        EditorUtility.SetDirty(table);
        AssetDatabase.SaveAssets();
    }

    /// <summary>
    /// 行を更新
    /// </summary>
    public static void Update(Class_Stage.Row row)
    {
        EditorUtility.SetDirty(table);
        AssetDatabase.SaveAssets();
    }

    /// <summary>
    /// 変更したテーブル内容を破棄
    /// </summary>
    public static void RejectChanges()
    {
        Resources.UnloadAsset(table);
    }
    
    static void setLastId()
    {
        if (table == null || table.Rows == null || table.Rows.Count == 0)
        {
            lastId = 0;
        }
        else
        {
            lastId = table.Rows.Max(r => r.ID);
        }
    }
#endif

    /// <summary>
    /// ID
    /// </summary>
    public static Class_Stage.Row FindRowByID(int val, bool errorLog = true)
    {
        return table?.FindRowByID(val, errorLog);
    }

}
