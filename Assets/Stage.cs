using System.Collections;
using System.Collections.Generic;
using UnityEngine;

/// <summary>
/// Stage: ScriptableObject singleton accessor
/// </summary>
public partial class Stage : MonoBehaviour
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

    /// <summary>
    /// awake
    /// </summary>
    void Awake()
    {
        table = Table;
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
    }

    /// <summary>
    /// 行を配列順で取得する
    /// </summary>
    /// <param name="index">配列番号</param>
    public static Class_Stage.Row GetRow(int index)
    {
        return table.GetRow(index);
    }

    /// <summary>
    /// ID
    /// </summary>
    public static Class_Stage.Row FindRowByID(int val, bool errorLog = true)
    {
        return table.FindRowByID(val, errorLog);
    }

}
