using System.Collections;
using System.Collections.Generic;
using UnityEngine;

/// <summary>
/// Character: ScriptableObject singleton accessor
/// </summary>
public partial class Character : MonoBehaviour
{
    [SerializeField]
    Class_Character        Table;

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
    public static List<Class_Character.Row> Rows
    {
        get
        {
            return table.Rows;
        }
    }

    static Class_Character table;

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
            table = Class_Character.CreateInstance();
            table.SetRowsByJson(obj.ToString());
        }
        else
        {
            // Scriptable Object
            table = obj as Class_Character;
        }
    }

    /// <summary>
    /// 行を配列順で取得する
    /// </summary>
    /// <param name="index">配列番号</param>
    public static Class_Character.Row GetRow(int index)
    {
        return table.GetRow(index);
    }

    /// <summary>
    /// ID
    /// </summary>
    public static Class_Character.Row FindRowByID(int val, bool errorLog = true)
    {
        return table.FindRowByID(val, errorLog);
    }

}
