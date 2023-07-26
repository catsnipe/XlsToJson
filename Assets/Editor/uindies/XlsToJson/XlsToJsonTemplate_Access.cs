﻿using System.Collections;
using System.Collections.Generic;
using UnityEngine;
#if UNITY_EDITOR
using System.Linq;
using UnityEditor;
#endif

/// <summary>
/// XlsToJsonTemplate_Access: ScriptableObject singleton accessor
/// </summary>
public partial class XlsToJsonTemplate_Access : MonoBehaviour
{
    [SerializeField]
    XlsToJsonTemplate_Class        Table;

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
    public static List<XlsToJsonTemplate_Class.Row> Rows
    {
        get
        {
            return table.Rows;
        }
    }

    static XlsToJsonTemplate_Class table;
#if UNITY_EDITOR
    static int lastId;
#endif

    /// <summary>
    /// awake
    /// </summary>
    void Awake()
    {
        table = Table;
#if UNITY_EDITOR
        setLastId();
#endif
    }

    /// <summary>
    /// テーブルを設定
    /// </summary>
    public static void SetTable(Object obj)
    {
        if (obj is TextAsset)
        {
            // Json
            table = XlsToJsonTemplate_Class.CreateInstance();
            table.SetRowsByJson(obj.ToString());
        }
        else
        {
            // Scriptable Object
            table = obj as XlsToJsonTemplate_Class;
        }
#if UNITY_EDITOR
        setLastId();
#endif
    }

    /// <summary>
    /// テーブルを取得
    /// </summary>
    public static XlsToJsonTemplate_Class GetTable()
    {
        return table;
    }

    /// <summary>
    /// 行を配列順で取得する
    /// </summary>
    /// <param name="index">配列番号</param>
    public static XlsToJsonTemplate_Class.Row GetRow(int index)
    {
        return table?.GetRow(index);
    }

#if UNITY_EDITOR
    /// <summary>
    /// 新規行を追加. ID 自動生成
    /// </summary>
    public static void Add(XlsToJsonTemplate_Class.Row row)
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
    public static void Update(XlsToJsonTemplate_Class.Row row)
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
        lastId = table.Rows.Max(r => r.ID);
    }
#endif

//$$REGION INDEX_FIND$$
//$$REGION END_INDEX_FIND$$
}
