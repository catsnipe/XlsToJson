using System.Collections;
using System.Collections.Generic;
using UnityEngine;

//$$REGION GLOBAL_ENUM$$
//$$REGION_END GLOBAL_ENUM$$
//$$REGION TABLE$$
/// <summary>
/// XlsToJsonTemplate_Class: ScriptableObject
/// </summary>
public class XlsToJsonTemplate_Class : ScriptableObject
{
//$$REGION ENUM$$
//$$REGION_END ENUM$$
//$$REGION CONST$$
//$$REGION_END CONST$$
//$$REGION CLASS$$
    /// <summary>
    /// １行分のフィールドリスト
    /// </summary>
    public class Row
    {
        public int ID;
    }
//$$REGION_END CLASS$$
//$$REGION CODE$$
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
//$$REGION INDEX$$
//$$REGION END_INDEX$$

    /// <summary>
    /// .ctor
    /// </summary>
    public XlsToJsonTemplate_Class()
    {
        Rows      = new List<Row>();
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

//$$REGION INDEX_FIND$$
//$$REGION END_INDEX_FIND$$
//$$REGION_END CODE$$
}
//$$REGION_END TABLE$$
