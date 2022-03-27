using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using NPOI.SS.UserModel;
using UnityEditor;
using UnityEngine;

public partial class XlsToJson : EditorWindow
{
    /// <summary>
    /// enum を解析
    /// </summary>
    static bool analyzeEnumsAndConsts(SheetEntity report)
    {
        Dictionary<string, EnumInfo>  enums   = report.Enums;
        Dictionary<string, ConstInfo> consts  = report.Consts;
        ISheet                        sheet   = report.Sheet;
        string[,]                     grid    = report.Grid;
        Dictionary<string, PosIndex>  posList = report.PosList;

        // enum
        foreach (var pair in posList)
        {
            if (pair.Key.Contains(TRIGGER_ID) == true)
            {
                continue;
            }
            if (pair.Key.Contains(TRIGGER_SUBCLASS) == true)
            {
                continue;
            }
            if (pair.Key.Contains(TRIGGER_ENUM) == true)
            {
                continue;
            }
            if (pair.Key.Contains(TRIGGER_CONST) == true)
            {
                continue;
            }
            
            PosIndex top = pair.Value;

            EnumInfo info = new EnumInfo();
            info.GroupName = top.Name;
            info.Comment = getCommentRight(grid, top.R, top.C);
            enums.Add(pair.Key, info);

            for (int r = top.R+1; r < grid.GetUpperBound(0)+1; r++)
            {
                int    c       = top.C;
                string name    = grid[r, c];
                string suffix  = null;

                if (string.IsNullOrEmpty(name) == true)
                {
                    break;
                }
                
                if (info.Members.ContainsKey(name) == true)
                {
                    // 同じメンバ名
                    LogError(MSG_SAMEMEMBER, report.SheetName, GetXLS_RC(r, c), name);
                    return false;
                }

                if (name.IndexOf("=") >= 0)
                {
                    string[] names = name.Split('=');
                    name   = names[0].Trim();
                    suffix = names[1].Trim();
                }

                Member row = new Member();
                row.Suffix  = suffix;
                row.Comment = getCommentRight(grid, r, c);
                info.Members.Add(name, row);
            }
        }

        // const
        foreach (var pair in posList)
        {
            if (pair.Key.Contains(TRIGGER_CONST) == false)
            {
                continue;
            }

            PosIndex top = pair.Value;
            
            string type = pair.Key.Replace($"{TRIGGER_CONST}.", "");

            ConstInfo info = new ConstInfo();
            consts.Add(pair.Key, info);

            for (int r = top.R+1; r < grid.GetUpperBound(0)+1; r++)
            {
                int    c       = top.C;
                string name    = grid[r, c];
                string suffix  = grid[r, c+1];

                if (string.IsNullOrEmpty(name) == true)
                {
                    break;
                }
                
                if (info.Members.ContainsKey(name) == true)
                {
                    // 同じメンバ名
                    LogError(MSG_SAMEMEMBER, report.SheetName, GetXLS_RC(r, c), name);
                    return false;
                }

                Member row = new Member();
                row.Type    = type;
                row.Suffix  = suffix;
                row.Comment = getCommentRight(grid, r, c+1);
                info.Members.Add(name, row);
            }
        }
        return true;
    }

    /// <summary>
    /// クラスコメントを解析
    /// </summary>
   static void analyzeClassComments(SheetEntity report)
    {
        Dictionary<string, ClassInfo> classes = report.Classes;
        string[,]                     grid    = report.Grid;
        Dictionary<string, PosIndex>  posList = report.PosList;

        if (classes.ContainsKey(CLASSTMPL_ROW) == true)
        {
            classes[CLASSTMPL_ROW].Comment = "Table Row";
        }
        if (posList.ContainsKey(TRIGGER_SUBCLASS) == false)
        {
            return;
        }

        PosIndex id = posList[TRIGGER_SUBCLASS];
        int c = id.C;

        for (int r = id.R+1; r < grid.GetUpperBound(0)+1; r++)
        {
            string name = grid[r, c];
            if (string.IsNullOrEmpty(name) == true)
            {
                break;
            }

            if (classes.ContainsKey(name) == true)
            {
                classes[name].Comment = getCommentRight(grid, r, c);
            }

            foreach (var cls in classes)
            {
                foreach (var member in cls.Value.Members)
                {
                    if (member.Key == name && member.Value.Type == name)
                    {
                        member.Value.Comment = getCommentRight(grid, r, c);
                    }
                }
            }
        }
    }

    /// <summary>
    /// クラス（テーブル）を解析。クラス内クラスも可能
    /// </summary>
    static bool analyzeClasses(SheetEntity report)
    {
        Dictionary<string, ClassInfo> classes = report.Classes;
        ISheet                        sheet   = report.Sheet;
        string[,]                     grid    = report.Grid;
        Dictionary<string, PosIndex>  posList = report.PosList;

        classes.Clear();
        classes.Add(CLASSTMPL_ROW, new ClassInfo());
        
        PosIndex id = posList[TRIGGER_ID];
        int      c;

        for (c = id.C; c < grid.GetUpperBound(1)+1; c++)
        {
            string name    = grid[id.R, c];
            bool   indexer = false;
            bool   isenum  = false;

            if (string.IsNullOrEmpty(name) == true)
            {
                break;
            }

            if (name.IndexOf(SIGN_INDEXER) >= 0)
            {
                name    = name.Replace(SIGN_INDEXER, "");
                indexer = true;
            }
            else
            if (name == TRIGGER_ID)
            {
                indexer = true;
            }

            string parentClassName = CLASSTMPL_ROW;

            if (name.IndexOf(".") > 0)
            {
                string[] names = name.Split('.');

                for (int i = 0; i < names.Length-1; i++)
                {
                    string classname = names[i];

                    if (classes[parentClassName].Members.ContainsKey(classname) == false)
                    {
                        Member row = new Member();
                        row.Type   = classname;
                        row.Suffix = $"= new {classname}()";
                        classes[parentClassName].Members.Add(classname, row);
                    }
                    if (classes.ContainsKey(classname) == false)
                    {
                        classes.Add(classname, new ClassInfo());
                    }
                    parentClassName = classname;
                }

                name = names[names.Length-1];
            }

            if (id.R+1 >= grid.GetUpperBound(0)+1)
            {
                // 最終行で型が取れない
                LogError(MSG_TYPE_NOTFOUND, report.SheetName, GetXLS_RC(id.R, c), "");
                return false;
            }

            string typestr = grid[id.R+1, c];
            bool   isList  = false;

            if (string.IsNullOrEmpty(typestr) == true)
            {
                // Default String. ID だけ int
                typestr = name == TRIGGER_ID ? "int" : "string";
            }
            else
            {
                typestr = typestr.Trim();

                if (typestr.IndexOf(SIGN_ENUM) == 0)
                {
                    typestr = typestr.Replace(SIGN_ENUM, "").Trim();
                    isenum  = true;
                    if (typestr.IndexOf('.') > 0)
                    {
                        // XXXXXX.eNum -> XXXXXX_Table.eNum
//typestr = typestr.Remove(typestr.IndexOf('.'));
//typestr = typestr.Insert(0, report.TableName);
                    }
                }

                if (typestr.IndexOf(SIGN_LIST) >= 0)
                {
                    isList = true;
                    // string[] -> List<string>
                    typestr = "List<" + typestr.Replace("[]", "").Trim() + ">";
                }
            }

            bool isContains = classes[parentClassName].Members.ContainsKey(name);
            if (isList == true)
            {
                if (isContains == true)
                {
                    // List 配列は複数登録の可能性があるので、エラーは出さずスルー
                }
                else
                {
                    // List
                    Member row  = new Member();
                    row.Type    = typestr;
                    row.Suffix  = $"= new {typestr}()";
                    row.Comment = getCommentUp(grid, id.R, c);
                    row.Indexer = indexer;
                    row.IsEnum  = isenum;
                    classes[parentClassName].Members.Add(name, row);
                }
            }
            else
            {
                if (isContains == true)
                {
                    // 配列ではないのに同じメンバ名
                    LogError(MSG_SAMEMEMBER, report.SheetName, GetXLS_RC(id.R, c), name);
                    return false;
                }
                else
                {
                    // Normal
                    Member row  = new Member();
                    row.Type    = typestr;
                    row.Suffix  = null;
                    row.Comment = getCommentUp(grid, id.R, c);
                    row.Indexer = indexer;
                    row.IsEnum  = isenum;
                    classes[parentClassName].Members.Add(name, row);
                }
            }
        }
        return true;
    }
    
}
