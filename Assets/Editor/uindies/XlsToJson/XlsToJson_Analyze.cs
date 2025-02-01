// Copyright (c) catsnipe
// Released under the MIT license

// Permission is hereby granted, free of charge, to any person obtaining a 
// copy of this software and associated documentation files (the 
// "Software"), to deal in the Software without restriction, including 
// without limitation the rights to use, copy, modify, merge, publish, 
// distribute, sublicense, and/or sell copies of the Software, and to 
// permit persons to whom the Software is furnished to do so, subject to 
// the following conditions:
   
// The above copyright notice and this permission notice shall be 
// included in all copies or substantial portions of the Software.
   
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, 
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF 
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND 
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using NPOI.SS.UserModel;
using UnityEditor;
using UnityEngine;

public partial class XlsToJson : EditorWindow
{
    static Dictionary<string, EnumInfo> allEnums;

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
            
            string key = pair.Key;
            if (key.Contains(TRIGGER_GLOBAL_ENUM) == false)
            {
                key = $"{PREFIX_CLASS}{pair.Key}";
            }

            PosIndex top = pair.Value;

            EnumInfo info = new EnumInfo();
            info.GroupName = top.Name;
            info.Comment = getCommentRight(grid, top.R, top.C);
            enums.Add(key, info);

            if (allEnums.ContainsKey(key) == true)
            {
                LogError(eMsg.SAMEMEMBER, report.SheetName, "", key);
                return false;
            }
            allEnums.Add(key, info);

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
                    LogError(eMsg.SAMEMEMBER, report.SheetName, GetXLS_RC(r, c), name);
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
                    LogError(eMsg.SAMEMEMBER, report.SheetName, GetXLS_RC(r, c), name);
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
            string field   = grid[id.R, c];
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
                        row.FieldName = field;
                        row.Type      = classname;
                        row.Suffix    = $"= new {classname}()";
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
                LogError(eMsg.TYPE_NOTFOUND, report.SheetName, GetXLS_RC(id.R, c), "");
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
                    if (typestr.IndexOf('.') <= 0)
                    {
                        // eTest -> Class_Test.eTest
                        typestr = getLocalEnum(typestr, PREFIX_CLASS + report.ClassName);
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
                    Member row    = new Member();
                    row.FieldName = field;
                    row.Type      = typestr;
                    row.Suffix    = $"= new {typestr}()";
                    row.Comment   = getCommentUp(grid, id.R, c);
                    row.Indexer   = indexer;
                    row.IsEnum    = isenum;
                    classes[parentClassName].Members.Add(name, row);
                }
            }
            else
            {
                if (isContains == true)
                {
                    // 配列ではないのに同じメンバ名
                    LogError(eMsg.SAMEMEMBER, report.SheetName, GetXLS_RC(id.R, c), name);
                    return false;
                }
                else
                {
                    // Normal
                    Member row    = new Member();
                    row.FieldName = field;
                    row.Type      = typestr;
                    row.Suffix    = null;
                    row.Comment   = getCommentUp(grid, id.R, c);
                    row.Indexer   = indexer;
                    row.IsEnum    = isenum;
                    classes[parentClassName].Members.Add(name, row);
                }
            }
        }
        return true;
    }

    static string getLocalEnum(string typestr, string currentClassName)
    {
        string isListString = "";
        string searchString = typestr;

        if (typestr.IndexOf("[]") >= 0)
        {
            isListString = "[]";
            searchString = searchString.Replace("[]", "");
        }

        foreach (var pair in allEnums)
        {
            if (pair.Key.IndexOf(TRIGGER_GLOBAL_ENUM) < 0 && pair.Value.GroupName == searchString)
            {
                if (string.IsNullOrEmpty(currentClassName) == false && pair.Key.IndexOf(currentClassName) >= 0)
                {
                    // 同じクラスなので省略できる
                    return typestr;
                }
                else
                {
                    return pair.Key + isListString;
                }
            }
        }

        return typestr;
    }
}
