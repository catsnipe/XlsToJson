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

using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Codice.CM.Client.Differences.Graphic;
using NPOI.SS.UserModel;
using UnityEditor;
using UnityEngine;
using static XlsToJson;

public partial class XlsToJson : EditorWindow
{
    /// <summary>
    /// StaticValue の基幹となるクラスを生成
    /// </summary>
    static string createTableStaticClass(SheetEntity report)
    {
        string template_file  = TEMPLATE_STATIC_CLASS;

        var    classes        = report.Classes;
        var    enums          = report.Enums;
        var    consts         = report.Consts;
        string tablename      = report.TableName;

        string import_dir     = searchXlsToJsonDirectory();
        string template_class = NAME_CLASS;
        string text           = null;

        if (import_dir != null)
        {
            text = File.ReadAllText(pathCombine(import_dir, template_file), Encoding.UTF8);
        }
        if (text == null)
        {
            // テンプレートが見つからない
            DialogError(eMsg.CLASSTMPL_NOTFOUND, template_file);
            return null;
        }

        int padding        = 0;
        int padding_key    = 0;

        // enum
        var sb_enum        = new StringBuilder();
        var sb_global_enum = new StringBuilder();
        foreach (var pair in enums)
        {
            StringBuilder sb;
            string        indent = "";

            if (pair.Key.IndexOf(TRIGGER_GLOBAL_ENUM) == 0)
            {
                sb = sb_global_enum;
            }
            else
            {
                sb = sb_enum;
                indent = "\t";
            }
            if (string.IsNullOrEmpty(pair.Value.Comment) == false)
            {
                addCommentText(sb, pair.Value.Comment, indent.Length == 0 ? 0 : 1);
            }
            sb.AppendLine($"{indent}public enum {pair.Value.GroupName}");
            sb.AppendLine($"{indent}{{");
            foreach (var member in pair.Value.Members)
            {
                Member row = member.Value;

                addCommentText(sb, row.Comment, indent.Length == 0 ? 1 : 2);
                if (string.IsNullOrEmpty(row.Suffix) == true)
                {
                    sb.AppendLine($"{indent}\t{member.Key},");
                }
                else
                {
                    sb.AppendLine($"{indent}\t{member.Key} = {row.Suffix},");
                }
            }
            sb.AppendLine($"{indent}}}");
            sb.AppendLine( "");
        }

        // const
        var sb_const = new StringBuilder();

        foreach (var pair in consts)
        {
            foreach (var member in pair.Value.Members)
            {
                if (padding < member.Value.Type.Length)
                {
                    padding = member.Value.Type.Length;
                }
                if (padding_key < member.Key.Length)
                {
                    padding_key = member.Key.Length;
                }
            }
        }

        foreach (var pair in consts)
        {
            StringBuilder sb     = sb_const;
            string        indent = "\t";

            foreach (var member in pair.Value.Members)
            {
                if (string.IsNullOrEmpty(member.Value.Comment) == false)
                {
                    addCommentText(sb, member.Value.Comment, indent.Length == 0 ? 0 : 1);
                }

                string value = "";
                switch (member.Value.Type)
                {
                    case "string":
                        value = $"\"{member.Value.Suffix}\"";
                        break;
                    case "float":
                        value = $"{member.Value.Suffix}f";
                        break;
                    case "double":
                        value = $"{member.Value.Suffix}d";
                        break;
                    case "decimal":
                        value = $"{member.Value.Suffix}m";
                        break;
                    case "int":
                    case "long":
                        value = $"{member.Value.Suffix}";
                        break;
                    default:
                        Debug.LogError($"const type error '{member.Value.Type}'");
                        Debug.LogError($"(possible type) int, float, string");
                        break;
                }
                sb.AppendLine($"{indent}public const {member.Value.Type.PadRight(padding)} {member.Key.PadRight(padding_key)} = {value};");
            }
        }
        sb_const.AppendLine("");

        var sb_class      = new StringBuilder();
        var tableNames    = new Dictionary<string, string>();

        foreach (var entity in sheetList)
        {
            if (tableNames.ContainsKey(entity.ClassName) == false)
            {
                tableNames.Add(entity.ClassName, entity.TableName);
            }
        }

        foreach (var pair in classes)
        {
            foreach (var member in pair.Value.Members)
            {
                if (padding < member.Value.Type.Length)
                {
                    padding = member.Value.Type.Length;
                }
            }
        }
        foreach (var pair in classes)
        {
            if (string.IsNullOrEmpty(pair.Value.Comment) == false)
            {
                addCommentText(sb_class, pair.Value.Comment, 1);
            }
            sb_class.AppendLine( "\t[System.Serializable]");
            sb_class.AppendLine($"\tpublic class {pair.Key}");
            sb_class.AppendLine( "\t{");
            foreach (var member in pair.Value.Members)
            {
                Member row  = member.Value;
                string type = row.Type;

                addCommentText(sb_class, row.Comment, 2);
                if (row.Suffix == null)
                {
                    sb_class.AppendLine($"\t\tpublic {type.PadRight(padding)} {member.Key};");
                }
                else
                {
                    sb_class.AppendLine($"\t\tpublic {type.PadRight(padding)} {member.Key} {row.Suffix};");
                }
            }
            sb_class.AppendLine( "\t}");
            sb_class.AppendLine( "");
        }

        int start;
        int end;

        text = text.Replace("\r\n", "\n");

        // 最初からある Row を消す
        start = text.IndexOf(CLASSTMPL_CLASS_SIGN) + CLASSTMPL_CLASS_SIGN.Length + "\n".Length;
        end   = text.IndexOf(CLASSTMPL_CLASS_ENDSIGN);
        text  = text.Remove(start, end - start);

        text  = text.Replace(CLASSTMPL_GENUM_ENDSIGN + "\n", sb_global_enum.ToString());
        text  = text.Replace(CLASSTMPL_ENUM_ENDSIGN + "\n", sb_enum.ToString());
        text  = text.Replace(CLASSTMPL_CONST_ENDSIGN + "\n", sb_const.ToString());
        text  = text.Replace(CLASSTMPL_CLASS_ENDSIGN + "\n", sb_class.ToString());
        text  = text.Replace(CLASSTMPL_GENUM_SIGN + "\n", "");
        text  = text.Replace(CLASSTMPL_ENUM_SIGN + "\n", "");
        text  = text.Replace(CLASSTMPL_CONST_SIGN + "\n", "");
        text  = text.Replace(CLASSTMPL_CLASS_SIGN + "\n", "");
        if (report.Classes.Count == 0 && sb_enum.Length == 0 && sb_const.Length == 0)
        {
            // テーブルがない場合、余計なコードを一切消す
            start = text.IndexOf(CLASSTMPL_TABLE_SIGN) + CLASSTMPL_TABLE_SIGN.Length + "\n".Length;
            end   = text.IndexOf(CLASSTMPL_TABLE_ENDSIGN);
            text  = text.Remove(start, end - start);
        }
        text  = text.Replace(CLASSTMPL_TABLE_SIGN + "\n", "");
        text  = text.Replace(CLASSTMPL_TABLE_ENDSIGN + "\n", "");
        text  = text.Replace("\t", "    ");
        text  = text.Replace(template_class, tablename);

        return text;
    }


    /// <summary>
    /// StaticValue をアクセスするシングルトンクラスを生成
    /// </summary>
    static string createTableStaticAccess(SheetEntity report)
    {
        string                        template_file = TEMPLATE_STATIC_ACCESS;
        
        ISheet                        sheet         = report.Sheet;
        Dictionary<string, ClassInfo> classes       = report.Classes;
        string                        accessorname  = report.AccessorName;
        string                        tablename     = report.TableName;

        string                        import_dir    = searchXlsToJsonDirectory();
        string                        text          = null;

        if (import_dir != null)
        {
            text = File.ReadAllText(pathCombine(import_dir, template_file), Encoding.UTF8);
        }
        if (text == null)
        {
            // テンプレートが見つからない
            DialogError(eMsg.CLASSTMPL_NOTFOUND, template_file);
            return null;
        }

        var sb_index      = new StringBuilder();
        var sb_index_null = new StringBuilder();
        var sb_index_find = new StringBuilder();
        var tableNames    = new Dictionary<string, string>();

        foreach (var entity in sheetList)
        {
            if (tableNames.ContainsKey(entity.ClassName) == false)
            {
                tableNames.Add(entity.ClassName, entity.TableName);
            }
        }

        foreach (var pair in classes)
        {
            foreach (var member in pair.Value.Members)
            {
                Member row  = member.Value;
                string type = row.Type;

                if (row.IsEnum == true)
                {
                    if (type.IndexOf('.') >= 0)
                    {
                        string classname = type.Substring(0, type.IndexOf('.'));
                        if (tableNames.ContainsKey(classname) == false)
                        {
                            Log(eMsg.NOT_FOUND_ENUMTBL, report.SheetName);
                        }
                        else
                        {
                            type = type.Replace(classname, tableNames[classname]);
                        }
                    }
                }

                if (row.Indexer == true)
                {
                    string name = $"{member.Key}Rows";

                    sb_index.AppendLine($"\tstatic Dictionary<{type}, Row> {name};");
                    sb_index_null.AppendLine($"\t\t{name} = null;");
                    sb_index_find.AppendLine( "\t/// <summary>");
                    sb_index_find.AppendLine($"\t/// {member.Key}");
                    sb_index_find.AppendLine( "\t/// </summary>");
                    sb_index_find.AppendLine($"\tpublic static Row FindRowBy{member.Key}({type} val, bool errorLog = true)");
                    sb_index_find.AppendLine( "\t{");
                    sb_index_find.AppendLine($"\t\tif ({name} == null)");
                    sb_index_find.AppendLine( "\t\t{");
                    sb_index_find.AppendLine($"\t\t\t{name} = new Dictionary<{type}, Row>();");
                    sb_index_find.AppendLine($"\t\t\tRows.ForEach( (row) => {{ if ({name}.ContainsKey(row.{member.Key}) == false) {name}.Add(row.{member.Key}, row); }} );");
                    sb_index_find.AppendLine( "\t\t}");
                    sb_index_find.AppendLine( "\t\t");
                    if (type == "string")
                    {
                        sb_index_find.AppendLine($"\t\tif (val == null || {name}.ContainsKey(val) == false)");
                    }
                    else
                    {
                        sb_index_find.AppendLine($"\t\tif ({name}.ContainsKey(val) == false)");
                    }
                    sb_index_find.AppendLine( "\t\t{");
                    sb_index_find.AppendLine( "\t\t\tif (errorLog == true)");
                    sb_index_find.AppendLine( "\t\t\t{");
                    sb_index_find.AppendLine( "\t\t\t\tDebug.LogError($\"cannot find: {val}\");");
                    sb_index_find.AppendLine( "\t\t\t}");
                    sb_index_find.AppendLine( "\t\t\treturn null;");
                    sb_index_find.AppendLine( "\t\t}");
                    sb_index_find.AppendLine($"\t\treturn {name}[val];");
                    sb_index_find.AppendLine( "\t}");
                    sb_index_find.AppendLine( "");
                }
            }
        }

        string template_class  = NAME_CLASS;
        string template_access = NAME_ACCESS;

        text = text.Replace("\r\n", "\n");
        text = text.Replace(template_class, tablename);
        text = text.Replace(template_access, accessorname);
        text = text.Replace(CTMPL_INDEX_ENDSIGN + "\n", sb_index.ToString());
        text = text.Replace(CTMPL_INDEX_SIGN + "\n", "");
        text = text.Replace(CTMPL_INDEX_FIND_ENDSIGN + "\n", sb_index_find.ToString());
        text = text.Replace(CTMPL_INDEX_FIND_SIGN + "\n", "");
        text = text.Replace("\t", "    ");

        return text;
    }







    /// <summary>
    /// ScriptableObject の基幹となるクラスを生成
    /// </summary>
    static string createTableClass(SheetEntity report, eImporterType importerType)
    {
        string template_file;

        if (importerType == eImporterType.Json || importerType == eImporterType.ScriptableObject)
        {
            template_file  = TEMPLATE_CLASS;
        }
        else
        {
            Debug.LogError($"unmatch method. [{importerType}]");
            return null;
        }

        var    classes        = report.Classes;
        var    enums          = report.Enums;
        var    consts         = report.Consts;
        string tablename      = report.TableName;

        string import_dir     = searchXlsToJsonDirectory();
        string template_class = Path.GetFileNameWithoutExtension(TEMPLATE_CLASS);
        string text           = null;

        if (import_dir != null)
        {
            text = File.ReadAllText(pathCombine(import_dir, template_file), Encoding.UTF8);
        }
        if (text == null)
        {
            // テンプレートが見つからない
            DialogError(eMsg.CLASSTMPL_NOTFOUND, template_file);
            return null;
        }

        int padding        = 0;
        int padding_key    = 0;

        // enum
        var sb_enum        = new StringBuilder();
        var sb_global_enum = new StringBuilder();
        foreach (var pair in enums)
        {
            StringBuilder sb;
            string        indent = "";

            if (pair.Key.IndexOf(TRIGGER_GLOBAL_ENUM) == 0)
            {
                sb = sb_global_enum;
            }
            else
            {
                sb = sb_enum;
                indent = "\t";
            }
            if (string.IsNullOrEmpty(pair.Value.Comment) == false)
            {
                addCommentText(sb, pair.Value.Comment, indent.Length == 0 ? 0 : 1);
            }
            sb.AppendLine($"{indent}public enum {pair.Value.GroupName}");
            sb.AppendLine($"{indent}{{");
            foreach (var member in pair.Value.Members)
            {
                Member row = member.Value;

                addCommentText(sb, row.Comment, indent.Length == 0 ? 1 : 2);
                if (string.IsNullOrEmpty(row.Suffix) == true)
                {
                    sb.AppendLine($"{indent}\t{member.Key},");
                }
                else
                {
                    sb.AppendLine($"{indent}\t{member.Key} = {row.Suffix},");
                }
            }
            sb.AppendLine($"{indent}}}");
            sb.AppendLine( "");
        }

        // const
        var sb_const = new StringBuilder();

        foreach (var pair in consts)
        {
            foreach (var member in pair.Value.Members)
            {
                if (padding < member.Value.Type.Length)
                {
                    padding = member.Value.Type.Length;
                }
                if (padding_key < member.Key.Length)
                {
                    padding_key = member.Key.Length;
                }
            }
        }

        foreach (var pair in consts)
        {
            StringBuilder sb     = sb_const;
            string        indent = "\t";

            foreach (var member in pair.Value.Members)
            {
                if (string.IsNullOrEmpty(member.Value.Comment) == false)
                {
                    addCommentText(sb, member.Value.Comment, indent.Length == 0 ? 0 : 1);
                }

                string value = "";
                switch (member.Value.Type)
                {
                    case "string":
                        value = $"\"{member.Value.Suffix}\"";
                        break;
                    case "float":
                        value = $"{member.Value.Suffix}f";
                        break;
                    case "double":
                        value = $"{member.Value.Suffix}d";
                        break;
                    case "decimal":
                        value = $"{member.Value.Suffix}m";
                        break;
                    case "int":
                    case "long":
                        value = $"{member.Value.Suffix}";
                        break;
                    default:
                        Debug.LogError($"const type error '{member.Value.Type}'");
                        Debug.LogError($"(possible type) int, float, string");
                        break;
                }
                sb.AppendLine($"{indent}public const {member.Value.Type.PadRight(padding)} {member.Key.PadRight(padding_key)} = {value};");
            }
        }
        sb_const.AppendLine("");

        var sb_class      = new StringBuilder();
        var sb_index      = new StringBuilder();
        var sb_index_null = new StringBuilder();
        var sb_index_find = new StringBuilder();
        var tableNames    = new Dictionary<string, string>();

        foreach (var entity in sheetList)
        {
            if (tableNames.ContainsKey(entity.ClassName) == false)
            {
                tableNames.Add(entity.ClassName, entity.TableName);
            }
        }

        foreach (var pair in classes)
        {
            foreach (var member in pair.Value.Members)
            {
                if (padding < member.Value.Type.Length)
                {
                    padding = member.Value.Type.Length;
                }
            }
        }
        foreach (var pair in classes)
        {
            if (string.IsNullOrEmpty(pair.Value.Comment) == false)
            {
                addCommentText(sb_class, pair.Value.Comment, 1);
            }
            sb_class.AppendLine( "\t[System.Serializable]");
            sb_class.AppendLine($"\tpublic class {pair.Key}");
            sb_class.AppendLine( "\t{");
            foreach (var member in pair.Value.Members)
            {
                Member row  = member.Value;
                string type = row.Type;

                addCommentText(sb_class, row.Comment, 2);
                if (row.Suffix == null)
                {
                    sb_class.AppendLine($"\t\tpublic {type.PadRight(padding)} {member.Key};");
                }
                else
                {
                    sb_class.AppendLine($"\t\tpublic {type.PadRight(padding)} {member.Key} {row.Suffix};");
                }

                if (row.Indexer == true)
                {
                    string name = $"{member.Key}Rows";

                    if (importerType == eImporterType.StaticValue)
                    {
                        sb_index.AppendLine($"\tstatic Dictionary<{type}, Row> {name};");
                    }
                    else
                    {
                        sb_index.AppendLine($"\tDictionary<{type}, Row> {name};");
                    }
                    sb_index_null.AppendLine($"\t\t{name} = null;");
                    sb_index_find.AppendLine( "\t/// <summary>");
                    sb_index_find.AppendLine($"\t/// {member.Key}");
                    sb_index_find.AppendLine( "\t/// </summary>");
                    sb_index_find.AppendLine($"\tpublic Row FindRowBy{member.Key}({type} val, bool errorLog = true)");
                    sb_index_find.AppendLine( "\t{");
                    sb_index_find.AppendLine($"\t\tif ({name} == null)");
                    sb_index_find.AppendLine( "\t\t{");
                    sb_index_find.AppendLine($"\t\t\t{name} = new Dictionary<{type}, Row>();");
                    sb_index_find.AppendLine($"\t\t\tRows.ForEach( (row) => {{ if ({name}.ContainsKey(row.{member.Key}) == false) {name}.Add(row.{member.Key}, row); }} );");
                    sb_index_find.AppendLine( "\t\t}");
                    sb_index_find.AppendLine( "\t\t");
                    if (type == "string")
                    {
                        sb_index_find.AppendLine($"\t\tif (val == null || {name}.ContainsKey(val) == false)");
                    }
                    else
                    {
                        sb_index_find.AppendLine($"\t\tif ({name}.ContainsKey(val) == false)");
                    }
                    sb_index_find.AppendLine( "\t\t{");
                    sb_index_find.AppendLine( "\t\t\tif (errorLog == true)");
                    sb_index_find.AppendLine( "\t\t\t{");
                    sb_index_find.AppendLine( "\t\t\t\tDebug.LogError($\"cannot find: {val}\");");
                    sb_index_find.AppendLine( "\t\t\t}");
                    sb_index_find.AppendLine( "\t\t\treturn null;");
                    sb_index_find.AppendLine( "\t\t}");
                    sb_index_find.AppendLine($"\t\treturn {name}[val];");
                    sb_index_find.AppendLine( "\t}");
                    sb_index_find.AppendLine( "");
                }
            }
            sb_class.AppendLine( "\t}");
            sb_class.AppendLine( "");
        }

        int start;
        int end;

        text = text.Replace("\r\n", "\n");

        // 最初からある Row を消す
        start = text.IndexOf(CLASSTMPL_CLASS_SIGN) + CLASSTMPL_CLASS_SIGN.Length + "\n".Length;
        end   = text.IndexOf(CLASSTMPL_CLASS_ENDSIGN);
        text  = text.Remove(start, end - start);

        text  = text.Replace(CLASSTMPL_GENUM_ENDSIGN + "\n", sb_global_enum.ToString());
        text  = text.Replace(CLASSTMPL_ENUM_ENDSIGN + "\n", sb_enum.ToString());
        text  = text.Replace(CLASSTMPL_CONST_ENDSIGN + "\n", sb_const.ToString());
        text  = text.Replace(CLASSTMPL_CLASS_ENDSIGN + "\n", sb_class.ToString());
        text  = text.Replace(CTMPL_INDEX_ENDSIGN + "\n", sb_index.ToString());
        text  = text.Replace(CTMPL_INDEX_NULL_ENDSIGN + "\n", sb_index_null.ToString());
        text  = text.Replace(CTMPL_INDEX_FIND_ENDSIGN + "\n", sb_index_find.ToString());
        text  = text.Replace(CLASSTMPL_GENUM_SIGN + "\n", "");
        text  = text.Replace(CLASSTMPL_ENUM_SIGN + "\n", "");
        text  = text.Replace(CLASSTMPL_CONST_SIGN + "\n", "");
        text  = text.Replace(CLASSTMPL_CLASS_SIGN + "\n", "");
        text  = text.Replace(CTMPL_INDEX_SIGN + "\n", "");
        text  = text.Replace(CTMPL_INDEX_NULL_SIGN + "\n", "");
        text  = text.Replace(CTMPL_INDEX_FIND_SIGN + "\n", "");
        if (report.Classes.Count == 0 && sb_enum.Length == 0 && sb_const.Length == 0)
        {
            // テーブルがない場合、余計なコードを一切消す
            start = text.IndexOf(CLASSTMPL_TABLE_SIGN) + CLASSTMPL_TABLE_SIGN.Length + "\n".Length;
            end   = text.IndexOf(CLASSTMPL_TABLE_ENDSIGN);
            text  = text.Remove(start, end - start);
        }
        else
        if (report.Classes.Count == 0)
        {
            // Row がない場合、余計なコードを一切消す
            start = text.IndexOf(CLASSTMPL_CODE_SIGN) + CLASSTMPL_CODE_SIGN.Length + "\n".Length;
            end   = text.IndexOf(CLASSTMPL_CODE_ENDSIGN);
            text  = text.Remove(start, end - start);
        }
        text  = text.Replace(CLASSTMPL_TABLE_SIGN + "\n", "");
        text  = text.Replace(CLASSTMPL_TABLE_ENDSIGN + "\n", "");
        text  = text.Replace(CLASSTMPL_CODE_SIGN + "\n", "");
        text  = text.Replace(CLASSTMPL_CODE_ENDSIGN + "\n", "");
        text  = text.Replace("\t", "    ");
        text  = text.Replace(template_class, tablename);

        return text;
    }

    /// <summary>
    /// ScriptableObject をアクセスするシングルトンクラスを生成
    /// </summary>
    static string createTableAccess(SheetEntity report, eImporterType importerType)
    {
        string template_file;

        if (importerType == eImporterType.Json || importerType == eImporterType.ScriptableObject)
        {
            template_file  = TEMPLATE_ACCESS;
        }
        else
        {
            Debug.LogError($"unmatch method. [{importerType}]");
            return null;
        }
        
        ISheet                        sheet         = report.Sheet;
        Dictionary<string, ClassInfo> classes       = report.Classes;
        string                        accessorname  = report.AccessorName;
        string                        tablename     = report.TableName;

        string                        import_dir    = searchXlsToJsonDirectory();
        string                        text          = null;

        if (import_dir != null)
        {
            text = File.ReadAllText(pathCombine(import_dir, template_file), Encoding.UTF8);
        }
        if (text == null)
        {
            // テンプレートが見つからない
            DialogError(eMsg.CLASSTMPL_NOTFOUND, template_file);
            return null;
        }

        var sb_index_find = new StringBuilder();
        var tableNames    = new Dictionary<string, string>();

        foreach (var entity in sheetList)
        {
            if (tableNames.ContainsKey(entity.ClassName) == false)
            {
                tableNames.Add(entity.ClassName, entity.TableName);
            }
        }

        foreach (var pair in classes)
        {
            foreach (var member in pair.Value.Members)
            {
                Member row  = member.Value;
                string type = row.Type;

                if (row.IsEnum == true)
                {
                    if (type.IndexOf('.') >= 0)
                    {
                        string classname = type.Substring(0, type.IndexOf('.'));
                        if (tableNames.ContainsKey(classname) == false)
                        {
                            Log(eMsg.NOT_FOUND_ENUMTBL, report.SheetName);
                        }
                        else
                        {
                            type = type.Replace(classname, tableNames[classname]);
                        }
                    }
                }

                if (row.Indexer == true)
                {
                    string name = $"{member.Key}Rows";

                    sb_index_find.AppendLine( "\t/// <summary>");
                    sb_index_find.AppendLine($"\t/// {member.Key}");
                    sb_index_find.AppendLine( "\t/// </summary>");
                    sb_index_find.AppendLine($"\tpublic static {tablename}.Row FindRowBy{member.Key}({type} val, bool errorLog = true)");
                    sb_index_find.AppendLine( "\t{");
                    sb_index_find.AppendLine($"\t\treturn table?.FindRowBy{member.Key}(val, errorLog);");
                    sb_index_find.AppendLine( "\t}");
                    sb_index_find.AppendLine( "");
                }
            }
        }

        string template_access = Path.GetFileNameWithoutExtension(TEMPLATE_ACCESS);
        string template_class  = Path.GetFileNameWithoutExtension(TEMPLATE_CLASS);

        text = text.Replace("\r\n", "\n");
        text = text.Replace(template_class, tablename);
        text = text.Replace(template_access, accessorname);
        text = text.Replace(CTMPL_INDEX_FIND_ENDSIGN + "\n", sb_index_find.ToString());
        text = text.Replace(CTMPL_INDEX_FIND_SIGN + "\n", "");
        text = text.Replace("\t", "    ");

        return text;
    }

    /// <summary>
    /// Importer（シートごと）
    /// Importer
    /// </summary>
    /// <param name="report">テーブル</param>
    /// <param name="datadir">出力ディレクトリ</param>
    /// <param name="import_file">From Xlsx</param>
    /// <param name="export_file">To Xlsx</param>
    /// <returns>From Xlsx の作成コード, To Xlsx の作成コード</returns>
    static (string importText, string exportText) createImporter(
        SheetEntity report,
        string datadir,
        eImporterType importerType
    )
    {
        string import_file;
        string export_file;
        string extension;

        if (importerType == eImporterType.Json)
        {
            import_file = IMPORT_TPL_JSON;
            export_file = EXPORT_TPL_JSON;
            extension = ".txt";
        }
        else
        if (importerType == eImporterType.ScriptableObject)
        {
            import_file = IMPORT_TPL_SCRIPTOBJ;
            export_file = EXPORT_TPL_SCRIPTOBJ;
            extension = ".asset";
        }
        else
        if (importerType == eImporterType.StaticValue)
        {
            import_file = IMPORT_TPL_STATIC;
            export_file = null;
            extension = ".cs";
        }
        else
        {
            LogError(eMsg.FREE, $"undefined type: '{importerType}'");
            return (null, null);
        }

        ISheet                       sheet         = report.Sheet;
        string[,]                    grid          = report.Grid;
        Dictionary<string, PosIndex> posList       = report.PosList;
        string                       sheetname     = report.SheetName;
        string                       sheetorg      = report.Sheet.SheetName;
        string                       classname     = report.ClassName;
        string                       tablename     = report.TableName;

        string                       import_dir    = searchXlsToJsonDirectory();
        string                       import_text   = null;
        string                       export_text   = null;

        string                       datafile;

        if (importerType == eImporterType.StaticValue)
        {
            datafile = report.AccessorName + extension;
        }
        else
        {
            datafile = report.DataName + extension;
        }

        if (import_dir != null)
        {
            import_text = File.ReadAllText(pathCombine(import_dir, import_file), Encoding.UTF8);
            if (importerType != eImporterType.StaticValue)
            {
                export_text = File.ReadAllText(pathCombine(import_dir, export_file), Encoding.UTF8);
            }
        }

        if (importerType == eImporterType.StaticValue)
        {
            if (import_text == null)
            {
                // テンプレートが見つからない
                DialogError(eMsg.CLASSTMPL_NOTFOUND, import_file, export_file);
                return (null, null);
            }
        }
        else
        {
            if (import_text == null || export_text == null)
            {
                // テンプレートが見つからない
                DialogError(eMsg.CLASSTMPL_NOTFOUND, import_file, export_file);
                return (null, null);
            }
        }

        var sb_import = new StringBuilder();
        var sb_export = new StringBuilder();

        if (importerType == eImporterType.StaticValue)
        {
            sb_import.AppendLine($"\t\t\tline += \"{_cr()}\";");
        }

        PosIndex id = posList[TRIGGER_ID];

        string preMember = null;
        int    listCount = 0;
        int    indent    = 3;

        HashSet<string> subClasses = new HashSet<string>();

        imports(
            report,
            grid,
            id.R,
            id.C,
            sb_import,
            sb_export,
            importerType,
            null,
            ref preMember,
            ref listCount,
            ref indent);


        if (importerType == eImporterType.StaticValue)
        {
            sb_import.AppendLine($"\t\t\tline += \"{_t(2)}\";");
        }
        
        import_text = import_text.Replace(IMPORTTMPL_SHEET_NAME, sheetname);
        import_text = import_text.Replace(IMPORTTMPL_SHEET_ORG, sheetorg);
        import_text = import_text.Replace(IMPORTTMPL_TABLE_NAME, tablename);
        import_text = import_text.Replace(IMPORTTMPL_EXPORT_PATH, datafile);
        import_text = import_text.Replace(IMPORTTMPL_IMPORT_ROW, sb_import.ToString());
        import_text = import_text.Replace("\t", "    ");

        if (importerType != eImporterType.StaticValue)
        {
            export_text = export_text.Replace(IMPORTTMPL_SHEET_NAME, sheetname);
            export_text = export_text.Replace(IMPORTTMPL_SHEET_ORG, sheetorg);
            export_text = export_text.Replace(IMPORTTMPL_TABLE_NAME, tablename);
            export_text = export_text.Replace(IMPORTTMPL_EXPORT_PATH, datafile);
            export_text = export_text.Replace(IMPORTTMPL_EXPORT_ROW, sb_export.ToString());
            export_text = export_text.Replace("\t", "    ");
        }

        return (import_text, export_text);
    }

    static int imports(
        SheetEntity report,
        string[,]  grid,
        int row,
        int col,
        StringBuilder sb_import,
        StringBuilder sb_export,
        eImporterType importerType,
        string subClassName,
        ref string preMember,
        ref int listCount,
        ref int indent
        )
    {
        if (importerType != eImporterType.StaticValue)
        {
            subClassName = null;
        }

        for (int c = col; c < grid.GetUpperBound(1)+1; c++)
        {
            string member  = grid[row, c];
            string typestr = grid[row+1, c];

            if (string.IsNullOrEmpty(member) == true)
            {
                break;
            }

            if (typestr == null)
            {
                typestr = "";
            }

            // delete '*'
            member = member.Replace(SIGN_INDEXER, "");

            // import
            bool   isList    = false;
            bool   isEnum    = false;
            bool   isNull    = false;
            string getFunc   = null;

            if (typestr.IndexOf(SIGN_LIST) >= 0)
            {
                isList = true;
                typestr = typestr.Replace(SIGN_LIST, "");
            }
            else
            {
                isList = false;
            }

            string typeShortEnum = "";

            if (typestr.IndexOf(SIGN_ENUM) >= 0)
            {
                getFunc = "GetEnum   ";
                typestr = typestr.Replace(SIGN_ENUM, "");
                typestr = getLocalEnum(typestr, null);

                string className = PREFIX_CLASS + report.ClassName + ".";

                if (typestr.IndexOf(className) >= 0)
                {
                    typeShortEnum = typestr.Replace(className, "");
                }
                else
                {
                    typeShortEnum = typestr;
                }
                isEnum  = true;
            }
            else
            {
                typeShortEnum = typestr;

                switch (typestr)
                {
                    case "bool":
                        getFunc = "GetBool   ";
                        break;
                    case "int":
                        getFunc = "GetInt    ";
                        break;
                    case "long":
                        getFunc = "GetLong   ";
                        break;
                    case "float":
                        getFunc = "GetFloat  ";
                        break;
                    case "double":
                        getFunc = "GetDouble ";
                        break;
                    case "decimal":
                        getFunc = "GetDecimal";
                        break;
                    case "string":
                        getFunc = "GetString ";
                        break;
                    default:
                        if (string.IsNullOrEmpty(typestr) == false)
                        {
                            Log(eMsg.NOT_IMPORTABLE_TYPE, typestr, member);
                        }
                        getFunc = "GetNull  ";
                        isNull  = true;
                        break;
                }
            }

            string checkSubName = member;

            if (importerType == eImporterType.StaticValue)
            {

                if (subClassName != null)
                {
                    if (checkSubName.IndexOf(subClassName + ".") < 0)
                    {
                        indent--;
                        sb_import.AppendLine($"\t\t\tline += $\"{_t(indent)}}}}},{_cr()}\";");

                        return c-1;
                    }

                    checkSubName = checkSubName.Replace($"{subClassName}.", "");
                }

                if (checkSubName.IndexOf(".") > 0)
                {
                    string name = checkSubName.Substring(0, checkSubName.LastIndexOf("."));

                    sb_import.AppendLine($"\t\t\tline += $\"{_t(indent)}{name} = {{{{{_cr()}\";");

                    if (subClassName != null)
                    {
                        name = $"{subClassName}.{name}";
                    }

                    indent++;

                    c = imports(
                        report,
                        grid,
                        row,
                        c,
                        sb_import,
                        sb_export,
                        importerType,
                        name, 
                        ref preMember,
                        ref listCount,
                        ref indent);

                    if (c < 0)
                    {
                        break;
                    }
                    continue;
                }
            }

            // import
            if (importerType == eImporterType.StaticValue)
            {
                if (isList == false)
                {
                    if (isNull == true)
                    {
                        sb_import.AppendLine($"\t\t\tif (import(r, c, cell, XlsToJson.{getFunc}()) == false) break;");
                    }
                    else
                    {
                        sb_import.AppendLine($"\t\t\tif (import(r, c, cell, XlsToJson.{getFunc}(cell, out row.{member})) == false) break;");
                        sb_import.AppendLine(createInitializeArgument(member, checkSubName, typeShortEnum, indent));
                    }
                }
                else
                {
                    if (preMember != member)
                    {
                        sb_import.AppendLine($"\t\t\tif (import(r, c, cell, XlsToJson.{getFunc}(cell, out {typestr} {member})) == false) break;");
                        sb_import.AppendLine(createInitializeArgumentListTop(member, checkSubName, typeShortEnum, indent));
                    }
                    else
                    {
                        sb_import.AppendLine($"\t\t\tif (import(r, c, cell, XlsToJson.{getFunc}(cell, out {member})) == false) break;");
                        sb_import.AppendLine(createInitializeArgumentList(member, typeShortEnum, indent));

                        if (c == grid.GetUpperBound(1) || member != grid[row, c+1])
                        {
                            sb_import.AppendLine($"\t\t\tline += \"{_t(indent)}}},{_cr()}\";");
                        }
                    }
                }
            }
            else
            {
                if (isList == false)
                {
                    if (isNull == true)
                    {
                        sb_import.AppendLine($"\t\t\tif (import(r, c, cell, XlsToJson.{getFunc}()) == false) break;");
                    }
                    else
                    {
                        sb_import.AppendLine($"\t\t\tif (import(r, c, cell, XlsToJson.{getFunc}(cell, out row.{member})) == false) break;");
                    }
                }
                else
                {
                    if (preMember != member)
                    {
                        sb_import.AppendLine($"\t\t\tif (import(r, c, cell, XlsToJson.{getFunc}(cell, out {typestr} {member})) == false) break;");
                    }
                    else
                    {
                        sb_import.AppendLine($"\t\t\tif (import(r, c, cell, XlsToJson.{getFunc}(cell, out {member})) == false) break;");
                    }
                    sb_import.AppendLine($"\t\t\trow.{member}.Add({member});");
                }
            }

            // export
            if (isEnum == true)
            {
                sb_export.AppendLine($"\t\t\terow.CellValue(c++, row.{member}.ToString());");
            }
            else
            {
                if (isList == false)
                {
                    sb_export.AppendLine($"\t\t\terow.CellValue(c++, row.{member});");
                }
                else
                {
                    if (preMember != member)
                    {
                        listCount = 0;
                    }
                    sb_export.AppendLine($"\t\t\terow.CellValue(c++, row.{member}[{listCount++}]);");
                }
            }

            if (c < grid.GetUpperBound(1))
            {
                sb_import.AppendLine($"\t\t\tcell = grid[r, ++c];");
            }

            preMember = member;
        }

        if (importerType == eImporterType.StaticValue)
        {
            if (subClassName != null)
            {
                string[] subClasses = subClassName.Split('.');

                for (int i = 0; i < subClasses.Length; i++)
                {
                    sb_import.AppendLine($"\t\t\t}},");
                }
            }
        }

        return -1;
    }

    static string createInitializeArgument(string member, string shortname, string typestr, int d)
    {
        string ret = null;

        switch (typestr)
        {
            case "float":
                ret = $"\t\t\tline += $\"{_t(d)}{shortname} = {{row.{member}}}f,{_cr()}\";";
                break;
            case "bool":
                ret = $"\t\t\tline += $\"{_t(d)}{shortname} = {{row.{member}.ToString().ToLower()}},{_cr()}\";";
                break;
            case "int":
            case "long":
            case "double":
            case "decimal":
                ret = $"\t\t\tline += $\"{_t(d)}{shortname} = {{row.{member}}},{_cr()}\";";
                break;
            case "string":
                ret = $"\t\t\tline += $\"{_t(d)}{shortname} = @\\\"{{row.{member}}}\\\",{_cr()}\";";
                break;
            default:
                ret = $"\t\t\tline += $\"{_t(d)}{shortname} = {typestr}.{{row.{member}}},{_cr()}\";";
                break;
        }

        return ret;
    }

    static string createInitializeArgumentListTop(string member, string shortname, string typestr, int d)
    {
        string ret = null;

        switch (typestr)
        {
            case "bool":
            case "int":
            case "long":
            case "float":
            case "double":
            case "decimal":
                ret = $"\t\t\tline += $\"{_t(d)}{shortname} = {{{{{_cr()}{_t(d+1)}{{{member}}},{_cr()}\";";
                break;
            case "string":
                ret = $"\t\t\tline += $\"{_t(d)}{shortname} = {{{{{_cr()}{_t(d+1)}@\\\"{{{member}}}\\\",{_cr()}\";";
                break;
            default:
                // enum
                ret = $"\t\t\tline += $\"{_t(d)}{shortname} = {{{{{_cr()}{_t(d+1)}{typestr}.{{{member}}},{_cr()}\";";
                break;
        }

        return ret;
    }

    static string createInitializeArgumentList(string member, string typestr, int d)
    {
        string ret = null;

        switch (typestr)
        {
            case "bool":
            case "int":
            case "long":
            case "float":
            case "double":
            case "decimal":
                ret = $"\t\t\tline += $\"{_t(d+1)}{{{member}}},{_cr()}\";";
                break;
            case "string":
                ret = $"\t\t\tline += $\"{_t(d+1)}@\\\"{{{member}}}\\\",{_cr()}\";";
                break;
            default:
                // enum
                ret = $"\t\t\tline += $\"{_t(d+1)}{typestr}.{{{member}}},{_cr()}\";";
                break;
        }

        return ret;
    }

    static string _t(int level)
    {
        if (prettyFormat == true)
        {
            return new string(' ', level*2);
        }
        else
        {
            return "";
        }
    }

    static string _cr()
    {
        if (prettyFormat == true)
        {
            return "\\r\\n";
        }
        else
        {
            return " ";
        }
    }

    /// <summary>
    /// 全ての Importer / Exporter をコールするクラスを生成する
    /// </summary>
    /// <param name="import_execlist">全ての Importer リスト</param>
    /// <param name="export_execlist">全ての Exporter リスト</param>
    /// <param name="import_file">Importer 生成クラスファイル</param>
    /// <param name="export_file">Exporter 生成クラスファイル</param>
    /// <returns></returns>
    static (string importText, string exportText) createAllImporter(
        string import_execlist,
        string export_execlist,
        eImporterType importerType
    )
    {
        string import_file = IMPORT_TPL_ALL;
        string export_file = EXPORT_TPL_ALL;

        string import_dir  = searchXlsToJsonDirectory();
        string import_text = null;
        string export_text = null;

        string type_name;
        string type_short_name;
        string type_cap_name;

        if (importerType == eImporterType.Json)
        {
            type_name = "Json";
            type_short_name = "Json";
            type_cap_name = "JSON";
        }
        else
        if (importerType == eImporterType.ScriptableObject)
        {
            type_name = "ScriptableObject";
            type_short_name = "ScriptObj";
            type_cap_name = "SCRIPTOBJ";
        }
        else
        {
            type_name = "StaticValue";
            type_short_name = "StaticValue";
            type_cap_name = "STATICVALUE";
        }

        if (import_dir != null)
        {
            import_text = File.ReadAllText(pathCombine(import_dir, import_file), Encoding.UTF8);
            if (importerType != eImporterType.StaticValue)
            {
                export_text = File.ReadAllText(pathCombine(import_dir, export_file), Encoding.UTF8);
            }
        }

        if (importerType == eImporterType.StaticValue)
        {
            if (import_text == null)
            {
                // テンプレートが見つからない
                DialogError(eMsg.CLASSTMPL_NOTFOUND, import_file, export_file);
                return (null, null);
            }
        }
        else
        {
            if (import_text == null || export_text == null)
            {
                // テンプレートが見つからない
                DialogError(eMsg.CLASSTMPL_NOTFOUND, import_file, export_file);
                return (null, null);
            }
        }

        string filenameOnly = Path.GetFileNameWithoutExtension(xlsPath);
        string filename     = Path.GetFileName(xlsPath);
        
        int priority = 0;
        for (int i = 0; i < filename.Length; i++)
        {
            priority += filename[i] * 2;
        }

        if (import_text.IndexOf("ScriptableObject") > 0)
        {
            priority += 10000;
        }

        import_text = import_text.Replace(IMPORTTMPL_EXCELL_NAME, filenameOnly);
        import_text = import_text.Replace(IMPORTTMPL_EXCELL_FILENAME, filename);
        if (importerType == eImporterType.StaticValue)
        {
            import_text = import_text.Replace(IMPORTTMPL_EXPORT_DIR, classDir);
        }
        else
        {
            import_text = import_text.Replace(IMPORTTMPL_EXPORT_DIR, dataDir);
        }
        import_text = import_text.Replace(IMPORTTMPL_IMPORT_EXECLIST, import_execlist);
        import_text = import_text.Replace(IMPORTTMPL_PRIORITY, priority.ToString());
        import_text = import_text.Replace(IMPORTTMPL_TYPE_NAME, type_name);
        import_text = import_text.Replace(IMPORTTMPL_TYPE_SHORTNAME, type_short_name);
        import_text = import_text.Replace(IMPORTTMPL_TYPE_CAPNAME, type_cap_name);

        if (importerType != eImporterType.StaticValue)
        {
            export_text = export_text.Replace(IMPORTTMPL_EXCELL_NAME, filenameOnly);
            export_text = export_text.Replace(IMPORTTMPL_EXCELL_FILENAME, filename);
            export_text = export_text.Replace(IMPORTTMPL_EXPORT_DIR, dataDir);
            export_text = export_text.Replace(IMPORTTMPL_EXPORT_EXECLIST, export_execlist);
            export_text = export_text.Replace(IMPORTTMPL_PRIORITY, (priority+1).ToString());
            export_text = export_text.Replace(IMPORTTMPL_TYPE_NAME, type_name);
            export_text = export_text.Replace(IMPORTTMPL_TYPE_SHORTNAME, type_short_name);
            export_text = export_text.Replace(IMPORTTMPL_TYPE_CAPNAME, type_cap_name);
        }

        return (import_text, export_text);
    }
}
