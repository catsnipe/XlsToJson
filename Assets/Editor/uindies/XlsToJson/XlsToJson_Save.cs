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
using UnityEditor;
using UnityEngine;

public partial class XlsToJson : EditorWindow
{
    /// <summary>
    /// シートを保存
    /// </summary>
    void saveSheet(List<SheetEntity> list, List<SheetEntity> alllist)
    {
        int cnt = 0;
        
        try
        {
            string nameNoExt    = Path.GetFileNameWithoutExtension(IF_XLS_TO_JSON_ACCESSOR);
            string ifInputPath  = pathCombine(searchXlsToJsonDirectory(), IF_XLS_TO_JSON_ACCESSOR);
            string ifOutputPath = pathCombine(classDir, nameNoExt + ".cs");

            string[] files = Directory.GetFiles(Application.dataPath, nameNoExt + ".cs", SearchOption.AllDirectories);

            if (files.Length > 0)
            {
                ifOutputPath = files[0];
            }
            File.Copy(ifInputPath, ifOutputPath, true);
            AssetDatabase.Refresh(ImportAssetOptions.ForceUpdate);

            fileWriteStart();

            var so_import_execlist = new StringBuilder();
            var so_export_execlist = new StringBuilder();
            var j_import_execlist  = new StringBuilder();
            var j_export_execlist  = new StringBuilder();
            var st_import_execlist = new StringBuilder();
            var st_export_execlist = new StringBuilder();

            var importerType = enumelate(importerTypeNo, eImporterType.Json);

            int max = list.Count*3;
            foreach (SheetEntity ent in list)
            {
                completePreSuffixName(ent);

                string sheetName = ent.SheetName;
                if (CancelableProgressBar(cnt++, max, eMsg.FREE, ent.TableName) == true)
                {
                    break;
                }
                saveTableClass(ent, classDir, importerType);

                if (accessor.Used == true)
                {
                    if (CancelableProgressBar(cnt++, max, eMsg.FREE, ent.AccessorName) == true)
                    {
                        break;
                    }
                    saveTableAccess(ent, classDir, importerType);
                }

                if (CancelableProgressBar(cnt++, max, eMsg.FREE, ent.DataName + IMPORT_FILENAME_SUFFIX) == true)
                {
                    break;
                }

                saveImporterExporter(ent, dataDir, eImporterType.Json, importerType == eImporterType.Json);
                saveImporterExporter(ent, dataDir, eImporterType.ScriptableObject, importerType == eImporterType.ScriptableObject);
                saveImporterExporter(ent, dataDir, eImporterType.StaticValue, importerType == eImporterType.StaticValue);
            }

            foreach (SheetEntity ent in alllist)
            {
                string sheetName = ent.SheetName;

                so_import_execlist.AppendLine($"                new {PREFIX_SCRIPTOBJ}{sheetName}_Import(),");
                so_export_execlist.AppendLine($"        {PREFIX_SCRIPTOBJ}{sheetName}_Export.Exec(srcbook, newbook, exportDirectory);");

                j_import_execlist.AppendLine($"                new {PREFIX_JSON}{sheetName}_Import(),");
                j_export_execlist.AppendLine($"        {PREFIX_JSON}{sheetName}_Export.Exec(srcbook, newbook, exportDirectory);");

                st_import_execlist.AppendLine($"                new {PREFIX_STATIC}{sheetName}_Import(),");
                st_export_execlist.AppendLine($"        {PREFIX_STATIC}{sheetName}_Export.Exec(srcbook, newbook, exportDirectory);");
            }

            saveAllImporterExporter(j_import_execlist.ToString(), j_export_execlist.ToString(), eImporterType.Json, importerType == eImporterType.Json);
            saveAllImporterExporter(so_import_execlist.ToString(), so_export_execlist.ToString(), eImporterType.ScriptableObject, importerType == eImporterType.ScriptableObject);
            saveAllImporterExporter(st_import_execlist.ToString(), st_export_execlist.ToString(), eImporterType.StaticValue, importerType == eImporterType.StaticValue);
        }
        finally
        {
            EditorUtility.ClearProgressBar();
        }
    }

    /// <summary>
    /// ScriptableObject の基幹となるクラスを生成し、保存
    /// </summary>
    static bool saveTableClass(SheetEntity report, string classdir, eImporterType importerType)
    {
        // Class は統一クラスとして既に出力されている
        if (report.ClassName != report.SheetName)
        {
            return true;
        }

        string text;

        if (importerType == eImporterType.StaticValue)
        {
            text = createTableStaticClass(report);
        }
        else
        {
            text = createTableClass(report, importerType);
        }

        if (text != null)
        {
            string classpath = pathCombine(classdir, report.TableName) + ".cs";

            fileWrite(classpath, text);
        }

        return text != null;
    }

    /// <summary>
    /// ScriptableObject をアクセスするシングルトンクラスを生成し、保存
    /// </summary>
    static bool saveTableAccess(SheetEntity report, string classdir, eImporterType importerType)
    {
        // テーブルがない
        if (report.Classes.Count == 0)
        {
            return true;
        }

        string text;

        if (importerType == eImporterType.StaticValue)
        {
            text = createTableStaticAccess(report);
        }
        else
        {
            text = createTableAccess(report, importerType);
        }

        if (text != null)
        {
            string classpath = pathCombine(classdir, report.AccessorName) + ".cs";

            fileWrite(classpath, text);
        }

        return text != null;
    }


    static bool saveImporterExporter(
        SheetEntity report,
        string datadir,
        eImporterType importerType,
        bool createOrNot)
    {
        string prefix;

        if (importerType == eImporterType.Json)
        {
            prefix = PREFIX_JSON;
        }
        else
        if (importerType == eImporterType.ScriptableObject)
        {
            prefix = PREFIX_SCRIPTOBJ;
        }
        else
        {
            prefix = PREFIX_STATIC;
        }

        // テーブルがない
        if (report.Classes.Count == 0)
        {
            return true;
        }

        string workdir   = pathCombine(searchXlsToJsonDirectory(), IMPORT_DIRECTORY);
        string importout = pathCombine(workdir, $"{prefix}{report.SheetName}{IMPORT_FILENAME_SUFFIX}.cs");
        string exportout = pathCombine(workdir, $"{prefix}{report.SheetName}{EXPORT_FILENAME_SUFFIX}.cs");

        if (workdir != null)
        {
            CompleteDirectory(workdir);

            if (File.Exists(importout) == true) AssetDatabase.DeleteAsset(importout);
            if (File.Exists(exportout) == true) AssetDatabase.DeleteAsset(exportout);

            if (createOrNot == true)
            {
                var    text       = createImporter(report, datadir, importerType);
                string importText = text.importText;
                string exportText = text.exportText;

                if (importText == null)
                {
                    return false;
                }

                fileWrite(importout, importText);
                if (importerType != eImporterType.StaticValue)
                {
                    fileWrite(exportout, exportText);
                }
            }
        }
        return true;
    }

    static bool saveAllImporterExporter(
        string import_execlist,
        string export_execlist,
        eImporterType importerType,
        bool createOrNot)
    {
        string prefix;

        if (importerType == eImporterType.Json)
        {
            prefix = PREFIX_JSON;
        }
        else
        if (importerType == eImporterType.ScriptableObject)
        {
            prefix = PREFIX_SCRIPTOBJ;
        }
        else
        {
            prefix = PREFIX_STATIC;
        }

        var    xlsName   = Path.GetFileNameWithoutExtension(xlsPath);
        string workdir   = pathCombine(searchXlsToJsonDirectory(), IMPORT_DIRECTORY);
        string importout = pathCombine(workdir, $"{PREFIXSIGN_ALL}{prefix}{xlsName}{IMPORT_FILENAME_SUFFIX}.cs");
        string exportout = pathCombine(workdir, $"{PREFIXSIGN_ALL}{prefix}{xlsName}{EXPORT_FILENAME_SUFFIX}.cs");

        if (workdir != null)
        {
            CompleteDirectory(workdir);

            if (File.Exists(importout) == true) AssetDatabase.DeleteAsset(importout);
            if (File.Exists(exportout) == true) AssetDatabase.DeleteAsset(exportout);

            if (createOrNot == true)
            {
                var    text       = createAllImporter(import_execlist, export_execlist, importerType);
                string importText = text.importText;
                string exportText = text.exportText;

                if (importText == null)
                {
                    return false;
                }

                fileWrite(importout, importText);
                if (importerType != eImporterType.StaticValue)
                {
                    fileWrite(exportout, exportText);
                }
            }
        }
        return true;
    }


    static List<string> writeFileList;

    static void fileWriteStart()
    {
        writeFileList = new List<string>();
    }

    static void fileWrite(string filename, string text)
    {
        // unix 形式に合わせる
        text = text.Replace("\r\n", "\n");

        File.WriteAllText(filename, text, Encoding.UTF8);
        AssetDatabase.ImportAsset(filename);
        
        writeFileList.Add(filename.Replace("Assets/", ""));
    }

    static string getWriteFileList()
    {
        writeFileList.Sort();

        var sb = new StringBuilder();
        foreach (var file in writeFileList)
        {
            sb.AppendLine(file);
        }

        return sb.ToString();
    }
}
