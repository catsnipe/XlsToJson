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

            int max = list.Count*3;
            foreach (SheetEntity ent in list)
            {
                completePreSuffixName(ent);

                string sheetName = ent.SheetName;
                if (CancelableProgressBar(cnt++, max, eMsg.FREE, ent.TableName) == true)
                {
                    break;
                }
                saveTableClass(ent, classDir);

                if (accessor.Used == true)
                {
                    if (CancelableProgressBar(cnt++, max, eMsg.FREE, ent.AccessorName) == true)
                    {
                        break;
                    }
                    saveTableAccess(ent, classDir);
                }

                if (CancelableProgressBar(cnt++, max, eMsg.FREE, ent.DataName + IMPORT_FILENAME_SUFFIX) == true)
                {
                    break;
                }

                saveJson(ent, dataDir, importerJson);
                saveScriptObj(ent, dataDir, importerScriptObj);
            }

            foreach (SheetEntity ent in alllist)
            {
                string sheetName = ent.SheetName;

                so_import_execlist.AppendLine($"                new {PREFIX_SCRIPTOBJ}{sheetName}_Import(),");
                so_export_execlist.AppendLine($"        {PREFIX_SCRIPTOBJ}{sheetName}_Export.Exec(srcbook, newbook, exportDirectory);");

                j_import_execlist.AppendLine($"                new {PREFIX_JSON}{sheetName}_Import(),");
                j_export_execlist.AppendLine($"        {PREFIX_JSON}{sheetName}_Export.Exec(srcbook, newbook, exportDirectory);");
            }

            saveAllJson(j_import_execlist.ToString(), j_export_execlist.ToString(), importerJson);
            saveAllScriptObj(so_import_execlist.ToString(), so_export_execlist.ToString(), importerScriptObj);
        }
        finally
        {
            EditorUtility.ClearProgressBar();
        }
    }

    /// <summary>
    /// ScriptableObject の基幹となるクラスを生成し、保存
    /// </summary>
    static bool saveTableClass(SheetEntity report, string classdir)
    {
        // Class は統一クラスとして既に出力されている
        if (report.ClassName != report.SheetName)
        {
            return true;
        }

        string text = createTableClass(report);

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
    static bool saveTableAccess(SheetEntity report, string classdir)
    {
        // テーブルがない
        if (report.Classes.Count == 0)
        {
            return true;
        }
        // Class は統一クラスとして既に出力されている
        if (report.ClassName != report.SheetName)
        {
            return true;
        }

        string text = createTableAccess(report);

        if (text != null)
        {
            string classpath = pathCombine(classdir, report.AccessorName) + ".cs";

            fileWrite(classpath, text);
        }

        return text != null;
    }

    /// <summary>
    /// ScriptableObject を自動生成する Editor クラスを生成し、保存
    /// </summary>
    static bool saveScriptObj(SheetEntity report, string datadir, bool createOrNot)
    {
        // テーブルがない
        if (report.Classes.Count == 0)
        {
            return true;
        }

        string workdir   = pathCombine(searchXlsToJsonDirectory(), IMPORT_DIRECTORY);
        string importout = pathCombine(workdir, $"{PREFIX_SCRIPTOBJ}{report.SheetName}{IMPORT_FILENAME_SUFFIX}.cs");
        string exportout = pathCombine(workdir, $"{PREFIX_SCRIPTOBJ}{report.SheetName}{EXPORT_FILENAME_SUFFIX}.cs");

        if (workdir != null)
        {
            CompleteDirectory(workdir);

            if (File.Exists(importout) == true) AssetDatabase.DeleteAsset(importout);
            if (File.Exists(exportout) == true) AssetDatabase.DeleteAsset(exportout);

            if (createOrNot == true)
            {
                var    text       = createScriptObj(report, datadir);
                string importText = text.importText;
                string exportText = text.exportText;

                if (importText == null || exportText == null)
                {
                    return false;
                }

                fileWrite(importout, importText);
                fileWrite(exportout, exportText);
            }
        }
        return true;
    }

    static bool saveAllScriptObj(string import_execlist, string export_execlist, bool createOrNot)
    {
        var    xlsName   = Path.GetFileNameWithoutExtension(xlsPath);
        string workdir   = pathCombine(searchXlsToJsonDirectory(), IMPORT_DIRECTORY);
        string importout = pathCombine(workdir, $"{PREFIXSIGN_ALL}{PREFIX_SCRIPTOBJ}{xlsName}{IMPORT_FILENAME_SUFFIX}.cs");
        string exportout = pathCombine(workdir, $"{PREFIXSIGN_ALL}{PREFIX_SCRIPTOBJ}{xlsName}{EXPORT_FILENAME_SUFFIX}.cs");

        if (workdir != null)
        {
            CompleteDirectory(workdir);

            if (File.Exists(importout) == true) AssetDatabase.DeleteAsset(importout);
            if (File.Exists(exportout) == true) AssetDatabase.DeleteAsset(exportout);

            if (createOrNot == true)
            {
                var    text       = createAllScriptObj(import_execlist, export_execlist);
                string importText = text.importText;
                string exportText = text.exportText;

                if (importText == null || exportText == null)
                {
                    return false;
                }

                fileWrite(importout, importText);
                fileWrite(exportout, exportText);
            }
        }
        return true;
    }

    /// <summary>
    /// Json を自動生成する Editor クラスを生成し、保存
    /// </summary>
    static bool saveJson(SheetEntity report, string datadir, bool createOrNot)
    {
        // テーブルがない
        if (report.Classes.Count == 0)
        {
            return true;
        }

        string workdir   = pathCombine(searchXlsToJsonDirectory(), IMPORT_DIRECTORY);
        string importout = pathCombine(workdir, $"{PREFIX_JSON}{report.SheetName}{IMPORT_FILENAME_SUFFIX}.cs");
        string exportout = pathCombine(workdir, $"{PREFIX_JSON}{report.SheetName}{EXPORT_FILENAME_SUFFIX}.cs");

        if (workdir != null)
        {
            CompleteDirectory(workdir);

            if (File.Exists(importout) == true) AssetDatabase.DeleteAsset(importout);
            if (File.Exists(exportout) == true) AssetDatabase.DeleteAsset(exportout);

            if (createOrNot == true)
            {
                var    text       = createJson(report, datadir);
                string importText = text.importText;
                string exportText = text.exportText;

                if (importText == null || exportText == null)
                {
                    return false;
                }

                fileWrite(importout, importText);
                fileWrite(exportout, exportText);
            }
        }
        return true;
    }

    static bool saveAllJson(string import_execlist, string export_execlist, bool createOrNot)
    {
        var    xlsName   = Path.GetFileNameWithoutExtension(xlsPath);
        string workdir   = pathCombine(searchXlsToJsonDirectory(), IMPORT_DIRECTORY);
        string importout = pathCombine(workdir, $"{PREFIXSIGN_ALL}{PREFIX_JSON}{xlsName}{IMPORT_FILENAME_SUFFIX}.cs");
        string exportout = pathCombine(workdir, $"{PREFIXSIGN_ALL}{PREFIX_JSON}{xlsName}{EXPORT_FILENAME_SUFFIX}.cs");

        if (workdir != null)
        {
            CompleteDirectory(workdir);

            if (File.Exists(importout) == true) AssetDatabase.DeleteAsset(importout);
            if (File.Exists(exportout) == true) AssetDatabase.DeleteAsset(exportout);

            if (createOrNot == true)
            {
                var    text       = createAllJson(import_execlist, export_execlist);
                string importText = text.importText;
                string exportText = text.exportText;

                if (importText == null || exportText == null)
                {
                    return false;
                }

                fileWrite(importout, importText);
                fileWrite(exportout, exportText);
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
