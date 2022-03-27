using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Security.Cryptography;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using UnityEditor;
using UnityEngine;
using System.Security.Policy;

public partial class XlsToJson : EditorWindow
{
    /// <summary>
    /// コメントを示すキーワード
    /// </summary>
    static readonly string[] SIGN_COMMENTS  = new string[] { "//", "///", "/*", "*/", "[[[", "]]]", "#" };
    /// <summary>
    /// ID のポジションを示すキーワード
    /// </summary>
    public const string TRIGGER_ID          = "ID";
    /// <summary>
    /// [CLASS] のポジションを示すキーワード
    /// </summary>
    const string TRIGGER_SUBCLASS           = "[SUBCLASS]";
    /// <summary>
    /// [ENUM] のポジションを示すキーワード
    /// </summary>
    const string TRIGGER_ENUM               = "[ENUM]";
    /// <summary>
    /// [CONST] のポジションを示すキーワード
    /// </summary>
    const string TRIGGER_CONST              = "[CONST]";
    /// <summary>
    /// [GLOBAL_ENUM] のポジションを示すキーワード
    /// </summary>
    const string TRIGGER_GLOBAL_ENUM        = "[GLOBAL_ENUM]";
    /// <summary>
    /// 最大行数
    /// </summary>
    const int    ROWS_MAX                   = 2000;
    /// <summary>
    /// 最大列数
    /// </summary>
    const int    COLS_MAX                   = 50;

    const int    MARGINROW_BETWEEN_CTG      = 1;
    const string SIGN_ENUM                  = "enum:";
    const string SIGN_LIST                  =  "[]";
    const string SIGN_INDEXER               =  "*";
    const string SIGN_COMMA                 =  ".";

    const string CLASSTMPL_GENUM_SIGN       = "//$$REGION GLOBAL_ENUM$$";
    const string CLASSTMPL_GENUM_ENDSIGN    = "//$$REGION_END GLOBAL_ENUM$$";
    const string CLASSTMPL_ENUM_SIGN        = "//$$REGION ENUM$$";
    const string CLASSTMPL_ENUM_ENDSIGN     = "//$$REGION_END ENUM$$";
    const string CLASSTMPL_CONST_SIGN       = "//$$REGION CONST$$";
    const string CLASSTMPL_CONST_ENDSIGN    = "//$$REGION_END CONST$$";
    const string CLASSTMPL_CLASS_SIGN       = "//$$REGION CLASS$$";
    const string CLASSTMPL_CLASS_ENDSIGN    = "//$$REGION_END CLASS$$";
    const string CLASSTMPL_CODE_SIGN        = "//$$REGION CODE$$";
    const string CLASSTMPL_CODE_ENDSIGN     = "//$$REGION_END CODE$$";
    const string CLASSTMPL_TABLE_SIGN       = "//$$REGION TABLE$$";
    const string CLASSTMPL_TABLE_ENDSIGN    = "//$$REGION_END TABLE$$";
    const string CTMPL_INDEX_SIGN           = "//$$REGION INDEX$$";
    const string CTMPL_INDEX_ENDSIGN        = "//$$REGION END_INDEX$$";
    const string CTMPL_INDEX_FIND_SIGN      = "//$$REGION INDEX_FIND$$";
    const string CTMPL_INDEX_FIND_ENDSIGN   = "//$$REGION END_INDEX_FIND$$";

    const string CLASSTMPL_ROW              = "Row";
    const string IMPORTTMPL_SHEET_NAME      = "$$SHEET_NAME$$";
    const string IMPORTTMPL_SHEET_ORG       = "$$SHEET_ORG$$";
    const string IMPORTTMPL_TABLE_NAME      = "$$TABLE_NAME$$";
    const string IMPORTTMPL_EXCELL_NAME     = "$$EXCELL_NAME$$";
    const string IMPORTTMPL_EXCELL_FILENAME = "$$EXCELL_FILENAME$$";
    const string IMPORTTMPL_EXPORT_DIR      = "$$EXPORT_DIRECTORY$$";
    const string IMPORTTMPL_EXPORT_PATH     = "$$EXPORT_PATH$$";
    const string IMPORTTMPL_IMPORT_ROW      = "$$IMPORT_ROW$$";
    const string IMPORTTMPL_IMPORT_EXECLIST = "$$IMPORT_EXEC_LIST$$";
    const string IMPORTTMPL_EXPORT_ROW      = "$$EXPORT_ROW$$";
    const string IMPORTTMPL_EXPORT_EXECLIST = "$$EXPORT_EXEC_LIST$$";

    const string IMPORT_DIRECTORY           = "importer/";
    const string IMPORT_TPL_SCRIPTOBJ       = "ScriptObjTemplate_Import.txt";
    const string IMPORT_TPL_SCRIPTOBJ_ALL   = "AllScriptObjTemplate_Import.txt";
    const string IMPORT_TPL_JSON            = "JsonTemplate_Import.txt";
    const string IMPORT_TPL_JSON_ALL        = "AllJsonTemplate_Import.txt";
    const string IMPORT_FILENAME_SUFFIX     = "_Import";
    const string EXPORT_TPL_SCRIPTOBJ       = "ScriptObjTemplate_Export.txt";
    const string EXPORT_TPL_SCRIPTOBJ_ALL   = "AllScriptObjTemplate_Export.txt";
    const string EXPORT_TPL_JSON            = "JsonTemplate_Export.txt";
    const string EXPORT_TPL_JSON_ALL        = "AllJsonTemplate_Export.txt";
    const string EXPORT_FILENAME_SUFFIX     = "_Export";
    const string ASSETS_HOME                = "Assets/";
    const string ASSETS_RESOURCE            = "Assets/Resources/";
    const string PREFIX_CLASS               = "Class_";
    const string PREFIX_ACCESS              = "";
    const string PREFIX_DATA                = "Data_";
    const string PREFIX_SCRIPTOBJ           = "ScriptObj_";
    const string PREFIX_JSON                = "Json_";
    const string PREFIXSIGN_ALL             = "All";

    const string PREFS_CLASS_DIRECTORY      = ".classdir";
    const string PREFS_DATA_DIRECTORY       = ".datadir";
    const string PREFS_IMPORTER_JSON        = ".importjson";
    const string PREFS_IMPORTER_SCRIPTOBJ   = ".importsobj";
    const string PREFS_TOGETHER_CLASS       = ".together";
    const string PREFS_SHEET_NO             = ".sheetno";
    const string PREFS_PREFIX_TABLE         = ".pretable";
    const string PREFS_SUFFIX_TABLE         = ".suftable";
    const string PREFS_PREFIX_ACCESSOR      = ".preaccess";
    const string PREFS_SUFFIX_ACCESSOR      = ".sufaccess";
    const string PREFS_USED_ACCESSOR        = ".usedaccess";
    const string PREFS_PREFIX_DATA          = ".predata";
    const string PREFS_SUFFIX_DATA          = ".sufdata";

    static readonly string CLASS_NAME                = $"{nameof(XlsToJson)}";
    static readonly string MSG_ROWMAXOVER            = "[{0}] 行数が作成可能最大数を超えています[{1}].\r\n- セルに見えない空白文字が含まれている可能性があります.\r\n- どうしても最大を増やしたい場合は、" + CLASS_NAME + ".ROWS_MAX の値を増やします.";
    static readonly string MSG_COLMAXOVER            = "[{0}] 列数が作成可能最大数を超えています[{1}].\r\n- セルに見えない空白文字が含まれている可能性があります.\r\n- どうしても最大を増やしたい場合は、" + CLASS_NAME + ".COLS_MAX の値を増やします.";
    static readonly string MSG_ID_ONLYONE            = "[{0}: {1}] ID は 1 テーブルに 1 つのみです.";
    static readonly string MSG_ENUMNAME_ONLYONE      = "[{0}: {1}] 同名の ENUM が既に存在します.";
    static readonly string MSG_ENUMNAME_NOTFOUND     = "[{0}: {1}] ENUM のグループ名がありません.";
    static readonly string MSG_CONSTNAME_ONLYONE     = "[{0}: {1}] 同名の CONST クラスが既に存在します.";
    static readonly string MSG_CONSTNAME_NOTFOUND    = "[{0}: {1}] CONST のクラス名がありません.";
    static readonly string MSG_SAMEMEMBER            = "[{0}: {1}] 既に同じメンバーがあります. '{2}'";
    static readonly string MSG_TYPE_NOTFOUND         = "[{0}: {1}] タイプがありません. '{2}'";
    static readonly string MSG_NEED_BLANKROW         = "[{0}: {1}] 各カテゴリ間は最低 {2} 行マージンを取る必要があります.";
    static readonly string MSG_CLASSTMPL_NOTFOUND    = "{0}({1}) が見つかりません.";
    static readonly string MSG_DIRECTORY_INVALID     = "ディレクトリは {0} から始まる相対パスを指定してください.";
    static readonly string MSG_CREATE_ENVIRONMENT    = "環境を作成しました.\r\nTools/XlsToJson メニューで使用可能です.\r\n\r\n作成ファイル:\r\n{0}";
    static readonly string MSG_USE_SAME_CLASS        = "{0} がベースクラスとして適用されます.";
    static readonly string MSG_SHEETNAME_CANT_JP     = "{0}: 日本語名のシートは無視します.";
    static readonly string MSG_NEED_AUTONUMBER_FIELD = "{0}: オートナンバーフィールド ID(int) がありません.";
    static readonly string MSG_NOT_FOUND_ENUMTBL     = "{0}: enum で存在しないシート名が指定されています.";
    static readonly string MSG_CREATE_ACCESS         = "*データのシングルトンアクセスを可能にします";
    static readonly string MSG_BASECLASS_NOTFOUND    = "{0}: ベースクラス [{1}] がありません.";
    static readonly string MSG_BASECLASS_UNMATCH     = "{0}: ベースクラス [{1}] とテーブルの型が異なります.";

    static readonly string MSG_CANCEL                = "ユーザーキャンセルされました。";

    public static readonly string MSG_JSON_EXPORT_CONFIRM
                                                     = "JsonData を [{0}] に書き出します. よろしいですか？";
    public static readonly string MSG_SCRIPTOBJ_EXPORT_CONFIRM
                                                     = "ScriptableObject を [{0}] に書き出します. よろしいですか？";
    public static readonly string MSG_JSON_IMPORT_CONFIRM
                                                     = "[{0}] から JsonData を作成します. よろしいですか？";
    public static readonly string MSG_SCRIPTOBJ_IMPORT_CONFIRM
                                                     = "[{0}] から ScriptableObject を作成します. よろしいですか？";

    /// <summary>
    /// ポジションインデックス
    /// </summary>
    public class PosIndex
    {
        /// <summary>行</summary>
        public int R;
        /// <summary>列</summary>
        public int C;
        /// <summary>(Enumなどの)名前</summary>
        public string Name;
    }
    
    /// <summary>
    /// クラスメンバ、enum メンバ、const メンバを示す
    /// </summary>
    class Member
    {
        /// <summary>型</summary>
        public string Type;
        /// <summary>追加後尾文字列</summary>
        public string Suffix;
        /// <summary>コメント</summary>
        public string Comment;
        /// <summary>FindRow 可能なインデクサは true</summary>
        public bool   Indexer;
        /// <summary>enum は true</summary>
        public bool   IsEnum;
    }
    
    /// <summary>
    /// クラス情報
    /// </summary>
    class ClassInfo
    {
        /// <summary></summary>
        public string Comment;
        /// <summary></summary>
        public Dictionary<string, Member> Members = new Dictionary<string, Member>();
    }

    /// <summary>
    /// enum 情報
    /// </summary>
    class EnumInfo
    {
        /// <summary></summary>
        public string GroupName;
        /// <summary></summary>
        public string Comment;
        /// <summary></summary>
        public Dictionary<string, Member> Members = new Dictionary<string, Member>();
    }
    
    /// <summary>
    /// const 情報
    /// </summary>
    class ConstInfo
    {
        /// <summary></summary>
        public Dictionary<string, Member> Members = new Dictionary<string, Member>();
    }
    
    /// <summary>
    /// シート情報
    /// </summary>
    class SheetEntity
    {
        public Dictionary<string, ClassInfo> Classes = new Dictionary<string, ClassInfo>();
        public Dictionary<string, EnumInfo>  Enums   = new Dictionary<string, EnumInfo>();
        public Dictionary<string, ConstInfo> Consts  = new Dictionary<string, ConstInfo>();
        public Dictionary<string, PosIndex>  PosList = new Dictionary<string, PosIndex>();
        public string[,]                     Grid;
        public ISheet                        Sheet;
        public string                        Text       = null;
        public string                        AccessText = null;
        public string                        Hash;
        public string                        SheetName;
        public string                        ClassName;
        public string                        TableName;
        public string                        AccessorName;
        public string                        DataName;

        /// <summary>
        /// クラスや enum の情報からハッシュを取得する
        /// </summary>
        public void CreateHash()
        {
            StringBuilder sb = new StringBuilder();
            if (Classes.Count == 0)
            {
                Hash = null;
            }
            else
            {
                foreach (var cls in Classes)
                {
                    sb.Append(cls.Key + " ");
                    foreach (var member in cls.Value.Members)
                    {
                        sb.Append(member.Key + " ");
                        sb.Append(member.Value.Type + " ");
                        sb.Append(member.Value.Suffix + " ");
                    }
                }

                SHA256CryptoServiceProvider hashProvider = new SHA256CryptoServiceProvider();
                Hash =
                    string.Join(
                        "",
                        hashProvider.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString()))
                        .Select(x => $"{x:x2}")
                    );
            }
        }
    }

    class FileNameEx
    {
        public string Prefix;
        public string Suffix;
        public bool   Used;
    }

    static string                     colTexts = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

    static List<SheetEntity>          sheetList;
    static string                     xlsPath;
    static string                     prefsKey;
    static string                     classDir;
    static string                     dataDir;
    static FileNameEx                 table;
    static FileNameEx                 accessor;
    static FileNameEx                 data;
    static bool                       importerJson;
    static bool                       importerScriptObj;
    static int                        preSheetNo;
    static int                        sheetNo;

    Vector2                           scroll;
    Vector2                           scroll2;

    /// <summary>
    /// GUI window
    /// </summary>
    void OnGUI()
    {
        if (sheetList == null || sheetList.Count == 0)
        {
            Close();
            return;
        }
        // 前回選択していたシート番号のページがない
        if (sheetNo >= sheetList.Count)
        {
            sheetNo = sheetList.Count - 1;
        }

        GUILayout.Space(20);

        GUILayout.BeginHorizontal();
        GUILayout.Label("base class dir:", GUILayout.Width(150));
        classDir = EditorGUILayout.TextField("", classDir);
        GUILayout.EndHorizontal();

        GUILayout.BeginHorizontal();
        GUILayout.Label("output dir:", GUILayout.Width(150));
        dataDir = EditorGUILayout.TextField("", dataDir);
        GUILayout.EndHorizontal();
        GUILayout.Space(20);

        GUILayout.BeginHorizontal();
        GUILayout.Label("json", GUILayout.Width(150));
        importerJson = EditorGUILayout.Toggle("", importerJson, GUILayout.Width(30));
        GUILayout.EndHorizontal();
        GUILayout.BeginHorizontal();
        GUILayout.Label("scriptable object", GUILayout.Width(150));
        importerScriptObj = EditorGUILayout.Toggle("", importerScriptObj, GUILayout.Width(30));
        GUILayout.EndHorizontal();
        GUILayout.Space(20);

        SheetEntity entity = sheetList[sheetNo];

        GUILayout.BeginHorizontal();
        GUI.skin.textField.alignment = TextAnchor.LowerRight;
        GUILayout.Space(155);
        GUILayout.Label("[prefix]", new GUILayoutOption[] { GUILayout.Width(100) });
        GUILayout.Label("[name]", new GUILayoutOption[] { GUILayout.Width(200) });
        GUILayout.Label("[suffix]", new GUILayoutOption[] { GUILayout.Width(100) });
        GUILayout.EndHorizontal();

        // table class name
        GUILayout.BeginHorizontal();
        GUILayout.Label("base class name:", GUILayout.Width(150));
        GUI.skin.textField.alignment = TextAnchor.MiddleRight;
        table.Prefix = EditorGUILayout.TextField(table.Prefix, new GUILayoutOption[] { GUILayout.Width(100) });
        GUI.skin.textField.alignment = TextAnchor.MiddleCenter;
        GUILayout.Label(entity.ClassName, GUILayout.Width(200));
        GUI.skin.label.alignment = TextAnchor.LowerLeft;
        table.Suffix = EditorGUILayout.TextField(table.Suffix, new GUILayoutOption[] { GUILayout.Width(100) });
        GUILayout.EndHorizontal();

        // accessor class name
        GUILayout.BeginHorizontal();
        GUILayout.Label("accessor class name:", GUILayout.Width(150));
        GUI.skin.textField.alignment = TextAnchor.MiddleRight;
        accessor.Prefix = EditorGUILayout.TextField(accessor.Prefix, new GUILayoutOption[] { GUILayout.Width(100) });
        GUI.skin.textField.alignment = TextAnchor.LowerLeft;
        GUILayout.Label(entity.ClassName, GUILayout.Width(200));
        GUI.skin.label.alignment = TextAnchor.LowerLeft;
        accessor.Suffix = EditorGUILayout.TextField(accessor.Suffix, new GUILayoutOption[] { GUILayout.Width(100) });
        GUILayout.Space(20);
        accessor.Used = EditorGUILayout.Toggle("", accessor.Used, GUILayout.Width(30));
        EditorGUILayout.LabelField(MSG_CREATE_ACCESS);
        GUILayout.EndHorizontal();

        // data class name
        GUILayout.BeginHorizontal();
        GUILayout.Label("output name:", GUILayout.Width(150));
        GUI.skin.textField.alignment = TextAnchor.MiddleRight;
        data.Prefix = EditorGUILayout.TextField(data.Prefix, new GUILayoutOption[] { GUILayout.Width(100) });
        GUI.skin.textField.alignment = TextAnchor.LowerLeft;
        GUILayout.Label(entity.SheetName, GUILayout.Width(200));
        GUI.skin.label.alignment = TextAnchor.LowerLeft;
        data.Suffix = EditorGUILayout.TextField(data.Suffix, new GUILayoutOption[] { GUILayout.Width(100) });
        GUILayout.EndHorizontal();
        GUILayout.Space(20);

        List<string> names = new List<string>();
        foreach (SheetEntity ent in sheetList)
        {
            names.Add(ent.Sheet.SheetName);
        }

        bool change = false;

        GUILayout.BeginHorizontal();
        sheetNo = EditorGUILayout.Popup(sheetNo, names.ToArray(), GUILayout.Width(150));

        if (GUILayout.Button("refresh", GUILayout.MaxWidth(150)))
        {
            change = true;
        }
        if (preSheetNo != sheetNo)
        {
            entity = sheetList[sheetNo];

            change = true;
            preSheetNo = sheetNo;
        }

        if (change == true)
        {
            completePreSuffixName(entity);

            entity.AccessText = "";

            if (entity.ClassName == entity.Sheet.SheetName)
            {
                entity.Text = createTableClass(entity);
                if (accessor.Used == true)
                {
                    entity.AccessText = createTableAccess(entity);
                }
            }
            else
            {
                entity.Text = string.Format(MSG_USE_SAME_CLASS, entity.ClassName);
            }
        }
        GUILayout.EndHorizontal();

        GUILayout.BeginHorizontal();
        GUI.skin.textField.alignment = TextAnchor.UpperLeft;
        scroll = EditorGUILayout.BeginScrollView(scroll);
        EditorGUILayout.TextArea("// [Table]\r\n\r\n" + entity.Text, GUILayout.ExpandHeight(true));
        EditorGUILayout.EndScrollView();
        scroll2 = EditorGUILayout.BeginScrollView(scroll2);
        EditorGUILayout.TextArea("// [Accessor]\r\n\r\n" + entity.AccessText, GUILayout.ExpandHeight(true));
        EditorGUILayout.EndScrollView();
        GUILayout.EndHorizontal();
        GUILayout.Space(20);

        using (new EditorGUILayout.HorizontalScope())
        {
            if (GUILayout.Button("CREATE Importer"))
            {
                completePreSuffixNameAll();
                if (checkSaveDirectory() == true)
                {
                    saveSheet(sheetList);

                    Dialog(MSG_CREATE_ENVIRONMENT, getWriteFileList());

                    Close();
                }
            }
            if (GUILayout.Button("close"))
            {
                checkSaveDirectory();
                Close();
            }
        }
    }
    
    /// <summary>
    /// XLS を右クリック - XlsToJson
    /// </summary>
    [MenuItem ("Assets/XlsToJson Settings...", priority = 999999)]
    static void Import()
    {
        if (Selection.objects.Length == 0)
        {
            LogError("no selecting excell");
            return;
        }
        var obj = Selection.objects[0];

        xlsPath  = AssetDatabase.GetAssetPath(obj);
        prefsKey = Path.GetFileNameWithoutExtension(xlsPath);

        // prefs からディレクトリを復帰（エクセル名単位）
        loadPrefs();

        using (FileStream stream = File.Open (xlsPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            var window = ScriptableObject.CreateInstance<XlsToJson>();
            sheetList = new List<SheetEntity>();

            IWorkbook book = null;
            if (Path.GetExtension(xlsPath) == ".xls")
            {
                book = new HSSFWorkbook(stream);
            }
            else
            {
                book = new XSSFWorkbook(stream);
            }

            bool error = false;

            for (int sheetno = 0; sheetno < book.NumberOfSheets; ++sheetno)
            {
                ISheet sheet = book.GetSheetAt(sheetno);

                // 日本語のシートは無視
                if (Encoding.GetEncoding("Shift_JIS").GetByteCount(sheet.SheetName) != sheet.SheetName.Length)
                {
                    LogWarning(MSG_SHEETNAME_CANT_JP, sheet.SheetName);
                    continue;
                }

                SheetEntity entity = new SheetEntity();

                entity.Sheet = sheet;
                entity.Grid = GetGrid(sheet, entity.PosList);
                // 解析失敗
                if (entity.Grid == null)
                {
                    continue;
                }

                setClassName(entity);

                // クラス解析
                if (entity.PosList.ContainsKey(TRIGGER_ID) == true)
                {
                    if (analyzeClasses(entity) == false)
                    {
                        continue;
                    }
                }
                else
                {
                    LogWarning(MSG_NEED_AUTONUMBER_FIELD, sheet.SheetName);
                    continue;
                }

                // enum解析
                if (analyzeEnumsAndConsts(entity) == false)
                {
                    continue;
                }

                // クラスコメント解析（あれば）
                analyzeClassComments(entity);

                entity.CreateHash();

                sheetList.Add(entity);
            }

            if (checkBaseClassVerified() == false)
            {
                return;
            }

            if (error == false)
            {
                window.Show();
            }
        }
    }

    /// <summary>
    /// シート名からベースクラス名を取得する
    /// Stage2:Stage というシートがある場合、Stage がベースクラスとなる
    /// </summary>
    static void setClassName(SheetEntity entity)
    {
        string name = entity.Sheet.SheetName;

        if (name.IndexOf("@") > 0)
        {
            string[] names = name.Split('@');

            entity.SheetName = names[0];
            entity.ClassName = names[1];
        }
        else
        {
            entity.SheetName = 
            entity.ClassName = name;
        }
    }

    /// <summary>
    /// ベースクラスが存在するか、テーブルの型が一致しているかを確認する
    /// </summary>
    /// <returns>true..問題なし, false..問題あり</returns>
    static bool checkBaseClassVerified()
    {
        Dictionary<string, string> hashClasses = new Dictionary<string, string>();

        foreach (SheetEntity ent in sheetList)
        {
            if (ent.SheetName == ent.ClassName)
            {
                hashClasses.Add(ent.ClassName, ent.Hash);
            }
        }

        foreach (SheetEntity ent in sheetList)
        {
            if (ent.SheetName != ent.ClassName)
            {
                if (hashClasses.ContainsKey(ent.ClassName) == false)
                {
                    LogError(MSG_BASECLASS_NOTFOUND, ent.SheetName, ent.ClassName);
                    return false;
                }
                if (hashClasses[ent.ClassName] != ent.Hash)
                {
                    LogError(MSG_BASECLASS_UNMATCH, ent.SheetName, ent.ClassName);
                    return false;
                }
            }
        }

        return true;
    }

    /// <summary>
    /// テーブルクラス名、アクセッサクラス名を作成
    /// </summary>
    static void completePreSuffixNameAll()
    {
        foreach (SheetEntity ent in sheetList)
        {
            completePreSuffixName(ent);
        }
    }

    /// <summary>
    /// テーブルクラス名、アクセッサクラス名を作成
    /// </summary>
    static void completePreSuffixName(SheetEntity entity)
    {
        entity.TableName    = table.Prefix + entity.ClassName + table.Suffix;
        entity.AccessorName = accessor.Prefix + entity.ClassName + accessor.Suffix;
        entity.DataName     = data.Prefix + entity.SheetName + data.Suffix;
    }
    
    /// <summary>
    /// コメント文字列を除去して返す
    /// </summary>
    static string removeSignComment(string str)
    {
        if (str == null)
        {
            return null;
        }
        foreach (string sign in SIGN_COMMENTS)
        {
            str = str.Replace(sign, "");
        }
        return str;
    }
    
    /// <summary>
    /// シートの最大行数、最大列数を調べる
    /// </summary>
    static bool getRowAndColumnMax(ISheet sheet, out int rowMax, out int colMax)
    {
        string name  = sheet.SheetName;

        rowMax = sheet.LastRowNum+1;
        colMax = 0;

        for (int r = 0; r < rowMax; r++)
        {
            IRow row = sheet.GetRow(r);
            if (checkRowIsNullOrEmpty(row) == true)
            {
                continue;
            }

            ICell cell = row.Cells[row.Cells.Count-1];
            if (colMax < cell.ColumnIndex+1)
            {
                colMax = cell.ColumnIndex+1;
            }
        }
        
        // 予め決めておいたバッファ最大量を超える場合、エラー
        if (rowMax >= ROWS_MAX)
        {
            LogError(MSG_ROWMAXOVER, name, $"{rowMax} >= {ROWS_MAX}");
            return false;
        }
        if (colMax >= COLS_MAX)
        {
            LogError(MSG_COLMAXOVER, name, $"{colMax} >= {COLS_MAX}");
            return false;
        }

        return true;
    }
    
    /// <summary>
    /// １行分のセルデータを取得
    /// </summary>
    /// <param name="sheet">シート</param>
    /// <param name="grid">グリッド情報</param>
    /// <param name="posList">ID, [CLASS], enum のデータポジションを示すリスト</param>
    /// <param name="r">行数</param>
    /// <param name="marginrow_between_category">カテゴリ先頭までに、この値が 0 である必要がある</param>
    /// <returns>false..取得失敗</returns>
    static bool getCells(ISheet sheet, string[,] grid, Dictionary<string, PosIndex> posList, int r, ref int marginrow_between_category)
    {
        string name  = sheet.SheetName;

        IRow   row   = sheet.GetRow(r);
        if (checkRowIsNullOrEmpty(row) == true)
        {
            if (--marginrow_between_category < 0)
            {
                marginrow_between_category = 0;
            }
            return true;
        }

        IRow   rowDown = null;
        if (r+1 <= sheet.LastRowNum)
        {
            rowDown = sheet.GetRow(r+1);
        }

        for (int c = 0; c < row.Cells.Count; c++)
        {
            ICell    cell     = row.Cells[c];
            string   cellstr  = cell.ToString();
            bool     category = false;
            CellType celltype = cell.CellType == CellType.Formula ? cell.CachedFormulaResultType : cell.CellType;

            switch (celltype)
            {
                case CellType.Numeric:
                    cellstr = cell.NumericCellValue.ToString();
                    break;
                case CellType.Boolean:
                    cellstr = cell.BooleanCellValue.ToString();
                    break;
                case CellType.String:
                    cellstr = cell.StringCellValue.ToString();
                    break;
            }

            int col = cell.ColumnIndex;
            grid[r, col] = cellstr;

            // テーブルのトリガー ID
            if (cellstr == TRIGGER_ID)
            {
                if (posList.ContainsKey(TRIGGER_ID) == true)
                {
                    // 既に ID がある
                    LogError(MSG_ID_ONLYONE, name, GetXLS_RC(r, col));
                    return false;
                }
                posList.Add(TRIGGER_ID, new PosIndex(){ R=r, C=col});

                category = true;
            }
            else
            // テーブルのサブクラスコメント
            if (cellstr == TRIGGER_SUBCLASS)
            {
                if (posList.ContainsKey(TRIGGER_SUBCLASS) == true)
                {
                    // 既に ID がある
                    LogError(MSG_ID_ONLYONE, name, GetXLS_RC(r, col));
                    return false;
                }
                posList.Add(TRIGGER_SUBCLASS, new PosIndex(){ R=r, C=col});

                category = true;
            }
            else
            // enum グループ
            if (cellstr == TRIGGER_ENUM)
            {
                if (rowDown == null)
                {
                    // enum 名がない
                    LogError(MSG_ENUMNAME_NOTFOUND, name, GetXLS_RC(r, col));
                    return false;
                }
                string ename = rowDown.GetCell(col).ToString();
                string key   = name + "." + ename;
                if (posList.ContainsKey(key) == true)
                {
                    // 既に同じ enum がある
                    LogError(MSG_ENUMNAME_ONLYONE, name, GetXLS_RC(r, col));
                    return false;
                }
                posList.Add(key, new PosIndex(){ R=r+1, C=col, Name=ename });

                category = true;
            }
            else
            // global enum グループ
            if (cellstr.ToLower() == TRIGGER_GLOBAL_ENUM.ToLower())
            {
                if (rowDown == null)
                {
                    // enum 名がない
                    LogError(MSG_ENUMNAME_NOTFOUND, name, GetXLS_RC(r, col));
                    return false;
                }
                string ename = rowDown.GetCell(col).ToString();
                string key   = TRIGGER_GLOBAL_ENUM + "." + ename;
                if (posList.ContainsKey(key) == true)
                {
                    // 既に同じ enum がある
                    LogError(MSG_ENUMNAME_ONLYONE, name, GetXLS_RC(r, col));
                    return false;
                }
                posList.Add(key, new PosIndex(){ R=r+1, C=col, Name=ename });

                category = true;
            }
            else
            // const グループ
            if (cellstr == TRIGGER_CONST)
            {
                if (rowDown == null)
                {
                    // const 名がない
                    LogError(MSG_CONSTNAME_NOTFOUND, name, GetXLS_RC(r, col));
                    return false;
                }
                string ename = rowDown.GetCell(col).ToString();
                string key   = TRIGGER_CONST + "." + ename;
                if (posList.ContainsKey(key) == true)
                {
                    // 既に同じ const がある
                    LogError(MSG_CONSTNAME_ONLYONE, name, GetXLS_RC(r, col));
                    return false;
                }
                posList.Add(key, new PosIndex(){ R=r+1, C=col, Name=ename });

                category = true;
            }

            if (category == true)
            {
                if (marginrow_between_category > 0)
                {
                    // 各カテゴリ間は最低マージン 2 行必要
                    LogError(MSG_NEED_BLANKROW, name, GetXLS_RC(r, col), MARGINROW_BETWEEN_CTG);
                    return false;
                }
                marginrow_between_category = MARGINROW_BETWEEN_CTG;
            }
        }

        return true;
    }
    
    /// <summary>
    /// Row が null または空行か確認する
    /// </summary>
    /// <param name="row">確認する行</param>
    /// <returns>true..null または空行</returns>
    static bool checkRowIsNullOrEmpty(IRow row)
    {
        if (row == null)
        {
            return true;
        }

        // エクセル上、見た目に何もないがデータとして "" だけ検出される行を null とみなし無視する
        // null ではないのに、Cells.Count = 0 の row を返すこともある…
        for (int c = 0; c < row.Cells.Count; c++)
        {
            // なにか入っていた
            if (string.IsNullOrEmpty(row.Cells[c].ToString()) == false)
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// １行上にコメントがある場合、その文字列を返す
    /// </summary>
    static string getCommentUp(string[,] grid, int r, int c)
    {
        string cell = null;
        if (r >= 1)
        {
            cell = grid[r-1, c];
        }
        if (CheckSignComment(cell) == true)
        {
            return removeSignComment(cell);
        }
        return cell;
    }

    /// <summary>
    /// １行右にコメントがある場合、その文字列を返す
    /// </summary>
    static string getCommentRight(string[,] grid, int r, int c)
    {
        string cell = null;
        if ((c+1) < (grid.GetUpperBound(1)+1))
        {
            cell = grid[r, c+1];
        }
        if (CheckSignComment(cell) == true)
        {
            return removeSignComment(cell);
        }
        return cell;
    }
    
    /// <summary>
    /// XlsToJson.cs のあるディレクトリを取得する
    /// </summary>
    static string searchXlsToJsonDirectory()
    {
        string[] files = Directory.GetFiles(Application.dataPath, $"{nameof(XlsToJson)}.cs", SearchOption.AllDirectories);
        if (files != null && files.Length == 1)
        {
            // フルパスから相対パスに
            string path = Path.GetDirectoryName(files[0]).Replace("\\", "/");
            path = ASSETS_HOME + path.Replace(Application.dataPath + "/", "");
            return path;
        }
        return null;
    }
    
    /// <summary>
    /// セーブディレクトリが適切かチェック＆修正. 問題なければ prefs に保存
    /// </summary>
    /// <returns>true..成功</returns>
    static bool checkSaveDirectory()
    {
        // windows と mac の違いを吸収
        classDir = classDir.Replace("\\", "/").Trim();
        dataDir  = dataDir.Replace("\\", "/").Trim();

        // パスは Assets/ から始まる相対パス
        if (classDir.IndexOf(ASSETS_HOME) != 0)
        {
            DialogError(MSG_DIRECTORY_INVALID, ASSETS_HOME);
            return false;
        }
        if (dataDir.IndexOf(ASSETS_HOME) != 0)
        {
            DialogError(MSG_DIRECTORY_INVALID, ASSETS_HOME);
            return false;
        }
        
        // フォルダを予め作成しておく
        CompleteDirectory(classDir);
        CompleteDirectory(dataDir);

        // prefs にディレクトリを保存（エクセル名単位）
        savePrefs();

        return true;
    }

    /// <summary>
    /// prefs にディレクトリを保存（エクセル名単位）
    /// </summary>
    static void savePrefs()
    {
        EditorPrefs.SetString(prefsKey + PREFS_CLASS_DIRECTORY, classDir);
        EditorPrefs.SetString(prefsKey + PREFS_DATA_DIRECTORY, dataDir);
        EditorPrefs.SetBool(prefsKey + PREFS_IMPORTER_JSON, importerJson);
        EditorPrefs.SetBool(prefsKey + PREFS_IMPORTER_SCRIPTOBJ, importerScriptObj);
        EditorPrefs.SetInt(prefsKey + PREFS_SHEET_NO, sheetNo);
        EditorPrefs.SetString(prefsKey + PREFS_PREFIX_TABLE, table.Prefix);
        EditorPrefs.SetString(prefsKey + PREFS_SUFFIX_TABLE, table.Suffix);
        EditorPrefs.SetString(prefsKey + PREFS_PREFIX_ACCESSOR, accessor.Prefix);
        EditorPrefs.SetString(prefsKey + PREFS_SUFFIX_ACCESSOR, accessor.Suffix);
        EditorPrefs.SetBool(prefsKey + PREFS_USED_ACCESSOR, accessor.Used);
        EditorPrefs.SetString(prefsKey + PREFS_PREFIX_DATA, data.Prefix);
        EditorPrefs.SetString(prefsKey + PREFS_SUFFIX_DATA, data.Suffix);
    }

    /// <summary>
    /// prefs からディレクトリを復帰（エクセル名単位）
    /// </summary>
    static void loadPrefs()
    {
        table     = new FileNameEx();
        accessor  = new FileNameEx();
        data      = new FileNameEx();

        classDir          = EditorPrefs.GetString(prefsKey + PREFS_CLASS_DIRECTORY);
        dataDir           = EditorPrefs.GetString(prefsKey + PREFS_DATA_DIRECTORY);
        importerJson      = EditorPrefs.GetBool(prefsKey + PREFS_IMPORTER_JSON, true);
        importerScriptObj = EditorPrefs.GetBool(prefsKey + PREFS_IMPORTER_SCRIPTOBJ, true);
        sheetNo           = EditorPrefs.GetInt(prefsKey + PREFS_SHEET_NO);
        preSheetNo        = -1;

        // 初期値
        if (string.IsNullOrEmpty(classDir) == true || classDir.IndexOf(ASSETS_HOME) != 0)
        {
            classDir = ASSETS_HOME;
        }
        if (string.IsNullOrEmpty(dataDir) == true || dataDir.IndexOf(ASSETS_HOME) != 0)
        {
            dataDir = ASSETS_RESOURCE;
        }

        table.Prefix    = EditorPrefs.GetString(prefsKey + PREFS_PREFIX_TABLE, PREFIX_CLASS);
        table.Suffix    = EditorPrefs.GetString(prefsKey + PREFS_SUFFIX_TABLE, "");
        accessor.Prefix = EditorPrefs.GetString(prefsKey + PREFS_PREFIX_ACCESSOR, PREFIX_ACCESS);
        accessor.Suffix = EditorPrefs.GetString(prefsKey + PREFS_SUFFIX_ACCESSOR, "");
        accessor.Used   = EditorPrefs.GetBool(prefsKey + PREFS_USED_ACCESSOR, true);
        data.Prefix     = EditorPrefs.GetString(prefsKey + PREFS_PREFIX_DATA, PREFIX_DATA);
        data.Suffix     = EditorPrefs.GetString(prefsKey + PREFS_SUFFIX_DATA, "");
    }

    /// <summary>
    /// コメントテキスト生成
    /// </summary>
    /// <param name="sb">Append する StringBuilder</param>
    /// <param name="description">説明</param>
    /// <param name="indent">タブインデントの数</param>
    static void addCommentText(StringBuilder sb, string description, int indent)
    {
        string tab = "".PadLeft(indent, '\t');

        if (string.IsNullOrEmpty(description) == true)
        {
//            sb.AppendLine($"{tab}///<summary>\r\n{tab}/// \r\n{tab}///</summary>");
        }
        else
        {
            string[] descs = description.Replace("\r", "").Split('\n');
            sb.AppendLine($"{tab}///<summary>");
            foreach (string desc in descs)
            {
                sb.AppendLine($"{tab}/// {desc}");
            }
            sb.AppendLine($"{tab}///</summary>");
        }
    }
    
    /// <summary>
    /// Path.Combine の後、フォルダ区切りを / にして返す
    /// </summary>
    static string pathCombine(string path0, string path1)
    {
        return Path.Combine(path0, path1).Replace("\\", "/");
    }

    /// <summary>
    /// ログ表示
    /// </summary>
    public static void Log(string msg, params object[] objs)
    {
        if (objs != null)
        {
            msg = string.Format(msg, objs);
        }
        Debug.Log($"{CLASS_NAME}:" + msg);
    }

    /// <summary>
    /// 警告表示
    /// </summary>
    public static void LogWarning(string msg, params object[] objs)
    {
        if (objs != null)
        {
            msg = string.Format(msg, objs);
        }
        Debug.LogWarning($"{CLASS_NAME}:" + msg);
    }

    /// <summary>
    /// エラー表示
    /// </summary>
    public static void LogError(string msg, params object[] objs)
    {
        if (objs != null)
        {
            msg = string.Format(msg, objs);
        }
        Debug.LogError($"{CLASS_NAME}:" + msg);
    }


    /// <summary>
    /// ダイアログ表示
    /// </summary>
    public static void Dialog(string msg, params object[] objs)
    {
        if (objs != null)
        {
            msg = string.Format(msg, objs);
        }
        EditorUtility.DisplayDialog($"{CLASS_NAME}", msg, "ok");
    }

    /// <summary>
    /// エラー表示
    /// </summary>
    public static void DialogError(string msg, params object[] objs)
    {
        if (objs != null)
        {
            msg = string.Format(msg, objs);
        }
        EditorUtility.DisplayDialog($"{CLASS_NAME}", $"[ERROR]\r\n{msg}", "ok");
    }

    /// <summary>
    /// ok/cancel ダイアログ表示
    /// </summary>
    public static bool DialogSelect(string msg, params object[] objs)
    {
        if (objs != null)
        {
            msg = string.Format(msg, objs);
        }
        return EditorUtility.DisplayDialog($"{CLASS_NAME}", msg, "ok", "cancel");
    }

    /// <summary>
    /// キャンセルつき進捗バー (index+1)/max %
    /// </summary>
    public static bool CancelableProgressBar(int index, int max, string msg)
    {
        float	perc = (float)(index+1) / (float)max;
        
        bool result =
            EditorUtility.DisplayCancelableProgressBar(
                nameof(XlsToJson),
                perc.ToString("00.0%") + "　" + msg,
                perc
            );
        if (result == true)
        {
            EditorUtility.ClearProgressBar();
            Dialog(MSG_CANCEL);
            return true;
        }
        return false;
    }
}
