using NPOI.SS.UserModel;

public interface iXlsToJsonImporter
{
    public void Exec(IWorkbook book, string exportDirectory);
    public string GetExportPath();
    public bool CheckSuccess();
}
