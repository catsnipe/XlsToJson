using NPOI.SS.UserModel;

public class ExcelRow
{
    IRow row;

    public ExcelRow(IRow _row)
    {
        row = _row;
    }

    public ExcelRow(ISheet sheet, int rowIndex)
    {
        row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
    }

    public void CellValue(int columnIndex, bool value)
    {
        getCell(columnIndex).SetCellValue(value);
    }

    public void CellValue(int columnIndex, System.DateTime value)
    {
        getCell(columnIndex).SetCellValue(value);
    }

    public void CellValue(int columnIndex, int value)
    {
        getCell(columnIndex).SetCellValue(value);
    }

    public void CellValue(int columnIndex, float value)
    {
        getCell(columnIndex).SetCellValue(value);
    }

    public void CellValue(int columnIndex, string value)
    {
        getCell(columnIndex).SetCellValue(value);
    }

    ICell getCell(int columnIndex)
    {
        return row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
    }

    public static void CopyRow(IRow srcrow, IRow newrow)
    {
        for (int c = 0; c <= srcrow.LastCellNum; c++)
        {
            ICell srccell = srcrow.GetCell(c);
            if (srccell == null)
            {
                continue;
            }
            ICell newcell = newrow.CreateCell(c);

            ICellStyle newcellStyle = newrow.Sheet.Workbook.CreateCellStyle();
            newcellStyle.CloneStyleFrom(srccell.CellStyle);
            newcell.CellStyle = newcellStyle;

            newcell.SetCellType(srccell.CellType);

            switch (srccell.CellType)
            {
                case CellType.Blank:
                    newcell.SetCellValue(srccell.StringCellValue);
                    break;
                case CellType.Boolean:
                    newcell.SetCellValue(srccell.BooleanCellValue);
                    break;
                case CellType.Error:
                    newcell.SetCellErrorValue(srccell.ErrorCellValue);
                    break;
                case CellType.Formula:
                    newcell.SetCellFormula(srccell.CellFormula);
                    break;
                case CellType.Numeric:
                    newcell.SetCellValue(srccell.NumericCellValue);
                    break;
                case CellType.String:
                    newcell.SetCellValue(srccell.RichStringCellValue);
                    break;
            }
        }

    }
}
