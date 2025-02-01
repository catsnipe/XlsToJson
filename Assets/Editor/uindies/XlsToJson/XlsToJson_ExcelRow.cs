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

    public ICell CellValue(int columnIndex, bool value, ICellStyle style = null)
    {
        ICell cell = GetCell(columnIndex);
        cell.SetCellValue(value);
        if (style != null)
        {
            cell.CellStyle = style;
        }
        return cell;
    }

    public ICell CellValue(int columnIndex, System.DateTime value, ICellStyle style = null)
    {
        ICell cell = GetCell(columnIndex);
        cell.SetCellValue(value);
        if (style != null)
        {
            cell.CellStyle = style;
        }
        return cell;
    }

    public ICell CellValue(int columnIndex, int value, ICellStyle style = null)
    {
        ICell cell = GetCell(columnIndex);
        cell.SetCellValue(value);
        if (style != null)
        {
            cell.CellStyle = style;
        }
        return cell;
    }

    public ICell CellValue(int columnIndex, float value, ICellStyle style = null)
    {
        ICell cell = GetCell(columnIndex);
        cell.SetCellValue(value);
        if (style != null)
        {
            cell.CellStyle = style;
        }
        return cell;
    }

    public ICell CellValue(int columnIndex, string value, ICellStyle style = null)
    {
        ICell cell = GetCell(columnIndex);
        cell.SetCellValue(value);
        if (style != null)
        {
            cell.CellStyle = style;
        }
        return cell;
    }

    public ICell GetCell(int columnIndex)
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
