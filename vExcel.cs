using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Runtime.InteropServices;

namespace vExcel
{
    public class vExcel:IDisposable
    {
        private readonly List<vWorksheet> _vWorksheets = new List<vWorksheet>();
        public Application ThisApplication { get; set; }
        private readonly Workbook Workbook;
        private bool _deletedFirstSheet;
        private bool _hasDisposed;
        private vExcel()
        {
            ThisApplication = new Application();
            Workbook = ThisApplication.Workbooks.Add(Type.Missing);
        }

        public static vExcel Factory()
        {
            return new vExcel();
        }

        public vWorksheet PushNewSheet(string Name)
        {
            CheckUniqueName(Name);
            ThisApplication.Worksheets.Add(Type.Missing);
            if (!_deletedFirstSheet)
            {
                _deletedFirstSheet = true;
                var excelSheet1 = (Worksheet)ThisApplication.Worksheets[1];
                excelSheet1.Delete();
            }
            var excelSheet = (Worksheet)ThisApplication.Worksheets[1];
            var sheet = new vWorksheet(excelSheet, Name);
            _vWorksheets.Add(sheet);
            return sheet;
        }

        /// <summary>
        /// Removes sheet by name and returns the previously created sheet
        /// </summary>
        /// <param name="TabLabel">Name of the sheet to be removed.</param>
        /// <returns></returns>
        public vWorksheet PopSheetByName(string TabLabel)
        {
            var sheet = _vWorksheets.Find(e => e.TabLabel == TabLabel);
            ThisApplication.DisplayAlerts = false;
            sheet.GetWorksheet().Delete();
            ThisApplication.DisplayAlerts = true;
            _vWorksheets.RemoveAll(e => e.TabLabel == TabLabel);
            return _vWorksheets.Last();
        }

        public vWorksheet GetSheetByName(string TabLabel)
        {
            var sheet = _vWorksheets.Find(e => e.TabLabel == TabLabel);
            sheet._isUnset = true;
            return sheet;
        }

        public vWorksheet RenameSheetByName(string current, string newName)
        {
            var sheet = _vWorksheets.Find(e => e.TabLabel == current);
            if(sheet == null) throw new Exception($"Cannot rename, {current} was not found.");
            sheet.GetWorksheet().Name = newName;
            sheet.TabLabel = newName;
            sheet._isUnset = true;
            return sheet;
        }

        public vWorksheet CopySheetByName(string current, string newName)
        {
            var sheetOriginal = _vWorksheets.Find(e => e.TabLabel == current);
            if (sheetOriginal == null) throw new Exception($"Cannot copy, {current} was not found.");
            sheetOriginal.GetWorksheet().Copy((Worksheet) ThisApplication.Worksheets[1]);
            var newSheet = (Worksheet) ThisApplication.Worksheets[1];
            var sheet = new vWorksheet(newSheet, newName);
            _vWorksheets.Add(sheet);
            return sheet;
        }

        /// <summary>
        /// Full path to save
        /// </summary>
        /// <param name="path">Example: Directory.GetCurrentDirectory() + "\\test.xlsx"</param>
        public void SaveOverride(string path)
        {
            if (File.Exists(path)) File.Delete(path);
            Workbook.SaveAs(path);
            Workbook.Close();
        }

        public void Close()
        {
            ThisApplication.Quit();
            Marshal.ReleaseComObject(Workbook);
            Marshal.ReleaseComObject(ThisApplication);
            _hasDisposed = true;
        }

        public static void OpenInExcel(string path)
        {
            Thread.Sleep(1000);
            Process process = new Process();
            process.StartInfo.FileName = path;
            process.Start();
        }

        private void CheckUniqueName(string NewName)
        {
            var names = _vWorksheets.Select(e => e.TabLabel).ToList();
            if(names.Contains(NewName)) throw new Exception("Worksheet tab names must be unique.");
        }

        public void Dispose()
        {
            if (_hasDisposed) return;
            Close();
        }
    }

    public class vWorksheet
    {
        private readonly Worksheet Worksheet;
        public string TabLabel { get; set; }
        private int[] _range = {-1,-1,-1,-1};
        private bool _isRange;
        internal bool _isUnset = true;

        public vWorksheet(Worksheet worksheet, string tabLabel)
        {
            Worksheet = worksheet;
            Worksheet.Name = tabLabel;
            TabLabel = tabLabel;
        }

        #region selector
        public vWorksheet SelectCells(int TopX, int TopY, int BottomX, int BottomY)
        {
            if (TopX < 1 || TopY < 1 || BottomX < 1 || BottomY < 1) throw new Exception("Range must start at 1");
            if (TopY > BottomY) throw new Exception("Top coordinate must be above bottom.");
            if (TopX > BottomX) throw new Exception("Right coordinate must be before left.");
            _range[0] = TopX;
            _range[1] = TopY;
            _range[2] = BottomX;
            _range[3] = BottomY;
            _isRange = true;
            if (_isRange && BottomY == -1) throw new Exception("Range input is incorrect");
            _isUnset = false;
            return this;
        }

        public vWorksheet SelectCell(int X, int Y)
        {
            if (X < 1 || X < 1) throw new Exception("Range must start at 1");
            _range[0] = X;
            _range[1] = Y;
            _range[2] = X;
            _range[3] = Y;
            _isRange = false;
            _isUnset = false;
            return this;
        }
        #endregion

        #region value
        public vWorksheet SetValue(string value)
        {
            CellRangeAnySelector()[2].Value = value;
            return this;
        }

        public vWorksheet ReplaceValue(string current, string newValue)
        {
            CheckIfSelected();
            for (int i = _range[0]; i < _range[2] + 1; i++)
            {
                for (int j = _range[1]; j < _range[3] + 1; j++)
                {
                    var cell = Worksheet.Cells[j, i];
                    if (cell.Value == current) cell.Value = newValue;
                }
            }
            return this;
        }

        /// <summary>
        /// Replace any value that contains a substring of the current value
        /// </summary>
        /// <param name="current"></param>
        /// <param name="newValue"></param>
        /// <returns></returns>
        public vWorksheet ReplaceValueContaining(string containValue, string newValue)
        {
            CheckIfSelected();
            for (int i = _range[0]; i < _range[2] + 1; i++)
            {
                for (int j = _range[1]; j < _range[3] + 1; j++)
                {
                    var cell = Worksheet.Cells[j, i];
                    var text = (string)cell.Value;
                    if (text.Contains(containValue)) cell.Value = newValue;
                }
            }
            return this;
        }

        public vWorksheet ClearValue()
        {
            CellRangeAnySelector()[2].Value = "";
            return this;
        }
        #endregion

        #region Font
        public vWorksheet SetFontSize(int size)
        {
            CellRangeAnySelector()[2].Font.Size = size;
            return this;
        }

        public vWorksheet SetFontColor(Color Color)
        {
            CellRangeAnySelector()[2].Font.Color = Color;
            return this;
        }

        public vWorksheet SetFontFamily(string family)
        {
            CellRangeAnySelector()[2].Font.Name = family;
            return this;
        }

        public vWorksheet SetFontBold(bool isBold)
        {
            CellRangeAnySelector()[2].Font.Bold = isBold;
            return this;
        }

        public vWorksheet SetFontItalic(bool isItalic)
        {
            CellRangeAnySelector()[2].Font.Italic = isItalic;
            return this;
        }

        public vWorksheet SetFontUnderline(bool isUnderline)
        {
            CellRangeAnySelector()[2].Font.Underline = isUnderline;
            return this;
        }

        public vWorksheet SetFontStrikethrough(bool isStrikethrough)
        {
            CellRangeAnySelector()[2].Font.Strikethrough = isStrikethrough;
            return this;
        }

        public vWorksheet SetFontHorizontalCenter()
        {
            CellRangeAnySelector()[2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            return this;
        }
        public vWorksheet SetFontHorizontalLeft()
        {
            CellRangeAnySelector()[2].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            return this;
        }
        public vWorksheet SetFontHorizontalRight()
        {
            CellRangeAnySelector()[2].HorizontalAlignment = XlHAlign.xlHAlignRight;
            return this;
        }
        public vWorksheet SetFontVerticalCenter()
        {
            CellRangeAnySelector()[2].VerticalAlignment = XlVAlign.xlVAlignCenter;
            return this;
        }
        public vWorksheet SetFontVerticalBottom()
        {
            CellRangeAnySelector()[2].VerticalAlignment = XlVAlign.xlVAlignBottom;
            return this;
        }
        public vWorksheet SetFontVerticalTop()
        {
            CellRangeAnySelector()[2].VerticalAlignment = XlVAlign.xlVAlignTop;
            return this;
        }

        /// <summary>
        /// AutoSizeColumns will have no effect on cells with textwrap enabled.
        /// </summary>
        /// <param name="isWrap"></param>
        /// <returns></returns>
        public vWorksheet SetTextwrap(bool isWrap)
        {
            CellRangeAnySelector()[1].WrapText = isWrap;
            return this;
        }
        #endregion

        #region Row and column
        /// <summary>
        /// AutoSize relative to the selected cells.
        /// </summary>
        /// <returns></returns>
        public vWorksheet AutoSizeColumns()
        {
            dynamic range = CellRangeAnySelector()[1];
            range.EntireColumn.AutoFit();
            return this;
        }

        public vWorksheet AutoSizeColumnsRelative()
        {
            dynamic range = CellRangeAnySelector()[1];
            range.Columns.AutoFit();
            return this;
        }

        public vWorksheet SetColumnWidth(int width)
        {
            dynamic range = CellRangeAnySelector()[1];
            range.EntireColumn.ColumnWidth = width;
            return this;
        }
        public vWorksheet SetRowHeight(int height)
        {
            dynamic range = CellRangeAnySelector()[1];
            range.EntireRow.RowHeight = height;
            return this;
        }
        public vWorksheet FreezePaneRow(bool toggle, int row = 0)
        {
            Worksheet.Application.ActiveWindow.SplitRow = row;
            Worksheet.Application.ActiveWindow.FreezePanes = toggle;
            return this;
        }
        public vWorksheet FreezePaneColumn(bool toggle, int column = 0)
        {
            Worksheet.Application.ActiveWindow.SplitColumn = column;
            Worksheet.Application.ActiveWindow.FreezePanes = toggle;
            return this;
        }
        #endregion

        #region Border
        /// <summary>
        /// 0d,1d,2d,3d,4d are acceptable values
        /// </summary>
        /// <param name="Top"></param>
        /// <param name="Right"></param>
        /// <param name="Bottom"></param>
        /// <param name="Left"></param>
        /// <returns></returns>
        public vWorksheet SetBorderWeights(double Top, double Right, double Bottom, double Left)
        {
            dynamic range = CellRangeAnySelector()[1];
            if (Top != 0d) range.Borders[XlBordersIndex.xlEdgeTop].Weight = Top;
            if (Right != 0d) range.Borders[XlBordersIndex.xlEdgeRight].Weight = Right;
            if (Bottom != 0d) range.Borders[XlBordersIndex.xlEdgeBottom].Weight = Bottom;
            if (Left != 0d) range.Borders[XlBordersIndex.xlEdgeLeft].Weight = Left;
            return this;
        }

        public vWorksheet SetBorderWeightsEach(double Top, double Right, double Bottom, double Left)
        {
            CheckIfSelected();
            for (int i = _range[0]; i < _range[2] + 1; i++)
            {
                for (int j = _range[1]; j < _range[3] + 1; j++)
                {
                    var range = Worksheet.Range[Worksheet.Cells[j, i], Worksheet.Cells[j, i]];
                    if (Top != 0d) range.Borders[XlBordersIndex.xlEdgeTop].Weight = Top;
                    if (Right != 0d) range.Borders[XlBordersIndex.xlEdgeRight].Weight = Right;
                    if (Bottom != 0d) range.Borders[XlBordersIndex.xlEdgeBottom].Weight = Bottom;
                    if (Left != 0d) range.Borders[XlBordersIndex.xlEdgeLeft].Weight = Left;
                }
            }
            return this;
        }

        public vWorksheet SetBorderWeightEach(double weight)
        {
            CheckIfSelected();
            for (int i = _range[0]; i < _range[2] + 1; i++)
            {
                for (int j = _range[1]; j < _range[3] + 1; j++)
                {
                    var range = Worksheet.Range[Worksheet.Cells[j, i], Worksheet.Cells[j, i]];
                    range.Borders.Weight = weight;
                }
            }
            return this;
        }

        public vWorksheet SetBorderBottom(double Weight, Color Color)
        {
            dynamic range = CellRangeAnySelector()[1];
            range.Borders[XlBordersIndex.xlEdgeBottom].Weight = Weight;
            range.Borders[XlBordersIndex.xlEdgeBottom].Color = Color;
            return this;
        }

        /// <summary>
        /// 1d,2d,3d,4d are acceptable values
        /// </summary>
        /// <param name="thickness"></param>
        /// <returns></returns>
        public vWorksheet SetBorderWeight(double thickness)
        {
            SetBorderWeights(thickness, thickness, thickness, thickness);
            return this;
        }



        //int HexTop, int HexRight, int HexBottom, int HexLeft
        public vWorksheet SetBorderColors(Color Top, Color Right, Color Bottom, Color Left)
        {
            dynamic range = CellRangeAnySelector()[1];
            range.Borders[XlBordersIndex.xlEdgeTop].Color = Top;
            range.Borders[XlBordersIndex.xlEdgeRight].Color = Right;
            range.Borders[XlBordersIndex.xlEdgeBottom].Color = Bottom;
            range.Borders[XlBordersIndex.xlEdgeLeft].Color = Left;
            return this;
        }

        public vWorksheet SetBorderColorsEach(Color Top, Color Right, Color Bottom, Color Left)
        {
            CheckIfSelected();
            for (int i = _range[0]; i < _range[2] + 1; i++)
            {
                for (int j = _range[1]; j < _range[3] + 1; j++)
                {
                    var range = Worksheet.Range[Worksheet.Cells[j, i], Worksheet.Cells[j, i]];
                    range.Borders[XlBordersIndex.xlEdgeTop].Color = Top;
                    range.Borders[XlBordersIndex.xlEdgeRight].Color = Right;
                    range.Borders[XlBordersIndex.xlEdgeBottom].Color = Bottom;
                    range.Borders[XlBordersIndex.xlEdgeLeft].Color = Left;
                }
            }
            return this;
        }

        public vWorksheet SetBorderColorEach(Color Color)
        {
            CheckIfSelected();
            for (int i = _range[0]; i < _range[2] + 1; i++)
            {
                for (int j = _range[1]; j < _range[3] + 1; j++)
                {
                    var range = Worksheet.Range[Worksheet.Cells[j, i], Worksheet.Cells[j, i]];
                    range.Borders.Color = Color;
                }
            }
            return this;
        }

        public vWorksheet SetDefaultBorder()
        {
            CheckIfSelected();
            var range = Worksheet.Range[Worksheet.Cells[_range[0], _range[1]], Worksheet.Cells[_range[2], _range[3]]];
            range.Borders.Color = Color.LightGray;
            range.Borders.Weight = 2d;
            range.Borders.LineStyle = XlLineStyle.xlContinuous;
            return this;
        }

        public vWorksheet SetBlank()
        {
            CheckIfSelected();
            var range = Worksheet.Range[Worksheet.Cells[_range[0], _range[1]], Worksheet.Cells[_range[2], _range[3]]];
            range.Borders.Color = Color.Transparent;
            range.Interior.Color = Color.Transparent;
            range.Borders.LineStyle = XlLineStyle.xlLineStyleNone;
            return this;
        }

        public vWorksheet SetBorderColor(Color Color)
        {
            SetBorderColors(Color, Color, Color, Color);
            return this;
        }
        #endregion

        #region others
        public vWorksheet SetComment(string comment)
        {
            CheckIfSelected();
            for (int i = _range[0]; i < _range[2] + 1; i++)
            {
                for (int j = _range[1]; j < _range[3] + 1; j++)
                {
                    var cell = Worksheet.Cells[j, i];
                    cell.AddComment(comment);
                }
            }
            return this;
        }

        public vWorksheet RemoveComment()
        {
            CheckIfSelected();
            for (int i = _range[0]; i < _range[2] + 1; i++)
            {
                for (int j = _range[1]; j < _range[3] + 1; j++)
                {
                    var cell = Worksheet.Cells[j, i];
                    cell.Comment.Delete();
                }
            }
            return this;
        }

        public vWorksheet SetBackgroundColor(Color color)
        {
            dynamic range = CellRangeAnySelector()[1];
            range.Interior.Color = color;
            return this;
        }
        #endregion


        private dynamic[] CellRangeAnySelector()
        {
            CheckIfSelected();
            dynamic[] CellRangeAny = {null, null, null};
            CellRangeAny[0] = Worksheet.Cells[_range[1], _range[0]];
            CellRangeAny[1] = Worksheet.Range[
                Worksheet.Cells[_range[1], _range[0]], Worksheet.Cells[_range[3], _range[2]]
            ];
            CellRangeAny[2] = CellRangeAny[1];
            if (!_isRange) CellRangeAny[2] = CellRangeAny[0];
            return CellRangeAny;
        }

        private void CheckIfSelected()
        {
            if(_isUnset) throw new Exception($"No cell or range have been selected for worksheet {TabLabel}. Use .SelectCells(#,#,#,#) first.");
        }

        internal Worksheet GetWorksheet()
        {
            return Worksheet;
        }
    }
}
