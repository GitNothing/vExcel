<h2>About</h2>
A fluent API use to quickly build an Excel file. Use it as a .NET class library.

<h3>Prerequisite namespace</h3>

```csharp
using System.Drawing;
```

<h3>Methods for sheet operations.</h3>
<h4>vExcel Class</h4>
<table>
<tr><td>PushNewSheet(String Name)</td><td>Creates and returns the created sheet</td></tr>
<tr><td>PopSheetByName(String TabLabel)</td><td>Removes and returns the last created sheet</td></tr>
<tr><td>GetSheetByName(String TabLabel)</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>RenameSheetByName(String current, String newName)"</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>CopySheetByName(String current, String newName)"</td><td>Copies and returns the copy sheet</td></tr>
<tr><td>SaveOverride(String path)</td><td>Save Excel file</td></tr>
<tr><td>Close()</td><td>Use this or Using keyword to dispose resources</td></tr>
</table>

<h3>Methods for cell or range of cells manipulation.</h3>
<h4>vWorksheet Class</h4>
<table>
<tr><td>SelectCells(Int32 TopX, Int32 TopY, Int32 BottomX, Int32 BottomY)"</td><td>Coordinate of top-left and bottom-right cell</td></tr>
<tr><td>SelectCell(Int32 X, Int32 Y)"</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetValue(String value)</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>ReplaceValue(String current, String newValue)"</td><td>Replace value if entire string is matched</td></tr>
<tr><td>ReplaceValueContaining(String containValue, String newValue)"</td><td>Replace value from matching substring</td></tr>
<tr><td>ClearValue()</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetFontSize(Int32 size)</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetFontColor(Color Color)</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetFontFamily(String family)</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetFontBold(Boolean isBold)</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetFontItalic(Boolean isItalic)</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetFontUnderline(Boolean isUnderline)</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetFontStrikethrough(Boolean isStrikethrough)</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetFontHorizontalCenter()</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetFontHorizontalLeft()</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetFontHorizontalRight()</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetFontVerticalCenter()</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetFontVerticalBottom()</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetFontVerticalTop()</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>AutoSizeColumns()</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>AutoSizeColumnsRelative()</td><td>Autosize entire column based on the current selected cell.</td></tr>
<tr><td>SetColumnWidth(Int32 width)</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetRowHeight(Int32 height)</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>FreezePaneRow(Boolean toggle, Int32 row)</td><td>Size does not need to be specified if toggling false</td></tr>
<tr><td>FreezePaneColumn(Boolean toggle, Int32 column)</td><td>Size does not need to be specified if toggling false</td></tr>
<tr><td>SetBorderWeights(Double Top, Double Right, Double Bottom, Double Left)"</td><td>"Top, right, bottom, left</td></tr>
<tr><td>SetBorderWeightsEach(Double Top, Double Right, Double Bottom, Double Left)"</td><td>Applies per cell</td></tr>
<tr><td>SetBorderWeightEach(Double weight)</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetBorderBottom(Double Weight, Color Color)"</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetBorderWeight(Double thickness)</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetBorderColors(Color Top, Color Right, Color Bottom, Color Left)"</td><td>"Top, right, bottom, left</td></tr>
<tr><td>SetBorderColorsEach(Color Top, Color Right, Color Bottom, Color Left)"</td><td>Applies per cell</td></tr>
<tr><td>SetBorderColorEach(Color Color)</td><td>Applies per cell</td></tr>
<tr><td>SetDefaultBorder()</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetBlank()</td><td>Set background to default color and removes any presence of borders. Font is not affected.</td></tr>
<tr><td>SetBorderColor(Color Color)</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetComment(String comment)</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>RemoveComment()</td><td>SELF-EXPLANATORY</td></tr>
<tr><td>SetBackgroundColor(Color color)</td><td>SELF-EXPLANATORY</td></tr>
</table>

<h3>Sample code usage</h3>

```csharp
static void Main(string[] args)
        {
            var filepath = Directory.GetCurrentDirectory() + "\\test.xlsx";          
            using (var excel = vExcel.vExcel.Factory())
            {
                excel.PushNewSheet("First sheet")
                    .SelectCells(2, 2, 6, 6)
                    .SetValue("helloworld")
                    .SelectCells(2,5,6,6)
                    .SetValue("helloMS")
                    .SelectCells(2, 2, 6, 6)
                    .SetFontSize(16)
                    .SetBorderColors(Color.Blue, Color.Green, Color.Orange, Color.Red)
                    .SetBorderWeights(4d, 3d, 4d, 3d)
                    .ReplaceValue("helloworld", "hellocat")
                    .ReplaceValueContaining("cat", "hellochicken")
                    .SetFontItalic(true)
                    .AutoSizeColumns()

                    .SelectCells(2, 8, 6, 8)
                    .SetValue("foobar")
                    .SetFontSize(16)
                    .SetFontFamily("Arial Black")
                    .SetFontBold(true)
                    .SetFontStrikethrough(true)
                    .SetFontUnderline(true)
                    .SetFontHorizontalRight()
                    .SetBorderColor(Color.Blue)
                    .SetBorderWeights(4d, 4d, 4d, 4d)
                    .SetBorderBottom(4d, Color.BlueViolet)

                    .SelectCells(2, 2, 2, 8)
                    .SetBackgroundColor(Color.Black)
                    .SetFontColor(Color.White);

                //Copies and returns the copy sheet
                excel.CopySheetByName("First sheet", "Copy of first sheet")
                    .SelectCells(2, 2, 6, 6)
                    .SetBorderWeightEach(4d)
                    .SetBorderColorsEach(Color.Blue, Color.Green, Color.Orange, Color.Red)
                    .SelectCells(2, 4, 6, 6)
                    .SetBorderColorEach(Color.CornflowerBlue)
                    .SetBorderBottom(2d, Color.BlueViolet);

                excel.PushNewSheet("Second sheet");

                excel.PushNewSheet("Third sheet");

                //Removes "Second sheet" returns the last created sheet, which is "Third Sheet"
                excel.PopSheetByName("Second sheet")
                    .SelectCells(2, 5, 5, 5)
                    .SetValue("3nd sheet!!!")
                    .SetFontColor(Color.GreenYellow)
                    .SetBorderColor(Color.Blue)
                    .SetBackgroundColor(Color.Magenta)
                    .SetBorderWeight(4d);

                excel.RenameSheetByName("Third sheet", "3rd sheet")
                    .SelectCells(2, 2, 4, 4)
                    .SetValue("3nd")
                    .SetFontBold(true)
                    .AutoSizeColumnsRelative()
                    .SelectCell(1,1)
                    .SetValue("Notice it is autosize relative to the '3rd' cells")
                    .AutoSizeColumnsRelative();

                //Creates a colorful sheet
                var colorfulSheet = excel.PushNewSheet("Colorful!");
                var colors = new List<Color>()
                {
                    Color.Red,
                    Color.Aquamarine,
                    Color.Blue,
                    Color.Brown,
                    Color.Green,
                    Color.Orange,
                    Color.Magenta,
                    Color.DodgerBlue
                };
                var families = new List<String>()
                {
                    "Algerian", "Arial Black", "Arial Rounded MT Bold",
                    "Broadway", "Cooper Black", "Lucida Handwriting", "Magneto",
                    "Viner Hand ITC", "Vladimir Script"

                };
                var ran = new Random();
                var getColor = new Func<Color>(() => colors[ran.Next(0, colors.Count)]);
                var getFamily = new Func<String>(() => families[ran.Next(0, families.Count)]);
                for (int i = 1; i < 12; i++)
                {
                    for (int j = 1; j < 12; j++)
                    {
                        colorfulSheet
                            .SelectCell(i, j)
                            .SetValue("COLORZ!")
                            .SetFontSize(ran.Next(9, 18))
                            .SetFontBold(ran.Next(0,2) == 1)
                            .SetFontItalic(ran.Next(0, 2) == 1)
                            .SetFontStrikethrough(ran.Next(0, 2) == 1)
                            .SetFontUnderline(ran.Next(0, 2) == 1)
                            .SetFontColor(getColor())
                            .SetFontFamily(getFamily())
                            .SetBorderColors(getColor(), getColor(), getColor(), getColor())
                            .SetBorderWeight(4d)
                            .SetBackgroundColor(getColor())
                            .AutoSizeColumns();
                    }
                }
                //hollow out around the center
                colorfulSheet
                    .SelectCells(3, 3, 9, 10)
                    .SetBackgroundColor(Color.Transparent)
                    .ClearValue()
                    .SetDefaultBorder();

                //Saves the excel file, will override if already exist
                excel.SaveOverride(filepath);
            }

            //Opens the xlsx file in Excel
            vExcel.vExcel.OpenInExcel(filepath);
        }
    }
```

<h3>Result from sample Code</h3>
<img src="1.JPG" width="40%">
<img src="2.JPG" width="40%">
<img src="3.JPG" width="40%">
<img src="4.JPG" width="40%">
