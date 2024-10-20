using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace AwesomeExcel.BridgeOpenXML;

public static class AutoSizeColumns
{
    // https://stackoverflow.com/a/26180406/7327715

    public static Columns AutoSize(SheetData sheetData)
    {
        var maxColWidth = GetMaxCharacterWidth(sheetData);

        Columns columns = new Columns();
        //this is the width of my font - yours may be different
        double maxWidth = 7;
        foreach (var item in maxColWidth)
        {
            //width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
            double width = Math.Truncate((item.Value * maxWidth + 5) / maxWidth * 256) / 256;

            //pixels=Truncate(((256 * {width} + Truncate(128/{Maximum Digit Width}))/256)*{Maximum Digit Width})
            double pixels = Math.Truncate(((256 * width + Math.Truncate(128 / maxWidth)) / 256) * maxWidth);

            //character width=Truncate(({pixels}-5)/{Maximum Digit Width} * 100+0.5)/100
            double charWidth = Math.Truncate((pixels - 5) / maxWidth * 100 + 0.5) / 100;

            Column col = new Column() { BestFit = true, Min = (UInt32)(item.Key + 1), Max = (UInt32)(item.Key + 1), CustomWidth = true, Width = (DoubleValue)width };
            columns.Append(col);
        }

        return columns;
    }

    private static Dictionary<int, int> GetMaxCharacterWidth(SheetData sheetData)
    {
        //iterate over all cells getting a max char value for each column
        Dictionary<int, int> maxColWidth = new Dictionary<int, int>();
        var rows = sheetData.Elements<Row>();
        UInt32[] numberStyles = new UInt32[] { 5, 6, 7, 8 }; //styles that will add extra chars
        UInt32[] boldStyles = new UInt32[] { 1, 2, 3, 4, 6, 7, 8 }; //styles that will bold
        foreach (var r in rows)
        {
            var cells = r.Elements<Cell>().ToArray();

            //using cell index as my column
            for (int i = 0; i < cells.Length; i++)
            {
                var cell = cells[i];
                var cellValue = cell.CellValue == null ? string.Empty : cell.CellValue.InnerText;
                var cellTextLength = cellValue.Length;

                if (cell.StyleIndex != null && numberStyles.Contains(cell.StyleIndex))
                {
                    int thousandCount = (int)Math.Truncate((double)cellTextLength / 4);

                    //add 3 for '.00' 
                    cellTextLength += (3 + thousandCount);
                }

                if (cell.StyleIndex != null && boldStyles.Contains(cell.StyleIndex))
                {
                    //add an extra char for bold - not 100% acurate but good enough for what i need.
                    cellTextLength += 1;
                }

                if (maxColWidth.ContainsKey(i))
                {
                    var current = maxColWidth[i];
                    if (cellTextLength > current)
                    {
                        maxColWidth[i] = cellTextLength;
                    }
                }
                else
                {
                    maxColWidth.Add(i, cellTextLength);
                }
            }
        }

        return maxColWidth;
    }
}
