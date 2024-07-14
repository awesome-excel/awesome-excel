using _Excel = AwesomeExcel.Common.Models;
using _NPOI = NPOI.SS.UserModel;

namespace AwesomeExcel.BridgeNpoi;

public class StyleConverter
{
    private readonly _NPOI.IWorkbook npoiWorkbook;

    public StyleConverter(_NPOI.IWorkbook npoiWorkbook)
    {
        this.npoiWorkbook = npoiWorkbook ?? throw new ArgumentNullException(nameof(npoiWorkbook));
    }

    public virtual _NPOI.ICellStyle Convert(_Excel.Style excelStyle)
    {
        if (excelStyle is null)
        {
            throw new ArgumentNullException(nameof(excelStyle));
        }

        return CreateCellStyle(excelStyle);
    }

    public virtual _NPOI.IFont Convert(_Excel.FontStyle excelStyle)
    {
        if (excelStyle is null)
        {
            throw new ArgumentNullException(nameof(excelStyle));
        }

        return CreateFont(excelStyle);
    }

    private _NPOI.ICellStyle CreateCellStyle(_Excel.Style excelStyle)
    {
        _NPOI.ICellStyle npoiStyle = npoiWorkbook.CreateCellStyle();

        if (excelStyle == null)
            return npoiStyle;

        SetBorderTopColor();
        SetBorderBottomColor();
        SetBorderLeftColor();
        SetBorderRightColor();
        SetFillPattern();
        SetFillForegroundColor();
        SetFontStyle();
        SetHorizontalAlignment();
        SetVerticalAlignment();
        SetDateTimeFormat();

        return npoiStyle;

        void SetBorderTopColor()
        {
            if (excelStyle.BorderTopColor.HasValue)
            {
                npoiStyle.TopBorderColor = (short)excelStyle.BorderTopColor.Value;
                npoiStyle.BorderTop = _NPOI.BorderStyle.Thin;
            }
        }

        void SetBorderBottomColor()
        {
            if (excelStyle.BorderBottomColor.HasValue)
            {
                npoiStyle.BottomBorderColor = (short)excelStyle.BorderBottomColor.Value;
                npoiStyle.BorderBottom = _NPOI.BorderStyle.Thin;
            }
        }

        void SetBorderLeftColor()
        {
            if (excelStyle.BorderLeftColor.HasValue)
            {
                npoiStyle.LeftBorderColor = (short)excelStyle.BorderLeftColor.Value;
                npoiStyle.BorderLeft = _NPOI.BorderStyle.Thin;
            }
        }

        void SetBorderRightColor()
        {
            if (excelStyle.BorderRightColor.HasValue)
            {
                npoiStyle.RightBorderColor = (short)excelStyle.BorderRightColor.Value;
                npoiStyle.BorderRight = _NPOI.BorderStyle.Thin;
            }
        }

        void SetFillPattern()
        {
            if (excelStyle.FillPattern.HasValue)
            {
                npoiStyle.FillPattern = (_NPOI.FillPattern)excelStyle.FillPattern.Value;
            }
        }

        void SetFillForegroundColor()
        {
            if (excelStyle.FillForegroundColor.HasValue)
            {
                npoiStyle.FillForegroundColor = (short)excelStyle.FillForegroundColor.Value;

                if (npoiStyle.FillPattern == _NPOI.FillPattern.NoFill)
                    npoiStyle.FillPattern = _NPOI.FillPattern.SolidForeground;
            }
        }

        void SetFontStyle()
        {
            if (excelStyle.FontStyle != null)
            {
                _NPOI.IFont npoiFont = Convert(excelStyle.FontStyle);
                npoiStyle.SetFont(npoiFont);
            }
        }

        void SetHorizontalAlignment()
        {
            if (excelStyle.HorizontalAlignment.HasValue)
                npoiStyle.Alignment = (_NPOI.HorizontalAlignment)excelStyle.HorizontalAlignment.Value;
        }

        void SetVerticalAlignment()
        {
            if (excelStyle.VerticalAlignment.HasValue)
                npoiStyle.VerticalAlignment = (_NPOI.VerticalAlignment)excelStyle.VerticalAlignment.Value;
        }

        void SetDateTimeFormat()
        {
            if (!string.IsNullOrWhiteSpace(excelStyle.DateTimeFormat))
            {
                _NPOI.IDataFormat dataFormatService = npoiWorkbook.CreateDataFormat();
                string excelFormat = excelStyle.DateTimeFormat;
                short npoiFormat = dataFormatService.GetFormat(excelFormat);
                npoiStyle.DataFormat = npoiFormat;
            }
        }
    }

    private _NPOI.IFont CreateFont(_Excel.FontStyle fontStyle)
    {
        _NPOI.IFont npoiFont = npoiWorkbook.CreateFont();

        if (fontStyle == null)
            return npoiFont;

        SetName();
        SetColor();
        SetHeightInPoints();
        SetIsBold();

        return npoiFont;

        void SetName()
        {
            if (!string.IsNullOrWhiteSpace(fontStyle.Name))
                npoiFont.FontName = fontStyle.Name;
        }

        void SetColor()
        {
            if (fontStyle.Color.HasValue)
                npoiFont.Color = (short)fontStyle.Color.Value;
        }

        void SetHeightInPoints()
        {
            if (fontStyle.HeightInPoints.HasValue)
                npoiFont.FontHeightInPoints = fontStyle.HeightInPoints.Value;
        }

        void SetIsBold()
        {
            if (fontStyle.IsBold.HasValue)
                npoiFont.IsBold = fontStyle.IsBold.Value;
        }
    }
}
