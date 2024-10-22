﻿using _NPOI = NPOI.SS.UserModel;

namespace AwesomeExcel.BridgeNPOI;

internal class NpoiHelper
{
    public void SetCellValue(_NPOI.ICell cell, AwesomeExcel.Models.ColumnType columnType, object? value)
    {
        string _value = value?.ToString() ?? string.Empty;

        if (columnType == AwesomeExcel.Models.ColumnType.String)
        {
            cell.SetCellValue(_value);
        }
        else if (columnType == AwesomeExcel.Models.ColumnType.Numeric)
        {
            if (TryParseNumeric(columnType, value, _value, out double number))
            {
                cell.SetCellValue(number);
            }
            else
            {
                cell.SetCellValue("non numeric value");
            }
        }
        else if (columnType == AwesomeExcel.Models.ColumnType.DateTime)
        {
            if (TryParseDateTime(columnType, value, _value, out DateTime dt))
            {
                cell.SetCellValue(dt);
            }
            else
            {
                cell.SetCellValue("non datetime value");
            }
        }
        else
        {
            cell.SetCellValue(_value);
        }
    }

    private bool TryParseNumeric(AwesomeExcel.Models.ColumnType columnType, object value, string valueStr, out double number)
    {
        if (columnType != AwesomeExcel.Models.ColumnType.Numeric)
        {
            number = 0;
            return false;
        }

        switch (value)
        {
            case bool boo:
                number = Convert.ToDouble(boo);
                return true;
            case byte b:
                number = b;
                return true;
            case short s:
                number = s;
                return true;
            case ushort us:
                number = us;
                return true;
            case int i:
                number = i;
                return true;
            case uint ui:
                number = ui;
                return true;
            case long l:
                number = l;
                return true;
            case ulong ul:
                number = ul;
                return true;
            case float f:
                number = f;
                return true;
            case double d:
                number = d;
                return true;
            case decimal dec:
                number = Convert.ToDouble(dec);
                return true;

            default:
                return double.TryParse(valueStr, out number);
        }
    }

    private bool TryParseDateTime(AwesomeExcel.Models.ColumnType columnType, object value, string valueStr, out DateTime dt)
    {
        if (columnType != AwesomeExcel.Models.ColumnType.DateTime)
        {
            dt = default;
            return false;
        }

        if (value is DateTime _datetime)
        {
            dt = _datetime;
            return true;
        }
        else if (value is DateOnly date)
        {
            dt = date.ToDateTime(new TimeOnly(0, 0, 0, 0));
            return true;
        }
        else if (value is TimeOnly time)
        {
            DateTime minDateTime = DateTime.MinValue;
            dt = new DateTime(minDateTime.Year, minDateTime.Month, minDateTime.Day, time.Hour, time.Minute, time.Second);
            return true;
        }
        else
        {
            return DateTime.TryParse(valueStr, out dt);
        }
    }
}