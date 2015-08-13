using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelPOC
{
    [SuppressMessage("ReSharper", "PossiblyMistakenUseOfParamsMethod")]
    public class Sheet<TData> : ISheet
    {
        private readonly List<Mapping> mappings;

        private IEnumerable<TData> datasource;

        internal Sheet(string sheetname)
        {
            SheetName = sheetname;
            mappings = new List<Mapping>();
        }

        public string SheetName { get; }

        public Sheet<TData> Map(string columnName, Func<TData, dynamic> propertyExtractor)
        {
            return Map(columnName, null, propertyExtractor, null);
        }

        public Sheet<TData> Map(string columnName, string columnTitle, Func<TData, dynamic> propertyExtractor)
        {
            return Map(columnName, columnTitle, propertyExtractor, null);
        }

        public Sheet<TData> Map(string columnName, Func<TData, dynamic> propertyExtractor, string format)
        {
            return Map(columnName, null, propertyExtractor, format);
        }

        public Sheet<TData> Map(string columnName, string columnTitle, Func<TData, dynamic> propertyExtractor, string format)
        {
            mappings.Add(new Mapping(columnName, columnTitle, propertyExtractor, format));
            return this;
        }

        public void SetData(IEnumerable<TData> datasource)
        {
            this.datasource = datasource;
        }

        SheetData ISheet.GetSheetData()
        {
            uint rowIndex = 1;
            var sheetData = new SheetData();

            var hasTitle = mappings.Any(m => !string.IsNullOrEmpty(m.ColumnTitle));
            if (hasTitle)
            {
                sheetData.Append(GetHeader(rowIndex));
                rowIndex++;
            }

            foreach (var data in datasource)
            {
                sheetData.Append(GetRow(rowIndex, data));
                rowIndex++;
            }

            return sheetData;
        }

        private Row GetHeader(uint rowIndex)
        {
            var header = new Row()
            {
                RowIndex = rowIndex
            };

            foreach (var mapping in mappings)
            {
                header.Append(new Cell()
                {
                    CellReference = mapping.ColumnName + rowIndex,
                    CellValue = new CellValue(mapping.ColumnTitle),
                    DataType = CellValues.String
                });
            }
            return header;
        }

        private Row GetRow(uint rowIndex, TData data)
        {
            var row = new Row()
            {
                RowIndex = rowIndex
            };

            foreach (var mapping in mappings)
            {
                var dataProperty = mapping.DataExtractor(data);
                string value = mapping.Format == null
                    ? dataProperty.ToString()
                    : dataProperty.ToString(mapping.Format);

                row.Append(new Cell()
                {
                    CellReference = mapping.ColumnName + rowIndex,
                    CellValue = new CellValue(value),
                    DataType = CellValues.String //GetCellType(dataProperty)
                });
            }
            return row;
        }

        private static EnumValue<CellValues> GetCellType(dynamic data)
        {
            Type dataType = data.GetType();
            switch (dataType.Name)
            {
                case "DateTime":
                    return CellValues.Date;
                case "String":
                    return CellValues.InlineString;
                case "Boolean":
                    return CellValues.Boolean;
                default:
                    return CellValues.Number;
            }
        }

        private class Mapping
        {
            public Mapping(string columnName, string columnTitle, Func<TData, dynamic> dataExtractor, string format)
            {
                ColumnName = columnName;
                ColumnTitle = columnTitle;
                DataExtractor = dataExtractor;
                Format = format;
            }

            public string ColumnName { get; }
            public string ColumnTitle { get; }
            public Func<TData, dynamic> DataExtractor { get; }
            public string Format { get; }
        }
    }
}