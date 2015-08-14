using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelLibrary
{
    internal interface ISheet
    {
        string SheetName { get; }

        SheetData GetSheetData();

        Columns GetColumns();
    }

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

        public int ColumnCount => mappings.Count;

        #region Map Methods

        public Sheet<TData> Map(string columnName, Func<TData, dynamic> propertyExtractor)
        {
            return Map(columnName, string.Empty, propertyExtractor, null);
        }

        public Sheet<TData> Map(string columnName, string columnTitle, Func<TData, dynamic> propertyExtractor)
        {
            return Map(columnName, columnTitle, propertyExtractor, null);
        }

        public Sheet<TData> Map(string columnName, Func<TData, dynamic> propertyExtractor, string format)
        {
            return Map(columnName, string.Empty, propertyExtractor, format);
        }

        public Sheet<TData> Map(string columnName, string columnTitle, Func<TData, dynamic> propertyExtractor,
            string format)
        {
            mappings.Add(new Mapping(columnName, columnTitle, propertyExtractor, format));
            return this;
        }

        #endregion
        
        public void SetData(IEnumerable<TData> datasource)
        {
            this.datasource = datasource;
        }

        public Columns GetColumns()
        {
            var columns = new Columns();
            uint index = 1;
            foreach (var mapping in mappings)
            {
                var maxLength = 200;

                var optimalWidth = Math.Max(datasource.Max(d => GetValue(d, mapping).Length), mapping.ColumnTitle.Length);

                var column = new Column()
                {
                    Min = index,
                    Max = index,
                    Width = Math.Min(maxLength, optimalWidth + 5),
                    CustomWidth = true
                };
                columns.Append(column);
                index++;
            }

            return columns;
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
                    StyleIndex = 2,
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
                row.Append(new Cell()
                {
                    StyleIndex = 1,
                    CellReference = mapping.ColumnName + rowIndex,
                    CellValue = new CellValue(GetValue(data, mapping)),
                    DataType = CellValues.String //GetCellType(data)
            });
            }
            return row;
        }

        private static string GetValue(TData data, Mapping mapping)
        {
            var dataProperty = mapping.DataExtractor(data);
            return mapping.Format == null
                ? dataProperty.ToString(CultureInfo.InvariantCulture)
                : dataProperty.ToString(mapping.Format, CultureInfo.InvariantCulture);
        }

        private static CellValues GetCellType(TData data)
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