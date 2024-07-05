using System;
using System.Data;
using System.Linq;
using System.Collections.Generic;
using ClosedXML.Excel;
using System.Reflection;
using System.Diagnostics;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations.Schema;
using System.Globalization;

namespace ClosedXMLExtensions
{
    public static class IXLWorksheetExtensions
    {
        public static void SetCellValue(this IXLWorksheet worksheet, int rowIndex, int columnIndex, object value)
        {
            worksheet.Cell(rowIndex, columnIndex).SetValue(value);
        }

        public static void SetCellValue(this IXLWorksheet worksheet, int rowIndex, string column, object value) => worksheet.Cell(rowIndex, column).SetValue(value);

        public static void SetCellFormulaA1(this IXLWorksheet worksheet, int rowIndex, int columnIndex, string formulaA1)
        {
            worksheet.Cell(rowIndex, columnIndex).FormulaA1 = formulaA1;
        }

        public static void SetCellFormulaA1(this IXLWorksheet worksheet, int rowIndex, string column, string formulaA1)
            => worksheet.Cell(rowIndex, column).FormulaA1 = formulaA1;

        public static void SetCellNumberFormat(this IXLWorksheet worksheet, int rowIndex, int columnIndex, string format)
            => worksheet.Cell(rowIndex, columnIndex).Style.NumberFormat.Format = format;

        public static void SetCellNumberFormat(this IXLWorksheet worksheet, int rowIndex, string column, string format)
            => worksheet.Cell(rowIndex, column).Style.NumberFormat.Format = format;

        public static void SetCellBackgroundColor(this IXLWorksheet worksheet, int rowIndex, string column, XLColor color)
            => worksheet.Cell(rowIndex, column).Style.Fill.BackgroundColor = color;

        public static void SetCellBackgroundColor(this IXLWorksheet worksheet, int rowIndex, int columnIndex, XLColor color)
            => worksheet.Cell(rowIndex, columnIndex).Style.Fill.BackgroundColor = color;

        /// <summary>
        /// Set number format for columns
        /// </summary>
        /// <param name="worksheet">The worksheet</param>
        /// <param name="columnsExpression">The column expression that indicates which column(s) should be included, e.g. "H:J"</param>
        /// <param name="format">The number format string</param>
        public static void SetColumnsNumberFormat(this IXLWorksheet worksheet, string columnsExpression, string format)
        {
            worksheet.Columns(columnsExpression).Style.NumberFormat.Format = format;
        }

        public static void FreezeRowAndColumn(this IXLWorksheet worksheet, int rows, int columns)
        {
            worksheet.SheetView.Freeze(rows, columns);
        }

        public static void SetValue(this IXLCell cell, object value)
        {
            if (value == null)
                cell.Value = Blank.Value;
            else if (value is string str)
                cell.Value = str;
            else if (value is decimal d)
                cell.Value = d;
            else if (value is int i)
                cell.Value = i;
            else if (value is DateTime dt)
                cell.Value = dt;
            else if (value is bool b)
                cell.Value = b;
            else if (value is double dbl)
                cell.Value = dbl;
            else if (value is float f)
                cell.Value = f;
            else if (value is DateTimeOffset dto)
                cell.Value = dto.DateTime;
            else
                cell.Value = value.ToString();
        }

        /// <summary>
        /// Populate worksheet with data source.
        /// </summary>
        /// <typeparam name="T">The data type of data source</typeparam>
        /// <param name="worksheet">The worksheet that will be populated</param>
        /// <param name="dataSource">The data source</param>
        /// <param name="templateFieldRowIndex">The row index where property names reside, 1-based</param>
        /// <param name="newRowAdding">Invoke before adding a new row in worksheet. Func parameters: the worksheet; the row index (1-based) of the row that is going to add;
        /// the data item; the updated row index (1-based), should be the same row index if no row is added in Func body.
        /// </param>
        /// <param name="newRowAdded">Invoke after a new row is added in worksheet. Func parameters: the worksheet; the row index (1-based) of the current added row;
        /// the data item; the updated row index (1-based), should be the same row index if no row is added in Func body.</param>
        /// <param name="allDataPopulated">Invoke after all data is populated. Func parameters: the worksheet; the last row index (1-based) that data is populated to worksheet;
        /// the data set; the updated last row index (1-based), should be the same last row index if no row is added in Func body</param>
        /// <returns>The last row index that is filled in.</returns>
        public static int Populate<T>(this IXLWorksheet worksheet, IEnumerable<T> dataSource, int templateFieldRowIndex = 2
            , Func<IXLWorksheet, int, T, int> newRowAdding = null, Func<IXLWorksheet, int, T, int> newRowAdded = null
            , Func<IXLWorksheet, int, IEnumerable<T>, int> allDataPopulated = null)
        {
            Type dataSourceType = typeof(T);
            IEnumerable<PropertyInfo> properties = dataSourceType.GetProperties().Where(p => p.CanRead);
            var propertyMapping = new Dictionary<int, PropertyInfo>(properties.Count());
            var plainMapping = new Dictionary<int, object>();

            // read template fields, the format of template field is "{propertyName}", blank template field should be "{_}"
            IXLRow row = worksheet.Row(templateFieldRowIndex);
            int columnIndex = 1;
            string templateFieldName;
            PropertyInfo pi;

            while (true)
            {
                templateFieldName = worksheet.Cell(templateFieldRowIndex, columnIndex).Value.ToString();

                if (string.IsNullOrWhiteSpace(templateFieldName))
                    break;

                if (templateFieldName.StartsWith("{"))
                {
                    templateFieldName = templateFieldName.Trim('{', '}');

                    if (templateFieldName == "_")
                    {
                        plainMapping.Add(columnIndex, string.Empty);
                    }
                    else
                    {
                        pi = properties.SingleOrDefault(p => p.Name.Equals(templateFieldName, StringComparison.OrdinalIgnoreCase));

                        if (pi != null)
                            propertyMapping.Add(columnIndex, pi);
                    }
                }
                else
                {
                    plainMapping.Add(columnIndex, templateFieldName);
                }

                columnIndex++;
            }

            row.Delete();

            // populate worksheet with data
            int rowIndex = templateFieldRowIndex - 1;
            IEnumerable<int> allSheetColumnIndexes = propertyMapping.Keys.Concat(plainMapping.Keys);

            foreach (T data in dataSource)
            {
                rowIndex++;

                if (newRowAdding != null)
                    rowIndex = newRowAdding.Invoke(worksheet, rowIndex, data);

                foreach (int colIndex in allSheetColumnIndexes)
                {
                    if (propertyMapping.ContainsKey(colIndex))
                    {
                        worksheet.SetCellValue(rowIndex, colIndex, propertyMapping[colIndex].GetValue(data));
                        continue;
                    }

                    if (plainMapping.ContainsKey(colIndex))
                    {
                        worksheet.SetCellValue(rowIndex, colIndex, plainMapping[colIndex]);
                        continue;
                    }

                    worksheet.SetCellValue(rowIndex, colIndex, "(can not map)");
                }

                if (newRowAdded != null)
                    rowIndex = newRowAdded.Invoke(worksheet, rowIndex, data);
            }

            if (allDataPopulated != null)
                rowIndex = allDataPopulated.Invoke(worksheet, rowIndex, dataSource);

            return rowIndex;
        }

        /// <summary>
        /// /// Populate worksheet with data view.
        /// </summary>
        /// <param name="worksheet">The worksheet that will be populated</param>
        /// <param name="dataView">The data view</param>
        /// <param name="templateFieldRowIndex">The row index where property names reside, 1-based</param>
        /// <param name="newRowAdding">Invoke before adding a new row in worksheet. Func parameters: the worksheet; the row index (1-based) of the row that is going to add;
        /// the data item; the updated row index (1-based), should be the same row index if no row is added in Func body.
        /// </param>
        /// <param name="newRowAdded">Invoke after a new row is added in worksheet. Func parameters: the worksheet; the row index (1-based) of the current added row;
        /// the data item; the updated row index (1-based), should be the same row index if no row is added in Func body.</param>
        /// <param name="allDataPopulated">Invoke after all data is populated. Func parameters: the worksheet; the last row index (1-based) that data is populated to worksheet;
        /// the data set; the updated last row index (1-based), should be the same last row index if no row is added in Func body</param>
        /// <returns>The last row index that is filled in.</returns>
        public static int Populate(this IXLWorksheet worksheet, DataView dataView, int templateFieldRowIndex = 2
            , Func<IXLWorksheet, int, DataRowView, int> newRowAdding = null, Func<IXLWorksheet, int, DataRowView, int> newRowAdded = null
            , Func<IXLWorksheet, int, DataView, int> allDataPopulated = null)
        {
            var dataSourceColumnMapping = new Dictionary<int, DataColumn>(dataView.Table.Columns.Count);
            var plainMapping = new Dictionary<int, string>();

            // read template fields, the format of template field is "{Column Name}", blank template field should be "{_}"
            IXLRow sheetRow = worksheet.Row(templateFieldRowIndex);
            int sheetColumnIndex = 1;
            string templateFieldName;
            DataColumn dataViewColumn;

            while (true)
            {
                templateFieldName = worksheet.Cell(templateFieldRowIndex, sheetColumnIndex).Value.ToString();

                if (string.IsNullOrWhiteSpace(templateFieldName))
                    break;

                if (templateFieldName.StartsWith("{"))
                {
                    templateFieldName = templateFieldName.Trim('{', '}');

                    if (templateFieldName == "_")
                    {
                        plainMapping.Add(sheetColumnIndex, string.Empty);
                    }
                    else
                    {
                        dataViewColumn = dataView.Table.Columns[templateFieldName];

                        if (dataViewColumn != null)
                            dataSourceColumnMapping.Add(sheetColumnIndex, dataViewColumn);
                    }
                }
                else
                {
                    plainMapping.Add(sheetColumnIndex, templateFieldName);
                }

                sheetColumnIndex++;
            }

            sheetRow.Delete();

            // populate worksheet with data
            int sheetRowIndex = templateFieldRowIndex - 1;
            IEnumerable<int> allSheetColumnIndexes = dataSourceColumnMapping.Keys.Concat(plainMapping.Keys);
            object dataValue;

            foreach (DataRowView data in dataView)
            {
                sheetRowIndex++;

                if (newRowAdding != null)
                    sheetRowIndex = newRowAdding.Invoke(worksheet, sheetRowIndex, data);

                foreach (int sheetColIndex in allSheetColumnIndexes)
                {
                    if (dataSourceColumnMapping.ContainsKey(sheetColIndex))
                    {
                        dataValue = data[dataSourceColumnMapping[sheetColIndex].ColumnName];

                        if (dataValue is DBNull)
                            dataValue = null;

                        worksheet.SetCellValue(sheetRowIndex, sheetColIndex, dataValue);
                        continue;
                    }

                    if (plainMapping.ContainsKey(sheetColIndex))
                    {
                        worksheet.SetCellValue(sheetRowIndex, sheetColIndex, plainMapping[sheetColIndex]);
                        continue;
                    }

                    worksheet.SetCellValue(sheetRowIndex, sheetColIndex, "(can not map)");
                }

                if (newRowAdded != null)
                    sheetRowIndex = newRowAdded.Invoke(worksheet, sheetRowIndex, data);
            }

            if (allDataPopulated != null)
                sheetRowIndex = allDataPopulated.Invoke(worksheet, sheetRowIndex, dataView);

            return sheetRowIndex;
        }

        public static List<T> ConvertTo<T>(this IXLWorksheet worksheet, Func<string, string> mapColumnNameToPropertyName
            , bool throwExceptionIfMissingPropertyMappings = false, Func<string, bool> skipProperty = null
            , Func<string, string, XLCellValue, XLCellValue> modifyCellValue = null
            , bool readCachedValueWhenNotImplementedException = false, Action<Exception> caughtExceptionHandler = null) where T : new()
        {
            IEnumerable<PropertyInfo> pis = typeof(T).GetProperties();
            pis = pis.Where(pi => pi.GetCustomAttribute<NotMappedAttribute>() == null);

            if (skipProperty != null)
                pis = pis.Where(pi => !skipProperty(pi.Name));

            IEnumerable<ColumnNamePropertyName> colNamePropNames = pis.Select(pi =>
            {
                DisplayNameAttribute dna = pi.GetCustomAttribute<DisplayNameAttribute>();
                return new ColumnNamePropertyName(dna == null ? pi.Name : dna.DisplayName, pi.Name);
            });
                
            PropertyColumnMapping[] mappings = pis.Select(pi => new PropertyColumnMapping(pi.Name, null, -1, pi)).ToArray();

            string colName, mappedPropertyName;
            PropertyColumnMapping pcm;

            /* Set max column count to 256 to avoid infinite loop, column index is 1-based */
            for (int colIndex = 1; colIndex <= 256; colIndex++)
            {
                mappedPropertyName = null;
                XLCellValue v = worksheet.Cell(1, colIndex).Value; // row index is 1-based

                if (v.IsBlank)
                    break;

                colName = v.ToString();

                if (string.IsNullOrEmpty(colName))
                    break;

                if (mapColumnNameToPropertyName != null)
                    mappedPropertyName = mapColumnNameToPropertyName(colName);

                if (string.IsNullOrEmpty(mappedPropertyName))
                {
                    mappedPropertyName = colNamePropNames.Where(cp => cp.ColumnName.Equals(colName, StringComparison.OrdinalIgnoreCase))
                        .Select(cp => cp.PropertyName).FirstOrDefault();
                }

                if (!string.IsNullOrEmpty(mappedPropertyName))
                {
                    pcm = mappings.SingleOrDefault(m => m.PropertyName == mappedPropertyName);

                    if (pcm != null)
                    {
                        pcm.ColumnName = colName;
                        pcm.ColumnIndex = colIndex;
                    }
                }
            }

            if (throwExceptionIfMissingPropertyMappings)
            {
                IEnumerable<string> missingProperties = mappings.Where(m => m.ColumnIndex < 0).Select(m => m.PropertyName);

                if (missingProperties.Any())
                    throw new Exception($"The following property names are not mapped in Excel file: {string.Join(", ", missingProperties)}");
            }

            bool rowHasData;
            XLCellValue cellValue;
            var result = new List<T>();
            T currentT;
            Type propertyType;
            int blankRowCount = 0;

            /* set max row count to short.MaxValue to avoid infinite loop, row index is 1-based */
            for (int rowIndex = 2; rowIndex < short.MaxValue; rowIndex++)
            {
                rowHasData = false;

                if (blankRowCount >= 3)
                    break;

                currentT = new T();

                foreach (PropertyColumnMapping propertyMapping in mappings)
                {
                    try
                    {
                        cellValue = worksheet.Cell(rowIndex, propertyMapping.ColumnIndex).Value;
                    }
                    catch(NotImplementedException nex)
                    {
                        nex.Data.Add("RowIndex", rowIndex.ToString());
                        nex.Data.Add("MappingColumnIndex", propertyMapping.ColumnIndex.ToString());
                        nex.Data.Add("PropertyName", propertyMapping.PropertyName);
                        nex.Data.Add("ColumnName", propertyMapping.ColumnName);

                        if (readCachedValueWhenNotImplementedException && nex.Message.Equals("References from other files are not yet implemented."))
                        {
                            caughtExceptionHandler?.Invoke(nex);
                            cellValue = worksheet.Cell(rowIndex, propertyMapping.ColumnIndex).CachedValue;
                        }
                        else
                        {
                            throw;
                        }
                    }

                    if (rowHasData == false && !cellValue.IsBlank && !cellValue.IsError)
                        rowHasData = true;

                    if (modifyCellValue != null)
                        cellValue = modifyCellValue.Invoke(propertyMapping.ColumnName, propertyMapping.PropertyName, cellValue);

                    if (cellValue.IsBlank)
                        continue;

                    if (cellValue.IsError)
                    {
                        XLError err = cellValue.GetError();
                        throw new Exception($"Excel cell value had error: {err}. Property: {propertyMapping.PropertyName}, Column: {propertyMapping.ColumnName}.");
                    }

                    propertyType = propertyMapping.PropertyInfo.PropertyType;
                    
                    if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                        propertyType = Nullable.GetUnderlyingType(propertyType);

                    if (propertyType == typeof(string) && cellValue.IsText)
                    {
                        propertyMapping.PropertyInfo.SetValue(currentT, cellValue.GetText());
                    }
                    else if (propertyType == typeof(string) && cellValue.IsDateTime)
                    {
                        propertyMapping.PropertyInfo.SetValue(currentT, cellValue.GetDateTime().ToString("MM/dd/yyyy"));
                    }
                    else if (propertyType == typeof(string) && cellValue.IsBoolean)
                    {
                        propertyMapping.PropertyInfo.SetValue(currentT, cellValue.GetBoolean() ? "Y" : "N");
                    }
                    else if (propertyType == typeof(string) && cellValue.IsNumber)
                    {
                        propertyMapping.PropertyInfo.SetValue(currentT, cellValue.GetNumber().ToString());
                    }
                    else if (propertyType == typeof(string) && cellValue.IsTimeSpan)
                    {
                        propertyMapping.PropertyInfo.SetValue(currentT, cellValue.GetTimeSpan().ToString());
                    }
                    else if (propertyType == typeof(DateTime) && cellValue.IsDateTime)
                    {
                        propertyMapping.PropertyInfo.SetValue(currentT, cellValue.GetDateTime());
                    }
                    else if (propertyType == typeof(DateTime) && cellValue.IsText)
                    {
                        string str = cellValue.GetText();

                        if (!string.IsNullOrEmpty(str) && DateTime.TryParseExact(str
                            , new[] { "MM/dd/yyyy", "M/d/yyyy", "MM/dd/yy", "M/d/yy", "MM.dd.yyyy", "M.d.yyyy", "MM.dd.yy", "M.d.yy", "yyyy-MM-dd", "yyyy-M-d", "yy-MM-dd", "yyyM-d" }
                            , CultureInfo.CurrentCulture, DateTimeStyles.None, out DateTime dt))
                        {
                            propertyMapping.PropertyInfo.SetValue(currentT, dt);
                        }
                    }
                    else if (propertyType == typeof(bool) && cellValue.IsBoolean)
                    {
                        propertyMapping.PropertyInfo.SetValue(currentT, cellValue.GetBoolean());
                    }
                    else if (propertyType == typeof(TimeSpan) && cellValue.IsTimeSpan)
                    {
                        propertyMapping.PropertyInfo.SetValue(currentT, cellValue.GetTimeSpan());
                    }
                    else if (propertyType == typeof(int) && cellValue.IsText)
                    {
                        propertyMapping.PropertyInfo.SetValue(currentT, int.Parse(cellValue.GetText()));
                    }
                    else if (propertyType == typeof(decimal) && cellValue.IsText)
                    {
                        propertyMapping.PropertyInfo.SetValue(currentT, decimal.Parse(cellValue.GetText()));
                    }
                    else if (propertyType == typeof(float) && cellValue.IsText)
                    {
                        propertyMapping.PropertyInfo.SetValue(currentT, float.Parse(cellValue.GetText()));
                    }
                    else if (propertyType == typeof(double) && cellValue.IsText)
                    {
                        propertyMapping.PropertyInfo.SetValue(currentT, double.Parse(cellValue.GetText()));
                    }
                    else
                    {
                        // assume the cell value is a number.
                        double v = cellValue.GetNumber();
                        try
                        {
                            propertyMapping.PropertyInfo.SetValue(currentT, Convert.ChangeType(v, propertyType));
                        }
                        catch(Exception ex)
                        {
                            var nex = new Exception($"Failed to set column {propertyMapping.ColumnName} value to property {propertyMapping.PropertyName}: {ex.Message}", ex);
                            nex.Data.Add("PropertyName", propertyMapping.PropertyName);
                            nex.Data.Add("ColumnName", propertyMapping.ColumnName);
                            nex.Data.Add("PropertyType", propertyType.FullName);
                            nex.Data.Add("CellValue", v.ToString());
                            throw nex;
                        }
                    }
                }

                if (!rowHasData)
                {
                    blankRowCount++;
                    continue;
                }

                result.Add(currentT);
            }

            return result;
        }

        [DebuggerDisplay("Property Name: {PropertyName}, ColumnIndex: {ColumnIndex}, Type: {PropertyInfo.PropertyType.ToString()}")]
        private class PropertyColumnMapping
        {
            public string PropertyName { get; set; }
            public string ColumnName { get; set; }
            public int ColumnIndex { get; set; }
            public PropertyInfo PropertyInfo { get; set; }

            public PropertyColumnMapping(string propertyName, string columnName, int columnIndex, PropertyInfo propertyInfo)
            {
                PropertyName = propertyName;
                ColumnName = columnName;
                ColumnIndex = columnIndex;
                PropertyInfo = propertyInfo;
            }
        }

        [DebuggerDisplay("ColumnName: {ColumnName}, PropertyName: {PropertyName}")]
        private struct ColumnNamePropertyName
        {
            public string ColumnName { get; set; }
            public string PropertyName { get; set; }

            public ColumnNamePropertyName(string columnName, string propertyName)
            {
                ColumnName = columnName;
                PropertyName = propertyName;
            }
        }
    }
}
