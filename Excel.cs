using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Icen.Utils.Excel.Columns;
using Icen.Utils.Excel.Enumerations;
using Icen.Utils.Excel.Exceptions;

/**
 * @author <a href="mailto:moicen1988@gmail.com">Moicen</a>
 */
 
namespace Icen.Utils.Excel
{
    public class Excel
    {

        /// <summary>
        /// 列索引
        /// </summary>
        private struct ColumnIndex
        {
            public string Field { get; set; }

            public int Index { get; set; }
        }

        public static void Dispose()
        {
            Cache = null;
            Skipped = null;
            CellStyles.Clear();
        }

        #region Import
        /// <summary>
        /// 存储预读取字段的缓存数据
        /// </summary>
        private static Hashtable Cache { get; set; }

        private static List<KeyValuePair<string, string>> Skipped { get; set; }

        /// <summary>
        /// 当缓存字段的某行值为空，将键设为默认
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private static string Key4Null(string name)
        {
            if(string.IsNullOrEmpty(name)) throw new ArgumentNullException("name");
            return "NullKey_" + name;
        }

        /// <summary>
        /// 从文件流读取数据到DataTable
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <param name="table">表映射</param>
        /// <returns></returns>
        public static DataTable Read(Stream stream, Table table)
        {
            var dt = new DataTable(table.Name);
            var workbook = WorkbookFactory.Create(stream);
            var sheet = workbook.GetSheetAt(workbook.ActiveSheetIndex);
            var header = sheet.GetRow(sheet.FirstRowNum);
            var cis = Prepare(dt, table.Columns, header);
            if (sheet.FirstRowNum == sheet.LastRowNum) return dt;
            //初始化缓存
            Cache = new Hashtable();
            for (var i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
            {
                var dr = dt.NewRow();
                var status = ReadRow(sheet.GetRow(i), dr, table.Columns, cis);
                if (status == RowStates.Valid)
                    dt.Rows.Add(dr);
            }
            if (table.Validation != null) table.Validation(dt);
            return dt;
        }
        /// <summary>
        /// 从文件流读取数据到DataTable，并返回被跳过的行数据
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="table"></param>
        /// <param name="skipped"></param>
        /// <returns></returns>
        public static DataTable Read(Stream stream, Table table, out List<KeyValuePair<string, string>> skipped)
        {
            Skipped = new List<KeyValuePair<string, string>>();
            var dt = Read(stream, table);
            skipped = Skipped;
            return dt;
        }

        /// <summary>
        /// 读取预备设置
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="columns"></param>
        /// <param name="hearder"></param>
        /// <returns></returns>
        private static List<ColumnIndex> Prepare(DataTable dt, IEnumerable<Column> columns, IRow hearder)
        {
            var columnIndices = new List<ColumnIndex>();
            foreach (var column in columns)
            {
                dt.Columns.Add(column.Field);
                if (column.Intrinsic)
                {
                    var cell = hearder.Cells.FirstOrDefault(c => c.StringCellValue.Trim() == column.Caption.Trim());
                    if (cell == null) throw new ColumnNotExsitInExcelException(column.Caption);
                    columnIndices.Add(new ColumnIndex() { Field = column.Field, Index = cell.ColumnIndex });
                }
                else
                    columnIndices.Add(new ColumnIndex() { Field = column.Field, Index = -1 });
            }
            return columnIndices;
        }

        /// <summary>
        /// 读取行数据
        /// </summary>
        /// <param name="row"></param>
        /// <param name="dr"></param>
        /// <param name="columns"></param>
        /// <param name="columnIndices"></param>
        /// <returns></returns>
        private static RowStates ReadRow(IRow row, DataRow dr, IEnumerable<Column> columns, IEnumerable<ColumnIndex> columnIndices)
        {
            if(row == null) return RowStates.Empty;
            var status = RowStates.Valid;
            foreach (var columnIndex in columnIndices)
            {
                //因关联格式化而提前读取的项直接跳过
                if (dr[columnIndex.Field] != null && dr[columnIndex.Field] != DBNull.Value) continue;
                var column = columns.First(c => c.Field == columnIndex.Field);
                //格式化
                var value = Format(column, columnIndex, row, columnIndices, columns);
                //校验数据有效性
                status = Validate(columnIndex, column, value, row);
                //数据无效直接跳出
                if (status != RowStates.Valid) break;
                //将格式化后的值赋予DataTable
                dr[columnIndex.Field] = value;
            }
            return status;
        }

        /// <summary>
        /// 格式化
        /// </summary>
        /// <param name="column"></param>
        /// <param name="columnIndex"></param>
        /// <param name="row"></param>
        /// <param name="columnIndices"></param>
        /// <param name="columns"></param>
        /// <returns></returns>
        private static object Format(Column column, ColumnIndex columnIndex, IRow row, IEnumerable<ColumnIndex> columnIndices,
            IEnumerable<Column> columns)
        {
            //额外列
            if (!column.Intrinsic) return Format(column, row, columns, columnIndices);
            //读取原始单元格值
            var cellValue = GetCellValue(row.GetCell(columnIndex.Index), column.ResultType);
            //如果定义为预读取， 则优先读取缓存
            if (column.PreRead) return ReadCache(column, row, columnIndices, columns, cellValue);
            //返回格式化后的值
            return Format(column, row, columnIndices, columns, cellValue);
        }
        /// <summary>
        /// 数据有效性校验
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <param name="column"></param>
        /// <param name="value"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private static RowStates Validate(ColumnIndex columnIndex, Column column, object value, IRow row)
        {
            if ((value == null || string.IsNullOrEmpty(value.ToString())) && column.Required)
            //如果格式化的值为null且该列为必需，则本行数据为空或无效，跳过。
            {
                //非映射列直接返回空行定义
                if (columnIndex.Index < 0) return RowStates.Empty;
                object cellValue = GetCellValue(row.GetCell(columnIndex.Index));
                if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()))
                    //如果单元格的值为空，则定义为空行
                    return RowStates.Empty;
                //如果单元格的值不为空，则定义为无效行，添加到跳过的数据集中
                if (Skipped != null)
                    Skipped.Add(new KeyValuePair<string, string>(column.Caption, cellValue.ToString()));
                return RowStates.Invalid;
            }
            return RowStates.Valid;
        }
        /// <summary>
        /// 读取缓存
        /// </summary>
        /// <param name="column"></param>
        /// <param name="row"></param>
        /// <param name="columnIndices"></param>
        /// <param name="columns"></param>
        /// <param name="cellValue"></param>
        /// <returns></returns>
        private static object ReadCache(Column column, IRow row, IEnumerable<ColumnIndex> columnIndices,
            IEnumerable<Column> columns, object cellValue)
        {
            if (!Cache.ContainsKey(column.Field)) Cache.Add(column.Field, new Dictionary<object, object>());
            var dict = Cache[column.Field] as Dictionary<object, object> ??
                       new Dictionary<object, object>();
            var key = cellValue ?? Key4Null(column.Field);
            if (!dict.ContainsKey(key)) dict.Add(key, Format(column, row, columnIndices, columns, cellValue));
            return dict[key];
        }

        /// <summary>
        /// 格式化固有列
        /// </summary>
        /// <param name="column"></param>
        /// <param name="row"></param>
        /// <param name="columnIndices"></param>
        /// <param name="columns"></param>
        /// <param name="cellValue"></param>
        /// <returns></returns>
        private static object Format(Column column, IRow row, IEnumerable<ColumnIndex> columnIndices,
            IEnumerable<Column> columns, object cellValue)
        {
            switch (column.FormatType)
            {

                case ColumnFormatTypes.Normal:
                    return column.Generate(cellValue);
                case ColumnFormatTypes.Relate:
                    var relateColumnIndex = columnIndices.First(ci => ci.Field == column.RelateField);
                    var reqVal = Format(columns.First(cm => cm.Field == column.RelateField), relateColumnIndex, row,
                        columnIndices, columns);
                    return column.Generate(cellValue, reqVal);
                case ColumnFormatTypes.MultiRelate:
                    var reqVals = new object[column.RelateFields.Length];
                    for (var i = 0; i < column.RelateFields.Length; i++)
                    {
                        var relateField = column.RelateFields[i];
                        var relateIColumnIndex = columnIndices.First(ci => ci.Field == relateField);
                        reqVals[i] = Format(columns.First(cm => cm.Field == relateField), relateIColumnIndex, row,
                            columnIndices, columns);
                    }
                    return column.Generate(cellValue, reqVals);
                case ColumnFormatTypes.None:
                default:
                    return cellValue;
            }
        }
        /// <summary>
        /// 格式化额外列
        /// </summary>
        /// <param name="column"></param>
        /// <param name="row"></param>
        /// <param name="columns"></param>
        /// <param name="columnIndices"></param>
        /// <returns></returns>
        private static object Format(Column column, IRow row, IEnumerable<Column> columns, IEnumerable<ColumnIndex> columnIndices)
        {
            switch (column.FormatType)
            {

                case ColumnFormatTypes.Relate:
                    var relateColumnIndex = columnIndices.First(ci => ci.Field == column.RelateField);
                    var relateValue = Format(columns.First(cm => cm.Field == column.RelateField), relateColumnIndex, row,
                        columnIndices, columns);
                    return column.Generate(relateValue);
                case ColumnFormatTypes.MultiRelate:
                    var relateValues = new object[column.RelateFields.Length];
                    for (var i = 0; i < column.RelateFields.Length; i++)
                    {
                        var relateField = column.RelateFields[i];
                        var relateIColumnIndex = columnIndices.First(ci => ci.Field == relateField);
                        relateValues[i] = Format(columns.First(cm => cm.Field == relateField), relateIColumnIndex, row,
                            columnIndices, columns);
                    }
                    return column.Generate(relateValues);
                case ColumnFormatTypes.Normal:
                case ColumnFormatTypes.None:
                    return column.Generate();
                default:
                    return null;
            }
        }
        /// <summary>
        /// 根据单元格类型读取单元格的值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="cellType"></param>
        /// <returns></returns>
        private static object GetCellValue(ICell cell, CellType cellType = CellType.Unknown)
        {
            if (cell == null) return null;
            if (cellType == CellType.Unknown) cellType = cell.CellType;
            cell.SetCellType(cellType);
            switch (cellType)
            {
                case CellType.Blank:
                    return string.Empty;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Numeric:
                    return cell.NumericCellValue;
                case CellType.String:
                    return TrimString(cell.StringCellValue);
                case CellType.Formula:
                    return GetCellValue(cell, cell.CachedFormulaResultType);
                case CellType.Error:
                    return null;
                default:
                    return TrimString(cell.ToString());
            }
        }
        /// <summary>
        /// 根据ExcelResultType读取单元格的值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="columnType"></param>
        /// <returns></returns>
        private static object GetCellValue(ICell cell, ExcelResultTypes columnType)
        {
            if (cell == null || cell.CellType == CellType.Blank) return null;
            switch (columnType)
            {
                case ExcelResultTypes.String:
                    return GetCellValue(cell, CellType.String);
                case ExcelResultTypes.Boolean:
                    return GetCellValue(cell, CellType.Boolean);
                case ExcelResultTypes.Number:
                case ExcelResultTypes.Percent:
                    return GetCellValue(cell, CellType.Numeric);
                case ExcelResultTypes.DateTime:
                    return GetCellValue(cell, CellType.String);
                case ExcelResultTypes.Picture:
                    return null;
            }
            return GetCellValue(cell);
        }

        private static string TrimString(string str)
        {
            return str.TrimStart('\t', '\n', '\r', '\b', '\f', '\v').TrimEnd('\t', '\n', '\r', '\b', '\f', '\v').Trim();
        }

        #endregion

        #region Export

        /// <summary>
        /// 缓存CellStyle
        /// </summary>
        private static readonly Dictionary<ExcelResultTypes, ICellStyle> CellStyles =
            new Dictionary<ExcelResultTypes, ICellStyle>(6);

        private static ICellStyle GetCellStyle(ExcelResultTypes type, IWorkbook wb)
        {
            if (CellStyles.ContainsKey(type)) return CellStyles[type];
            var style = wb.CreateCellStyle();
            switch (type)
            {
                case ExcelResultTypes.DateTime:
                    style.DataFormat = wb.CreateDataFormat().GetFormat("yyyy-MM-dd");
                    break;
                case ExcelResultTypes.Percent:
                    style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00%");
                    break;
            }
            return CellStyles[type] = style;
        }


        #region From Generic
        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <typeparam name="TEntity">泛型类型</typeparam>
        /// <param name="list">泛型数据集</param>
        /// <param name="advance">是否生成高级版本（07及以上）Excel， 默认false，生成2003版</param>
        /// <returns></returns>
        public static MemoryStream Export<TEntity>(IEnumerable<TEntity> list, bool advance = false)
        {
            IWorkbook workbook = new HSSFWorkbook();
            if (advance) workbook = new XSSFWorkbook();
            Read(workbook, list);
            var stream = new MemoryStream();
            workbook.Write(stream);
            return stream;
        }
        /// <summary>
        /// 生成Excel workbook
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="workbook"></param>
        /// <param name="list"></param>
        private static void Read<TEntity>(IWorkbook workbook, IEnumerable<TEntity> list)
        {
            var sheet = workbook.CreateSheet();
            //构造标题行
            var caption = sheet.CreateRow(0);
            var props = typeof(TEntity).GetProperties();
            var columnIndices = Prepare(caption, props);
            int index = 1;
            foreach (var item in list)
            {
                ReadRow(item, sheet.CreateRow(index), columnIndices, props);
                index++;
            }
        }
        /// <summary>
        /// 读取预定义配置
        /// </summary>
        /// <param name="caption"></param>
        /// <param name="propertyInfos"></param>
        /// <returns></returns>
        private static List<ColumnIndex> Prepare(IRow caption, IEnumerable<PropertyInfo> propertyInfos)
        {
            var columnIndices = new List<ColumnIndex>();
            int index = 0;
            foreach (var prop in propertyInfos)
            {
                var nameAttrs = prop.GetCustomAttributes(typeof(DisplayNameAttribute), false);
                if (nameAttrs.Length == 0) continue;
                var nameAttr = nameAttrs[0] as DisplayNameAttribute;
                if(nameAttr == null) continue;
                caption.CreateCell(index).SetCellValue(nameAttr.DisplayName);
                columnIndices.Add(new ColumnIndex()
                {
                    Field = prop.Name,
                    Index = index
                });
                index++;
            }
            return columnIndices;
        }
        /// <summary>
        /// 读取行数据
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="entity"></param>
        /// <param name="row"></param>
        /// <param name="columnIndices"></param>
        /// <param name="propertyInfos"></param>
        private static void ReadRow<TEntity>(TEntity entity, IRow row, IEnumerable<ColumnIndex> columnIndices, IEnumerable<PropertyInfo> propertyInfos)
        {
            foreach (var prop in propertyInfos)
            {
                var columnIndex = columnIndices.FirstOrDefault(ci => ci.Field == prop.Name);
                if (columnIndex.Equals(default(ColumnIndex))) continue;
                row.CreateCell(columnIndex.Index).SetCellValue(GetPropertyValue(prop, prop.GetValue(entity, null)));
            }
        }
        /// <summary>
        /// 读取属性值
        /// </summary>
        /// <param name="property"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        private static string GetPropertyValue(PropertyInfo property, object value)
        {
            var attrs = property.GetCustomAttributes(typeof(PropertyFormatAttribute), false);
            if (attrs.Length > 0)
            {
                var formatAttr = attrs[0] as PropertyFormatAttribute;
                if (formatAttr != null)
                {
                    value = Format(formatAttr.Type, value);
                }
            }

            return value == null ? string.Empty : value.ToString();
        }
        /// <summary>
        /// 格式化
        /// </summary>
        /// <param name="formatType"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        private static string Format(FormatTypes formatType, object value)
        {
            switch (formatType)
            {
                case FormatTypes.Percent:
                    return ((decimal)value).ToString("P");
                case FormatTypes.Date:
                    return value == null ? string.Empty : (Convert.ToDateTime(value)).ToString("yyyy-MM-dd");
                case FormatTypes.DateTime:
                    return value == null ? string.Empty : (Convert.ToDateTime(value)).ToString("yyyy-MM-dd HH:mm:ss");
                case FormatTypes.None:
                default:
                    return value == null ? string.Empty : value.ToString();
            }
        }

        #endregion

        #region From DataTable
        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <param name="table">表映射</param>
        /// <param name="advance">是否生成高级版本（07及以上）Excel， 默认false，生成2003版</param>
        /// <returns>包含导出数据的Excel文件流</returns>
        public static MemoryStream Export(DataTable dt, Table table, bool advance = false)
        {
            if (dt == null) throw new ArgumentNullException("dt");
            IWorkbook workbook = new HSSFWorkbook();
            if (advance) workbook = new XSSFWorkbook();
            var sheetName = string.IsNullOrEmpty(table.Caption) ? table.Name : table.Caption;
            Read(workbook, sheetName, dt, table.Columns, advance);
            var stream = new MemoryStream();
            workbook.Write(stream);
            return stream;
        }
        /// <summary>
        /// 生成Excel workbook
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetName"></param>
        /// <param name="dt"></param>
        /// <param name="columns"></param>
        /// <param name="advance">是否生成高级版本（07及以上）Excel， 默认false，生成2003版</param>
        /// <returns></returns>
        private static void Read(IWorkbook workbook, string sheetName, DataTable dt, List<Column> columns, bool advance)
        {
            var sheet = workbook.CreateSheet(sheetName);
            var header = sheet.CreateRow(0);
            var columnIndices = Prepare(header, columns);
            sheet.CreateFreezePane(1, 1);
            for (var i = 0; i < dt.Rows.Count; i++)
            {
                ReadRow(dt.Rows[i], sheet.CreateRow(i + 1), columns, columnIndices, dt.Columns, advance);
            }
        }
        /// <summary>
        /// 设置标题行样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        private static ICellStyle RenderHeaderStyle(IWorkbook workbook)
        {
            var style = workbook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.FillForegroundColor = HSSFColor.Lime.Index;
            style.FillPattern = FillPattern.SolidForeground;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderBottom = BorderStyle.Thin;
            var font = workbook.CreateFont();
            font.Boldweight = (short)FontBoldWeight.Bold;
            font.Color = HSSFColor.Green.Index;
            font.FontHeightInPoints = 12;
            font.FontName = "微软雅黑";
            style.SetFont(font);
            return style;
        }
        /// <summary>
        /// 设置各列默认单元格样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        private static ICellStyle RenderColumnStyle(IWorkbook workbook, Column column)
        {
            var style = workbook.CreateCellStyle();
            style.Alignment = (column.ResultType == ExcelResultTypes.Number ||
                               column.ResultType == ExcelResultTypes.Percent)
                ? HorizontalAlignment.Right
                : HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            var font = workbook.CreateFont();
            font.FontHeightInPoints = 11;
            font.FontName = "微软雅黑";
            style.SetFont(font);
            return style;
        }

        /// <summary>
        /// 读取预备设置
        /// </summary>
        /// <param name="header"></param>
        /// <param name="columns"></param>
        /// <returns></returns>
        private static List<ColumnIndex> Prepare(IRow header, IEnumerable<Column> columns)
        {
            var columnIndices = new List<ColumnIndex>();
            int index = 0;
            var headerCellStyle = RenderHeaderStyle(header.Sheet.Workbook);
            foreach (var column in columns)
            {
                var cell = header.CreateCell(index);
                cell.CellStyle = headerCellStyle;
                cell.SetCellValue(column.Caption);
                //设置列宽
                header.Sheet.SetColumnWidth(index, (Math.Max(column.Caption.Length, 5) * 2 + 4) * 256);
                header.Sheet.SetDefaultColumnStyle(index, RenderColumnStyle(header.Sheet.Workbook, column));
                columnIndices.Add(column.Intrinsic
                    ? new ColumnIndex() { Field = column.Field, Index = index }
                    : new ColumnIndex() { Field = column.Field, Index = -1 });
                index++;
            }
            return columnIndices;
        }

        /// <summary>
        /// 读取行数据
        /// </summary>
        /// <param name="dr"></param>
        /// <param name="row"></param>
        /// <param name="columns"></param>
        /// <param name="columnIndices"></param>
        /// <param name="dcs"></param>
        /// <param name="advance">是否生成高级版本（07及以上）Excel， 默认false，生成2003版</param>
        private static void ReadRow(DataRow dr, IRow row, List<Column> columns, List<ColumnIndex> columnIndices, DataColumnCollection dcs, bool advance)
        {
            int index = 0;
            foreach (var columnIndex in columnIndices)
            {
                var column = columns.First(c => c.Field == columnIndex.Field);
                var value = Format(column, columnIndex, dr, columnIndices, columns, dcs);
                SetCellValue(row.CreateCell(index), value, column.ResultType, advance);
                index++;
            }
        }
        /// <summary>
        /// 根据列数据类型设置单元格值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <param name="resultType"></param>
        /// <param name="advance">是否生成高级版本（07及以上）Excel， 默认false，生成2003版</param>
        private static void SetCellValue(ICell cell, object value, ExcelResultTypes resultType, bool advance)
        {
            var wb = cell.Sheet.Workbook;
            switch (resultType)
            {
                case ExcelResultTypes.Boolean:
                    cell.SetCellValue(Convert.ToBoolean(value));
                    return;
                case ExcelResultTypes.Number:
                    cell.SetCellValue(Convert.ToDouble(value));
                    return;
                case ExcelResultTypes.DateTime:
                    if (value == null)
                    {
                        cell.SetCellValue("");
                        return;
                    }
                    var date = Convert.ToDateTime(value);
                    if (!advance)
                    {
                        cell.SetCellValue(date.ToString("yyyy-MM-dd"));
                        return;
                    }
                    cell.SetCellValue(date);
                    cell.CellStyle = GetCellStyle(resultType, wb);
                    return;
                case ExcelResultTypes.Percent:
                    if (value == null)
                    {
                        cell.SetCellValue("N/A");
                        return;
                    }
                    var num = Convert.ToDouble(value);
                    if (!advance)
                    {
                        cell.SetCellValue(num.ToString("P2"));
                        return;
                    }
                    cell.SetCellValue(Convert.ToDouble(value));
                    cell.CellStyle = GetCellStyle(resultType, wb);
                    return;
                case ExcelResultTypes.Picture:
                    throw new NotSupportedException("暂不支持图片导出。");
                case ExcelResultTypes.String:
                default:
                    cell.SetCellValue(Convert.ToString(value));
                    return;
            }
        }

        /// <summary>
        /// 格式化
        /// </summary>
        /// <param name="column"></param>
        /// <param name="columnIndex"></param>
        /// <param name="dr"></param>
        /// <param name="columnIndices"></param>
        /// <param name="columns"></param>
        /// <param name="dcs"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        private static object Format(Column column, ColumnIndex columnIndex, DataRow dr,
            IEnumerable<ColumnIndex> columnIndices, IEnumerable<Column> columns, DataColumnCollection dcs,
            object value = null)
        {
            //非DataTable映射列
            if (!column.Intrinsic) return Format(column, dr, columnIndices, columns, dcs);
            //读取DataTable中的值
            if (value == null) value = dr[columnIndex.Field];
            //将DBNull转换为null
            if (value == DBNull.Value) value = null;
            //格式化
            switch (column.FormatType)
            {
                case ColumnFormatTypes.Relate:
                    var relateColumnIndex = columnIndices.First(ci => ci.Field == column.RelateField);
                    var relateValue = Format(columns.First(cm => cm.Field == column.RelateField), relateColumnIndex, dr,
                        columnIndices, columns, dcs);
                    return column.Generate(value, relateValue);
                case ColumnFormatTypes.MultiRelate:
                    var relateValues = new object[column.RelateFields.Length];
                    for (var i = 0; i < column.RelateFields.Length; i++)
                    {
                        var relateField = column.RelateFields[i];
                        var relateIColumnIndex = columnIndices.First(ci => ci.Field == relateField);
                        relateValues[i] = Format(columns.First(cm => cm.Field == relateField), relateIColumnIndex, dr,
                            columnIndices, columns, dcs);
                    }
                    return column.Generate(value, relateValues);
                case ColumnFormatTypes.Normal:
                    return column.Generate(value);
                case ColumnFormatTypes.None:
                default:
                    return value;
            }
        }
        /// <summary>
        /// 格式化额外列
        /// </summary>
        /// <param name="column"></param>
        /// <param name="dr"></param>
        /// <param name="columnIndices"></param>
        /// <param name="columns"></param>
        /// <param name="dcs"></param>
        /// <returns></returns>
        private static object Format(Column column, DataRow dr, IEnumerable<ColumnIndex> columnIndices,
            IEnumerable<Column> columns, DataColumnCollection dcs)
        {
            switch (column.FormatType)
            {
                case ColumnFormatTypes.Relate:
                    var relateColumnIndex = columnIndices.First(ci => ci.Field == column.RelateField);
                    var relateValue = Format(columns.First(cm => cm.Field == column.RelateField), relateColumnIndex, dr,
                        columnIndices, columns, dcs);
                    return column.Generate(relateValue);
                case ColumnFormatTypes.MultiRelate:
                    var relateValues = new object[column.RelateFields.Length];
                    for (var i = 0; i < column.RelateFields.Length; i++)
                    {
                        var relateField = column.RelateFields[i];
                        var relateIColumnIndex = columnIndices.First(ci => ci.Field == relateField);
                        relateValues[i] = Format(columns.First(cm => cm.Field == relateField), relateIColumnIndex, dr,
                            columnIndices, columns, dcs);
                    }
                    return column.Generate(relateValues);
                case ColumnFormatTypes.Normal:
                case ColumnFormatTypes.None:
                    return column.Generate();
                default:
                    return null;
            }
        }
        #endregion

        #region From DataSet / Multiple tables
        /// <summary>
        /// 从DataSet导出Excel
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="tables"></param>
        /// <param name="advance"></param>
        /// <returns></returns>
        public static MemoryStream Export(DataSet ds, List<Table> tables, bool advance = false)
        {
            if (ds == null) throw new ArgumentNullException("ds");
            IWorkbook workbook = new HSSFWorkbook();
            if (advance) workbook = new XSSFWorkbook();
            tables.ForEach(t =>
            {
                var sheetName = string.IsNullOrEmpty(t.Caption) ? t.Name : t.Caption;
                Read(workbook, sheetName, ds.Tables[t.Name], t.Columns, advance);
            });
            var stream = new MemoryStream();
            workbook.Write(stream);
            return stream;
        }

        /// <summary>
        /// 从多个DataTable列表导出Excel
        /// </summary>
        /// <param name="dts"></param>
        /// <param name="tables"></param>
        /// <param name="advance"></param>
        /// <returns></returns>
        public static MemoryStream Export(IEnumerable<DataTable> dts, List<Table> tables, bool advance = false)
        {
            if (dts == null) throw new ArgumentNullException("dts");
            IWorkbook workbook = new HSSFWorkbook();
            if (advance) workbook = new XSSFWorkbook();
            tables.ForEach(t =>
            {
                var sheetName = string.IsNullOrEmpty(t.Caption) ? t.Name : t.Caption;
                Read(workbook, sheetName, dts.First(dt => dt.TableName == t.Name), t.Columns, advance);
            });
            var stream = new MemoryStream();
            workbook.Write(stream);
            return stream;
        }

        #endregion


        #endregion
    }

}
