using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Icen.Utils.Excel.Columns;
using Icen.Utils.Excel.Enumerations;

namespace Icen.Utils.Excel
{
    public static class ColumnFactory
    {
        #region ExtrinsicColumn

        #region Import

        /// <summary>
        /// 构建额外列（导入时使用）
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="value">写入值</param>
        /// <returns></returns>
        public static ExtrinsicColumn CreatExtrinsicColumn(string field, object value)
        {
            return new ExtrinsicColumn(field, value);
        }
        /// <summary>
        /// 构建额外列（导入时使用）
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="format">自定义格式化方法</param>
        /// <returns></returns>
        public static ExtrinsicColumn CreatExtrinsicColumn(string field, Func<object> format)
        {
            return new ExtrinsicColumn(field, format);
        }

        /// <summary>
        /// 构建额外列（导入时使用）
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="relateField">相关字段名</param>
        /// <param name="format">自定义格式化方法</param>
        /// <returns></returns>
        public static ExtrinsicColumn CreatExtrinsicColumn(string field, string relateField, Func<object, object> format)
        {
            return new ExtrinsicColumn(field, relateField, format);
        }
        /// <summary>
        /// 构建额外列（导入时使用）
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="relateFields">相关字段名数组</param>
        /// <param name="format">自定义格式化方法</param>
        /// <returns></returns>
        public static ExtrinsicColumn CreatExtrinsicColumn(string field, string[] relateFields, Func<object[], object> format)
        {
            return new ExtrinsicColumn(field, relateFields, format);
        }

        #endregion

        #region Export
        /// <summary>
        /// 构建额外列（导出时使用）
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="value">写入值</param>
        /// <param name="caption">标题</param>
        /// <param name="type">Excel列类型</param>
        /// <returns></returns>
        public static ExtrinsicColumn CreatExtrinsicColumn(string field, object value, string caption, ExcelResultTypes type)
        {
            return new ExtrinsicColumn(field, value, caption, type);
        }
        /// <summary>
        /// 构建额外列（导出时使用）
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="format">自定义格式化方法</param>
        /// <param name="caption">标题</param>
        /// <param name="type">Excel列类型</param>
        /// <returns></returns>
        public static ExtrinsicColumn CreatExtrinsicColumn(string field, Func<object> format, string caption, ExcelResultTypes type)
        {
            return new ExtrinsicColumn(field, format, caption, type);
        }
        /// <summary>
        /// 构建额外列（导出时使用）
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="relateField">相关字段名</param>
        /// <param name="format">自定义格式化方法</param>
        /// <param name="caption">标题</param>
        /// <param name="type">Excel列类型</param>
        /// <returns></returns>
        public static ExtrinsicColumn CreatExtrinsicColumn(string field, string relateField, Func<object, object> format,
            string caption, ExcelResultTypes type)
        {
            return new ExtrinsicColumn(field, relateField, format, caption, type);
        }
        /// <summary>
        /// 构建额外列（导出时使用）
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="relateFields">相关字段名数组</param>
        /// <param name="format">自定义格式化方法</param>
        /// <param name="caption">标题</param>
        /// <param name="type">Excel列类型</param>
        /// <returns></returns>
        public static ExtrinsicColumn CreatExtrinsicColumn(string field, string[] relateFields, Func<object[], object> format,
            string caption, ExcelResultTypes type)
        {
            return new ExtrinsicColumn(field, relateFields, format, caption, type);
        }
        #endregion

        #endregion

        #region IntrinsicColumn

        #region Import

        /// <summary>
        /// 构建固有列（导入时使用）
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="required">是否必需项</param>
        /// <param name="caption">列标题</param>
        /// <param name="type">Excel列类型</param>
        /// <returns></returns>
        public static IntrinsicColumn CreatIntrinsicColumn(string field, bool required, string caption,
            ExcelResultTypes type = ExcelResultTypes.String)
        {
            return new IntrinsicColumn(field, required, caption, type);
        }

        /// <summary>
        /// 构建固有列（导入时使用）
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="format">自定义格式化方法</param>
        /// <param name="required">是否必需项</param>
        /// <param name="caption">列标题</param>
        /// <param name="type">Excel列类型</param>
        /// <returns></returns>
        public static IntrinsicColumn CreatIntrinsicColumn(string field, Func<object, object> format, bool required,
            string caption, ExcelResultTypes type = ExcelResultTypes.String)
        {
            return new IntrinsicColumn(field, format, required, caption, type);
        }

        /// <summary>
        /// 构建固有列（导入时使用）
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="format">自定义格式化方法</param>
        /// <param name="caption">列标题</param>
        /// <param name="preRead">是否预读取</param>
        /// <param name="type">Excel列类型</param>
        /// <returns></returns>
        public static IntrinsicColumn CreatIntrinsicColumn(string field, Func<object, object> format, string caption,
            bool preRead, ExcelResultTypes type = ExcelResultTypes.String)
        {
            return new IntrinsicColumn(field, format, caption, preRead, type);
        }

        /// <summary>
        /// 构建固有列（导入时使用）
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="format">自定义格式化方法</param>
        /// <param name="required">是否必需项</param>
        /// <param name="caption">列标题</param>
        /// <param name="preRead">是否预读取</param>
        /// <param name="type">Excel列类型</param>
        /// <returns></returns>
        public static IntrinsicColumn CreatIntrinsicColumn(string field, Func<object, object> format, bool required,
            string caption, bool preRead, ExcelResultTypes type = ExcelResultTypes.String)
        {
            return new IntrinsicColumn(field, format, required, caption, preRead, type);
        }

        /// <summary>
        /// 构建固有列（导入时使用）
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="relateField">相关字段名</param>
        /// <param name="format">自定义格式化方法</param>
        /// <param name="required">是否必需项</param>
        /// <param name="caption">列标题</param>
        /// <param name="type">Excel列类型</param>
        /// <returns></returns>
        public static IntrinsicColumn CreatIntrinsicColumn(string field, string relateField,
            Func<object, object, object> format, bool required,
            string caption, ExcelResultTypes type = ExcelResultTypes.String)
        {
            return new IntrinsicColumn(field, relateField, format, required, caption, type);
        }

        /// <summary>
        /// 构建固有列（导入时使用）
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="relateField">相关字段名</param>
        /// <param name="format">自定义格式化方法</param>
        /// <param name="caption">列标题</param>
        /// <param name="preRead">是否预读取</param>
        /// <param name="type">Excel列类型</param>
        /// <returns></returns>
        public static IntrinsicColumn CreatIntrinsicColumn(string field, string relateField,
            Func<object, object, object> format, string caption,
            bool preRead, ExcelResultTypes type = ExcelResultTypes.String)
        {
            return new IntrinsicColumn(field, relateField, format, caption, preRead, type);
        }

        /// <summary>
        /// 构建固有列（导入时使用）
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="relateField">相关字段名</param>
        /// <param name="format">自定义格式化方法</param>
        /// <param name="required">是否必需项</param>
        /// <param name="caption">列标题</param>
        /// <param name="preRead">是否预读取</param>
        /// <param name="type">Excel列类型</param>
        /// <returns></returns>
        public static IntrinsicColumn CreatIntrinsicColumn(string field, string relateField,
            Func<object, object, object> format, bool required,
            string caption, bool preRead, ExcelResultTypes type = ExcelResultTypes.String)
        {
            return new IntrinsicColumn(field, relateField, format, required, caption, preRead, type);
        }

        /// <summary>
        /// 构建固有列（导入时使用）
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="relateFields">相关字段名数组</param>
        /// <param name="format">自定义格式化方法</param>
        /// <param name="required">是否必需项</param>
        /// <param name="caption">列标题</param>
        /// <param name="type">Excel列类型</param>
        /// <returns></returns>
        public static IntrinsicColumn CreatIntrinsicColumn(string field, string[] relateFields,
            Func<object, object[], object> format,
            bool required, string caption, ExcelResultTypes type = ExcelResultTypes.String)
        {
            return new IntrinsicColumn(field, relateFields, format, required, caption, type);
        }
        /// <summary>
        /// 构建固有列（导入时使用）
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="relateFields">相关字段名数组</param>
        /// <param name="format">自定义格式化方法</param>
        /// <param name="caption">列标题</param>
        /// <param name="preRead">是否预读取</param>
        /// <param name="type">Excel列类型</param>
        /// <returns></returns>
        public static IntrinsicColumn CreatIntrinsicColumn(string field, string[] relateFields,
            Func<object, object[], object> format,
            string caption, bool preRead, ExcelResultTypes type = ExcelResultTypes.String)
        {
            return new IntrinsicColumn(field, relateFields, format, caption, preRead, type);
        }
        /// <summary>
        /// 构建固有列（导入时使用）
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="relateFields">相关字段名数组</param>
        /// <param name="format">自定义格式化方法</param>
        /// <param name="required">是否必需项</param>
        /// <param name="caption">列标题</param>
        /// <param name="preRead">是否预读取</param>
        /// <param name="type">Excel列类型</param>
        /// <returns></returns>
        public static IntrinsicColumn CreatIntrinsicColumn(string field, string[] relateFields,
            Func<object, object[], object> format,
            bool required, string caption, bool preRead, ExcelResultTypes type = ExcelResultTypes.String)
        {
            return new IntrinsicColumn(field, relateFields, format, required, caption, preRead, type);
        }

        #endregion

        #region Export

        /// <summary>
        /// 构建固有列
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="caption">列标题</param>
        /// <param name="type">Excel列类型</param>
        /// <returns></returns>
        public static IntrinsicColumn CreatIntrinsicColumn(string field, string caption,
            ExcelResultTypes type = ExcelResultTypes.String)
        {
            return new IntrinsicColumn(field, caption, type);
        }

        /// <summary>
        /// 构建固有列
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="format">自定义格式化方法</param>
        /// <param name="caption">列标题</param>
        /// <param name="type">Excel列类型</param>
        /// <returns></returns>
        public static IntrinsicColumn CreatIntrinsicColumn(string field, Func<object, object> format, string caption,
            ExcelResultTypes type = ExcelResultTypes.String)
        {
            return new IntrinsicColumn(field, format, caption, type);
        }

        /// <summary>
        /// 构建固有列
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="relateField">相关字段名</param>
        /// <param name="format">自定义格式化方法</param>
        /// <param name="caption">列标题</param>
        /// <param name="type">Excel列类型</param>
        /// <returns></returns>
        public static IntrinsicColumn CreatIntrinsicColumn(string field, string relateField,
            Func<object, object, object> format, string caption,
            ExcelResultTypes type = ExcelResultTypes.String)
        {
            return new IntrinsicColumn(field, relateField, format, caption, type);
        }

        /// <summary>
        /// 构建固有列
        /// </summary>
        /// <param name="field">字段名</param>
        /// <param name="relateFields">相关字段名数组</param>
        /// <param name="format">自定义格式化方法</param>
        /// <param name="caption">列标题</param>
        /// <param name="type">Excel列类型</param>
        /// <returns></returns>
        public static IntrinsicColumn CreatIntrinsicColumn(string field, string[] relateFields,
            Func<object, object[], object> format,
            string caption, ExcelResultTypes type = ExcelResultTypes.String)
        {
            return new IntrinsicColumn(field, relateFields, format, caption, type);
        }

        #endregion

        #endregion
    }
}
