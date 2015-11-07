using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Icen.Utils.Excel.Enumerations;

namespace Icen.Utils.Excel.Columns
{
    public class ExtrinsicColumn : Column
    {
        #region Properties
        /// <summary>
        /// 固有列，false
        /// </summary>
        public override bool Intrinsic { get { return false; } }
        /// <summary>
        /// 外部值
        /// </summary>
        private object Value { get; set; }

        /// <summary>
        /// 外部格式化方法
        /// </summary>
        public Func<object> Format { get; set; }

        /// <summary>
        /// 使用必需相关字段的外部格式化方法
        /// 参数为用来格式化的关联字段的值
        /// </summary>
        public Func<object, object> RelateFormat { get; set; }

        /// <summary>
        /// 使用必需相关字段的外部格式化方法
        /// 参数为用来格式化的关联字段的值
        /// </summary>
        public Func<object[], object> RelatesFormat { get; set; }

        #endregion

        #region Methods

        internal override object Generate()
        {
            if (Value != null) return Value;
            return Format();
        }

        internal override object Generate(object relate)
        {
            return RelateFormat(relate);
        }

        internal override object Generate(object[] relates)
        {
            return RelatesFormat(relates);
        }

        internal override object Generate(object origin, object relate)
        {
            throw new NotSupportedException("ExtrinsicColumn do not support Generator with origin value!");
        }

        internal override object Generate(object origin, object[] relates)
        {
            throw new NotSupportedException("ExtrinsicColumn do not support Generator with origin value!");
        }

        #endregion

        #region Constructors

        /// <summary>
        /// 基本列
        /// </summary>
        /// <param name="field"></param>
        private ExtrinsicColumn(string field)
        {
            if (field == null || string.IsNullOrEmpty(field))
                throw new ArgumentNullException("field");
            this.Field = field;
            this.Value = null;
            this.Caption = string.Empty;
            this.ResultType = ExcelResultTypes.String;
            this.RelateField = string.Empty;
            this.PreRead = false;
            this.Required = false;
        }

        /// <summary>
        /// 恒值列
        /// 导入时使用，将DataTable指定列写入恒定值
        /// </summary>
        /// <param name="field"></param>
        /// <param name="value"></param>
        internal ExtrinsicColumn(string field, object value)
            : this(field)
        {
            this.Value = value;
            this.FormatType = ColumnFormatTypes.None;
        }

        /// <summary>
        /// 基本格式化列
        /// 导入时使用给定格式化方法将DataTable列写入
        /// </summary>
        /// <param name="field"></param>
        /// <param name="format"></param>
        internal ExtrinsicColumn(string field, Func<object> format)
            : this(field)
        {
            if (format == null) throw new ArgumentNullException("format");
            this.Format = format;
            this.FormatType = ColumnFormatTypes.Normal;
        }
        /// <summary>
        /// 关联格式化列
        /// 导入使用
        /// </summary>
        /// <param name="field"></param>
        /// <param name="relateField"></param>
        /// <param name="format"></param>
        internal ExtrinsicColumn(string field, string relateField, Func<object, object> format)
            : this(field)
        {
            if (string.IsNullOrEmpty(relateField)) throw new ArgumentNullException("relateField");
            if (format == null) throw new ArgumentNullException("format");
            this.RelateField = relateField;
            this.RelateFormat = format;
            this.FormatType = ColumnFormatTypes.Relate;
        }

        /// <summary>
        /// 关联（多个）格式化列
        /// 导入
        /// </summary>
        /// <param name="field"></param>
        /// <param name="relateFields"></param>
        /// <param name="format"></param>
        internal ExtrinsicColumn(string field, string[] relateFields, Func<object[], object> format)
            : this(field)
        {
            if (relateFields == null || relateFields.Length == 0) throw new ArgumentNullException("relateFields");
            if (format == null) throw new ArgumentNullException("format");
            this.RelateFields = relateFields;
            this.RelatesFormat = format;
            this.FormatType = ColumnFormatTypes.MultiRelate;
        }

        /// <summary>
        /// 基本列
        /// 导出
        /// </summary>
        /// <param name="field"></param>
        /// <param name="caption"></param>
        private ExtrinsicColumn(string field, string caption)
            : this(field)
        {
            if (string.IsNullOrEmpty(caption)) throw new ArgumentNullException("caption");
            this.Caption = caption;
            this.FormatType = ColumnFormatTypes.None;
        }
        /// <summary>
        /// 指定类型恒值列
        /// 导出时使用，将恒定值写入指定类型列
        /// </summary>
        /// <param name="field"></param>
        /// <param name="value"></param>
        /// <param name="type"></param>
        /// <param name="caption"></param>
        internal ExtrinsicColumn(string field, object value, string caption, ExcelResultTypes type)
            : this(field, caption)
        {
            this.Value = value;
            this.ResultType = type;
        }

        /// <summary>
        /// 基本格式化列
        /// 导出时使用给定格式化方法生成值写入Excel指定列
        /// </summary>
        /// <param name="field"></param>
        /// <param name="format"></param>
        /// <param name="type"></param>
        /// <param name="caption"></param>
        internal ExtrinsicColumn(string field, Func<object> format, string caption, ExcelResultTypes type)
            : this(field, caption)
        {
            if (format == null) throw new ArgumentNullException("format");
            this.Format = format;
            this.ResultType = type;
            this.FormatType = ColumnFormatTypes.Normal;
        }



        /// <summary>
        /// 关联格式化列
        /// 导出使用
        /// </summary>
        /// <param name="field"></param>
        /// <param name="relateField"></param>
        /// <param name="format"></param>
        /// <param name="type"></param>
        /// <param name="caption"></param>
        internal ExtrinsicColumn(string field, string relateField, Func<object, object> format,
            string caption, ExcelResultTypes type)
            : this(field, relateField, format)
        {
            if (string.IsNullOrEmpty(caption)) throw new ArgumentNullException("caption");
            this.Caption = caption;
            this.ResultType = type;
        }


        /// <summary>
        /// 关联（多个）格式化
        /// 导出
        /// </summary>
        /// <param name="field"></param>
        /// <param name="relateFields"></param>
        /// <param name="format"></param>
        /// <param name="type"></param>
        /// <param name="caption"></param>
        internal ExtrinsicColumn(string field, string[] relateFields, Func<object[], object> format,
            string caption, ExcelResultTypes type)
            : this(field, relateFields, format)
        {
            if (string.IsNullOrEmpty(caption)) throw new ArgumentNullException("caption");
            this.Caption = caption;
            this.ResultType = type;
        }

        #endregion
    }
}
