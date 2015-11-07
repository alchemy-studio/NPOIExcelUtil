using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Icen.Utils.Excel.Enumerations;

namespace Icen.Utils.Excel.Columns
{
    public class IntrinsicColumn : Column
    {

        #region Properties
        /// <summary>
        /// 是否固有列,true
        /// </summary>
        public override bool Intrinsic { get { return true; } }
        /// <summary>
        /// 格式化方法
        /// 参数为当前原始值
        /// </summary>
        public Func<object, object> Format { get; set; }
        /// <summary>
        /// 使用必需相关字段的格式化方法
        /// 第一个参数为需要格式化的原始值，第二个参数为用来格式化的关联字段的值
        /// </summary>
        public Func<object, object, object> RelateFormat { get; set; }
        /// <summary>
        /// 使用必需相关字段(多个)的格式化方法
        /// 第一个参数为需要格式化的原始值，第二个参数为用来格式化的关联字段的值
        /// </summary>
        public Func<object, object[], object> RelatesFormat { get; set; }

        #endregion

        #region Methods

        internal override object Generate(object origin)
        {
            return Format(origin);
        }

        internal override object Generate(object origin, object relate)
        {
            return RelateFormat(origin, relate);
        }

        internal override object Generate(object origin, object[] relates)
        {
            return RelatesFormat(origin, relates);
        }

        internal override object Generate()
        {
            throw new NotSupportedException("IntrinsicColumn do not support Generator without origin value!");
        }

        internal override object Generate(object[] relates)
        {
            throw new NotSupportedException("IntrinsicColumn do not support Generator without origin value!");
        }

        #endregion

        #region Constructors

        /// <summary>
        /// 基本列
        /// </summary>
        /// <param name="field"></param>
        /// <param name="caption"></param>
        /// <param name="type"></param>
        internal IntrinsicColumn(string field, string caption, ExcelResultTypes type = ExcelResultTypes.String)
        {
            if (field == null || string.IsNullOrEmpty(field))
                throw new ArgumentNullException("field");
            this.Field = field;
            if (string.IsNullOrEmpty(caption))
                throw new ArgumentNullException("caption");
            this.Caption = caption;
            this.ResultType = type;
            this.FormatType = ColumnFormatTypes.None;
            this.RelateField = string.Empty;
            this.PreRead = false;
            this.Required = false;
        }

        /// <summary>
        /// 非空列
        /// </summary>
        /// <param name="field"></param>
        /// <param name="required"></param>
        /// <param name="caption"></param>
        /// <param name="type"></param>
        internal IntrinsicColumn(string field, bool required, string caption,
            ExcelResultTypes type = ExcelResultTypes.String)
            : this(field, caption, type)
        {
            this.Required = required;
        }

        /// <summary>
        /// 基本格式化列
        /// </summary>
        /// <param name="field"></param>
        /// <param name="format"></param>
        /// <param name="caption"></param>
        /// <param name="type"></param>
        internal IntrinsicColumn(string field, Func<object, object> format, string caption,
            ExcelResultTypes type = ExcelResultTypes.String)
            : this(field, caption, type)
        {
            if (format == null) throw new ArgumentNullException("format");
            this.Format = format;
            this.FormatType = ColumnFormatTypes.Normal;
        }

        /// <summary>
        /// 非空格式化列
        /// </summary>
        /// <param name="field"></param>
        /// <param name="format"></param>
        /// <param name="required"></param>
        /// <param name="caption"></param>
        /// <param name="type"></param>
        internal IntrinsicColumn(string field, Func<object, object> format, bool required, string caption,
            ExcelResultTypes type = ExcelResultTypes.String)
            : this(field, format, caption, type)
        {
            this.Required = required;
        }

        /// <summary>
        /// 预读取格式化列
        /// </summary>
        /// <param name="field"></param>
        /// <param name="format"></param>
        /// <param name="caption"></param>
        /// <param name="preRead"></param>
        /// <param name="type"></param>
        internal IntrinsicColumn(string field, Func<object, object> format, string caption, bool preRead,
            ExcelResultTypes type = ExcelResultTypes.String)
            : this(field, format, caption, type)
        {
            this.PreRead = preRead;
        }

        /// <summary>
        /// 预读取格式化非空列
        /// </summary>
        /// <param name="field"></param>
        /// <param name="format"></param>
        /// <param name="required"></param>
        /// <param name="caption"></param>
        /// <param name="preRead"></param>
        /// <param name="type"></param>
        internal IntrinsicColumn(string field, Func<object, object> format, bool required, string caption, bool preRead,
            ExcelResultTypes type = ExcelResultTypes.String)
            : this(field, format, required, caption, type)
        {
            this.PreRead = preRead;
        }

        /// <summary>
        /// 关联格式化列
        /// </summary>
        /// <param name="field"></param>
        /// <param name="relateField"></param>
        /// <param name="format"></param>
        /// <param name="caption"></param>
        /// <param name="type"></param>
        internal IntrinsicColumn(string field, string relateField, Func<object, object, object> format, string caption,
            ExcelResultTypes type = ExcelResultTypes.String)
            : this(field, caption, type)
        {
            if (string.IsNullOrEmpty(relateField)) throw new ArgumentNullException("relateField");
            if (format == null) throw new ArgumentNullException("format");
            this.RelateField = relateField;
            this.RelateFormat = format;
            this.FormatType = ColumnFormatTypes.Relate;
        }

        /// <summary>
        /// 非空关联格式化列
        /// </summary>
        /// <param name="field"></param>
        /// <param name="relateField"></param>
        /// <param name="format"></param>
        /// <param name="required"></param>
        /// <param name="caption"></param>
        /// <param name="type"></param>
        internal IntrinsicColumn(string field, string relateField, Func<object, object, object> format, bool required,
            string caption, ExcelResultTypes type = ExcelResultTypes.String)
            : this(field, relateField, format, caption, type)
        {
            this.Required = required;
        }

        /// <summary>
        /// 预读取关联格式化列
        /// </summary>
        /// <param name="field"></param>
        /// <param name="relateField"></param>
        /// <param name="format"></param>
        /// <param name="caption"></param>
        /// <param name="preRead"></param>
        /// <param name="type"></param>
        internal IntrinsicColumn(string field, string relateField, Func<object, object, object> format, string caption,
            bool preRead, ExcelResultTypes type = ExcelResultTypes.String)
            : this(field, relateField, format, caption, type)
        {
            this.PreRead = preRead;
        }
        /// <summary>
        /// 非空预读取关联格式化列
        /// </summary>
        /// <param name="field"></param>
        /// <param name="relateField"></param>
        /// <param name="format"></param>
        /// <param name="required"></param>
        /// <param name="caption"></param>
        /// <param name="preRead"></param>
        /// <param name="type"></param>
        internal IntrinsicColumn(string field, string relateField, Func<object, object, object> format, bool required,
            string caption, bool preRead, ExcelResultTypes type = ExcelResultTypes.String)
            : this(field, relateField, format, caption, preRead, type)
        {
            this.Required = required;
        }

        /// <summary>
        /// 关联（多个）格式化列
        /// </summary>
        /// <param name="field"></param>
        /// <param name="relateFields"></param>
        /// <param name="format"></param>
        /// <param name="caption"></param>
        /// <param name="type"></param>
        internal IntrinsicColumn(string field, string[] relateFields, Func<object, object[], object> format,
            string caption, ExcelResultTypes type = ExcelResultTypes.String)
            : this(field, caption, type)
        {
            if (relateFields == null || relateFields.Length == 0) throw new ArgumentNullException("relateFields");
            if (format == null) throw new ArgumentNullException("format");
            this.RelateFields = relateFields;
            this.RelatesFormat = format;
            this.FormatType = ColumnFormatTypes.MultiRelate;
        }

        /// <summary>
        /// 关联（多个）非空格式化列
        /// </summary>
        /// <param name="field"></param>
        /// <param name="relateFields"></param>
        /// <param name="format"></param>
        /// <param name="required"></param>
        /// <param name="caption"></param>
        /// <param name="type"></param>
        internal IntrinsicColumn(string field, string[] relateFields, Func<object, object[], object> format,
            bool required, string caption, ExcelResultTypes type = ExcelResultTypes.String)
            : this(field, relateFields, format, caption, type)
        {
            this.Required = required;
        }
        /// <summary>
        /// 关联（多个）预读取格式化列
        /// </summary>
        /// <param name="field"></param>
        /// <param name="relateFields"></param>
        /// <param name="format"></param>
        /// <param name="caption"></param>
        /// <param name="preRead"></param>
        /// <param name="type"></param>
        internal IntrinsicColumn(string field, string[] relateFields, Func<object, object[], object> format,
            string caption, bool preRead, ExcelResultTypes type = ExcelResultTypes.String)
            : this(field, relateFields, format, caption, type)
        {
            this.PreRead = preRead;
        }
        /// <summary>
        /// 关联（多个）非空预读取格式化列
        /// </summary>
        /// <param name="field"></param>
        /// <param name="relateFields"></param>
        /// <param name="format"></param>
        /// <param name="required"></param>
        /// <param name="caption"></param>
        /// <param name="preRead"></param>
        /// <param name="type"></param>
        internal IntrinsicColumn(string field, string[] relateFields, Func<object, object[], object> format,
            bool required, string caption, bool preRead, ExcelResultTypes type = ExcelResultTypes.String)
            : this(field, relateFields, format, required, caption, type)
        {
            this.PreRead = preRead;
        }

        #endregion
    }
}
