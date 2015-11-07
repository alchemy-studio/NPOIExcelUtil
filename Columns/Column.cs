using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Icen.Utils.Excel.Enumerations;

namespace Icen.Utils.Excel.Columns
{
    public abstract class Column
    {
        /// <summary>
        /// 字段名
        /// </summary>
        public string Field { get; set; }
        /// <summary>
        /// 列标题
        /// </summary>
        public string Caption { get; set; }
        /// <summary>
        /// 是否固有列
        /// </summary>
        public abstract bool Intrinsic { get; }
        /// <summary>
        /// Excel中的列类型
        /// </summary>
        public ExcelResultTypes ResultType { get; set; }
        /// <summary>
        /// 格式化方式
        /// </summary>
        public ColumnFormatTypes FormatType { get; set; }
        /// <summary>
        /// 是否预读取
        /// </summary>
        public bool PreRead { get; set; }
        /// <summary>
        /// 是否必需（不能为空，为空则该行跳过，不进行导入）
        /// 如果Format或者RelateFormat不为null，则认为Required为true。
        /// </summary>
        public bool Required { get; set; }
        /// <summary>
        /// 可能需要的相关字段
        /// </summary>
        public string RelateField { get; set; }
        /// <summary>
        /// 可能需要的相关字段
        /// </summary>
        public string[] RelateFields { get; set; }
        /// <summary>
        /// 生成最终值
        /// </summary>
        /// <returns></returns>
        internal abstract object Generate();
        /// <summary>
        /// 根据输入值生成最终值
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        internal abstract object Generate(object input);
        /// <summary>
        /// 根据原始值和相关值生成最终值
        /// </summary>
        /// <param name="origin"></param>
        /// <param name="relate"></param>
        /// <returns></returns>
        internal abstract object Generate(object origin, object relate);
        /// <summary>
        /// 根据原始值和相关值（多个）生成最终值
        /// </summary>
        /// <param name="origin"></param>
        /// <param name="relates"></param>
        /// <returns></returns>
        internal abstract object Generate(object origin, object[] relates);
        /// <summary>
        /// 根据相关值（多个）生成最终值
        /// </summary>
        /// <param name="relates"></param>
        /// <returns></returns>
        internal abstract object Generate(object[] relates);

    }
}
