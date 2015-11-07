using System;

namespace Icen.Utils.Excel.Exceptions
{
    /// <summary>
    /// 列在Excel中不存在时引发的异常
    /// </summary>
    [Serializable]
    public class ColumnNotExsitInExcelException : Exception
    {
        /// <summary>
        /// 用引发异常的列名称，初始化 ColumnNotExsitInExcelException 类型的新实例。
        /// </summary>
        /// <param name="name">引发异常的列名称。</param>
        public ColumnNotExsitInExcelException(string name) : base(string.Format("Excel中不存在名为 \"{0}\" 的列。", name)) { }

        /// <summary>
        /// 用引发异常的列名称，和内部异常，初始化 ColumnNotExsitInExcelException 类型的新实例。
        /// </summary>
        /// <param name="name">引发异常的列名称。</param>
        /// <param name="innerException">内部异常。</param>
        public ColumnNotExsitInExcelException(string name, Exception innerException) : base(string.Format("Excel中不存在名为 \"{0}\" 的列。", name), innerException) { }
    }
}
