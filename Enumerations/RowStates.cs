
namespace Icen.Utils.Excel.Enumerations
{
    /// <summary>
    /// 行状态
    /// </summary>
    internal enum RowStates
    {
        /// <summary>
        /// 空行
        /// </summary>
        Empty = 0,
        /// <summary>
        /// 有效行
        /// </summary>
        Valid = 1,
        /// <summary>
        /// 无效行（不可为空的列存在空值）
        /// </summary>
        Invalid = -1
    }
}
