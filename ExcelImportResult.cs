using System.Collections.Generic;

namespace Icen.Utils.Excel
{
    /// <summary>
    /// 导入结果
    /// </summary>
    public class ExcelImportResult
    {
        /// <summary>
        /// 导入的记录总数
        /// </summary>
        public int Records { get; set; }
        /// <summary>
        /// 跳过的记录
        /// </summary>
        public List<KeyValuePair<string, string>> Skipped { get; set; }
    }
}
