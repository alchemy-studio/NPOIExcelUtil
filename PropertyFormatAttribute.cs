using System;
using Icen.Utils.Excel.Enumerations;

namespace Icen.Utils.Excel
{
    /// <summary>
    /// 属性值格式化方式
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class PropertyFormatAttribute : Attribute
    {
        public FormatTypes Type { get; private set; }

        public PropertyFormatAttribute(FormatTypes type)
        {
            this.Type = type;
        }
    }
}
