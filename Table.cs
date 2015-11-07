using System;
using System.Collections.Generic;
using System.Data;
using Icen.Utils.Excel.Columns;

namespace Icen.Utils.Excel
{
    public class Table
    {
        /// <summary>
        /// (数据库)表名
        /// </summary>
        public string Name { get; private set; }
        /// <summary>
        /// Excel Sheet 名称
        /// </summary>
        public string Caption { get; private set; }
        /// <summary>
        /// 列映射
        /// </summary>
        public List<Column> Columns { get; set; }
        /// <summary>
        /// 数据校验方法
        /// </summary>
        public Action<DataTable> Validation { get; set; }

        /// <summary>
        /// 添加列
        /// </summary>
        /// <param name="column"></param>
        public void AddColumn(Column column)
        {
            if (this.Columns == null) this.Columns = new List<Column>();
            this.Columns.Add(column);
        }
        /// <summary>
        /// 添加多列
        /// </summary>
        /// <param name="columns"></param>
        public void AddColumns(IEnumerable<Column> columns)
        {
            if (this.Columns == null) this.Columns = new List<Column>();
            this.Columns.AddRange(columns);
        }

        public Table(string name)
        {
            Name = name;
        }
        public Table(string name, string caption)
        {
            Name = name;
            Caption = caption;
        }

        public Table()
            : this("DefalutTable")
        {
        }

        public Table(string name, List<Column> columns)
            : this(name)
        {
            if(columns == null) throw new ArgumentNullException("columns");
            Columns = columns;
        }

        public Table(string name, List<Column> columns, string caption)
            : this(name, caption)
        {
            if (columns == null) throw new ArgumentNullException("columns");
            Columns = columns;
        }

        public Table(string name, Action<DataTable> validation)
            : this(name)
        {
            if(validation == null) throw new ArgumentNullException("validation");
            Validation = validation;
        }
        public Table(string name, Action<DataTable> validation, string caption)
            : this(name, caption)
        {
            if (validation == null) throw new ArgumentNullException("validation");
            Validation = validation;
        }

        public Table(string name, Action<DataTable> validation, List<Column> columns)
            : this(name, validation)
        {
            if (columns == null) throw new ArgumentNullException("columns");
            Columns = columns;
        }

        public Table(string name, Action<DataTable> validation, List<Column> columns, string caption)
            : this(name, validation, caption)
        {
            if (columns == null) throw new ArgumentNullException("columns");
            Columns = columns;
        }

    }
}
