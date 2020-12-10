using System;

namespace XlsxCreator.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class XlsxTable : Attribute
    {
        public const string COLUMN_NAME_PROPERTY = nameof(ColumnName);
        public const string POSITION_PROPERTY = nameof(Priority);
        private const int DEFAULT_PRIORITY = 99999;

        public string ColumnName { get; }
        public int Priority { get; }

        public XlsxTable(string columnName)
        {
            ColumnName = columnName;
            Priority = DEFAULT_PRIORITY;
        }

        public XlsxTable(string columnName, int priority)
        {
            ColumnName = columnName;
            Priority = priority;
        }
    }
}
