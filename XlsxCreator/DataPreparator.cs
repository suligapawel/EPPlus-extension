using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using XlsxCreator.Attributes;

namespace XlsxCreator
{
    public class DataPreparator<T> : IDataPreparator
    {
        private readonly IReadOnlyCollection<T> _collection;

        public DataPreparator(IEnumerable<T> collection)
        {
            if (collection?.Any() != true)
                throw new ArgumentNullException(nameof(collection));

            _collection = collection
                .ToList()
                .AsReadOnly();
        }

        public IEnumerable<IEnumerable<object>> Prepare() => PrepareContent();
        public IEnumerable<IEnumerable<object>> PrepareWithHeaders()
        {
            var headers = PrepareHeaders();
            var content = PrepareContent();

            content.Insert(0, headers);
            return content;
        }

        private List<object> PrepareHeaders() => GetColumnNamesFrom(_collection.First());
        private List<List<object>> PrepareContent()
        {
            var result = new List<List<object>>();

            foreach (var item in _collection)
            {
                result.Add(GetPropertyValuesFrom(item));
            }

            return result;
        }

        private List<object> GetColumnNamesFrom(T item)
        {
            return GetAllProperties(item)
                .Where(propertyInfo => GetAttributeInfoFrom(propertyInfo) != null)
                .OrderBy(propertyInfo => GetAttributeInfoFrom(propertyInfo)?.Priority)
                .Select(propertyInfo => (object)GetAttributeInfoFrom(propertyInfo).ColumnName)
                .ToList();
        }

        private List<object> GetPropertyValuesFrom(T item)
        {
            var allItemProperties = GetAllProperties(item);
            var result = new List<PropertyInfo>();

            foreach (var property in allItemProperties)
            {
                var xlsxTableAttribute = GetAttributeInfoFrom(property);
                if (xlsxTableAttribute == null) continue;

                result.Add(property);
            }

            return result
                .OrderBy(propertyInfo => GetAttributeInfoFrom(propertyInfo)?.Priority)
                .Select(propertyInfo => propertyInfo.GetValue(item))
                .ToList();
        }

        private static PropertyInfo[] GetAllProperties(T item)
            => item.GetType().GetProperties();

        private XlsxTable GetAttributeInfoFrom(PropertyInfo property)
            => property.GetCustomAttributes(typeof(XlsxTable), false).FirstOrDefault() as XlsxTable;
    }
}
