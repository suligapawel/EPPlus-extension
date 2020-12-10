using System.Collections.Generic;

namespace XlsxCreator
{
    public interface IDataPreparator
    {
        IEnumerable<IEnumerable<object>> Prepare();
        IEnumerable<IEnumerable<object>> PrepareWithHeaders();
    }
}