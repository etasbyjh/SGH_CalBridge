using System;
namespace Udilovich.ExcelClient
{
    public interface IInputReader
    {
        string LookupInputValue(string ValueKey);
    }
}
