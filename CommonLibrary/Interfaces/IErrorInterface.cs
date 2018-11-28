using System;

namespace CommonLibrary.Interfaces
{
    public interface IErrorInterface
    {
        event Action<ResultExecute> Errors;
    }
}
