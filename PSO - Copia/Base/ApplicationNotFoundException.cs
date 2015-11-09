using System;

namespace Iren.ToolsExcel.Base
{
    public class RepositoryUpdateException : Exception
    {
        public RepositoryUpdateException()
        {
        }

        public RepositoryUpdateException(string message)
            : base(message)
        {
        }

        public RepositoryUpdateException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
