using System;

namespace Iren.FrontOffice.Core
{
    public class ApplicationNotFoundException : Exception
    {
        public ApplicationNotFoundException()
        {
        }

        public ApplicationNotFoundException(string message)
            : base(message)
        {
        }

        public ApplicationNotFoundException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
