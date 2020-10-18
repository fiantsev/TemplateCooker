using System;

namespace PluginAbstraction.Exceptions
{
    public class PluginUnsupportedException : Exception
    {
        public PluginUnsupportedException(string message) : base(message)
        {
        }
    }
}