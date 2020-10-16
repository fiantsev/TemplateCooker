using System;

namespace TemplateCooker.Domain.ResourceObjects
{
    public class TextResourceObject : ResourceObject
    {
        public string Object { get; }

        public TextResourceObject(string text)
        {
            if (text == null)
                throw new NullReferenceException();

            Object = text;
        }
    }
}