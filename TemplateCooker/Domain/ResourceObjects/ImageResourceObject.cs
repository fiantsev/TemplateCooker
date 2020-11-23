﻿using System;

namespace TemplateCooking.Domain.ResourceObjects
{
    public class ImageResourceObject : ResourceObject
    {
        public byte[] Object { get; }

        public ImageResourceObject(byte[] image)
        {
            if (image == null)
                throw new NullReferenceException();

            Object = image;
        }
    }
}