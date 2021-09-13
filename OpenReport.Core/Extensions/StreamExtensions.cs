using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace System.IO
{
    public static class StreamExtensions
    {
        public static Stream Clone(this Stream stream)
        {
            var clonedStream = new MemoryStream();

            stream.CopyTo(clonedStream);

            return clonedStream;
        }

        public static Stream ToStream(this byte[] bytes)
        {
            return new MemoryStream(bytes);
        }
    }
}
