using System;
using System.IO;

namespace VbaModuleBundler
{
	internal static class Streamer
	{
		internal static void CopyStream(System.IO.Stream inputStream, System.IO.Stream outputStream)
		{
			object @lock = new object();
			if (!inputStream.CanRead)
			{
				throw (new Exception("Can not read from inputstream"));
			}
			if (!outputStream.CanWrite)
			{
				throw (new Exception("Can not write to outputstream"));
			}
			if (inputStream.CanSeek)
			{
				inputStream.Seek(0, SeekOrigin.Begin);
			}

			const int bufferLength = 8096;
			var buffer = new Byte[bufferLength];
			lock (@lock)
			{
				int bytesRead = inputStream.Read(buffer, 0, bufferLength);
				// write the required bytes
				while (bytesRead > 0)
				{
					outputStream.Write(buffer, 0, bytesRead);
					bytesRead = inputStream.Read(buffer, 0, bufferLength);
				}
				outputStream.Flush();
			}
		}
	}
}
