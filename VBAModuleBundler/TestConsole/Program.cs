using System;
using System.IO;
using VbaModuleBundler;

namespace TestConsole
{
	class Program
	{
		static void Main(string[] args)
		{
			Console.WriteLine("Started");

			if (args.Length != 2)
				throw new ArgumentException("A source workbook must be presented to merge referenced modules/classes from, and a target workbook to merge to.");

			string source = args[0];
			string target = args[1];

			if (source.Equals(target, StringComparison.InvariantCultureIgnoreCase))
				throw new ArgumentException("Source and Target cannot be the same file.");

			var bundler = new Bundler(new Logger());

			bundler.TryGetFileInfo(source, out var sourceInfo);
			bundler.TryGetExcelPackage(sourceInfo, out var package);
			bundler.TryBundleProjects(ref package, true);

			var newFileName = Path.GetFileNameWithoutExtension(source) + DateTime.Now.Ticks;
			var fi = new FileInfo(target);

			package.SaveAs(fi);
			Console.WriteLine($"File \"{fi.FullName}\" saved.");
			Console.ReadLine();
		}
	}

	class Logger : ILogger
	{
		public void Log(string message)
		{
			Console.WriteLine(message);
		}
	}
}
