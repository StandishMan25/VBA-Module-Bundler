using System;
using System.IO;
using System.Linq;
using System.Text;
using VbaModuleBundler;

namespace TestConsole
{
	class Program
	{
		static void Main(string[] args)
		{
			var logger = new Logger();
			logger.Log("Started");

			var bundler = new Bundler(logger);

			for (var i = 0; i < args.Length; i++)
			{
				switch (args[i].Trim().ToLower())
				{
					case "/source":
						if (i < args.Length - 1)
							bundler.Source = args[i + 1];
						break;
					case "/target":
						if (i < args.Length - 1)
							bundler.Target = args[i + 1];
						break;
					case "/recurse":
						if (i < args.Length - 1)
						{
							bool.TryParse(args[i + 1], out var recurse);
							bundler.RecurseReferences = recurse;
						}
						break;
					case "/use-source":
						if (i < args.Length - 1)
						{
							bool.TryParse(args[i + 1], out var alwaysUseSource);
							bundler.AlwaysUseSource = alwaysUseSource;
						}
						break;
					case "/only-merge-used":
						if (i < args.Length - 1)
						{
							bool.TryParse(args[i + 1], out var onlyMergeUsed);
							bundler.OnlyMergeUsed = onlyMergeUsed;
						}
						break;
					case "/?":
					case "/h":
					case "/hlp":
					case "/help":
						DisplayHelp(logger);
						Console.ReadLine();
						break;
				}
			}

			if (String.IsNullOrWhiteSpace(bundler.Source) || String.IsNullOrWhiteSpace(bundler.Target))
				return;

			if (bundler.Source.Equals(bundler.Target, StringComparison.InvariantCultureIgnoreCase))
				throw new ArgumentException("Source and Target cannot be the same file.");

			bundler.TryGetFileInfo(bundler.Source, out var sourceInfo);
			bundler.TryGetExcelPackage(sourceInfo, out var package);
			bundler.TryBundleProjects(ref package);

			var newFileName = Path.GetFileNameWithoutExtension(bundler.Source) + DateTime.Now.Ticks;
			var fi = new FileInfo(bundler.Target);

			package.SaveAs(fi);
			Console.WriteLine($"File \"{fi.FullName}\" saved.");
			Console.ReadLine();
		}

		/// <summary>
		/// Super non-professional but informative help message.
		/// </summary>
		/// <param name="logger"></param>
		private static void DisplayHelp(ILogger logger)
		{
			var builder = new StringBuilder();
			builder.AppendLine("Parameter names start with /");
			builder.AppendLine("Parameters are source, target, recurse, use-source, and only-import-used");
			builder.AppendLine("Source: The path to the file you wish to pull all references from and merge into.");
			builder.AppendLine("Target: The path to the resulting file after the merge is complete. File does not need to exist before running (path might).");
			builder.AppendLine("Recurse: If true, will go down the chain of references until none are left, bubbling the merges.");
			builder.AppendLine("Use-Source: If true, will default to using the source modules on any conflict. If false, you will either be prompted or an exception will be thrown.");
			builder.AppendLine("Only-Import-Used: If true, will search through the code and determine which modules are required for functionality. If false, will include everything.");

			builder.AppendLine();
			logger.Log(builder.ToString());
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
