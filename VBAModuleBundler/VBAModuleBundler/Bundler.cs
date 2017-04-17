using OfficeOpenXml;
using OfficeOpenXml.VBA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace VbaModuleBundler
{
	public class Bundler
	{
		ILogger _logger;

		public Bundler(ILogger logger)
		{
			_logger = logger;
		}

		/// <summary>
		/// Attempts to create a <see cref=" FileInfo"/> from the provided <paramref name="source"/> path.
		/// </summary>
		/// <param name="source">Path to the file to use for bundling.</param>
		/// <param name="info">Generated <see cref="FileInfo"/>.</param>
		/// <remarks>The file at <paramref name="source"/> must be closed to function.</remarks>
		/// <returns></returns>
		public bool TryGetFileInfo(string source, out FileInfo info)
		{
			try
			{
				info = new FileInfo(source);
				_logger.Log($"Created fileinfo from \"{source}\"");
				return true;
			}
			catch (Exception ex)
			{
				throw new InvalidOperationException($"Unable to create file info object from source \"{source}\"\nEx:\n{ex.ToString()}");
			}
		}

		/// <summary>
		/// Attempts to create the <see cref="ExcelPackage"/> from the <paramref name="source"/>.
		/// </summary>
		/// <param name="source">Source <see cref="FileInfo"/> to generate an <see cref="ExcelPackage"/> from.</param>
		/// <param name="package">Generated <see cref="ExcelPackage"/>.</param>
		/// <returns></returns>
		public bool TryGetExcelPackage(FileInfo source, out ExcelPackage package)
		{
			try
			{
				package = new ExcelPackage(source);
				_logger.Log($"Created ExcelPackage from \"{source.FullName}\"");
				return true;
			}
			catch (Exception ex)
			{
				throw new InvalidOperationException($"Unable to create ExcelPackage from source \"{source}\"\nEx:\n{ex.ToString()}");
			}
		}

		/// <summary>
		/// Returns two distinct sets of references from the <see cref="ExcelVbaProject"/> <paramref name="project"/>, system references and Excel file references.
		/// </summary>
		/// <param name="project"><see cref="ExcelVbaProject"/> to extract references from.</param>
		/// <returns></returns>
		public (IEnumerable<ExcelVbaReference> System, IEnumerable<ExcelVbaReference> Excel) GetReferences(ExcelVbaProject project)
		{
			var systemReferences = project.References.Where(x => x.ReferenceRecordID != 14).ToList();
			var excelReferences = project.References.Where(x => x.ReferenceRecordID == 14).ToList();
			if (excelReferences.Count > 0)
				_logger.Log($"Found References:\n\t{String.Join("\n\t", excelReferences.Select(x => x.Name))}");
			else
				_logger.Log($"No references found for {project.Name}");
			return (systemReferences, excelReferences);
		}

		/// <summary>
		/// Returns the file path derived from the <see cref="ExcelVbaReference"/> Libid.
		/// </summary>
		/// <param name="reference"><see cref="ExcelVbaReference"/> to extract file path from.</param>
		/// <returns></returns>
		public string GetReferencePath(ExcelVbaReference reference)
		{
			var path = reference.Libid.Substring(3);
			_logger.Log($"Found reference path \"{path}\"");
			return path;
		}

		/// <summary>
		/// Gathers all modules and class modules from the <paramref name="sourceProject"/> and <paramref name="targetProject"/> and merges them together.
		/// </summary>
		/// <param name="sourceProject">Project containing modules to be merged into <paramref name="targetProject"/>.</param>
		/// <param name="targetProject">Project containing existing modules.</param>
		/// <param name="modules">Combined modules.</param>
		/// <remarks>This will replace early binding references with <paramref name="sourceProject"/> module name if found, and will prompt if a module with the same name but different code exists.</remarks>
		/// <returns>True if merge was successful.</returns>
		public bool TryMergeModules(ExcelVbaProject sourceProject, ExcelVbaProject targetProject, out IEnumerable<ExcelVBAModule> modules)
		{
			try
			{
				////	Get all modules, classes, and user forms from each project
				var sourceItems = sourceProject.Modules.Where(x => x.Type == eModuleType.Module || x.Type == eModuleType.Class).ToList();
				var targetItems = targetProject.Modules.Where(x => x.Type == eModuleType.Module || x.Type == eModuleType.Class || x.Type == eModuleType.Designer).ToList();

				var removeFromTarget = new List<ExcelVBAModule>();

				foreach (var targetItem in targetItems)
				{
					if (targetItem.Code.Contains($"{sourceProject.Name}."))
						targetItem.Code = targetItem.Code.Replace($"{sourceProject.Name}.", "");

					if (sourceItems.Any(x => x.Name == targetItem.Name && x.Code == targetItem.Code))
						sourceItems.Remove(sourceItems.Single(x => x.Name == targetItem.Name));

					if (sourceItems.Any(x => x.Name == targetItem.Name && x.Code != targetItem.Code))
					{
						try
						{
							var consolePresent = Console.WindowHeight > 0;
						}
						catch
						{
							throw new ArgumentException($"The source \"{sourceProject.Name}\" and target \"{targetProject.Name}\" have a {targetItem.Type} with the same name \"{targetItem.Name}\" and different code. Please remove from one or the other and run again.");
						}

						ConsoleColor backColor = Console.BackgroundColor, foreColor = Console.ForegroundColor;

						Console.BackgroundColor = ConsoleColor.Yellow;
						Console.ForegroundColor = ConsoleColor.Black;
						Console.WriteLine($"The source \"{sourceProject.Name}\" and target \"{targetProject.Name}\" have a {targetItem.Type} with the same name \"{targetItem.Name}\" and different code. Please advise: 0 to use source, 1 to keep target.");
						if (Console.ReadLine().ToString() == "0")
							removeFromTarget.Add(targetItem);
						else
							sourceItems.Remove(sourceItems.Single(x => x.Name == targetItem.Name));
						Console.BackgroundColor = backColor;
						Console.ForegroundColor = foreColor;
					}
				}

				foreach (var item in removeFromTarget)
					targetItems.Remove(item);

				////	Change the name of the source objects to contain the source's VBAProject name
				//foreach (var item in sourceItems)
				//{
				//	var name = $"{sourceProject.Name}_{item.Name}";
				//	item.Name = name;
				//}

				_logger.Log($"Merging source:\n\t{String.Join("\n\t", sourceItems.Select(x => x.Name))}\nWith Target:\n\t{String.Join("\n\t", targetItems.Select(x => x.Name))}");

				//	Put them together
				modules = targetItems.Concat(sourceItems);

				_logger.Log($"Merge complete:\n\t{String.Join("\n\t", modules.Select(x => x.Name))}");

				return true;
			}
			catch (Exception ex)
			{
				throw;
			}
		}

		/// <summary>
		/// Attempts to add the <paramref name="modules"/> to the <paramref name="targetProject"/>. This overwrites the <paramref name="targetProject"/>'s internal <see cref="ExcelVbaProject.Modules"/> collection, as it doesn't support a "Merge" operation.
		/// </summary>
		/// <param name="targetProject"><see cref="ExcelVbaProject"/> to "merge" <paramref name="modules"/> into.</param>
		/// <param name="modules">Collection of <see cref="ExcelVBAModule"/>s to merge into <paramref name="targetProject"/></param>
		/// <returns></returns>
		public bool TryAddToProject(ref ExcelVbaProject targetProject, IEnumerable<ExcelVBAModule> modules)
		{
			//	Could create instance of _list and actually add to the collection, rather than replacing it entirely.
			//	This would require a rewrite of TryMergeModules as it currently collects the targetProject's modules as well.
			try
			{
				//	Let's get reflecting! Set the internal collection of modules on the target.
				targetProject.Modules.GetType().GetField("_list", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)?.SetValue(targetProject.Modules, modules.ToList());
				_logger.Log($"Added the following to target \"{targetProject.Name}\":\n\t{String.Join("\n\t", modules.Select(x => x.Name))}");
				return true;
			}
			catch (Exception ex)
			{
				throw;
			}
		}

		/// <summary>
		/// Removes the dependency of the referenced file from the <paramref name="targetProject"/>.
		/// </summary>
		/// <param name="targetProject"><see cref="ExcelVbaProject"/> to remove reference from.</param>
		/// <param name="reference"><see cref="ExcelVbaReference"/> to remove.</param>
		/// <returns></returns>
		public bool TryRemoveReference(ref ExcelVbaProject targetProject, ExcelVbaReference reference)
		{
			try
			{
				_logger.Log($"Removing reference \"{reference.Name}\" from \"{targetProject.Name}\"");
				targetProject.References.Remove(reference);
				return true;
			}
			catch (Exception ex)
			{
				throw;
			}
		}
		
		/// <summary>
		/// Takes the <see cref="ExcelPackage"/> and bundles the referenced projects into a single package, recursing references if desired.
		/// </summary>
		/// <param name="package"><see cref="ExcelPackage"/> to gather referenced modules from, and bundle with.</param>
		/// <param name="recurseReferences">If true, will recurse each referenced file's references, bubbling up the merge references.</param>
		/// <returns></returns>
		public bool TryBundleProjects(ref ExcelPackage package, bool recurseReferences)
		{
			var project = package.Workbook.VbaProject;
			return TryBundleProjects(ref package, recurseReferences, ref project);
		}

		/// <summary>
		/// Actually executes the merging operations. 
		/// </summary>
		/// <param name="package"><see cref="ExcelPackage"/> to gather referenced modules from, and bundle with.</param>
		/// <param name="recurseReferences">If true, will recurse each referenced file's references, bubbling up the merge references.</param>
		/// <param name="project"><see cref="ExcelVbaProject"/> project that will be modified with merged modules.</param>
		/// <returns></returns>
		private bool TryBundleProjects(ref ExcelPackage package, bool recurseReferences, ref ExcelVbaProject project)
		{
			project = package.Workbook.VbaProject;
			var references = this.GetReferences(package.Workbook.VbaProject);

			/*	
			 *	Iterate through the references of the dependent Excel projects, recursing if desired.
			 *	This merges dependent modules into the target project, and adds system libraries as well,
			 *	to assist in preventing invalid projects on load.
			*/
			foreach (var reference in new List<ExcelVbaReference>(references.Excel))
			{
				var path = this.GetReferencePath(reference);
				this.TryGetFileInfo(path, out var referenceInfo);
				this.TryGetExcelPackage(referenceInfo, out var referencePackage);
				var referenceProject = referencePackage.Workbook.VbaProject;

				if (recurseReferences && this.GetReferences(referenceProject).Excel.Count() > 0)
				{
					_logger.Log($"Recursing references for \"{reference}\"");
					TryBundleProjects(ref referencePackage, true, ref referenceProject);
				}
				this.TryMergeModules(referenceProject, project, out var modules);
				this.TryAddToProject(ref project, modules);
				this.TryRemoveReference(ref project, reference);
			}

			//	Add the system 
			this.TryMergeSystemReferences(ref project, references.System);

			return true;
		}

		/// <summary>
		/// Merges required non-Excel file references that are referenced by referenced files.
		/// </summary>
		/// <param name="project"></param>
		/// <param name="references"></param>
		/// <remarks>For example, if a reference needs Microsoft Scripting Runtime 5.3, we want the resulting bundled project to have it as well.</remarks>
		/// <returns></returns>
		private bool TryMergeSystemReferences(ref ExcelVbaProject project, IEnumerable<ExcelVbaReference> references)
		{
			foreach (var reference in references)
			{
				//	Should not be adding Excel files
				if (reference.ReferenceRecordID == 14)
					continue;

				//	No need to add the same reference.
				if (project.References.Contains(reference))
					continue;

				project.References.Add(reference);
			}
			return true;
		}
	}
}
