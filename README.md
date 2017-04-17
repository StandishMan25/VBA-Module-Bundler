# VBA Module Bunder
Imports source's referenced xlam/xlsm file's classes and modules and bundles it into one singular project.

## Usage
Add a reference to the library in the target application and follow the below format. 

    var bundler = new Bundler(new Logger());

    bundler.TryGetFileInfo(source, out var sourceInfo);
    bundler.TryGetExcelPackage(sourceInfo, out var package);
    bundler.TryBundleProjects(ref package);

`TryBundleProjects` is the simple way, but if desired, you could replicate the private method and inject your own logic if desired.

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
				var referencedReferences = this.GetReferences(referenceProject);

				if (_recurseReferences && referencedReferences.Excel.Count() > 0)
				{
					_logger.Log($"Recursing references for \"{reference}\"");
					TryBundleProjects(ref referencePackage, ref referenceProject);
				}
				this.TryMergeModules(referenceProject, project, out var modules);
				this.TryMergeSystemReferences(ref project, referencedReferences.System);
				this.TryAddToProject(ref project, modules);
				this.TryRemoveReference(ref project, reference);
			}
            
##  Required References
This project uses [EPPlus](https://github.com/pruiz/EPPlus/tree/master/EPPlus) and the temporary [System.ValueTuple](https://www.nuget.org/packages/System.ValueTuple/) library from Microsoft, and Visual Studio 2017 for C# 7 syntax.
