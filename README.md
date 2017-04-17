# VBA Module Bunder
Imports source's referenced file's classes and modules and bundles it into one singular project.

##  Required References for Building
This project uses the following references:
* [EPPlus](https://github.com/pruiz/EPPlus/tree/master/EPPlus) 
* [System.ValueTuple](https://www.nuget.org/packages/System.ValueTuple/) library from Microsoft
* Visual Studio 2017 for C# 7 syntax.

## Usage
### Command Line Parameters using the Test Console
There are 5 properties exposed from the Bundler API, _italic_ are optional:
* __Source__: The path to the file you wish to pull all references from and merge into.
* __Target__: The path to the resulting file after the merge is complete.
* _Recurse_: If true, will go down the chain of references until none are left, bubbling the merges.
* _Use-Source_: If true, will default to using the source modules on any conflict. If false, will prompt or throw exception.
* _Only-Import-Used_: If true, will search through the code and determine which modules are required for functionality, else will include everything.
* _Help_: Displays a help message similar to this list. Can be invoked with ?, h, hlp, help.

You pass arguments with a '/' in front of the parameter name, and a space between the name and value, like so:
`C:\>VbaModuleBunder.exe /source "C:\some\path\to\file.xlam" /target "C:\some\path\to\anotherFile.xlsm" /recurse true /use-source true /only-include-used true`

### Other
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
			TryBundleProjects(ref referencePackage, ref referenceProject);
		}
		this.TryMergeModules(referenceProject, project, out var modules);
		this.TryMergeSystemReferences(ref project, referencedReferences.System);
		this.TryAddToProject(ref project, modules);
		this.TryRemoveReference(ref project, reference);
	}
            
## Caveats
* EPPlus cannot transfer/create Designer Modules (UserForms)
* EPPlus cannot save a file as an `xlam` file, thus no "hidden" libraries could be created from this. You could of course save this as an `xlsm` extension, then open and save as `xlam`.
* Ribbon methods (customUI) appear to not function properly when merged.
