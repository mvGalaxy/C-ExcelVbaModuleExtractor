void Main()
{
	string file=string.Empty;
	OpenFileDialog openDialog = new OpenFileDialog();
	openDialog.Title = "Select A File";
	openDialog.Filter = "Excel OpenXml (*.xlsx;*.xlsm)|*.xlsx;*.xlsm";

	if (openDialog.ShowDialog() == DialogResult.OK)
	{
		file = openDialog.FileName;
	}

    var vb = new VBAModuleExtractor();
	var e= new ExcelContainerExtractor();
	
	vb.Extract(file,Path.GetDirectoryName(file));
	e.UnzipContainer(new FileInfo(file),new DirectoryInfo(new FileInfo(file).DirectoryName));    
}

public class ExcelContainerExtractor
{

	public void ChangeExtensionToZip() {
		
	}

	public void UnzipContainer(FileInfo excelFile, DirectoryInfo targetExtractionDir)
	{
		string initialExtension = excelFile.Extension.ToString();
		try
		{
			if (excelFile.Extension == ".xlsx" || excelFile.Extension == ".xlsm")
			{
				excelFile.MoveTo(Path.ChangeExtension(excelFile.FullName, ".zip"));
				Z.Compression.Extensions.Extensions.ExtractZipFileToDirectory(excelFile, targetExtractionDir);

			}
			else
			{
				throw new FileLoadException("Excel file must have extension of xlsx or xlsm");
			}
		}
		finally{
			excelFile.MoveTo(Path.ChangeExtension(excelFile.FullName,initialExtension));
		}
		 	
	}
		
}

public class VBAModuleExtractor
{
	Microsoft.Office.Interop.Excel.Application excel = null;
	Microsoft.Office.Interop.Excel.Workbooks workbooks = null;
	Microsoft.Office.Interop.Excel.Workbook workbook = null;
	VBProject project = null;

	public void Extract(string fileName, string exportPath)
	{
		int excelProcessID = 0;
		try
		{

			excel = new Microsoft.Office.Interop.Excel.Application();
			excelProcessID = Process.GetProcessesByName("Excel").OrderByDescending(p => p.Id).Select(p => p.Id)
				.ToArray()[0];
			workbooks = excel.Workbooks;
			workbook = workbooks.Open(fileName, false, true, Type.Missing, Type.Missing, Type.Missing, true,
				Type.Missing, Type.Missing, false, false, Type.Missing, false, true, Type.Missing);

			project = workbook.VBProject;
			string projectName = project.Name;
			var procedureType = Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Proc;
			VBComponents components = project.VBComponents;
			foreach (VBComponent component in components)
			{

				VBComponent vbComponent = (VBComponent)component;
				Properties props = vbComponent.Properties;
				vbComponent.Export($@"{exportPath}\{vbComponent.Name}{GetFileExtensionForDocumentType(component)}");

			}

		}
		catch (Exception e)
		{
			Console.WriteLine("Failed to export VBA Modules");
			throw;
		}
		finally
		{
			excel.Quit();
			Process.GetProcessById(excelProcessID).Kill();
		}

	}
	
	private static string GetFileExtensionForDocumentType(VBComponent component)
	{
		return component.Type==vbext_ComponentType.vbext_ct_Document?".txt"
		                         :component.Type==vbext_ComponentType.vbext_ct_StdModule?".bas"
								 :component.Type==vbext_ComponentType.vbext_ct_ClassModule?".cls"
								 :".txt";
	}
}
