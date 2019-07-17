using System;
using System.Diagnostics;
using System.Linq;
using Microsoft.Vbe.Interop;


namespace ExcelVba
{

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

                    VBComponent vbComponent = (VBComponent) component;
                    Properties props = vbComponent.Properties;

                    vbComponent.Export($@"C:\Users\mvaysman\Source\Repos\vbaExtensions\{vbComponent.Name}.TXT");

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
    }
}
