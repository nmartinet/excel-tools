namespace VBAExtractor {
  using System;
  using System.Collections.Generic;
  using System.IO;
  using System.Linq;
  using Microsoft.Vbe.Interop;
  using Excel = Microsoft.Office.Interop.Excel;
  using VB = Microsoft.Vbe.Interop;

  class Program {
    /// <summary>
    ///     Exctracts the non empty vb modules from excel workbooks.
    ///     TODO    Add a way to reinject modules in the workbook
    ///     TODO    Add more infos when extracting - print current file, module etc
    ///     TODO    Add more checks/validatinos - valid xl file, path accessible, writable etc.
    /// </summary>
    /// <param name="args"></param>
    static void Main(string[] args) {
      //if no specific file a given, go through 
      if (args.Length == 0) {
        //only look for valid excel extentions
        //When opening excel workbooks, a file temporary file with a ~$ is
        //created - exclude them.
        var exts = new List<string> { ".xls", ".xlsx", ".xlsm", ".xlsb" };
        var fldr = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        foreach (var file in Directory.GetFiles(fldr).Where(f => exts.Any(f.EndsWith) && !f.Contains("~$")))
          ExtractVbaModules(file);
      } else {
        foreach (var wb in args)
          ExtractVbaModules(wb);
      };
    }

    /// <summary>
    ///     Extracts all non empty modules from a given workbook
    ///     to a folder in the same directory as the workbook
    ///     with the same name as the wokbook.
    ///     All excel objects get a module, sheets, workbooks etc.
    ///     checks for non empty to only get modified modules.
    /// </summary>
    /// <param name="wbPath">Full path to the targeted workbook</param>
    private static void ExtractVbaModules(string wbPath) {
      var xl = new Excel.Application { ScreenUpdating = false, EnableEvents = false, DisplayAlerts = false };
      var wb = xl.Workbooks.Open(wbPath);

      foreach (VB.VBComponent comp in wb.VBProject.VBComponents)
        ExportModule(wbPath, comp);

      wb.Close();
      xl.Quit();
    }

    /// <summary>
    ///     Saves a module to an external file if it is non empty
    /// </summary>
    /// <param name="wbPath">Path to the workbook. Used to find the appropriate save path</param>
    /// <param name="comp">VBE Module</param>
    private static void ExportModule(string wbPath, VBComponent comp) {
      if (comp.CodeModule.CountOfLines == 0) return;
      var basePath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
      var wbFolder = System.IO.Directory.CreateDirectory(Path.Combine(basePath, Path.GetFileNameWithoutExtension(wbPath)));

      comp.Export(String.Format("{0}\\{1}.vb", wbFolder.FullName.ToString(), comp.CodeModule.Name.ToString()));

    }
  }
}