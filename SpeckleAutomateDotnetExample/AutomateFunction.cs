using Microsoft.Office.Interop.Excel;
using Objects;
using Speckle.Automate.Sdk;
using Speckle.Core.Models.Extensions;

public static class AutomateFunction
{
  public static async Task Run(
    AutomationContext automationContext,
    FunctionInputs functionInputs
  )
  {
    Console.WriteLine("Starting execution");
    _ = typeof(ObjectsKit).Assembly; // INFO: Force objects kit to initialize

    Console.WriteLine("Receiving version");
    var commitObject = await automationContext.ReceiveVersion();

    Console.WriteLine("Received version: " + commitObject);

    var count = commitObject
      .Flatten()
      .Count(b => b.speckle_type == functionInputs.SpeckleTypeToCount);

    Console.WriteLine($"Counted {count} objects");
    Console.WriteLine("Finding excel?");

        Application excel = new Application();

        // Open the workbook
       // Workbook wb = excel.Workbooks.Open(@"C:\\temp\\mcheck.xlsx");

        // ...

        FileInfo fi = new FileInfo("C:\temp\\mcheck.xlsx");
        if (fi.Exists)
        {
            System.Diagnostics.Process.Start("C:\temp\\mcheck.xlsx");
            Console.WriteLine("where is my excel?");
        }

        if (count < functionInputs.SpeckleTypeTargetCount) {
      automationContext.MarkRunFailed($"Counted {count} objects where {functionInputs.SpeckleTypeTargetCount} were expected");
      return;
    }

    automationContext.MarkRunSuccess($"Counted {count} objects");
  }
}
