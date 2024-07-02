using Objects;
using Speckle.Automate.Sdk;
using Speckle.Core.Models.Extensions;
using Microsoft.Office.Interop.Excel;

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
        string target = "https://www.microsoft.com";
        System.Diagnostics.Process.Start(target);

        // Instantiate a Workbook object.
        Workbook workbook = new Workbook();
       
       

        // Add a new worksheet to the Excel object.
        Worksheet worksheet = workbook.Worksheets.Add("MySheet");
        worksheet.Cells["A1"].PutValue("Hello, World!");
    if (count < functionInputs.SpeckleTypeTargetCount) {
      automationContext.MarkRunFailed($"Counted {count} objects where {functionInputs.SpeckleTypeTargetCount} were expected");
      return;
    }

    automationContext.MarkRunSuccess($"Counted {count} objects");
  }
}
