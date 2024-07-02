using Objects;
using Speckle.Automate.Sdk;
using Speckle.Core.Models.Extensions;
using System.Data.OleDb;

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

        

        // ...

        FileInfo fi = new FileInfo("C:\\Users\\921907\\github\\ModelCheck_Classifications_Naming\\mcheck.xlsx");
        if (fi.Exists)
        {
            System.Diagnostics.Process.Start(@"C:\Users\921907\github\ModelCheck_Classifications_Naming\mcheck.xlsx");
        }

        if (count < functionInputs.SpeckleTypeTargetCount) {
      automationContext.MarkRunFailed($"Counted {count} objects where {functionInputs.SpeckleTypeTargetCount} were expected");
      return;
    }

    automationContext.MarkRunSuccess($"Counted {count} objects");
  }
}
