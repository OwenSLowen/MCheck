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

    Console.WriteLine($"Counted {count} objects whut");
    Console.WriteLine("Finding excel?");
    automationContext.MarkRunSuccess($"yep");
    automationContext.MarkRunSuccess("no$yep");

        try
            {
                string filePath = @"C:\temp\checkm.txt"; // Replace with your actual file path
                string content = File.ReadAllText(filePath);

                Console.WriteLine("File content:");
                Console.WriteLine(content);
                automationContext.MarkRunSuccess(content);
        }
            catch (Exception ex)
            {
            automationContext.MarkRunSuccess("nope");
        }
     
  
        

        if (count < functionInputs.SpeckleTypeTargetCount) {
      automationContext.MarkRunFailed($"Counted {count} objects where {functionInputs.SpeckleTypeTargetCount} were expected");
      return;
    }

    }
}
