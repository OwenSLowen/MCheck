using Microsoft.Office.Interop.Excel;
using Objects;
using Speckle.Automate.Sdk;
using Speckle.Core.Models.Extensions;
using System;
using System.IO;

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

        string filePath = @"C:\temp\checkm.txt"; // Replace with your actual file path
        string content = File.ReadAllText(filePath);
        try
            {
               
                Console.WriteLine("File content:");
                Console.WriteLine(content);
                automationContext.MarkRunSuccess(content);
        }
            catch (Exception ex)
            {
            automationContext.MarkRunSuccess(content);
        }
     
  
        

        if (count < functionInputs.SpeckleTypeTargetCount) {
      automationContext.MarkRunFailed($"Counted {count} objects where {functionInputs.SpeckleTypeTargetCount} were expected");
      return;
    }

    }
}
