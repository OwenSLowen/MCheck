using Objects;
using Objects.Geometry;
using Speckle.Automate.Sdk;
using Speckle.Core.Logging;
using Speckle.Core.Models.Extensions;
using ClosedXML;
using ClosedXML.Excel;
using System.Data;

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

        Console.WriteLine("Creating datatable");
        var streamDataTable = new DataTable();
        streamDataTable.TableName = "StreamObjects";

        var columns = new[] { "Id", "Name", "Speckle Type" };

        foreach (var column in columns)
        {
            streamDataTable.Columns.Add(column, typeof(string));
        }

        foreach (var commit in commitObject.Flatten())
        {
            var row = streamDataTable.NewRow();

            foreach (var column in columns)
            {
                var property = column.ToLowerInvariant().Replace(' ', '_');

                row[column] = commit[property]?.ToString() ?? "";
            }
            streamDataTable.Rows.Add(row);
        }

        var outputFile = $"out/SpeckleObjects.xlsx";


        Console.WriteLine("Creating workbook");
        using (var workbook = new XLWorkbook())
        {
            var streams = workbook.Worksheets.Add(streamDataTable);

            Console.WriteLine($"Saving workbook at '{Path.GetFullPath(outputFile)}'");
            workbook.SaveAs(outputFile);
        }

        Console.WriteLine("Storing excel file online");
        await automationContext.StoreFileResult(outputFile);

        automationContext.MarkRunSuccess($"Created report");

        //if (count < functionInputs.SpeckleTypeTargetCount) {
      //automationContext.MarkRunFailed($"Counted {count} objects where {functionInputs.SpeckleTypeTargetCount} were expected");
      //return;
    //}

    }
}
