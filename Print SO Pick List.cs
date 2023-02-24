// Allocate Stock and Print SO Pick list.
// Triggered by CallContextBpmData.ShortChar01 = "AllocatePrint"
// which is set in Sales Order Entry and is also used by -
//    Credit Checking - Sales Order Acknowledgement Print

System.IO.StreamWriter STDOUT = (System.IO.StreamWriter)null;
const string Method = "BPM-SalesOrder.InvoiceExists - Allocation And Print (2) - Post";
string OutPutFile = string.Format(@"C:\Temp\{0}.txt", Method);

//Erp.Contracts.SOPickListReportSvcContract hMyProc2 = null;
// Store the Current Plant before we change it
var CurrPlant = Session.PlantID;
var url = new Uri(Session.AppServerURL);

//toggle for debugging
bool Debug = false;

try
{
    if (Debug)
        /**/
        STDOUT = new System.IO.StreamWriter
      (new System.IO.FileStream(OutPutFile, System.IO.FileMode.Append));
    else
        /**/
        STDOUT = new System.IO.StreamWriter(System.IO.Stream.Null);

    STDOUT.WriteLine(string.Format("{0} - Start.", Method));
    STDOUT.WriteLine($"Original Plant = {CurrPlant}");

    string gAgentID = "SystemAgent";

    var gWorkStationID = callContextBpmData.ShortChar02;
    string gMaintProg = "Erp.Internal.OM.sopicklistreport.dll";
    int ReportStyle = 1001;
    DateTime? lowdate = DateTime.Parse("01/01/1900");
    DateTime? highdate = DateTime.Parse("31/12/2199");

    STDOUT.WriteLine($"WorkStationID =<{gWorkStationID}>");

    var PlantList = (from rowPartAlloc in this.Db.PartAlloc
                     join rowOrderRel in this.Db.OrderRel on new { Key0 = rowPartAlloc.Company, Key1 = rowPartAlloc.OrderNum, Key2 = rowPartAlloc.OrderLine, Key3 = rowPartAlloc.OrderRelNum } equals new { Key0 = rowOrderRel.Company, Key1 = rowOrderRel.OrderNum, Key2 = rowOrderRel.OrderLine, Key3 = rowOrderRel.OrderRelNum }
                     where (rowPartAlloc.Company == (this.BaqConstants.CompanyID)
                             && rowPartAlloc.OrderNum == (this.orderNum)
                             && rowPartAlloc.AllocatedQty > (0)
                             && (rowPartAlloc.Company == null
                             || rowPartAlloc.Company == ""
                             || rowPartAlloc.Company == (this.Session.CompanyID)))
                             && (rowOrderRel.OpenRelease == (true)
                             && rowOrderRel.Make == (false)
                             && (rowOrderRel.Company == null
                             || rowOrderRel.Company == ""
                             || rowOrderRel.Company == (this.Session.CompanyID)))
                     select new
                     {
                         rowOrderRel.Plant
                     }).Distinct();

    foreach (var Plant in PlantList)
    {
        STDOUT.WriteLine($"Changing to Plant: <{Plant.Plant}>");
        Session.PlantID = Plant.Plant;

        using (var DbLookup = new Erp.ErpContext())   // This looks to stop error 'The underlying provider failed on EnlistTransaction'
        {
            using (var hMyProc2 = Ice.Assemblies.ServiceRenderer.GetService<Erp.Contracts.SOPickListReportSvcContract>(DbLookup))
            {
                SOPickListReportTableset soplts = hMyProc2.GetNewParameters();
                foreach (var row in soplts.SOPickListReportParam)
                {
                    row.FromDate = lowdate;
                    row.FromDateToken = string.Empty;
                    row.ToDate = highdate;
                    row.ToDateToken = string.Empty;
                    row.OrderList = orderNum.ToString();
                    row.PartList = string.Empty;
                    row.NewPage = true;
                    row.BarCodes = true;
                    row.PrintKitComponents = true;
                    row.SysRowID = new Guid();
                    row.AgentSchedNum = 0;
                    row.AgentID = gAgentID;
                    row.AgentTaskNum = 0;
                    row.RecurringTask = false;

                    // Mod1 - Start of Block
                    row.WorkstationID = gWorkStationID;
                    row.ReportStyleNum = ReportStyle;
                    row.RptVersion = string.Empty;

                    //Always Print Preview
                    STDOUT.WriteLine("Using SSRSPREVIEW");
                    row.WorkstationID = callContextBpmData.ShortChar02;
                    row.AutoAction = "SSRSPREVIEW";
                    row.PrinterName = string.Empty;
                    row.RptPageSettings = string.Empty;
                    row.RptPrinterSettings = string.Empty;

                    row.TaskNote = string.Empty;
                    row.ArchiveCode = 0;
                    row.DateFormat = "dd/mm/yyyy";
                    row.NumericFormat = ",.";
                    row.AgentCompareString = string.Empty;
                    row.ProcessID = string.Empty;
                    row.ProcessCompany = string.Empty;
                    row.ProcessSystemCode = string.Empty;
                    row.ProcessTaskNum = 0;
                    row.DecimalsGeneral = 0;
                    row.DecimalsCost = 0;
                    row.DecimalsPrice = 0;
                    row.GlbDecimalsGeneral = 0;
                    row.GlbDecimalsCost = 0;
                    row.GlbDecimalsPrice = 0;
                    row.FaxSubject = string.Empty;
                    row.FaxTo = string.Empty;
                    row.FaxNumber = string.Empty;
                    row.EMailTo = string.Empty;
                    row.EMailCC = string.Empty;
                    row.EMailBCC = string.Empty;
                    row.EMailBody = string.Empty;
                    row.AttachmentType = "PDF";
                    row.ReportCurrencyCode = "GBP";
                    row.ReportCultureCode = "en-GB";
                    row.SSRSRenderFormat = string.Empty;
                    row.UIXml = string.Empty;
                    row.PrintReportParameters = false;
                    row.SSRSEnableRouting = false;
                    row.RowMod = "A";
                }

                // Now submit to Agent
                hMyProc2.SubmitToAgent(soplts, gAgentID, 0, 0, gMaintProg);
            }
        }
    }
}
catch (Exception ex)
{
    STDOUT.WriteLine(string.Format("*** {0} ***"), DateTime.Now.ToString());
    STDOUT.WriteLine(ex.Message);
}

finally
{
    // Reset the Plant
    try
    {
        Session.PlantID = CurrPlant;
    }
    catch
    {
        STDOUT.WriteLine("*** Failed to set Plant in finally block ****");
    }
    STDOUT.WriteLine(string.Format("{0} ----END---", Method));
    STDOUT.Close();
}
