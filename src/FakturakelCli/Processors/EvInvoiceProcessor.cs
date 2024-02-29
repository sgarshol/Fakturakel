using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace FakturakelCli.Processors;

public class EvInvoiceProcessor : IInvoiceProcessor
{
    public void ProcessInvoice(string path)
    {
        using var pdfReader = new PdfReader(path);

        var sb = new StringBuilder();

        for (var pageNumber = 1; pageNumber <= pdfReader.NumberOfPages; pageNumber++)
        {
            sb.Append(PdfTextExtractor.GetTextFromPage(pdfReader, pageNumber));
        }

        var text = sb.ToString();

        Match match;

        match = Regex.Match(text, @"6970631402831026");

        if (match.Success)
        {
            var report = new EVChargeReport();

            match = Regex.Match(text, @"Målepunkt ID: 29763505  - Målernummer: 6970631402831026  - Forbruk: (?<UsageKWhTotal>\d+) kWh  - Referanse:");

            if (match.Success)
            {
                report.UsageKWhTotal = match.Groups["UsageKWhTotal"]?.ToDecimal() ?? throw new Exception("No");
            }

            match = Regex.Match(text, @"(?<Kid>\d+) 1503.19.34610 (?<DueDate>\d{2}.\d{2}.\d{4}) (?<InvoiceSum>.*)");

            if (match.Success)
            {
                report.Kid = match.Groups["Kid"].Value ?? throw new Exception("No");
                report.DueDate = match.Groups["DueDate"].Value ?? throw new Exception("No");
                report.InvoiceSum = match.Groups["InvoiceSum"]?.ToDecimal() ?? throw new Exception("No");
            }

            match = Regex.Match(text, @"Månedlig fastbeløp \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} \d{2} dager (?<FixedCostMonth>.*) kr/mnd (?<FixedCost>.*)");

            if (match.Success)
            {
                report.FixedCostMonth = match.Groups["FixedCostMonth"]?.ToDecimal() ?? throw new Exception("No");
                report.FixedCost = match.Groups["FixedCost"]?.ToDecimal() ?? throw new Exception("No");
            }

            match = Regex.Match(text, @"Administrasjonsgebyr 25 % (?<AdminFee>.*)");

            if (match.Success)
            {
                report.AdminFee = match.Groups["AdminFee"]?.ToDecimal() ?? throw new Exception("No");
            }

            match = Regex.Match(text, @"Enovaavgift \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} \d{2} dager (?<EnovaTaxMonth>.*) nok (?<EnovaTax>.*)");

            if (match.Success)
            {
                report.EnovaTax = match.Groups["EnovaTax"]?.ToDecimal() ?? throw new Exception("No");
                report.EnovaTaxMonth = match.Groups["EnovaTaxMonth"]?.ToDecimal() ?? throw new Exception("No");
            }

            match = Regex.Match(text, @"Kapasitet 10-15 kW \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} \d{2} dager (?<Capacity_10_15_Month>.*) kr/mnd (?<Capacity_10_15>.*)");

            if (match.Success)
            {
                report.Capacity_10_15 = match.Groups["Capacity_10_15"]?.ToDecimal() ?? throw new Exception("No");
                report.Capacity_10_15_Month = match.Groups["Capacity_10_15_Month"]?.ToDecimal() ?? throw new Exception("No");
            }

            match = Regex.Match(text, @"Kapasitet 15-20 kW \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} \d{2} dager (?<Capacity_15_20_Month>.*) kr/mnd (?<Capacity_15_20>.*)");

            if (match.Success)
            {
                report.Capacity_15_20 = match.Groups["Capacity_15_20"]?.ToDecimal() ?? throw new Exception("No");
                report.Capacity_15_20_Month = match.Groups["Capacity_15_20_Month"]?.ToDecimal() ?? throw new Exception("No");
            }

            match = Regex.Match(text, @"Kapasitet 20-25 kW \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} \d{2} dager (?<Capacity_20_25_Month>.*) kr/mnd (?<Capacity_20_25>.*)");

            if (match.Success)
            {
                report.Capacity_20_25 = match.Groups["Capacity_20_25"]?.ToDecimal() ?? throw new Exception("No");
                report.Capacity_20_25_Month = match.Groups["Capacity_20_25_Month"]?.ToDecimal() ?? throw new Exception("No");
            }

            match = Regex.Match(text, @"Energirapportering \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} \d{2} dager (?<EnergyReportingMonth>.*) kr/mnd (?<EnergyReporting>.*)");

            if (match.Success)
            {
                report.EnergyReporting = match.Groups["EnergyReporting"]?.ToDecimal() ?? throw new Exception("No");
                report.EnergyReportingMonth = match.Groups["EnergyReportingMonth"]?.ToDecimal() ?? throw new Exception("No");
            }

            match = Regex.Match(text, @"Din strømpris \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} (?<UsageKWh>.*) kWh (?<UsagePrice>.*) øre/kWh (?<UsageCost>.*)");

            if (match.Success)
            {
                report.UsageKWh = match.Groups["UsageKWh"]?.ToDecimal() ?? throw new Exception("No");
                report.UsagePrice = match.Groups["UsagePrice"]?.ToDecimal() ?? throw new Exception("No");
                report.UsageCost = match.Groups["UsageCost"]?.ToDecimal() ?? throw new Exception("No");
            }

            match = Regex.Match(text, @"Energiledd dag \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} (?<EnergyLinkDayKWh>.*) kWh (?<EnergyLinkDayPrice>.*) øre/kWh (?<EnergyLinkDayCost>.*)");

            if (match.Success)
            {
                report.EnergyLinkDayKWh = match.Groups["EnergyLinkDayKWh"]?.ToDecimal() ?? throw new Exception("No");
                report.EnergyLinkDayPrice = match.Groups["EnergyLinkDayPrice"]?.ToDecimal() ?? throw new Exception("No");
                report.EnergyLinkDayCost = match.Groups["EnergyLinkDayCost"]?.ToDecimal() ?? throw new Exception("No");
            }

            match = Regex.Match(text, @"Energiledd natt/helg \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} (?<EnergyLinkNightKWh>.*) kWh (?<EnergyLinkNightPrice>.*) øre/kWh (?<EnergyLinkNightCost>.*)");

            if (match.Success)
            {
                report.EnergyLinkNightKWh = match.Groups["EnergyLinkNightKWh"]?.ToDecimal() ?? throw new Exception("No");
                report.EnergyLinkNightPrice = match.Groups["EnergyLinkNightPrice"]?.ToDecimal() ?? throw new Exception("No");
                report.EnergyLinkNightCost = match.Groups["EnergyLinkNightCost"]?.ToDecimal() ?? throw new Exception("No");
            }

            match = Regex.Match(text, @"Sikringsresultat hver mnd \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2}  (?<Sikringsresultat>.*)");

            if (match.Success)
            {
                report.Sikringsresultat = match.Groups["Sikringsresultat"]?.ToDecimal() ?? throw new Exception("No");
            }

            match = Regex.Match(text, @"Forbruksavgift \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} (?<UsageTaxKWh>.*) kWh (?<UsageTaxPrice>.*) øre/kWh (?<UsageTaxCost>.*)");

            if (match.Success)
            {
                report.UsageTaxKWh = match.Groups["UsageTaxKWh"]?.ToDecimal() ?? throw new Exception("No");
                report.UsageTaxPrice = match.Groups["UsageTaxPrice"]?.ToDecimal() ?? throw new Exception("No");
                report.UsageTaxCost = match.Groups["UsageTaxCost"]?.ToDecimal() ?? throw new Exception("No");
            }

            match = Regex.Match(text, @"Midlert. strømstønad \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} (?<SubsidiesKWh>.*) kWh (?<SubsidiesPrice>.*) øre/kWh (?<SubsidiesCost>.*)");

            if (match.Success)
            {
                report.SubsidiesKWh = match.Groups["SubsidiesKWh"]?.ToDecimal() ?? throw new Exception("No");
                report.SubsidiesPrice = match.Groups["SubsidiesPrice"]?.ToDecimal() ?? throw new Exception("No");
                report.SubsidiesCost = match.Groups["SubsidiesCost"]?.ToDecimal() ?? throw new Exception("No");
            }

            Console.WriteLine(JsonSerializer.Serialize(report, new JsonSerializerOptions() { WriteIndented = true }));

        }
    }
}

public interface IInvoiceProcessor
{
    void ProcessInvoice(string path);
}