
using FakturakelCli.Processors;

if (!args.Any())
{
    Console.WriteLine("Bruk: FakturakelCli {FILSTI}");

    return;
}

var path = args[0];

if (!File.Exists(path))
{
    Console.WriteLine($"FILSTI {path} eksiterer ikke.");
}

IInvoiceProcessor processor = new ChargeInvoiceProcessor();

processor.ProcessInvoice(path);

