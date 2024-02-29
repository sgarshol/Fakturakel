Add-Type -Path C:\temp\itextsharp.dll

$item = Get-Item 'C:\temp\Fjordkraft 2023.02.pdf'

function Get-FjordkraftFields($item)
{
    $pdf = New-Object -TypeName iTextSharp.text.pdf.PdfReader -ArgumentList $item.FullName

    $text = $null

    for ($pageNumber = 1; $pageNumber -le $($pdf.NumberOfPages); $pageNumber++)
    {
        $text += [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf, $pageNumber)
    }

    $report = @{}

    ### Forbruk
    if ($text -match "Målepunkt ID: 29763505  - Målernummer: 6970631402831026  - Forbruk: (?<Forbruk_kWh>\d+) kWh  - Referanse:")
    {
        $report["Forbruk_kWh"] = [decimal]::Parse($Matches["Forbruk_kWh"])
    }
    else
    {
        "Could not resolve Forbruk_kWh" | Write-Host -ForegroundColor Yellow

        return
    }

    ### FakturaSum
    if ($text -match "(?<Kid>\d+) 1503.19.34610 (?<Forfallsdato>\d{2}.\d{2}.\d{4}) (?<FakturaSum>.*)")
    {
        $report["Kid"] = $Matches["Kid"]
        $report["Forfallsdato"] = $Matches["Forfallsdato"]
        $report["FakturaSum"] = [decimal]::Parse($Matches["FakturaSum"])
    }
    else
    {
        "Could not resolve FakturaSum" | Write-Host -ForegroundColor Yellow

        return
    }

    ### Fastpris
    if ($text -match "Månedlig fastbeløp \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} \d{2} dager (?<FastprisMnd>.*) kr/mnd (?<Fastpris>.*)")
    {
        $report["FastprisMnd"] = [decimal]::Parse($Matches["FastprisMnd"])
        $report["Fastpris"] = [decimal]::Parse($Matches["Fastpris"])
    }
    else
    {
        "Could not resolve Fastpris" | Write-Host -ForegroundColor Yellow

        return
    }

    ### Administrasjonsgebyr
    if ($text -match "Administrasjonsgebyr 25 % (?<Administrasjonsgebyr>.*)")
    {
        $report["Administrasjonsgebyr"] = [decimal]::Parse($Matches["Administrasjonsgebyr"])
    }
    else
    {
        "Could not resolve Administrasjonsgebyr" | Write-Host -ForegroundColor Yellow

        return
    }

    ### Enovaavgift
    if ($text -match "Enovaavgift \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} \d{2} dager (?<EnovaavgiftMnd>.*) nok (?<Enovaavgift>.*)")
    {
        $report["EnovaavgiftMnd"] = [decimal]::Parse($Matches["EnovaavgiftMnd"])
        $report["Enovaavgift"] = [decimal]::Parse($Matches["Enovaavgift"])
    }
    else
    {
        "Could not resolve Enovaavgift" | Write-Host -ForegroundColor Yellow

        return
    }

    ### Kapasitet_10_15
    if ($text -match "Kapasitet 10-15 kW \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} \d{2} dager (?<Kapasitet_10_15_Mnd>.*) kr/mnd (?<Kapasitet_10_15>.*)")
    {
        $report["Kapasitet_10_15_Mnd"] = [decimal]::Parse($Matches["Kapasitet_10_15_Mnd"])
        $report["Kapasitet_10_15"] = [decimal]::Parse($Matches["Kapasitet_10_15"])
    }
    else
    {
        $report["Kapasitet_10_15"] = 0
    }

    ### Kapasitet_15_20
    if ($text -match "Kapasitet 15-20 kW \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} \d{2} dager (?<Kapasitet_15_20_Mnd>.*) kr/mnd (?<Kapasitet_15_20>.*)")
    {
        $report["Kapasitet_15_20_Mnd"] = [decimal]::Parse($Matches["Kapasitet_15_20_Mnd"])
        $report["Kapasitet_15_20"] = [decimal]::Parse($Matches["Kapasitet_15_20"])
    }
    else
    {
        $report["Kapasitet_15_20"] = 0
    }

    ### Kapasitet_20_25
    if ($text -match "Kapasitet 20-25 kW \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} \d{2} dager (?<Kapasitet_20_25_Mnd>.*) kr/mnd (?<Kapasitet_20_25>.*)")
    {
        $report["Kapasitet_20_25_Mnd"] = [decimal]::Parse($Matches["Kapasitet_20_25_Mnd"])
        $report["Kapasitet_20_25"] = [decimal]::Parse($Matches["Kapasitet_20_25"])
    }
    else
    {
        $report["Kapasitet_20_25"] = 0
    }

    ### EnergiOgKlima
    if ($text -match "Energi og Klimarapportering \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} \d{2} dager (?<EnergiOgKlimaMnd>.*) kr/mnd (?<EnergiOgKlima>.*)")
    {
        $report["EnergiOgKlimaMnd"] = [decimal]::Parse($Matches["EnergiOgKlimaMnd"])
        $report["EnergiOgKlima"] = [decimal]::Parse($Matches["EnergiOgKlima"])
    }
    else
    {
        $report["EnergiOgKlima"] = 0
    }

    ### StrømForbruk
    if ($text -match "Din strømpris \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} (?<StrømForbruk_kWh>.*) kWh (?<StrømForbruk_pris>.*) øre/kWh (?<StrømForbruk>.*)")
    {
        $report["StrømForbruk_kWh"] = [decimal]::Parse($Matches["StrømForbruk_kWh"])
        $report["StrømForbruk_pris"] = [decimal]::Parse($Matches["StrømForbruk_pris"])
        $report["StrømForbruk"] = [decimal]::Parse($Matches["StrømForbruk"])
    }
    else
    {
        "Could not resolve StrømForbruk" | Write-Host -ForegroundColor Yellow

        return
    }

    ### Energiledd_Dag
    if ($text -match "Energiledd dag \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} (?<Energiledd_Dag_kWh>.*) kWh (?<Energiledd_Dag_pris>.*) øre/kWh (?<Energiledd_Dag>.*)")
    {
        $report["Energiledd_Dag_kWh"] = [decimal]::Parse($Matches["Energiledd_Dag_kWh"])
        $report["Energiledd_Dag_pris"] = [decimal]::Parse($Matches["Energiledd_Dag_pris"])
        $report["Energiledd_Dag"] = [decimal]::Parse($Matches["Energiledd_Dag"])
    }
    else
    {
        "Could not resolve Energiledd_Dag" | Write-Host -ForegroundColor Yellow

        return
    }

    ### Energiledd_Natt
    if ($text -match "Energiledd natt/helg \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} (?<Energiledd_Natt_kWh>.*) kWh (?<Energiledd_Natt_pris>.*) øre/kWh (?<Energiledd_Natt>.*)")
    {
        $report["Energiledd_Natt_kWh"] = [decimal]::Parse($Matches["Energiledd_Natt_kWh"])
        $report["Energiledd_Natt_pris"] = [decimal]::Parse($Matches["Energiledd_Natt_pris"])
        $report["Energiledd_Natt"] = [decimal]::Parse($Matches["Energiledd_Natt"])
    }
    else
    {
        "Could not resolve Energiledd_Natt" | Write-Host -ForegroundColor Yellow

        return
    }

    ### Sikringsresultat
    if ($text -match "Sikringsresultat hver mnd \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2}  (?<Sikringsresultat>.*)")
    {
        $report["Sikringsresultat"] = [decimal]::Parse($Matches["Sikringsresultat"])
    }
    else
    {
        "Could not resolve Sikringsresultat" | Write-Host -ForegroundColor Yellow

        return
    }

    ### Forbruksavgift
    if ($text -match "Forbruksavgift \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} (?<Forbruksavgift_kWh>.*) kWh (?<Forbruksavgift_pris>.*) øre/kWh (?<Forbruksavgift>.*)")
    {
        $report["Forbruksavgift_kWh"] = [decimal]::Parse($Matches["Forbruksavgift_kWh"])
        $report["Forbruksavgift_pris"] = [decimal]::Parse($Matches["Forbruksavgift_pris"])
        $report["Forbruksavgift"] = [decimal]::Parse($Matches["Forbruksavgift"])
    }
    else
    {
        "Could not resolve Forbruksavgift" | Write-Host -ForegroundColor Yellow

        return
    }

    ### MidlertidigStrømstønad
    if ($text -match "Midlert. strømstønad \d{2}.\d{2}.\d{2} - \d{2}.\d{2}.\d{2} (?<MidlertidigStrømstønad_kWh>.*) kWh (?<MidlertidigStrømstønad_pris>.*) øre/kWh (?<MidlertidigStrømstønad>.*)")
    {
        $report["MidlertidigStrømstønad_kWh"] = [decimal]::Parse($Matches["MidlertidigStrømstønad_kWh"])
        $report["MidlertidigStrømstønad_pris"] = [decimal]::Parse($Matches["MidlertidigStrømstønad_pris"])
        $report["MidlertidigStrømstønad"] = [decimal]::Parse($Matches["MidlertidigStrømstønad"])
    }
    else
    {
        "Could not resolve MidlertidigStrømstønad" | Write-Host -ForegroundColor Yellow

        return
    }

    $items = [System.Collections.ArrayList]@()
}

function Validate-FjordkraftFields($report)
{
    $fields = [PSCustomObject]$report

    ($fields.Energiledd_Dag_kWh * $fields.Energiledd_Dag_pris / 100) - $fields.Energiledd_Dag
    ($fields.Energiledd_Natt_kWh * $fields.Energiledd_Natt_pris / 100) - $fields.Energiledd_Natt
}

function ConvertTo-FjordkraftFieldsList($report)
{
    $items = $report.Keys | 
        % {
            $key = $_
            $value = $report[$key]
        
            return [PSCustomObject]@{ Name = $key; Value = $value }
        } |
        Sort-Object -Property Name

    return $items
}


