# Install-Module PowerHTML -Scope CurrentUser -ErrorAction Stop

[CmdletBinding()] # for -Verbose
param(
  $year = (Get-Date).Year,
  [Parameter(mandatory=$true)]$structure,
  [Parameter(mandatory=$true)]$cookieSessionName = "SSESS<hash>",
  [Parameter(mandatory=$true)]$cookieSessionValue = "xxxxxxxxxxxxxxxxxx",
  $filterBestQualification = $true, # Afficher uniquement la meilleure qualification ; à $false le pilote apparaît dans chaque qualification validée
  $addOtherMembers = $false # Egalement prendre en compte les adhérents du club qui sont déjà licenciés dans un autre club
)

Write-Host "Year = $year"

$baseUrl = "https://intranet.ffvl.fr"
$url = "$baseUrl/structure/$structure/licences/$year"

# $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
# $cookie = New-Object System.Net.Cookie 
# $cookie.Name = $cookieSessionName
# $cookie.Value = $cookieSessionValue
# $cookie.Domain = ".ffvl.fr"
# $session.Cookies.Add($cookie);
# $res = Invoke-WebRequest -UseBasicParsing -Uri $url -WebSession $session
# $content = $res.Content

Write-Host -ForegroundColor Cyan "Téléchargement des licenciés... $url"
$wc = New-Object System.Net.WebClient
$wc.Headers.Add('Cookie', "$cookieSessionName=$cookieSessionValue")
$content = $wc.DownloadString($url)

$html = ConvertFrom-Html -Content $content
$tables = $html.SelectNodes("//div[@id='content']//table")  # Note: no js so we don't have the <table class="sticky-header">

$licencies_club = $tables[0]
$adherents_club = $tables[1]

function Get-Licencie ($firstname, $lastname, $licencie_url, $licence_type)
{
  $licencie_qualifications = @{
    pro = $false
    moniteur = $false
    biplace = $false
    bpc = $false
    bpc_theorique = $false
    bpc_pratique = $false
    bp = $false
    bp_theorique = $false
    bp_pratique = $false
    bpi = $false
    bp_speedriding = $false
    bpc_speedriding = $false
    juge = $false
    treuil = $false
    accompagnateur = $false
    animateur = $false
  }

  $licencie_content = $wc.DownloadString($licencie_url)
  $licencie_html = ConvertFrom-Html -Content $licencie_content

  $first_link = $licencie_html.SelectSingleNode("//div[@id='block-system-main']/a"); # [starts-with(@href, '/licences/')]
  $first_licence_year = $first_link.InnerText

  $options = $licencie_html.SelectSingleNode("//div[@id='block-system-main']/div[@class='item-list'][1]/ul");
  $licencie_options = @()
  foreach ($li in $options.SelectNodes('li'))
  {
    $option = $li.InnerText
    $licencie_options += $option

    # Possible options:
    # -----------------
    # Pack Essentiel Volant
    # Pack Tranquillité Volant
    # Pack Sérénité Volant
    # Pack Prémium Volant
    # Extension sports de pleine nature
    # Extension Assistance Rapatriement Monde entier
    # IJ Accidentelles 25?/jour (réservé Travailleurs Non Salariés)
    # Pack Essentiel Passager Volant
    # Extension Assistance Rapatriement Passager Monde entier
    # IJ Accidentelles 50?/jour (réservé Travailleurs Non Salariés)
    # Pack Essentiel Passager Catakite, Buggy ou Tandemkite (qualif. Tandemkitesurf obligatoire)
    # Extension RC - ULM classe 2 et 3
    # Extension RC - ULM biplace classe 2
    # Protection Juridique
    # -----------------
    # Formule A - Dommages accidentels seuls Valeur assurée 1000?
    # Formule A - Dommages accidentels seuls Valeur assurée 2000?
    # Formule A - Dommages accidentels seuls Valeur assurée 3000?
    # Formule B - Dommages accid/vol/perte Valeur assurée 1000?
    # Formule B - Dommages accid/vol/perte Valeur assurée 2000?
    # Formule B - Dommages accid/vol/perte Valeur assurée 3000?
    # -----------------
    # Option carte compétiteur parapente et speed-riding
    # Option carte compétiteur delta
    # Option carte compétiteur Kite
    # Option carte compétiteur cerf-volant
    # Option carte compétiteur boomerang
  }

  $h3 = $licencie_html.SelectSingleNode("//div[@class='item-list']/h3[text()='Qualifications']")
  if ($null -eq $h3)
  {
    $nodes = @()
  }
  else
  {
    $ul = $h3.SelectSingleNode("following-sibling::ul")
    $nodes = $ul.SelectNodes('li')
  }

  $obtained_this_year = @()
  foreach ($li in $nodes)
  {
    $qualification = $li.InnerText
    Write-Verbose $qualification

    if ($qualification.StartsWith("BPJEPS - option Parapente") -or
      $qualification.StartsWith("Moniteur BEES") -or
      $qualification.StartsWith("DEJEPS"))
    {
      $licencie_qualifications.pro = $true
      $licencie_qualifications.moniteur = $true
    }
    elseif ($qualification.StartsWith("Moniteur fédéral de vol libre option Parapente"))
    {
      $licencie_qualifications.moniteur = $true
    }
    elseif ($qualification.StartsWith("Qualification Biplace Parapente"))
    {
      $licencie_qualifications.biplace = $true
    }
    elseif ($qualification.StartsWith("Brevet de pilote confirmé Parapente"))
    {
      $licencie_qualifications.bpc = $true
      $licencie_qualifications.bpc_theorique = $true
      $licencie_qualifications.bpc_pratique = $true
    }
    elseif ($qualification.StartsWith("Brevet de pilote Parapente"))
    {
      $licencie_qualifications.bp = $true
      $licencie_qualifications.bp_theorique = $true
      $licencie_qualifications.bp_pratique = $true
    }
    elseif ($qualification.StartsWith("Brevet de pilote initial"))
    {
      $licencie_qualifications.bpi = $true
      $licencie_qualifications.bpi_theorique = $true
      $licencie_qualifications.bpi_pratique = $true
    }
    # Partie
    elseif ($qualification.StartsWith("Partie théorique du brevet de pilote confirmé Parapente"))
    {
      $licencie_qualifications.bpc_theorique = $true
    }
    elseif ($qualification.StartsWith("Partie pratique du brevet de pilote confirmé Parapente"))
    {
      $licencie_qualifications.bpc_pratique = $true
    }
    elseif ($qualification.StartsWith("Partie théorique du brevet de pilote Parapente"))
    {
      $licencie_qualifications.bp_theorique = $true
    }
    elseif ($qualification.StartsWith("Partie pratique du brevet de pilote Parapente"))
    {
      $licencie_qualifications.bp_pratique = $true
    }
    elseif ($qualification.StartsWith("Partie théorique du brevet de pilote initial Parapente"))
    {
      $licencie_qualifications.bpi_theorique = $true
    }
    elseif ($qualification.StartsWith("Partie pratique du brevet de pilote initial Parapente"))
    {
      $licencie_qualifications.bpi_pratique = $true
    }
    # Others qualifications
    elseif ($qualification.StartsWith("Accompagnateur fédéral Parapente"))
    {
      $licencie_qualifications.accompagnateur = $true
    }
    elseif ($qualification.StartsWith("Animateur fédéral Parapente"))
    {
      $licencie_qualifications.animateur = $true
    }
    elseif ($qualification.StartsWith("Brevet de pilote Speed-Riding n&deg;"))
    {
      $licencie_qualifications.bp_speedriding = $true
    }
    elseif ($qualification.StartsWith("Brevet de pilote confirmé Speed-Riding n&deg;"))
    {
      $licencie_qualifications.bpc_speedriding = $true
    }
    elseif ($qualification.StartsWith("Juge de précision d'atterrissage  Parapente"))
    {
      $licencie_qualifications.juge = $true
    }
    elseif ($qualification.StartsWith("Qualification fixe parapente Treuil"))
    {
      $licencie_qualifications.treuil = $true
    }
    elseif ($qualification.StartsWith("Parrainage ou préformation biplace  Parapente") -or
      $qualification.StartsWith("Entrée en formation biplace Parapente") -or
      $qualification.StartsWith("Statut aspirant biplaceur Parapente") -or
      $qualification.StartsWith("Sensibilisation à la pédagogie spécifique de ") -or
      $qualification.StartsWith("Elève Moniteur Fédéral (36 mois) Parapente") -or
      $qualification.StartsWith("UCC enseignement Speed-Riding") -or
      $qualification.StartsWith("UCC Hand'icare biplace Parapente"))
    {
      # ignore
    }
    else
    {
      Write-Host -ForegroundColor Yellow "Type de qualification inconnu : $qualification"
    }

    if ($qualification.EndsWith($year))
    {
      $obtained_this_year += $qualification
    }
  }

  # Fix state
  if ($licencie_qualifications.bpc_theorique -and $licencie_qualifications.bpc_pratique)
  {
    $licencie_qualifications.bpc = $true
  }
  if ($licencie_qualifications.bp_theorique -and $licencie_qualifications.bp_pratique)
  {
    $licencie_qualifications.bp = $true
  }
  if ($licencie_qualifications.bpi_theorique -and $licencie_qualifications.bpi_pratique)
  {
    $licencie_qualifications.bpi = $true
  }

  $best_qualification = ""
  if ($licencie_qualifications.pro)
  {
    $best_qualification = "pro"
  }
  elseif ($licencie_qualifications.moniteur)
  {
    $best_qualification = "moniteur"
  }
  elseif ($licencie_qualifications.biplace)
  {
    $best_qualification = "biplace"
  }
  elseif ($licencie_qualifications.bpc)
  {
    $best_qualification = "bpc"
  }
  elseif ($licencie_qualifications.bp)
  {
    $best_qualification = "bp"
  }
  elseif ($licencie_qualifications.bpi)
  {
    $best_qualification = "bpi"
  }

  $result = @{
    firstname = $firstname
    lastname = $lastname
    licence_link = $licence_link
    licence_type = $licence_type
    options = $licencie_options
    qualifications = $licencie_qualifications
    best_qualification = $best_qualification
    first_licence_year = $first_licence_year
    obtained_this_year = $obtained_this_year
  }
  return $result
}

$licencies = @()

Write-Host -ForegroundColor Cyan "Téléchargement de chaque licencié :"

$trs = $licencies_club.SelectNodes('tbody/tr')
if ($addOtherMembers)
{
  $trs += $adherents_club.SelectNodes('tbody/tr')
}
foreach ($tr in $trs)
{
  $firstname = $tr.SelectSingleNode('td[2]').InnerText
  $lastname = $tr.SelectSingleNode('td[3]').InnerText
  $licence_link = $tr.SelectSingleNode('td[4]/a').Attributes["href"].Value
  $licence_type = $tr.SelectSingleNode('td[6]').InnerText

  $licencie_url = $baseUrl + $licence_link
  Write-Host "$firstname $lastname | $licencie_url"
  $licencies += Get-Licencie $firstname $lastname $licencie_url $licence_type
}

function Get-Qualified ($licencies, $level, $ignoreIfHasBetterQualification)
{
  $qualified = @()
  foreach ($licencie in $licencies)
  {
    if ($licencie.qualifications[$level] -eq $true)
    {
      if ($ignoreIfHasBetterQualification -and ($licencie.best_qualification -ne $level))
      {
        continue
      }

      $qualified += $licencie.firstname + " " + $licencie.lastname
    }
  }
  return $qualified
}

function Get-UnQualified ($licencies)
{
  $unqualified = @()
  foreach ($licencie in $licencies)
  {
    if ($licencie.best_qualification -eq "")
    {
      $unqualified += "(" + $licencie.first_licence_year + ") " + $licencie.firstname + " " + $licencie.lastname
    }
  }
  return $unqualified
}

function Get-ObtainedQualificationsThisYear ($licencies)
{
  $qualified = @()
  $stats = @{
    bpi = @()
    bp = @()
    bpc = @()
    biplace = @()
    moniteur = @()
    pro = @()
    bp_speedriding = @()
    bpc_speedriding = @()
    juge = @()
    treuil = @()
    animateur = @()
    accompagnateur = @()
  }

  foreach ($licencie in $licencies)
  {
    $user = $licencie.firstname + " " + $licencie.lastname
    foreach ($obtained in $licencie.obtained_this_year)
    {
      # Gestion des brevets obtenus en plusieurs parties
      $got_brevet = $false
      $brevet = $null
      if ($obtained.StartsWith("Partie théorique du brevet de pilote confirmé Parapente") -or $obtained.StartsWith("Partie pratique du brevet de pilote confirmé Parapente"))
      {
        $brevet = "Brevet de pilote confirmé Parapente"
        $msg = $user + " - " + $brevet
        if ($licencie.qualifications.bpc -and ($msg -notin $qualified))
        {
          $qualified += $msg
          $got_brevet = $true
        }
      }
      elseif ($obtained.StartsWith("Partie théorique du brevet de pilote Parapente") -or $obtained.StartsWith("Partie pratique du brevet de pilote Parapente"))
      {
        $brevet = "Brevet de pilote Parapente"
        $msg = $user + " - " + $brevet
        if ($licencie.qualifications.bp -and ($msg -notin $qualified))
        {
          $qualified += $msg
          $got_brevet = $true
        }
      }
      elseif ($obtained.StartsWith("Partie théorique du brevet de pilote initial Parapente") -or $obtained.StartsWith("Partie pratique du brevet de pilote initial Parapente"))
      {
        $brevet = "Brevet de pilote initial Parapente"
        $msg = $user + " - " + $brevet
        if ($licencie.qualifications.bpi -and ($msg -notin $qualified))
        {
          $qualified += $msg
          $got_brevet = $true
        }
      }
      else
      {
        $brevet = $obtained
        $qualified += $user + " - " + $brevet
        $got_brevet = $true
      }

      if ($got_brevet -eq $false)
      {
        continue
      }

      if ($brevet.StartsWith("BPJEPS - option Parapente"))
      {
        $stats.pro += $user
      }
      if ($brevet.StartsWith("Moniteur fédéral de vol libre option Parapente"))
      {
        $stats.moniteur += $user
      }
      elseif ($brevet.StartsWith("Qualification Biplace Parapente"))
      {
        $stats.biplace += $user
      }
      elseif ($brevet.StartsWith("Brevet de pilote confirmé Parapente"))
      {
        $stats.bpc += $user
      }
      elseif ($brevet.StartsWith("Brevet de pilote Parapente"))
      {
        $stats.bp += $user
      }
      elseif ($brevet.StartsWith("Brevet de pilote initial"))
      {
        $stats.bpi += $user
      }
      elseif ($brevet.StartsWith("Accompagnateur fédéral Parapente"))
      {
        $stats.accompagnateur += $user
      }
      elseif ($brevet.StartsWith("Animateur fédéral Parapente"))
      {
        $stats.animateur += $user
      }
      elseif ($brevet.StartsWith("Brevet de pilote Speed-Riding"))
      {
        $stats.bp_speedriding += $user
      }
      elseif ($brevet.StartsWith("Brevet de pilote confirmé Speed-Riding"))
      {
        $stats.bpc_speedriding += $user
      }
      elseif ($brevet.StartsWith("Juge de précision d'atterrissage  Parapente"))
      {
        $stats.juge += $user
      }
      elseif ($brevet.StartsWith("Qualification fixe parapente Treuil"))
      {
        $stats.treuil += $user
      }
    }
  }
  return @{stats = $stats ; qualified = $qualified}
}

function Get-LicenceTypes ($licencies)
{
  $types = @{}
  foreach ($licencie in $licencies)
  {
    if ($types.ContainsKey($licencie.licence_type))
    {
      $types[$licencie.licence_type]++
    }
    else
    {
      $types[$licencie.licence_type] = 1
    }
  }
  return $types
}

function Get-LicenceOptions ($licencies)
{
  $options = @{}
  foreach ($licencie in $licencies)
  {
    foreach ($option in $licencie.options)
    {
      if ($options.ContainsKey($option))
      {
        $options[$option]++
      }
      else
      {
        $options[$option] = 1
      }
    }
  }
  return $options
}

Write-Host "`n----------`n"

Write-Host -ForegroundColor Cyan "Licences :"
Get-LicenceTypes $licencies | Format-Table -Autosize

$licence_options = Get-LicenceOptions $licencies
Write-Host -ForegroundColor Cyan "Options - Packs:"
$packs = $licence_options.GetEnumerator() | Where-Object { $_.Name.StartsWith("Pack ") -and ($_.Name -ne "Pack Essentiel Passager Volant") }
$packs | Format-Table
$total = 0
$packs | ForEach-Object { $total += $_.Value } | Out-Null
$nbWithoutIA = $licencies.Count - $total
$percent = ($nbWithoutIA/$licencies.Count).ToString("P1")
Write-Host -ForegroundColor Yellow "=> Sans IA = $nbWithoutIA (sur $($licencies.Count) soit $percent%)`n"
Write-Host -ForegroundColor Green "Note: d'autres assurances couvrant le risque parapente non ffvl existent"

Write-Host -ForegroundColor Cyan "Options - Formules:"
$licence_options.GetEnumerator() | Where-Object { $_.Name.StartsWith("Formule ") } | Format-Table -Autosize

Write-Host -ForegroundColor Cyan "Options - Autres:"
$others = "Extension sports de pleine nature","Protection Juridique","Option carte compétiteur parapente et speed-riding"
$licence_options.GetEnumerator() | Where-Object Name -in $others | Format-Table -Autosize

Write-Host "`n----------`n"

Write-Host -ForegroundColor Cyan "Légende :"
Write-Host "bpi = brevet de pilote initial"
Write-Host "bp = brevet de pilote"
Write-Host "bpc = brevet de pilote confirmé"
Write-Host "moniteur = moniteur fédéral"
Write-Host "pro = moniteur d'état"

Write-Host "`n----------`n"

Write-Host -ForegroundColor Cyan "Qualifications :"
$results = @{}
$categories = @("bpi", "bp", "bpc", "biplace", "moniteur", "pro")
foreach ($category in $categories)
{
  $res = Get-Qualified $licencies $category $filterBestQualification | Sort-Object
  $nb = $res.Count
  $results[$category] = $nb

  Write-Host -ForegroundColor Blue "`n${category} ($nb) :"
  $res
}

$res = Get-UnQualified $licencies | Sort-Object
$nb = $res.Count
$results["sans brevet"] = $nb
Write-Host -ForegroundColor Blue "`nsans brevet ($nb) :"
Write-Host -ForegroundColor Blue "(première année de licence ffvl entre parenthèses)"
$res | Format-Table


$now = Get-Date
$currentYear = $now.Year
if ($currentYear -eq $year)
{
  $extraYearWarning = " (au " + $now.ToString("dd-MM-yyyy") + ")"
}
Write-Host -ForegroundColor Cyan "`nStats Qualifications $year$extraYearWarning :"
$results | Format-Table -AutoSize


# biplaceur sans IA Passager
$biplaceursWithoutIAPassager = @()
$biplaceurs = $licencies | Where-Object { $_.qualifications.biplace -and $_.licence_type.StartsWith("Pratiquant biplace associatif") }
foreach ($biplaceur in $biplaceurs)
{
  $hasIAPassager = "Pack Essentiel Passager Volant" -in $biplaceur.options
  if (-not $hasIAPassager)
  {
    $biplaceursWithoutIAPassager += $biplaceur.firstname + " " + $biplaceur.lastname
  }
}
if ($biplaceursWithoutIAPassager.Count -gt 0)
{
  Write-Host -ForegroundColor Yellow "Biplaceurs avec licence biplace associatif sans IA passager :"
  $biplaceursWithoutIAPassager | Format-Table
}


Write-Host -ForegroundColor Cyan "`nVotre club comporte aussi :"
$categories = @("accompagnateur", "animateur", "juge")
foreach ($category in $categories)
{
  $res = Get-Qualified $licencies $category $false | Sort-Object
  Write-Host -ForegroundColor Blue "`n${category}:"
  $res | Format-Table
}

Write-Host -ForegroundColor Cyan "`nQualifications obtenues dans l'année $year$extraYearWarning :"
$res = Get-ObtainedQualificationsThisYear $licencies
# Write-Host -ForegroundColor Blue "Brut :"
# $res.qualified | Format-Table
# Write-Host -ForegroundColor Blue "Par groupe :"
foreach ($key in $res.stats.Keys)
{
  $nb = $res.stats[$key].Count
  if ($nb -gt 0)
  {
    Write-Host -ForegroundColor Green "$key ($nb) :"
    $res.stats[$key] | Sort-Object | Format-Table
  }
}


# Primo-pratiquant et BPI obtenu dans l'année
$primos = $licencies | Where-Object licence_type -eq "Primo-pratiquant en autonomie parapente, delta, speed-riding"
$primos_bpi = @()
foreach ($primo in $primos)
{
  if ($primo.qualifications.bpi)
  {
    $primos_bpi += $primo.firstname + " " + $primo.lastname
  }
}
if ($primos_bpi.Count -gt 0)
{
  Write-Host -ForegroundColor Cyan "`nPrimo-pratiquant et BPI obtenu dans l'année :"
  $primos_bpi | Sort-Object | Format-Table
}
