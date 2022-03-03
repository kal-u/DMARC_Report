######################################
# Nom du script : dmarc-report.ps1
# Utilité: ce script sert à traiter les rapports XML Dmarc reçus dans une boite aux lettres Exchange
# Le script alimente un fichier CSV. Il est alors très facile de créer un rapport Excel à base de tableaux et graphiques croisés dynamiques
# Usage: dmarc-report.ps1 (aucun argument)
# Auteur: Gabriel F 
# Mise à jour le: 03/03/2022
######################################


# Ajout du module de communication EWS (Exchange Web Services)
Import-Module "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

# Création de l'objet Service
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016)


# Définition de la boite aux lettres à scanner
$MailboxName = "utilisateur@domaine.fr"

# Récupération des credentials de l'utilisateur courant (qui doit avoir les droits sur la boite aux lettres
$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()

# Définition du répertoire de travail
$downloadDirectory = "C:\DMARC\"

# Définir du fichier de sortie
$outfile = $downloadDirectory+"dmarc.csv"

# Déclaration de l'URL d'accès aux Web Services Exchange
$uri=[system.URI] "https://SERVER/ews/exchange.asmx"

# Connexion au dossier Inbox = "Boite de réception" de la boite aux lettres
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind
$service.Url = $uri
$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)
$InboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)

# Définition du filtre de recherche dans la boite mail
$Subject = "Report Domain:"
$Sfsub = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject, $Subject)
$Sfha = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $true)
$sfCollection = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
$sfCollection.add($Sfha)
$sfCollection.add($Sfsub)

# Création du fichier de sortie (avec son entête) s'il n'existe pas encore
if (Test-Path $outfile -PathType Leaf) {
    write-host "Le fichier $outfile existe"
}
else {
    write-host "Le fichier $outfile n'existe pas."
    write-host "Création du fichier"
    New-Item -Path $outfile -ItemType File
    
    $delim = ';'
    $line = "DATE"+$delim+"DOMAINE"+$delim+"EXPEDITEUR"+$delim+"IP"+$delim+"SPF AUTH"+$delim+"SPF"+$delim+"DKIM"+$delim+"NB MAILS"+$delim+"EMETTEUR"+$delim+"DKIM AUTH"
    write-host $line
    Add-Content -Path $outfile -Value $line
}
Start-Sleep -Seconds 3

# Application du filtre de recherche sur les 2000 derniers mails
$view = new-object Microsoft.Exchange.WebServices.Data.ItemView(2000)
$frFolderResult = $InboxFolder.FindItems($sfCollection,$view)

# Traitement des mails
foreach ($miMailItems in $frFolderResult.Items)
{
    $miMailItems.Subject
    $miMailItems.Load()

    # Pour chaque pièce jointe, on l'enregistre dans le répertoire de travail en préfixant le nom du fichier avec la date du jour
    foreach($attach in $miMailItems.Attachments)
    {
          $attach.Load()
          $fiFile = new-object System.IO.FileStream(($downloadDirectory + "\" + (Get-Date -Format "yyMMdd") + "_" + $attach.Name.ToString()),[System.IO.FileMode]::Create)
          $fiFile.Write($attach.Content, 0, $attach.Content.Length)
          $fiFile.Close()
          write-host "Downloaded Attachment : " + (($downloadDirectory + "\" + (Get-Date -Format "yyMMdd") + "_" + $attach.Name.ToString()))
    }

    # Puis on passe le mail en "lu" et on le met dans la corbeille
    $miMailItems.isread = $true
    $miMailItems.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
    $miMailItems.delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems)
}

# Fonction pour extraire les fichiers zip
function Expand-ZIPFile($file, $destination)
{
      $shell = new-object -com shell.application
      $zip = $shell.NameSpace($file)
      foreach($item in $zip.items())
     {
           $shell.Namespace($destination).copyhere($item)
     }
}

# Définition du répertoire de travail et du répertoire des fichiers décompressés
$path=$downloadDirectory
$read_folder="$path"+"read_xml_files"
write-host $read_folder

# Définition du répertoire des fichiers XML traités
$path=$downloadDirectory
$processed_folder="$path"+"processed_xml_files"
write-host $processed_folder


# Liste de fichiers zip à traiter
$zipfiles=get-childitem("$path\*.Zip")

# Extration des fichiers xml contenus dans les zip
foreach($zipfile in $zipfiles.name)
{
    $source=$path+$zipfile
    write-host "ZIP SOURCE : $source"
    write-host "ZIP DESTINATION : $path+$zipfile"
    Expand-ZIPFile -file $source -destination $path
    Move-Item $source $read_folder
}


# Fonction de géolocation des adresses IP
function Get-MvaIpLocation {
    <#
.SYNOPSIS
    Retrieves Geo IP location data
.DESCRIPTION
    This command retrieves the Geo IP Location data for one or more IP addresses
.PARAMETER IPAddress <String[]>
    Specifies one or more IP Addresses for which you want to retrieve data for.
.EXAMPLE
    Get-MvaIpLocation -ipaddress '124.26.123.240','123.25.96.8'
.EXAMPLE
    '124.26.123.240','123.25.96.8' | Get-MvaIpLocation
.LINK
    https://get-note.net/2019/01/18/use-powershell-to-find-ip-geolocation
.INPUTS
    System.String
.OUTPUTS
    System.Management.Automation.PSCustomObject
.NOTES
    Author: Mario van Antwerpen
    Website: https://get-note.net
#>
    [cmdletbinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    Param (
        [Parameter(ValueFromPipeline, Mandatory, Position = 0, HelpMessage = "Enter an IP Address")]
        [ValidateScript({
            if ($_ -match '^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$') {
                $true
            } else {
                Throw "$_ is not a valid IPv4 Address!"
            }
        })]
        [string[]]$ipaddress
    )

    begin {
        Write-Verbose -message "Starting $($MyInvocation.Mycommand)"
    }

    process {
        foreach ($entry in $ipaddress) {
            $restUrl = "http://ip-api.com/json/$entry"

            try {
                Write-Verbose -Message "Connecting to rest endpoint"
                $result = Invoke-RestMethod -Method get -Uri $restUrl
                Write-output $result
            }
            catch {
                Write-Verbose -Message "Catched and error"
                $PSCmdlet.ThrowTerminatingError($PSitem)
            }
        }
    }

    end {
        Write-Verbose -message "Ending $($MyInvocation.Mycommand)"
    }
}


# Fonction de décompression des fichiers GZIP (extension .gz)
Function Ungzip {
    Param(
        $infile,
        $outfile = ($infile -replace '\.gz$','')
        )

    $input = New-Object System.IO.FileStream $inFile, ([IO.FileMode]::Open), ([IO.FileAccess]::Read), ([IO.FileShare]::Read)
    $output = New-Object System.IO.FileStream $outFile, ([IO.FileMode]::Create), ([IO.FileAccess]::Write), ([IO.FileShare]::None)
    $gzipStream = New-Object System.IO.Compression.GzipStream $input, ([IO.Compression.CompressionMode]::Decompress)

    $buffer = New-Object byte[](1024)
    while($true){
        $read = $gzipstream.Read($buffer, 0, 1024)
        if ($read -le 0){break}
        $output.Write($buffer, 0, $read)
        }

    $gzipStream.Close()
    $output.Close()
    $input.Close()
}

# Liste des fichiers GZIP
$gzfiles=get-childitem("$path\*.gz")

# Extraction des fichiers xml contenus dans les fichiers GZIP
foreach($gzfile in $gzfiles.name)
{
    $source=$path+$gzfile
    write-host "GZ SOURCE : $source"
    $dest = $source.Substring(0,$source.length-3)
    write-host "GZ DESTINATION :  $dest"
    Ungzip $source $dest
    Move-Item $source $read_folder
}


# Liste des fichiers XML dans le répertoire de travail
$files =Get-ChildItem("$path\*.xml")

# Traitement pour chaque fichier
Foreach($file in $files) {
    
    # $file_date=$file.LastAccessTime.Year.ToString()+"-"+$file.LastAccessTime.Month.ToString()+"-"+$file.LastAccessTime.Day.ToString()
    
    $WebConfigFile = "$path"+$file.name
    [xml] $xml = Get-Content $WebConfigFile
    $sender=$xml.SelectNodes("//report_metadata")
    $entries=$xml.SelectNodes("//record")
    $affected_domain=$xml.SelectNodes("//policy_published")
    $analyse_time = (Get-Date 01.01.1970)+([System.TimeSpan]::fromseconds($sender.date_range.begin))
    $formated_analyse_time = $analyse_time.Year.ToString()+"-"+$analyse_time.Month.ToString()+"-"+$analyse_time.Day.ToString()
    

    # Définition du délimiteur du CSV
    $delim = ';'

    # Traitement de chaque entrée du fichier XML
    foreach ($entry in $entries) {

        # Définition de la localisation de l'expéditeur
        $domain_emetteur = ""
        if(($entry.row.source_IP -eq "10.10.10.10") -or ($entry.row.source_IP -eq "10.10.10.11")) {
            $domain_emetteur = "MON IP PUBLIQUE"
        }
        elseif($entry.row.source_IP -eq "10.10.10.12") {
            $domain_emetteur = "MA SOLUTION d EMAILING DE MASSE"
        }
        else {
			try {  
				$hostname = [System.Net.Dns]::GetHostByAddress($_).HostName
				$domain_emetteur = $hostname.split('.')[-2,-1] -Join '.'
			} catch [Exception]{
					$domain_emetteur = ""
			}
			if ($domain_emetteur -eq "") { $domain_emetteur = "Inconnu" }
        }
        
        $line = $formated_analyse_time+"$delim"+$affected_domain.domain+"$delim"+$sender.org_name+"$delim"+$entry.row.source_IP+"$delim"+$entry.auth_results.spf.result+"$delim"+$entry.row.policy_evaluated.spf+"$delim"+$entry.row.policy_evaluated.dkim+"$delim"+$entry.row.count+"$delim"+$domain_emetteur+"$delim"+$entry.auth_results.spf.result
        write-host $line

        # Ajout de la ligne au fichier de résultat
        Add-Content -Path $outfile -Value $line 

      if($entry.auth_results.spf.result -eq "fail" -or $entry.auth_results.dkim.result -eq "fail") # Usurpation d'identité détectée
      {
         Write-host("Sender:", $sender.org_name,"`tSource IP:",$entry.row.source_IP,"`tSPF Auth Result:",$entry.auth_results.spf.result,"`tDKIM Auth Result:",$entry.auth_results.dkim.result,"`tSPF:",$entry.row.policy_evaluated.spf,"`tDKIM:",$entry.row.policy_evaluated.dkim,"`tNombre de mails",$entry.row.count,"`tDomaine : ",$affected_domain.domain,"`tDate: ",$formated_analyse_time,"`tEmetteur: ",$domain_emetteur) –ForegroundColor Yellow
      } else
      {
         Write-host("Sender:", $sender.org_name,"`tSource_IP:",$entry.row.source_IP,"`tSPF Auth Result:",$entry.auth_results.spf.result,"`tDKIM Auth Result:",$entry.auth_results.dkim.result,"`tSPF:",$entry.row.policy_evaluated.spf,"`tDKIM:",$entry.row.policy_evaluated.dkim,"`tNombre de mails",$entry.row.count,"`tDomaine : ",$affected_domain.domain,"`tDate: ",$formated_analyse_time,"`tEmetteur: ",$domain_emetteur) –ForegroundColor Green
      }
    }
    # Déplacement du fichier XML dans le répertoire des fichiers traités
    Move-Item $file $processed_folder
}
