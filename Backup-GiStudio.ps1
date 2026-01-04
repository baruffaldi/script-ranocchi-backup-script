<#
	Author: Filippo Baruffaldi <filippo@baruffaldi.info>
    Backup-GiStudio.ps1
    - Esegue il batch SQL preventivo
    - Copia le cartelle richieste in Z:\Backup-ANNO-MESE-GIORNO-ORA-MINUTO
    - Esclude AMeCO\tmp
    - Elimina i backup più vecchi di N giorni (se N>0)
    - Invia e-mail con tutto lo stdout (ed eventuali errori)
#>

# =======================
#   CONFIGURAZIONE
# =======================
$Debug                = $false       # se true visualizza più informazioni
$KeepLogs             = $true      # se true copia dentro la cartella LOGS il log
$RetentionDays        = 10          # N giorni; se 0, non cancella nulla
$LocalRetentionDays   = 3          # N giorni per retention locale (Archives); default 3
$DestFolderWithDate   = $true      # se true usa una cartella con la data nel nome
$BackupSQL            = $false       # se true effettua il backup SQL

# --- PATHS ---
$BackupRoot           = '\\NAS4FENIX\BackupRanocchiIncremental\'
$SrcDocs              = 'C:\RANOCCHI\GISTUDIO\gisbil\docs'
$SrcAmeco             = 'C:\RANOCCHI\GISTUDIO\AMeCO'
$AmecoTmp             = 'C:\RANOCCHI\GISTUDIO\AMeCO\tmp'
$PreBackupBat         = 'C:\RANOCCHI\GISTUDIO\gisbil\docs\backupSQL.bat'

# --- COMPRESSION ---
$CompressBck          = $true      # se true comprime il backup
$Prefer7Zip           = $true      # se true e 7zip presente, usa 7zip (comprime di più)

# --- SMTP ---
$SmtpServer = ''
$SmtpPort   = 587
$SmtpUser   = ''
$SmtpPass   = ''
$From       = ''
$To         = ''
$UseSsl     = $true

# =======================
#   INIZIO SCRIPT
# =======================
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# Determina cartella script (necessaria per Staging e Logs)
$ScriptDir =
    if ($PSScriptRoot) { $PSScriptRoot }
    elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath }
    elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
    else { (Get-Location).Path }

$script:startTime    = Get-Date
$script:lastTime     = $script:startTime
$timestamp            = Get-Date -Format 'yyyyMMdd-HHmm'

# Backup su cartella (nome base)
$backupFolder    = "Backup-$timestamp"
if ($DestFolderWithDate) {
	$DestDir         = Join-Path $BackupRoot $backupFolder
} else {
	$DestDir         = $BackupRoot
}

# Backup su archivio
$archiveName    = "$backupFolder.zip"
$archivePath    = Join-Path $BackupRoot $archiveName
# Local archive path (Archives subfolder)
$ArchivesDir    = Join-Path $ScriptDir 'Archives'
$localArchivePath = Join-Path $ArchivesDir $archiveName

# Staging persistente (nella cartella dello script)
$StagingDirName = 'BackupStaging'
$StagingDir     = Join-Path $ScriptDir $StagingDirName

$LogsRoot = Join-Path $ScriptDir 'LOGS'
$logsDate = Get-Date -Format 'yyyyMMdd-HHmmss'
$transcriptPath  = Join-Path $LogsRoot ("backup-transcript-{0}.log" -f ($logsDate))
$debugPath  = Join-Path $LogsRoot ("backup-debug-{0}.log" -f ($logsDate))
$hadError = $false


function Write-Header($text) {
    Write-Output ""
    Write-Output "========== $text =========="
}

function Invoke-RoboCopy {
    param(
        [Parameter(Mandatory)] [string] $Source,
        [Parameter(Mandatory)] [string] $Destination,
        [string[]] $ExtraArgs = @()
    )

    if (-not (Test-Path $Source)) { throw "Sorgente non trovata: $Source" }
    if (-not (Test-Path $Destination)) { New-Item -ItemType Directory -Path $Destination -Force | Out-Null }

    # Normalizza in una List (mantiene ordine e non perde niente)
    $argsList = New-Object System.Collections.Generic.List[string]
    foreach ($a in $ExtraArgs) { if ($null -ne $a -and $a -ne '') { [void]$argsList.Add($a) } }

    if ($Debug) {
        if (-not $debugPath) { throw "Debug attivo ma `$debugPath non è definito" }

        # Rimuove /NFL e /NDL (case-insensitive) mantenendo gli altri extraargs
        for ($i = $argsList.Count - 1; $i -ge 0; $i--) {
            if ($argsList[$i] -ieq '/NFL' -or $argsList[$i] -ieq '/NDL') {
                $argsList.RemoveAt($i)
            }
        }

        # Aggiunge solo se assente (case-insensitive)
        $need = @('/V','/TS','/FP','/BYTES') #,'/TEE')
        foreach ($n in $need) {
            if (-not ($argsList | Where-Object { $_ -ieq $n })) {
                [void]$argsList.Add($n)
            }
        }

        # /LOG+:$debugPath solo se non esiste già /LOG: o /LOG+
        if (-not ($argsList | Where-Object { $_ -match '^(?i)/LOG(\+|:)' })) {
            [void]$argsList.Add("/LOG+:$debugPath")
        }
    }

    # Torna a string[]
    $ExtraArgs = $argsList.ToArray()

    Write-Output "Eseguo RoboCopy:"
    Write-Output "  FROM: $Source"
    Write-Output "  TO  : $Destination"
    Write-Output "  ARGS: $($ExtraArgs -join ' ')"

    $args = @($Source, $Destination) + $ExtraArgs
    $proc = Start-Process -FilePath 'robocopy.exe' -ArgumentList $args -NoNewWindow -Wait -PassThru
    $code = $proc.ExitCode

    if ($code -ge 8) {
        throw "Robocopy ha restituito codice $code per '$Source' -> '$Destination'"
    } else {
        Write-Output "Robocopy completato con codice $code (OK / non fatale)."
    }
}

function Get-SevenZipPath {
    if (-not $Prefer7Zip) { return $null }
    $c = Get-Command '7z.exe' -ErrorAction SilentlyContinue
    if ($c) { return $c.Path }
    $paths = @(
        'C:\Program Files\7-Zip\7z.exe',
        'C:\Program Files (x86)\7-Zip\7z.exe'
    )
    foreach ($p in $paths) { if (Test-Path $p) { return $p } }
    return $null
}

function Compress-ItemToZip {
    param(
        [Parameter(Mandatory)] [string] $ItemPath,   # Cartella o file da zippare
        [Parameter(Mandatory)] [string] $ZipPath
    )
    # Calcolo parent e nome foglia per eseguire il comando dalla directory padre
    # Questo assicura che lo ZIP contenga la cartella radice (es. Backup-...) e non solo il contenuto.
    $parentDir = Split-Path -Parent $ItemPath
    $itemName  = Split-Path -Leaf $ItemPath

    $sevenZip = Get-SevenZipPath
    if ($sevenZip) {
        Write-Output "Comprimo con 7-Zip (massimo, -mx=9): $ZipPath"
        if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }

        # Eseguo 7z dalla cartella padre, includendo $itemName
        $args = @('a','-tzip','-mx=9','-y', $ZipPath, $itemName)
        $proc = Start-Process -FilePath $sevenZip -ArgumentList $args -WorkingDirectory $parentDir -NoNewWindow -Wait -PassThru
        if ($proc.ExitCode -ne 0) { throw "7-Zip ha restituito codice $($proc.ExitCode)" }
    } else {
        Write-Output "7-Zip non trovato: uso Compress-Archive (CompressionLevel=Optimal)."
        if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }

        # Compress-Archive preserva la struttura se si punta alla cartella
        Compress-Archive -Path $ItemPath -DestinationPath $ZipPath -CompressionLevel Optimal -Force
    }
}

function Calculate-Time {
    param(
        [bool] $Checkpoint = $false
    )
	
    $thisTime  = Get-Date
	
	if ($CheckPoint) {
		$when = "Orario"
		$elapsed  = $thisTime - $lastTime
	} else {
		$when = "Fine"
		$elapsed  = $thisTime - $script:startTime
		Write-Output ("Inizio: "    + $script:startTime.ToString('yyyy-MM-dd HH:mm:ss'))
	}
	
	Write-Output ($when + ": "    + $thisTime.ToString('yyyy-MM-dd HH:mm:ss'))
    Write-Output ("Durata: "  + ('{0:hh\:mm\:ss}' -f $elapsed))
	
	$script:lastTime = $thisTime
}

try {
    Start-Transcript -Path $transcriptPath -Append | Out-Null
    Write-Header "ORARIO"
    Write-Output ("Inizio: " + $script:startTime.ToString('yyyy-MM-dd HH:mm:ss'))
	
	<#
	$lockFile = Join-Path $StagingDir '.lock'
	if (Test-Path $lockFile) {
		Write-Error "Script già in esecuzione"
		exit 1
	}
	New-Item $lockFile -ItemType File | Out-Null
	#>

    Write-Header "VERIFICHE PRELIMINARI"
    if (-not (Test-Path $BackupRoot)) { throw "La radice destinazione non esiste: $BackupRoot" }
    Write-Output "Destinazione: $BackupRoot"

	if ($CompressBck) {
        # Assicuro che esista la StagingDir (o creata da zero o ripristinata/esistente)
		if (-not (Test-Path $StagingDir)) {
            New-Item -ItemType Directory -Path $StagingDir -Force | Out-Null
        }
        # Assicuro che esista la cartella Archives
        if (-not (Test-Path $ArchivesDir)) {
            New-Item -ItemType Directory -Path $ArchivesDir -Force | Out-Null
        }

		Write-Output "Staging persistente: $StagingDir"
        Write-Output "Cartella archivi locali: $ArchivesDir"
		Calculate-Time $true
	}

	if ($BackupSQL) {
		Write-Header "ESECUZIONE BACKUP SQL (batch preventivo)"
		if (-not (Test-Path $PreBackupBat)) {
			throw "File batch non trovato: $PreBackupBat"
		}
		Write-Output "Eseguo: $PreBackupBat"
		$p = Start-Process -FilePath $PreBackupBat -NoNewWindow -Wait -PassThru
		$exit = $p.ExitCode
		Write-Output "backupSQL.bat ExitCode: $exit"
		if ($exit -ne 0) {
			throw "backupSQL.bat terminato con codice $exit (errore)."
		}
		Calculate-Time $true
	}
	
	$destDocs  = Join-Path $StagingDir 'docs'
	$destAmeco = Join-Path $StagingDir 'AMeCO'

	Write-Header "SINCRONIZZAZIONE STAGING (docs)"
	# Uso /MIR per mirror (copia nuovi, aggiorna modificati, elimina cancellati)
	Invoke-RoboCopy -Source $SrcDocs -Destination $destDocs -ExtraArgs @(
		'/MIR','/COPY:DAT','/R:1','/W:5','/NFL','/NDL','/NP'
	)
	Calculate-Time $true

	Write-Header "SINCRONIZZAZIONE STAGING (AMeCO, escluso tmp)"
	Invoke-RoboCopy -Source $SrcAmeco -Destination $destAmeco -ExtraArgs @(
		'/MIR','/COPY:DAT','/R:1','/W:5','/NFL','/NDL','/NP',
		'/XD', $AmecoTmp
	)
	Calculate-Time $true

	if ($CompressBck) {
        Write-Header "PULIZIA ARCHIVI LOCALI VECCHI"
        if ($LocalRetentionDays -gt 0) {
            $localCutoff = (Get-Date).AddDays(-$LocalRetentionDays+1)
            Write-Output "Retention locale: $LocalRetentionDays giorni (elimino file < $($localCutoff.ToString('yyyy-MM-dd HH:mm')))"

            $oldLocalZips = Get-ChildItem -Path $ArchivesDir -Filter 'Backup-*.zip' -File -ErrorAction SilentlyContinue |
                            Where-Object { $_.LastWriteTime -lt $localCutoff }
            foreach ($lz in $oldLocalZips) {
                Write-Output "Elimino ZIP locale: $($lz.FullName)"
				try {
					Remove-Item $lz.FullName -Force -ErrorAction Stop
				} catch {
					Write-Output "ZIP $($z.FullName) NON ELIMINATO"
				}
            }
			Calculate-Time $true
        }

        Write-Header "COMPRESSIONE"
        # Comprimo direttamente la cartella di staging
        Compress-ItemToZip -ItemPath $StagingDir -ZipPath $localArchivePath
        Write-Output "Creato archivio locale: $localArchivePath"
		Calculate-Time $true

        # COPIA ARCHIVIO SU RETE
        if (Test-Path $localArchivePath) {
            Write-Header "COPIA ARCHIVIO SU RETE"
            Write-Output "Copio '$localArchivePath' in '$archivePath'..."
            Copy-Item -Path $localArchivePath -Destination $archivePath -Force
            Write-Output "Archivio copiato correttamente in destinazione."
			Calculate-Time $true
        }
	} else {
		Write-Header "CREAZIONE CARTELLA DI BACKUP FINALE"
		if ($DestFolderWithDate) {
			New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
		}
		
		Write-Header "COPIA STAGING → DESTINAZIONE FINALE"
		Invoke-RoboCopy -Source $StagingDir -Destination $DestDir -ExtraArgs @(
			'/MIR',
			'/COPY:DAT',
			'/R:1','/W:5',
			'/NFL','/NDL','/NP'
		)
		Calculate-Time $true
	}

    Write-Header "PULIZIA BACKUP VECCHI"
    if ($RetentionDays -gt 0) {
        $cutoff = (Get-Date).AddDays(-$RetentionDays)
        Write-Output "Retention: $RetentionDays giorni (elimino file < $($cutoff.ToString('yyyy-MM-dd HH:mm')))"

		# Archivi
        $oldZips = Get-ChildItem -Path $BackupRoot -Filter 'Backup-*.zip' -File -ErrorAction SilentlyContinue |
                   Where-Object { $_.LastWriteTime -lt $cutoff }
        foreach ($z in $oldZips) {
            Write-Output "Elimino ZIP: $($z.FullName)"
			try {
				Remove-Item $z.FullName -Force -ErrorAction Stop
			} catch {
				Write-Output "ZIP $($z.FullName) NON ELIMINATO"
			}
        }
		Calculate-Time $true

		# Cartelle
        $old = Get-ChildItem -Path $BackupRoot -Directory -ErrorAction SilentlyContinue |
               Where-Object { $_.Name -like 'Backup-*' -and $_.LastWriteTime -lt $cutoff }

        if ($old) {
            foreach ($dir in $old) {
                Write-Output "Elimino: $($dir.FullName)"
				try {
					Remove-Item -Path $dir.FullName -Recurse -Force -ErrorAction Stop
				} catch {
					Write-Output "FOLDER $($dir.FullName) NON ELIMINATA"
				}
            }
        } else {
            Write-Output "Nessuna cartella da eliminare."
        }
		Calculate-Time $true
    } else {
        Write-Output "Retention impostata a 0: salto la cancellazione dei vecchi backup."
    }

    Write-Header "ORARIO (FINE E DURATA)"
	Calculate-Time $false


    Write-Header "FINE OPERAZIONI"
	if ($CompressBck) {
		Write-Output "Backup ZIP completato senza eccezioni: $archivePath"
	} else {
		Write-Output "Backup completato senza eccezioni. Cartella: $DestDir"
	}
}
catch {
    $hadError = $true
    Write-Error ("ERRORE: " + $_.Exception.Message)
    if ($_.InvocationInfo.PositionMessage) {
        Write-Error ("Dettagli: " + $_.InvocationInfo.PositionMessage.Trim())
    }
}
finally {
    Stop-Transcript | Out-Null
    <#Remove-Item $lockFile -Force -ErrorAction SilentlyContinue#>
}

# =======================
#   INVIO E-MAIL
# =======================
if ($SmtpServer -ne '') {
	try {
		# Leggo tutto lo stdout/stderr catturato dal transcript
		$logRaw = Get-Content -Path $transcriptPath -Raw

		# Codifico per HTML e impacchetto in <pre> (così si vede tutto identico)
		Add-Type -AssemblyName System.Web
		$encoded = [System.Web.HttpUtility]::HtmlEncode($logRaw)
		$intro   = if (-not $hadError) { "<h1>BACKUP RIUSCITO</h1>" } else { "" }
		$esito   = if ($hadError) { "ERRORI" } else { "OK" }

		$body = @"
	$intro
	<p><strong>Esito:</strong> $esito</p>
	<p><strong>Cartella destinazione:</strong> $DestDir</p>
	<hr>
	<pre>$encoded</pre>
"@

		$subject = if ($hadError) { "Backup terminato con errori" } else { "Backup terminato" }

		$mail = New-Object System.Net.Mail.MailMessage
		$mail.From = $From
		$mail.To.Add($To)
		$mail.Subject = $subject
		$mail.IsBodyHtml = $true
		$mail.Body = $body

		$client = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
		$client.EnableSsl = $UseSsl
		if ($SmtpUser -and $SmtpPass) {
			$client.Credentials = New-Object System.Net.NetworkCredential($SmtpUser, $SmtpPass)
		} else {
			$client.UseDefaultCredentials = $true
		}

		$client.Send($mail)
		Write-Output "E-mail inviata a $To con oggetto: '$subject'."
	}
	catch {
		# Se fallisce l'invio e-mail, lo segnalo in console (rimane nel transcript)
		Write-Error ("ERRORE invio e-mail: " + $_.Exception.Message)
	}
}

<#
# Copia del transcript nella radice dei backup
if ($KeepLogs) {
	if ($transcriptPath -and (Test-Path $transcriptPath)) {
        # $ScriptDir è già calcolato all'inizio

		# LOGS dentro la cartella dello script
		$LogsRoot = Join-Path $ScriptDir 'LOGS'

		# Assicurati che esista
		if (-not (Test-Path $LogsRoot)) {
			New-Item -ItemType Directory -Path $LogsRoot -Force | Out-Null
		}

		$destTranscript = Join-Path $LogsRoot ("{0}-transcript.log" -f $backupFolder)
		Copy-Item -Path $transcriptPath -Destination $destTranscript -Force
		Write-Output "Transcript copiato in: $destTranscript"
	} else {
		Write-Warning "Transcript non trovato: $transcriptPath"
	}
}

# (Opzionale) Rimuovi il transcript temporaneo
# Remove-Item -Path $transcriptPath -Force -ErrorAction SilentlyContinue
#>
