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
$KeepLogs      = $true              # se true copia dentro la cartella LOGS il log
$RetentionDays = 20                 # N giorni; se 0, non cancella nulla
$BackupRoot    = '\\NAS4FENIX\BackupRanocchi\'
$SrcDocs       = 'C:\RANOCCHI\GISTUDIO\gisbil\docs'
$SrcAmeco      = 'C:\RANOCCHI\GISTUDIO\AMeCO'
$AmecoTmp      = 'C:\RANOCCHI\GISTUDIO\AMeCO\tmp'
$PreBackupBat  = 'C:\RANOCCHI\GISTUDIO\gisbil\docs\backupSQL.bat'

# --- COMPRESSION ---
$CompressBck   = $true              # se true comprime il backup
$Prefer7Zip    = $true              # se true e 7zip presente, usa 7zip (comprime di più)

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

$startTime       = Get-Date
$timestamp       = Get-Date -Format 'yyyyMMdd-HHmm'

# Backup su cartella (nome base)
$backupFolder    = "Backup-$timestamp"
$DestDir         = Join-Path $BackupRoot $backupFolder

# Backup su archivio
$archiveName    = "$backupFolder.zip"
$archivePath    = Join-Path $BackupRoot $archiveName

# Staging persistente (nella cartella dello script)
$StagingDirName = 'BackupStaging'
$StagingDir     = Join-Path $ScriptDir $StagingDirName

# Logica di ripristino: se non trovo BackupStaging, cerco cartelle "Backup-xxxxxxxx-xxxx"
# rimaste da un'esecuzione fallita (che aveva rinominato lo staging).
if (-not (Test-Path $StagingDir)) {
    $leftovers = Get-ChildItem -Path $ScriptDir -Directory | Where-Object { $_.Name -match '^Backup-\d{8}-\d{4}$' }
    # Se ce ne sono, prendiamo la più recente (o una qualsiasi) e la rinominiamo
    if ($leftovers) {
        $leftover = $leftovers | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        Write-Warning "Trovata cartella staging residua: '$($leftover.Name)'. Ripristino in '$StagingDirName'."
        Rename-Item -Path $leftover.FullName -NewName $StagingDirName -Force
    }
}

$transcriptPath  = Join-Path $env:TEMP ("backup-transcript-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
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
    # RoboCopy crea la destinazione se manca, ma meglio essere espliciti
    if (-not (Test-Path $Destination)) { New-Item -ItemType Directory -Path $Destination -Force | Out-Null }

    Write-Output "Eseguo RoboCopy:"
    Write-Output "  FROM: $Source"
    Write-Output "  TO  : $Destination"
    Write-Output "  ARGS: $($ExtraArgs -join ' ')"

    $args = @($Source, $Destination) + $ExtraArgs
    $proc = Start-Process -FilePath 'robocopy.exe' -ArgumentList $args -NoNewWindow -Wait -PassThru
    $code = $proc.ExitCode

    # Robocopy exit codes: 0 e 1 = successo; 2..7 = warning (accettabili); >=8 = errori.
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

try {
    Start-Transcript -Path $transcriptPath -Append | Out-Null
    Write-Header "ORARIO"
    Write-Output ("Inizio: " + $startTime.ToString('yyyy-MM-dd HH:mm:ss'))

    Write-Header "VERIFICHE PRELIMINARI"
    if (-not (Test-Path $BackupRoot)) { throw "La radice destinazione non esiste: $BackupRoot" }
    Write-Output "Destinazione: $BackupRoot"

	if ($CompressBck) {
        # Assicuro che esista la StagingDir (o creata da zero o ripristinata/esistente)
		if (-not (Test-Path $StagingDir)) {
            New-Item -ItemType Directory -Path $StagingDir -Force | Out-Null
        }
		Write-Output "Staging persistente: $StagingDir"
	}

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

	if ($CompressBck) {
        $destDocs  = Join-Path $StagingDir 'docs'
        $destAmeco = Join-Path $StagingDir 'AMeCO'

		Write-Header "SINCRONIZZAZIONE STAGING (docs)"
        # Uso /MIR per mirror (copia nuovi, aggiorna modificati, elimina cancellati)
		Invoke-RoboCopy -Source $SrcDocs -Destination $destDocs -ExtraArgs @(
			'/MIR','/COPY:DAT','/R:1','/W:5','/NFL','/NDL','/NP'
		)

		Write-Header "SINCRONIZZAZIONE STAGING (AMeCO, escluso tmp)"
		Invoke-RoboCopy -Source $SrcAmeco -Destination $destAmeco -ExtraArgs @(
			'/MIR','/COPY:DAT','/R:1','/W:5','/NFL','/NDL','/NP',
			'/XD', $AmecoTmp
		)

        # Rinomina temporanea per lo ZIP
        $tempStagingPath = Join-Path $ScriptDir $backupFolder

        # Se per assurdo esistesse già una cartella con quel nome (conflitto?), la rimuoviamo?
        # Non dovrebbe accadere col timestamp, ma per sicurezza:
        if (Test-Path $tempStagingPath) {
            Write-Warning "Cartella temporanea $tempStagingPath esistente? Rimozione forzata."
            Remove-Item $tempStagingPath -Recurse -Force
        }

        Write-Header "RINOMINA E COMPRESSIONE"
        Write-Output "Rinomino '$StagingDirName' in '$backupFolder'..."
        Rename-Item -Path $StagingDir -NewName $backupFolder

        # Percorso locale temporaneo per lo ZIP (nella stessa cartella dello script)
        $localArchivePath = Join-Path $ScriptDir $archiveName

        try {
            # Ora $StagingDir non esiste più con quel nome, esiste $tempStagingPath
            Write-Output "Comprimo in locale: $localArchivePath"
            Compress-ItemToZip -ItemPath $tempStagingPath -ZipPath $localArchivePath

            Write-Output "Sposto archivio in destinazione: $archivePath"
            Move-Item -Path $localArchivePath -Destination $archivePath -Force
            Write-Output "Archivio spostato con successo."
        }
        finally {
            # Ripristino nome
            Write-Output "Ripristino nome '$StagingDirName'..."
            if (Test-Path $tempStagingPath) {
                Rename-Item -Path $tempStagingPath -NewName $StagingDirName
            }
            # Pulizia di sicurezza: se il move fallisce, potremmo avere ancora il file locale?
            # Se ha fallito il move, l'utente potrebbe volerlo recuperare da locale,
            # quindi meglio non cancellarlo automaticamente qui, ma lasciare che se ne occupi la run successiva (overwrite)
            # o l'intervento manuale.
        }

        # Non cancello più lo staging alla fine
	} else {
		Write-Output "Creazione cartella di backup: $DestDir"
		New-Item -ItemType Directory -Path $DestDir -Force | Out-Null

		Write-Header "COPIA CARTELLA DOCS"
		$destDocs = Join-Path $DestDir 'docs'
		Invoke-RoboCopy -Source $SrcDocs -Destination $destDocs -ExtraArgs @(
			'/E',                  # include sottocartelle (anche vuote)
			'/COPY:DAT',           # copia Data, Attributi, Timestamp
			'/R:1','/W:5',         # 1 retry, 5s attesa
			'/NFL','/NDL','/NP'    # meno rumore
		)

		Write-Header "COPIA CARTELLA AMeCO (escludendo tmp)"
		$destAmeco = Join-Path $DestDir 'AMeCO'
		Invoke-RoboCopy -Source $SrcAmeco -Destination $destAmeco -ExtraArgs @(
			'/E',
			'/COPY:DAT',
			'/R:1','/W:5',
			'/NFL','/NDL','/NP',
			'/XD', $AmecoTmp      # esclude la sottocartella tmp
		)
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
            Remove-Item $z.FullName -Force -ErrorAction Stop
        }

		# Cartelle
        $old = Get-ChildItem -Path $BackupRoot -Directory -ErrorAction SilentlyContinue |
               Where-Object { $_.Name -like 'Backup-*' -and $_.LastWriteTime -lt $cutoff }

        if ($old) {
            foreach ($dir in $old) {
                Write-Output "Elimino: $($dir.FullName)"
                Remove-Item -Path $dir.FullName -Recurse -Force -ErrorAction Stop
            }
        } else {
            Write-Output "Nessuna cartella da eliminare."
        }
    } else {
        Write-Output "Retention impostata a 0: salto la cancellazione dei vecchi backup."
    }

    Write-Header "ORARIO (FINE E DURATA)"
    $endTime  = Get-Date
    $elapsed  = $endTime - $startTime
    Write-Output ("Fine: "    + $endTime.ToString('yyyy-MM-dd HH:mm:ss'))
    Write-Output ("Durata: "  + ('{0:hh\:mm\:ss}' -f $elapsed))


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
