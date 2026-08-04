<#
 .SYNOPSIS
   Argus - Active Directory Export Tool (LDAP Edition)

   Grafisches Tool (WPF) und Kommandozeilen-Tool zum Auslesen und Exportieren
   von Active-Directory-Benutzerobjekten nach Excel und/oder CSV - über reines
   LDAP (Port 389/636/3268/3269).

   Argus benötigt bewusst KEIN ADWS (Port 9389) und KEIN RSAT /
   ActiveDirectory-PowerShell-Modul. Es läuft damit auch in Umgebungen, in
   denen ADWS blockiert ist oder RSAT nicht installiert werden darf.
   Der CSV-Export hat keinerlei Abhängigkeiten; für den Excel-Export wird das
   PowerShell-Modul 'ImportExcel' benötigt (wird im GUI angeboten, falls es
   fehlt - alternativ einfach CSV exportieren).

 .DESCRIPTION
   Funktionsumfang:
     - Quelle wählbar: alle Benutzer, Mitglieder von Gruppe(n) (rekursiv,
       inklusive primärer Gruppe) oder eine Liste spezifischer Benutzer
       (sAMAccountName/UPN/Mail/CN) mit Abgleichbericht (nicht gefundene IDs)
     - Dynamischer Attribut-Filter-Builder (UND/ODER, typgerechte Operatoren
       für Datums- und Statusfelder, RAW-LDAP-Zusatzfilter für Power-User)
     - Frei wählbare Export-Spalten aus einem Attribut-Katalog plus beliebige
       zusätzliche LDAP-Attribute (mit Schema-Prüfung)
     - Optionale True/False-Spalten für Gruppenmitgliedschaften (rekursive
       Prüfung über verschachtelte Gruppen inkl. primärer Gruppe)
     - Ergebnis-Vorschau (DataGrid) vor dem Export
     - Export als formatierte Excel-Tabelle (ImportExcel) und/oder CSV
       (ohne Abhängigkeiten, UTF-8 mit BOM, Trennzeichen wählbar)
     - LDAPS (SSL/TLS), Signing+Sealing auf Port 389, Global Catalog
       (3268/3269), alternative Anmeldedaten, Zeitlimit
     - Profile (JSON) zum Speichern/Laden aller Einstellungen
     - Headless-Modus (-NoGui) für geplante Aufgaben / Automatisierung
     - Protokollierung in Datei (Audit-Trail) und Lauf-Protokoll im GUI
     - Nicht blockierende UI: Suche und Export laufen in eigenen Runspaces,
       Fortschritt und Protokoll werden live angezeigt, Abbruch jederzeit

 .PARAMETER NoGui
   Startet ohne Oberfläche (Headless-Modus). Einstellungen kommen aus
   -ProfilePath und/oder den übrigen Parametern. Exit-Codes:
   0 = Export erstellt, 2 = keine Treffer (kein Export), 1 = Fehler.

 .PARAMETER ProfilePath
   Pfad zu einer Profil-Datei (JSON), wie sie das GUI über
   "Profil > Speichern..." erzeugt. Im GUI-Modus wird das Profil beim Start
   geladen; im Headless-Modus dient es als Basis, die von explizit gesetzten
   Parametern überschrieben wird.

 .PARAMETER Server
   Domain Controller (Hostname/FQDN/IP). Leer = serverlose Bindung über den
   DC-Locator.

 .PARAMETER Port
   LDAP-Port: 389 (Standard), 636 (LDAPS), 3268 (Global Catalog),
   3269 (GC über SSL) oder ein benutzerdefinierter Port.

 .PARAMETER UseLdaps
   Erzwingt SSL/TLS (bei Port 636/3269 automatisch aktiv).

 .PARAMETER BaseDn
   Such-Basis, z. B. "OU=Benutzer,DC=firma,DC=local".
   Leer = gesamte Domäne (defaultNamingContext) bzw. bei Global Catalog der
   gesamte Forest (rootDomainNamingContext).

 .PARAMETER Mode
   Quelle: 'all' (alle Benutzer), 'group' (Mitglieder von Gruppen, rekursiv)
   oder 'users' (Liste spezifischer Benutzer).

 .PARAMETER Group
   Gruppen für Mode 'group' (CN, sAMAccountName oder DN).

 .PARAMETER GroupMatchAll
   Nur Benutzer, die in ALLEN angegebenen Gruppen sind (Standard: in
   mindestens einer).

 .PARAMETER User
   Benutzer-Identifikatoren für Mode 'users'
   (sAMAccountName, UPN, E-Mail oder CN).

 .PARAMETER Attributes
   Zu exportierende Spalten: Katalog-Schlüssel (z. B. GivenName, Surname,
   Mail, Enabled) oder beliebige LDAP-Attributnamen. Ohne Angabe gelten die
   Standard-Spalten des Katalogs bzw. die Auswahl aus dem Profil.

 .PARAMETER OutputPath
   Export-Ordner. Standard: Dokumente-Ordner des Benutzers.

 .PARAMETER FileName
   Dateiname ohne/mit Erweiterung; die passenden Endungen (.xlsx/.csv)
   werden automatisch gesetzt.

 .PARAMETER Format
   'xlsx', 'csv' oder 'both'. Standard: xlsx (Headless: csv, wenn
   ImportExcel fehlt, wird mit Warnung auf csv ausgewichen ist nicht der
   Fall - xlsx ohne ImportExcel bricht mit Fehler ab).

 .PARAMETER Credential
   Alternative Anmeldedaten für die LDAP-Verbindung.

 .PARAMETER LogPath
   Protokolldatei. Standard: %APPDATA%\Argus\Logs\Argus_<Datum>.log

 .PARAMETER Force
   Headless: vorhandene Export-Dateien ohne Rückfrage überschreiben.

 .PARAMETER AddTimestamp
   Hängt einen Zeitstempel (_JJJJMMTT-HHMMSS) an den Dateinamen an.

 .PARAMETER RequiredGroup
   Optionales Berechtigungs-Gate: Nur Mitglieder dieser Gruppe (Format
   "DOMAIN\Gruppe") dürfen das Tool ausführen.

 .PARAMETER TimeoutSec
   Zeitlimit (Sekunden) für LDAP-Suchen. 0 = Systemstandard.

 .EXAMPLE
   .\Argus.ps1
   Startet die grafische Oberfläche.

 .EXAMPLE
   .\Argus.ps1 -NoGui -ProfilePath C:\Argus\Profile\hr-export.json -Force
   Führt den im Profil gespeicherten Export ohne Oberfläche aus
   (z. B. für die Aufgabenplanung).

 .EXAMPLE
   .\Argus.ps1 -NoGui -Mode group -Group "VPN-Benutzer" -Format csv `
               -Server dc01.firma.local -UseLdaps -OutputPath C:\Export
   Exportiert alle (rekursiven) Mitglieder der Gruppe als CSV über LDAPS.

 .NOTES
   Verifizierte LDAP-Fallstricke, die das ActiveDirectory-Modul sonst für
   einen versteckt und die hier explizit behandelt werden:

   1) PageSize-Falle: DirectorySearcher.FindAll() liefert OHNE gesetztes
      PageSize nur die ersten 1000 Treffer zurück - ohne Fehler.
      -> PageSize = 1000 erzwingt Server-Side-Paging.
   2) FILETIME-"Never"-Falle: accountExpires/pwdLastSet/lastLogonTimestamp
      sind Int64. 0 und 9223372036854775807 bedeuten "nie/ungesetzt";
      FromFileTime() wirft dann bzw. liefert Jahr 1601/30828.
      -> Sentinel-Guard in Convert-ArgusFileTime.
   3) Rekursive Gruppenmitgliedschaft: memberOf:1.2.840.113556.1.4.1941:=<DN>
      (LDAP_MATCHING_RULE_IN_CHAIN) - der DN wird dabei RFC-4515-escaped,
      sonst brechen Klammern im Gruppennamen den Filter.
   4) Primäre Gruppe: memberOf enthält die primäre Gruppe NICHT
      (primaryGroupID-Falle, betrifft z. B. "Domänen-Benutzer").
      -> Gruppenfilter und -spalten prüfen zusätzlich (primaryGroupID=<RID>).
   5) Deaktivierte Konten: userAccountControl:1.2.840.113556.1.4.803:=2
      (Bit-AND-Matching-Rule).
   6) LDAP-Attributnamen != AD-Modul-Property-Namen (z. B. l=City, sn=Surname,
      physicalDeliveryOfficeName=Office, co=Land-Name, c=Land-ISO).
   7) Pipeline-Falle: PowerShell "entrollt" zurückgegebene Collections;
      ein leeres HashSet würde zu $null - daher "return ,$set".
   8) lastLogonTimestamp wird nur alle 9-14 Tage repliziert und ist damit
      bewusst "ungefähr" (dokumentiert, Spaltenname weist darauf hin).
   9) memberOf-Ranged-Retrieval: bei sehr vielen Gruppen (>1500) liefert der
      Server memberOf;range=0-1499 - wird erkannt und nachgeladen.

   Sicherheit / Betrieb:
     - Auf Port 389 wird Signing+Sealing ausgehandelt (kompatibel mit
       "LDAP signing required"-Härtung), LDAPS über 636/3269.
     - Exportierte Zellen werden gegen Formel-Injektion geschützt
       (führendes '=' wird neutralisiert; strenger CSV-Schutz optional).
     - Excel-Export ohne Zahlen-Autokonvertierung (führende Nullen in
       PLZ/EmployeeID bleiben erhalten).
     - Vor produktivem Einsatz Skript signieren (AllSigned/WDAC).

   Owner:   Semih Danyeri
   Kontakt: Microsoft Teams -> "Semih Danyeri"
#>
#Requires -Version 5.1
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Bewusste Konsolenausgabe im Headless-Modus')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Interne Hilfsfunktionen; Plural ist hier fachlich korrekt (Settings, Targets, Rows, ...)')]
[CmdletBinding()]
param(
    [switch]$NoGui,
    [string]$ProfilePath,
    [string]$Server,
    [ValidateRange(0, 65535)][int]$Port = 0,
    [switch]$UseLdaps,
    [string]$BaseDn,
    [ValidateSet('', 'all', 'group', 'users')][string]$Mode = '',
    [string[]]$Group,
    [switch]$GroupMatchAll,
    [string[]]$User,
    [string[]]$Attributes,
    [string]$OutputPath,
    [string]$FileName,
    [ValidateSet('', 'xlsx', 'csv', 'both')][string]$Format = '',
    [System.Management.Automation.PSCredential]$Credential,
    [string]$LogPath,
    [switch]$Force,
    [switch]$AddTimestamp,
    [string]$RequiredGroup,
    [ValidateRange(0, 86400)][int]$TimeoutSec = 120
)

$script:ArgusVersion = '2.0.0'
$script:ArgusName    = 'Argus'

# --- Anwendungsverzeichnisse ------------------------------------------------
$script:AppDataRoot = if ($env:APPDATA) { Join-Path $env:APPDATA 'Argus' } else { Join-Path ([IO.Path]::GetTempPath()) 'Argus' }
$script:LogDir      = Join-Path $script:AppDataRoot 'Logs'
$script:LastSettingsPath = Join-Path $script:AppDataRoot 'last-settings.json'
$script:DefaultExportDir = try { [Environment]::GetFolderPath('MyDocuments') } catch { 'C:\Temp' }
if (-not $script:DefaultExportDir) { $script:DefaultExportDir = 'C:\Temp' }

# Effektive Protokolldatei (eine Datei pro Tag; -LogPath überschreibt)
$script:LogFilePath = if ($LogPath) { $LogPath } else { Join-Path $script:LogDir ("Argus_{0}.log" -f (Get-Date -Format 'yyyy-MM-dd')) }

# --- Optionales Berechtigungs-Gate ------------------------------------------
# Kann alternativ zu -RequiredGroup hier fest hinterlegt werden ('DOMAIN\Gruppe').
$script:DefaultRequiredGroup = ''
$effectiveRequiredGroup = if ($RequiredGroup) { $RequiredGroup } else { $script:DefaultRequiredGroup }
if ($effectiveRequiredGroup) {
    $identity  = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    $inRole = $false
    try { $inRole = $principal.IsInRole($effectiveRequiredGroup) } catch { $inRole = $false }
    if (-not $inRole) {
        $msg = "Keine Berechtigung. Erforderliche Gruppe: $effectiveRequiredGroup`nKontakt: Teams -> Semih Danyeri"
        if ($NoGui) { Write-Error $msg; exit 1 }
        Add-Type -AssemblyName PresentationFramework
        [System.Windows.MessageBox]::Show($msg, "$script:ArgusName - Zugriff verweigert", 'OK', 'Warning') | Out-Null
        return
    }
}

# --- STA-Guard für WPF ------------------------------------------------------
if (-not $NoGui -and [Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    if ($PSVersionTable.PSEdition -eq 'Desktop' -and $PSCommandPath) {
        Start-Process -FilePath 'powershell.exe' -ArgumentList @('-STA', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath))
        return
    }
    Write-Error 'Die grafische Oberfläche benötigt einen STA-Thread. Bitte mit "powershell.exe -STA" starten oder -NoGui verwenden.'
    return
}

# --- Assemblies -------------------------------------------------------------
Add-Type -AssemblyName System.DirectoryServices
if (-not $NoGui) {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    Add-Type -AssemblyName System.Windows.Forms
}

# --- Attribut-Katalog (Single Source of Truth) ------------------------------
# Type: string | uac | date | datedirect | manager | memberof | guid | sid
#   string     -> Wert direkt aus LDAP-Property (Mehrfachwerte mit '; ' verbunden)
#   uac        -> Bool aus userAccountControl-Bit (UacBit; UacInvert=true kehrt um)
#   date       -> Int64 FILETIME -> DateTime (mit Never-Sentinel-Guard)
#   datedirect -> bereits DateTime (whenCreated/whenChanged via ADSI)
#   manager    -> DN -> CN extrahieren
#   memberof   -> Gruppen-CNs sortiert verbinden (+ primäre Gruppe)
#   guid/sid   -> Binärwert -> lesbare Zeichenkette
$script:AttributeCatalog = @(
    @{ Key='GivenName';       Header='Vorname';                 Group='Identität';    Ldap='givenName';                  Type='string';     Default=$true  }
    @{ Key='Surname';         Header='Nachname';                Group='Identität';    Ldap='sn';                         Type='string';     Default=$true  }
    @{ Key='DisplayName';     Header='Anzeigename';             Group='Identität';    Ldap='displayName';                Type='string';     Default=$false }
    @{ Key='CN';              Header='Name (CN)';               Group='Identität';    Ldap='cn';                         Type='string';     Default=$false }
    @{ Key='SamAccountName';  Header='SamAccountName';          Group='Identität';    Ldap='sAMAccountName';             Type='string';     Default=$true  }
    @{ Key='UPN';             Header='UserPrincipalName';       Group='Identität';    Ldap='userPrincipalName';          Type='string';     Default=$false }
    @{ Key='Mail';            Header='E-Mail';                  Group='Identität';    Ldap='mail';                       Type='string';     Default=$true  }
    @{ Key='ProxyAddresses';  Header='Proxy-Adressen';          Group='Identität';    Ldap='proxyAddresses';             Type='string';     Default=$false; Tip='Alle SMTP-Aliasse, mit ; getrennt.' }
    @{ Key='EmployeeID';      Header='EmployeeID';              Group='Identität';    Ldap='employeeID';                 Type='string';     Default=$true  }
    @{ Key='EmployeeNumber';  Header='EmployeeNumber';          Group='Identität';    Ldap='employeeNumber';             Type='string';     Default=$false }
    @{ Key='EmployeeType';    Header='Mitarbeitertyp';          Group='Identität';    Ldap='employeeType';               Type='string';     Default=$false }

    @{ Key='Department';      Header='Abteilung';               Group='Organisation'; Ldap='department';                 Type='string';     Default=$false }
    @{ Key='Title';           Header='Titel';                   Group='Organisation'; Ldap='title';                      Type='string';     Default=$false }
    @{ Key='Company';         Header='Firma';                   Group='Organisation'; Ldap='company';                    Type='string';     Default=$false }
    @{ Key='Manager';         Header='Manager (Name)';          Group='Organisation'; Ldap='manager';                    Type='manager';    Default=$false; Tip='Zeigt den CN des Managers. Filter erwartet einen Distinguished Name.' }
    @{ Key='Office';          Header='Büro';                    Group='Organisation'; Ldap='physicalDeliveryOfficeName'; Type='string';     Default=$false }
    @{ Key='Description';     Header='Beschreibung';            Group='Organisation'; Ldap='description';                Type='string';     Default=$false }

    @{ Key='OfficePhone';     Header='Telefon';                 Group='Kontakt';      Ldap='telephoneNumber';            Type='string';     Default=$false }
    @{ Key='MobilePhone';     Header='Mobil';                   Group='Kontakt';      Ldap='mobile';                     Type='string';     Default=$false }

    @{ Key='Street';          Header='Straße';                  Group='Adresse';      Ldap='streetAddress';              Type='string';     Default=$true  }
    @{ Key='City';            Header='Stadt';                   Group='Adresse';      Ldap='l';                          Type='string';     Default=$true  }
    @{ Key='PostCode';        Header='PLZ';                     Group='Adresse';      Ldap='postalCode';                 Type='string';     Default=$true  }
    @{ Key='State';           Header='Bundesland';              Group='Adresse';      Ldap='st';                         Type='string';     Default=$false }
    @{ Key='CountryName';     Header='Land';                    Group='Adresse';      Ldap='co';                         Type='string';     Default=$true  }
    @{ Key='CountryISO';      Header='Land (ISO)';              Group='Adresse';      Ldap='c';                          Type='string';     Default=$false }

    @{ Key='Enabled';         Header='Aktiviert';               Group='Konto/Status'; Ldap='userAccountControl';         Type='uac';        Default=$false; UacBit=2;     UacInvert=$true  }
    @{ Key='PwdNeverExpires'; Header='Passwort läuft nie ab';   Group='Konto/Status'; Ldap='userAccountControl';         Type='uac';        Default=$false; UacBit=65536; UacInvert=$false }
    @{ Key='LastLogonDate';   Header='Letzte Anmeldung (repl.)';Group='Konto/Status'; Ldap='lastLogonTimestamp';         Type='date';       Default=$false; Tip='lastLogonTimestamp wird nur alle 9-14 Tage repliziert - Wert ist bewusst ungefähr.' }
    @{ Key='PwdLastSet';      Header='Passwort gesetzt am';     Group='Konto/Status'; Ldap='pwdLastSet';                 Type='date';       Default=$false }
    @{ Key='Created';         Header='Erstellt';                Group='Konto/Status'; Ldap='whenCreated';                Type='datedirect'; Default=$false }
    @{ Key='Changed';         Header='Geändert';                Group='Konto/Status'; Ldap='whenChanged';                Type='datedirect'; Default=$false }
    @{ Key='AcctExpires';     Header='Konto läuft ab';          Group='Konto/Status'; Ldap='accountExpires';             Type='date';       Default=$false; Tip='Leere Zelle = läuft nie ab (oder nicht gesetzt).' }
    @{ Key='MemberOf';        Header='Gruppen (inkl. primärer)';Group='Konto/Status'; Ldap='memberOf';                   Type='memberof';   Default=$false; Tip='Direkte Gruppen plus primäre Gruppe (z. B. Domänen-Benutzer). Bei >1500 Gruppen wird automatisch nachgeladen.' }
    @{ Key='DN';              Header='Distinguished Name';      Group='Konto/Status'; Ldap='distinguishedName';          Type='string';     Default=$false }
    @{ Key='Guid';            Header='Objekt-GUID';             Group='Konto/Status'; Ldap='objectGUID';                 Type='guid';       Default=$false }
    @{ Key='Sid';             Header='SID';                     Group='Konto/Status'; Ldap='objectSid';                  Type='sid';        Default=$false }
)

# Operatoren für den Filter-Builder (Label -> interner Code)
$script:FilterOps = [ordered]@{
    'ist gleich'                    = 'equals'
    'ist nicht gleich'              = 'notequals'
    'enthält'                       = 'contains'
    'beginnt mit'                   = 'starts'
    'endet mit'                     = 'ends'
    'ist vorhanden'                 = 'present'
    'ist leer'                      = 'notpresent'
    'größer/gleich (>=)'            = 'ge'
    'kleiner/gleich (<=)'           = 'le'
    'innerhalb der letzten N Tage'  = 'lastdays'
    'älter als N Tage (inkl. nie)'  = 'olderdays'
    'RAW (LDAP-Wert)'               = 'raw'
}

# ============================================================================
#  ENGINE - gemeinsame Kernlogik für GUI-Worker (Runspace) und Headless-Modus
#  Wird per Dot-Sourcing in die Hauptsession geladen UND als Text in die
#  Worker-Runspaces injiziert. Darf daher NUR auf eigene Parameter und die
#  dokumentierten Kontext-Callbacks ($Ctx) zugreifen.
# ============================================================================
$script:ArgusEngine = {

    # RFC-4515-Escaping für Filter-WERTE. AllowWildcard=$false escaped auch '*'.
    function Convert-ArgusLdapValue {
        param([string]$Value, [bool]$AllowWildcard = $false)
        if ($null -eq $Value) { return '' }
        $sb = New-Object System.Text.StringBuilder
        foreach ($ch in $Value.ToCharArray()) {
            switch ($ch) {
                '\'  { [void]$sb.Append('\5c') }
                '('  { [void]$sb.Append('\28') }
                ')'  { [void]$sb.Append('\29') }
                "`0" { [void]$sb.Append('\00') }
                '*'  { if ($AllowWildcard) { [void]$sb.Append('*') } else { [void]$sb.Append('\2a') } }
                default { [void]$sb.Append($ch) }
            }
        }
        $sb.ToString()
    }

    # RFC-4515-Escaping für DNs, die als Filterwert eingebettet werden
    # (z. B. memberOf:...:=<DN>). Klammern in Gruppennamen brechen sonst den
    # Filter, Backslash-Escapes im DN würden als Hex-Escapes fehlgedeutet.
    function Convert-ArgusDnForFilter {
        param([string]$Dn)
        if ($null -eq $Dn) { return '' }
        $Dn.Replace('\', '\5c').Replace('(', '\28').Replace(')', '\29').Replace("`0", '\00')
    }

    # '/' hat in ADsPath Sonderbedeutung und muss im DN escaped werden.
    function Convert-ArgusDnForPath {
        param([string]$Dn)
        if ($null -eq $Dn) { return '' }
        $Dn.Replace('/', '\/')
    }

    function Get-ArgusAuthType {
        param([bool]$UseSsl)
        $t = [System.DirectoryServices.AuthenticationTypes]::Secure
        if ($UseSsl) {
            $t = $t -bor [System.DirectoryServices.AuthenticationTypes]::SecureSocketsLayer
        } else {
            # Kompatibel mit DC-Härtung "LDAP signing required" (2020+)
            $t = $t -bor [System.DirectoryServices.AuthenticationTypes]::Signing -bor [System.DirectoryServices.AuthenticationTypes]::Sealing
        }
        return $t
    }

    # Zentraler DirectoryEntry-Konstruktor. $Password ist SecureString oder $null;
    # der Klartext existiert nur für den Konstruktor-Aufruf und wird sofort genullt.
    function New-ArgusDirectoryEntry {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
        param([string]$Path, [string]$Username, [securestring]$Password, $AuthType)
        if ($Username) {
            $bstr = [IntPtr]::Zero
            try {
                $plain = $null
                if ($Password -is [securestring]) {
                    $bstr  = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
                    $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
                }
                return New-Object System.DirectoryServices.DirectoryEntry($Path, $Username, $plain, $AuthType)
            }
            finally {
                if ($bstr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
            }
        }
        return New-Object System.DirectoryServices.DirectoryEntry($Path, $null, $null, $AuthType)
    }

    # Baut den ADsPath-Präfix ("LDAP://server:port/" bzw. "LDAP://" serverlos).
    function Get-ArgusPathPrefix {
        param([hashtable]$Conn)
        $prefix = 'LDAP://'
        if ($Conn.Server) {
            $prefix += $Conn.Server
            if ($Conn.Port -and $Conn.Port -ne 389) { $prefix += ':' + $Conn.Port }
            $prefix += '/'
        }
        return $prefix
    }

    # Erzeugt die Such-Wurzel(n). Liefert ein Hashtable mit:
    #   Search   = DirectoryEntry der Such-Basis (BaseDn)
    #   Domain   = DirectoryEntry der Domänen-/Forest-Wurzel (für Gruppenauflösung,
    #              damit Gruppen außerhalb der Base-OU gefunden werden)
    #   BaseDn, DomainDn, DnsHost, SchemaDn, IsGc, PathPrefix, Username, Password, Auth
    function New-ArgusRoot {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
        param([hashtable]$Conn, [hashtable]$Ctx)
        $useSsl = [bool]$Conn.UseLdaps -or $Conn.Port -eq 636 -or $Conn.Port -eq 3269
        $isGc   = ($Conn.Port -eq 3268 -or $Conn.Port -eq 3269)
        $auth   = Get-ArgusAuthType $useSsl
        $prefix = Get-ArgusPathPrefix $Conn
        if (-not $Conn.Server -and $Conn.Port -and $Conn.Port -ne 389) {
            & $Ctx.Log "WARN: Port $($Conn.Port) ohne Server-Angabe wird ignoriert (serverlose Bindung nutzt Port 389)."
        }

        $rootDse = New-ArgusDirectoryEntry ($prefix + 'RootDSE') $Conn.Username $Conn.Password $auth
        $defaultNc = $null; $rootNc = $null; $dnsHost = $null; $schemaNc = $null
        try {
            $null = $rootDse.NativeObject   # erzwingt echte Bindung -> wirft früh
            $defaultNc = [string]$rootDse.Properties['defaultNamingContext'].Value
            $rootNc    = [string]$rootDse.Properties['rootDomainNamingContext'].Value
            $dnsHost   = [string]$rootDse.Properties['dnsHostName'].Value
            $schemaNc  = [string]$rootDse.Properties['schemaNamingContext'].Value
        }
        finally {
            try { $rootDse.Dispose() } catch { $null = $_ }
        }

        $baseDn = $Conn.BaseDn
        if ([string]::IsNullOrWhiteSpace($baseDn)) {
            $baseDn = if ($isGc -and $rootNc) { $rootNc } else { $defaultNc }
        }
        if (-not $baseDn) { throw 'defaultNamingContext konnte nicht ermittelt werden (Domäne erreichbar? Anmeldedaten korrekt?).' }
        $domainDn = if ($isGc -and $rootNc) { $rootNc } else { $defaultNc }
        if (-not $domainDn) { $domainDn = $baseDn }

        $searchEntry = New-ArgusDirectoryEntry ($prefix + (Convert-ArgusDnForPath $baseDn)) $Conn.Username $Conn.Password $auth
        try { $null = $searchEntry.NativeObject }
        catch {
            try { $searchEntry.Dispose() } catch { $null = $_ }
            throw
        }

        $domainEntry = $searchEntry
        $sameEntry   = $true
        if ($domainDn -ne $baseDn) {
            $domainEntry = New-ArgusDirectoryEntry ($prefix + (Convert-ArgusDnForPath $domainDn)) $Conn.Username $Conn.Password $auth
            $sameEntry = $false
        }

        return @{
            Search = $searchEntry; Domain = $domainEntry; SameEntry = $sameEntry
            BaseDn = $baseDn; DomainDn = $domainDn; DnsHost = $dnsHost; SchemaDn = $schemaNc
            IsGc = $isGc; UseSsl = $useSsl; PathPrefix = $prefix
            Username = $Conn.Username; Password = $Conn.Password; Auth = $auth
            TimeoutSec = [int]$Conn.TimeoutSec
        }
    }

    function Close-ArgusRoot {
        param($Root)
        if (-not $Root) { return }
        try { $Root.Search.Dispose() } catch { $null = $_ }
        if (-not $Root.SameEntry) { try { $Root.Domain.Dispose() } catch { $null = $_ } }
    }

    function New-ArgusBoundEntry {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
        param($Root, [string]$Dn)
        $entry = New-ArgusDirectoryEntry ($Root.PathPrefix + (Convert-ArgusDnForPath $Dn)) $Root.Username $Root.Password $Root.Auth
        try { $null = $entry.NativeObject }
        catch {
            try { $entry.Dispose() } catch { $null = $_ }
            throw
        }
        return $entry
    }

    function New-ArgusSearcher {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
        param($RootEntry, [string]$Filter, [string[]]$Load, [int]$TimeoutSec = 0)
        $ds = New-Object System.DirectoryServices.DirectorySearcher($RootEntry)
        $ds.Filter      = $Filter
        $ds.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
        $ds.PageSize    = 1000      # KRITISCH: ohne Paging max. 1000 Treffer
        $ds.SizeLimit   = 0
        # Referral-Chasing aus: schneller und deterministisch; Cross-Domain
        # deckt der Global-Catalog-Modus (Port 3268/3269) ab.
        $ds.ReferralChasing = [System.DirectoryServices.ReferralChasingOption]::None
        if ($TimeoutSec -gt 0) {
            $ds.ClientTimeout   = [TimeSpan]::FromSeconds($TimeoutSec)
            $ds.ServerTimeLimit = [TimeSpan]::FromSeconds($TimeoutSec)
        }
        $ds.PropertiesToLoad.Clear()
        foreach ($p in $Load) { [void]$ds.PropertiesToLoad.Add($p) }
        return $ds
    }

    # Property-Wert lesen; Mehrfachwerte werden sortiert mit '; ' verbunden,
    # Binärwerte lesbar markiert (statt sie stillschweigend abzuschneiden).
    function Get-ArgusPropValue {
        param($Props, [string]$Name)
        if (-not $Props.Contains($Name) -or $Props[$Name].Count -eq 0) { return $null }
        if ($Props[$Name].Count -eq 1) {
            $v = $Props[$Name][0]
            if ($v -is [byte[]]) { return ('[Binärdaten, {0} Bytes]' -f $v.Length) }
            return $v
        }
        $vals = foreach ($v in $Props[$Name]) {
            if ($v -is [byte[]]) { '[Binärdaten]' } else { [string]$v }
        }
        return (($vals | Sort-Object) -join '; ')
    }

    # Erstes Element roh (für Int64-FILETIME, byte[], primaryGroupID etc.).
    function Get-ArgusPropRaw {
        param($Props, [string]$Name)
        if ($Props.Contains($Name) -and $Props[$Name].Count -gt 0) { return $Props[$Name][0] }
        return $null
    }

    # FILETIME Int64 -> lokale DateTime, mit Never-Sentinel-Guard (0 / Int64::MaxValue).
    function Convert-ArgusFileTime {
        param($Raw)
        if ($null -eq $Raw) { return $null }
        try { $i = [int64]$Raw } catch { return $null }
        if ($i -le 0 -or $i -ge 9223372036854775807) { return $null }
        try { return [DateTime]::FromFileTime($i) } catch { return $null }
    }

    # CN aus einem DN extrahieren. Case-insensitiv, mit Auflösung von
    # \XX-Hex-Escapes und \<Zeichen>-Escapes in EINEM Durchgang.
    function Get-ArgusCnFromDn {
        param([string]$Dn)
        if ([string]::IsNullOrWhiteSpace($Dn)) { return $null }
        $m = [regex]::Match($Dn, '(?i)^\s*CN=(?<cn>(?:\\.|[^,])*)')
        if (-not $m.Success) { return $Dn }
        return [regex]::Replace($m.Groups['cn'].Value, '\\([0-9A-Fa-f]{2})|\\(.)', {
            param($mm)
            if ($mm.Groups[1].Success) { [string][char][Convert]::ToInt32($mm.Groups[1].Value, 16) }
            else { $mm.Groups[2].Value }
        })
    }

    function Get-ArgusGuidString {
        param($Raw)
        if ($null -eq $Raw) { return $null }
        try { return (New-Object Guid (, [byte[]]$Raw)).ToString() } catch { return $null }
    }

    function Get-ArgusSidString {
        param($Raw)
        if ($null -eq $Raw) { return $null }
        try { return (New-Object System.Security.Principal.SecurityIdentifier([byte[]]$Raw, 0)).Value } catch { return $null }
    }

    # Baut eine einzelne LDAP-Filterklausel - typgerecht. Wirft bei
    # Operator/Typ-Kombinationen, die stillschweigend leere Ergebnisse liefern
    # würden (Datum als Text, Substring auf DN-Attributen etc.).
    function New-ArgusClause {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
        param([string]$Ldap, [string]$Op, [string]$Value, [string]$Type, [int]$UacBit = 0, [bool]$UacInvert = $false)

        if ($Type -eq 'uac') {
            if ($Op -notin @('equals', 'notequals')) {
                throw "Feld '$Ldap' (Statusfeld): Operator wird nicht unterstützt - bitte 'ist gleich' oder 'ist nicht gleich' mit wahr/falsch verwenden."
            }
            $truthy = @('true', 'wahr', '1', 'ja', 'yes', 'aktiv', 'enabled', 'an')
            $want = $truthy -contains ([string]$Value).Trim().ToLower()
            if ($Op -eq 'notequals') { $want = -not $want }
            $bitSet = $want
            if ($UacInvert) { $bitSet = -not $want }
            if ($bitSet) { return "(userAccountControl:1.2.840.113556.1.4.803:=$UacBit)" }
            else         { return "(!(userAccountControl:1.2.840.113556.1.4.803:=$UacBit))" }
        }

        if ($Op -eq 'lastdays' -or $Op -eq 'olderdays') {
            if ($Type -notin @('date', 'datedirect')) {
                throw "Operator 'N Tage' ist nur für Datumsfelder möglich (Feld: $Ldap)."
            }
            $days = 0
            if (-not [int]::TryParse(([string]$Value).Trim(), [ref]$days) -or $days -lt 0) {
                throw "Bitte eine Anzahl Tage (>= 0) angeben (Feld: $Ldap)."
            }
            $cut = (Get-Date).AddDays(-$days)
            $lit = if ($Type -eq 'datedirect') { $cut.ToUniversalTime().ToString('yyyyMMddHHmmss') + '.0Z' }
                   else { [string]$cut.ToFileTimeUtc() }
            if ($Op -eq 'lastdays') { return "($Ldap>=$lit)" }
            # "älter als" inkl. nie gesetzt: NOT(>=Cutoff) trifft auch fehlende Werte
            return "(!($Ldap>=$lit))"
        }

        if ($Type -in @('date', 'datedirect')) {
            if ($Op -eq 'present')    { return "($Ldap=*)" }
            if ($Op -eq 'notpresent') { return "(!($Ldap=*))" }
            throw "Datumsfeld '$Ldap': bitte 'innerhalb der letzten N Tage', 'älter als N Tage', 'ist vorhanden' oder 'ist leer' verwenden."
        }

        if ($Type -in @('manager', 'memberof')) {
            if ($Op -eq 'present')    { return "($Ldap=*)" }
            if ($Op -eq 'notpresent') { return "(!($Ldap=*))" }
            if ($Op -in @('equals', 'notequals')) {
                if ($Value -notmatch '(?i)^\s*CN=') {
                    throw "Feld '$Ldap' erwartet einen Distinguished Name (CN=...,OU=...,DC=...). Substring-Suche wird von AD auf DN-Attributen nicht unterstützt."
                }
                $e = Convert-ArgusDnForFilter $Value.Trim()
                if ($Op -eq 'equals') { return "($Ldap=$e)" }
                return "(!($Ldap=$e))"
            }
            throw "Feld '$Ldap' (DN-Attribut): AD unterstützt hier nur 'ist gleich', 'ist nicht gleich', 'ist vorhanden', 'ist leer'."
        }

        if ($Type -in @('guid', 'sid')) {
            throw "Auf '$Ldap' kann nicht gefiltert werden (Binärattribut)."
        }

        $e = Convert-ArgusLdapValue $Value $false
        switch ($Op) {
            'equals'     { "($Ldap=$e)" }
            'notequals'  { "(!($Ldap=$e))" }
            'contains'   { "($Ldap=*$e*)" }
            'starts'     { "($Ldap=$e*)" }
            'ends'       { "($Ldap=*$e)" }
            'present'    { "($Ldap=*)" }
            'notpresent' { "(!($Ldap=*))" }
            'ge'         { "($Ldap>=$e)" }
            'le'         { "($Ldap<=$e)" }
            'raw'        { "($Ldap=$Value)" }   # bewusst NICHT escaped - Power-User
            default      { "($Ldap=$e)" }
        }
    }

    # Gruppe auflösen -> @{ Dn; Rid; Name } oder $null.
    # Wirft bei mehrdeutigen Namen (statt stillschweigend die "erste" zu nehmen).
    function Resolve-ArgusGroup {
        param($Root, [string]$NameOrDn, [hashtable]$Ctx)
        $dn = $null; $sidBytes = $null; $name = $NameOrDn

        if ($NameOrDn -match '(?i)^\s*(CN|OU|DC)=') {
            # Bereits ein DN -> direkt binden, um SID (für primaryGroupID) zu holen
            $entry = $null
            try {
                $entry = New-ArgusBoundEntry $Root $NameOrDn.Trim()
                $dn       = [string]$entry.Properties['distinguishedName'].Value
                $sidBytes = $entry.Properties['objectSid'].Value
                $name     = [string]$entry.Properties['name'].Value
            }
            catch {
                & $Ctx.Log "WARN: DN konnte nicht gebunden werden: $NameOrDn ($($_.Exception.Message))"
                return $null
            }
            finally {
                if ($entry) { try { $entry.Dispose() } catch { $null = $_ } }
            }
        }
        else {
            $e  = Convert-ArgusLdapValue $NameOrDn $false
            $ds = New-ArgusSearcher $Root.Domain "(&(objectClass=group)(|(cn=$e)(sAMAccountName=$e)(name=$e)))" @('distinguishedName', 'objectSid', 'name') $Root.TimeoutSec
            $ds.SizeLimit = 2
            $res = $null
            try {
                $res = $ds.FindAll()
                if ($res.Count -gt 1) {
                    $hits = @(foreach ($r in $res) { [string]$r.Properties['distinguishedname'][0] })
                    throw "Gruppenname '$NameOrDn' ist mehrdeutig ($($hits -join ' | ')). Bitte den vollständigen DN angeben."
                }
                if ($res.Count -eq 1) {
                    $r  = $res[0]
                    $dn = [string]$r.Properties['distinguishedname'][0]
                    $sidBytes = Get-ArgusPropRaw $r.Properties 'objectsid'
                    if ($r.Properties.Contains('name') -and $r.Properties['name'].Count -gt 0) { $name = [string]$r.Properties['name'][0] }
                }
            }
            finally {
                if ($res) { try { $res.Dispose() } catch { $null = $_ } }
                try { $ds.Dispose() } catch { $null = $_ }
            }
            if (-not $dn) { return $null }
        }

        $rid = $null
        $sidStr = Get-ArgusSidString $sidBytes
        if ($sidStr) { $rid = [int]($sidStr.Split('-')[-1]) }
        return @{ Dn = $dn; Rid = $rid; Name = $name }
    }

    # Filterklausel "Mitglied (rekursiv) von Gruppe" inkl. primärer Gruppe.
    function Get-ArgusGroupClause {
        param([hashtable]$GroupInfo)
        $dnEsc = Convert-ArgusDnForFilter $GroupInfo.Dn
        $chain = "(memberOf:1.2.840.113556.1.4.1941:=$dnEsc)"
        if ($GroupInfo.Rid) { return "(|$chain(primaryGroupID=$($GroupInfo.Rid)))" }
        return $chain
    }

    # Rekursive Mitglieder (User) einer Gruppe als GUID-Set - inkl. primärer
    # Gruppe. WICHTIG: ",$set" verhindert das Pipeline-Entrollen; ein leeres
    # HashSet käme sonst als $null zurück und ein einzelnes Element als String.
    function Get-ArgusGroupMemberGuids {
        param($Root, [hashtable]$GroupInfo, [hashtable]$Ctx)
        $set = New-Object 'System.Collections.Generic.HashSet[string]'
        $f   = '(&(objectCategory=person)(objectClass=user)' + (Get-ArgusGroupClause $GroupInfo) + ')'
        $ds  = New-ArgusSearcher $Root.Domain $f @('objectGUID') $Root.TimeoutSec
        $res = $null
        try {
            $res = $ds.FindAll()
            foreach ($r in $res) {
                if (& $Ctx.IsCancelled) { throw (New-Object System.OperationCanceledException 'Abgebrochen durch Benutzer.') }
                $g = Get-ArgusGuidString (Get-ArgusPropRaw $r.Properties 'objectguid')
                if ($g) { [void]$set.Add($g) }
            }
        }
        finally {
            if ($res) { try { $res.Dispose() } catch { $null = $_ } }
            try { $ds.Dispose() } catch { $null = $_ }
        }
        return , $set
    }

    # Prüft, ob ein LDAP-Attribut im Schema existiert (Warnung bei Tippfehlern
    # in benutzerdefinierten Attributen - LDAP selbst meldet keinen Fehler).
    function Test-ArgusAttributeExists {
        param($Root, [string]$LdapName)
        if (-not $Root.SchemaDn) { return $true }   # nicht prüfbar -> nicht blockieren
        $entry = $null; $ds = $null; $res = $null
        try {
            $entry = New-ArgusDirectoryEntry ($Root.PathPrefix + (Convert-ArgusDnForPath $Root.SchemaDn)) $Root.Username $Root.Password $Root.Auth
            $e  = Convert-ArgusLdapValue $LdapName $false
            $ds = New-ArgusSearcher $entry "(&(objectClass=attributeSchema)(lDAPDisplayName=$e))" @('lDAPDisplayName') $Root.TimeoutSec
            $ds.SizeLimit = 1
            $res = $ds.FindAll()
            return ($res.Count -gt 0)
        }
        catch { return $true }   # Schema nicht lesbar -> nicht blockieren
        finally {
            if ($res)   { try { $res.Dispose() }   catch { $null = $_ } }
            if ($ds)    { try { $ds.Dispose() }    catch { $null = $_ } }
            if ($entry) { try { $entry.Dispose() } catch { $null = $_ } }
        }
    }

    # memberOf vollständig lesen - erkennt Ranged Retrieval (>1500 Gruppen:
    # der Server liefert 'memberof;range=0-1499' statt 'memberof') und lädt
    # die restlichen Bereiche über den DirectoryEntry nach.
    function Get-ArgusMemberOfAll {
        param($Root, $Props, [string]$Dn)
        $result = New-Object System.Collections.ArrayList
        if ($Props.Contains('memberof')) {
            foreach ($v in $Props['memberof']) { [void]$result.Add([string]$v) }
        }
        $rangeKey = $null
        foreach ($pn in $Props.PropertyNames) {
            if ($pn -like 'memberof;range=*') { $rangeKey = $pn; break }
        }
        if (-not $rangeKey) { return , $result }

        foreach ($v in $Props[$rangeKey]) { [void]$result.Add([string]$v) }
        if ($rangeKey -match 'range=\d+-\*$') { return , $result }   # bereits vollständig

        $entry = $null
        try {
            $entry = New-ArgusBoundEntry $Root $Dn
            for ($guard = 0; $guard -lt 100; $guard++) {
                $low  = $result.Count
                $attr = "memberOf;range=$low-*"
                $entry.RefreshCache(@($attr))
                $foundKey = $null
                foreach ($pn in $entry.Properties.PropertyNames) {
                    if ($pn -like 'memberOf;range=*') { $foundKey = $pn; break }
                }
                if (-not $foundKey) { break }
                foreach ($v in $entry.Properties[$foundKey]) { [void]$result.Add([string]$v) }
                if ($foundKey -match 'range=\d+-\*$') { break }   # letzter Block erreicht
            }
        }
        catch { $null = $_ }   # Teilergebnis ist besser als Abbruch
        finally {
            if ($entry) { try { $entry.Dispose() } catch { $null = $_ } }
        }
        return , $result
    }

    function Split-ArgusList {
        param([object[]]$Items, [int]$Size)
        $chunks = New-Object System.Collections.ArrayList
        for ($i = 0; $i -lt $Items.Count; $i += $Size) {
            $end = [Math]::Min($i + $Size - 1, $Items.Count - 1)
            [void]$chunks.Add(@($Items[$i..$end]))
        }
        return , $chunks
    }

    # Zellen-Schutz gegen Formel-Injektion beim Export.
    #   Xlsx: nur führendes '=' wird neutralisiert (Export-Excel interpretiert es als Formel).
    #   CSV:  '=' und Steuerzeichen immer; '+', '-', '@' nur im strengen Modus,
    #         damit z. B. Telefonnummern (+49...) lesbar bleiben.
    function Protect-ArgusCell {
        param([string]$Value, [string]$Target, [bool]$Strict)
        if ([string]::IsNullOrEmpty($Value)) { return $Value }
        $c = $Value[0]
        if ($Target -eq 'xlsx') {
            if ($c -eq '=') { return "'" + $Value }
            return $Value
        }
        if ($c -eq '=' -or $c -eq [char]9 -or $c -eq [char]13 -or $c -eq [char]10) { return "'" + $Value }
        if ($Strict -and ($c -eq '+' -or $c -eq '-' -or $c -eq '@')) { return "'" + $Value }
        return $Value
    }

    function Write-ArgusFileLog {
        param([string]$Path, [string]$Message)
        if (-not $Path) { return }
        try {
            $dir = Split-Path -Parent $Path
            if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
            Add-Content -LiteralPath $Path -Value $Message -Encoding UTF8
        }
        catch { $null = $_ }   # Protokollierung darf den Export nie verhindern
    }

    # ------------------------------------------------------------------------
    # Kern: LDAP-Suche ausführen und DataTable aufbauen.
    # $P    : Parameter-Hashtable (siehe Aufrufer)
    # $Ctx  : @{ Log = {param($m)}; Progress = {param($pct,$status)}; IsCancelled = {[bool]} }
    # Liefert @{ Table; Count; Unmatched; DurationSec; Filter }
    # ------------------------------------------------------------------------
    function Invoke-ArgusSearch {
        param([hashtable]$P, [hashtable]$Ctx)
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $root = $null
        try {
            & $Ctx.Progress 0 'Verbinde via LDAP...'
            if ([string]::IsNullOrWhiteSpace($P.Server)) { & $Ctx.Log 'Kein DC angegeben -> serverlose Bindung (DC-Locator).' }
            else { & $Ctx.Log ("Ziel: {0}{1}" -f $P.Server, $(if ($P.Port -and $P.Port -ne 389) { ":$($P.Port)" } else { '' })) }

            $root = New-ArgusRoot @{
                Server = $P.Server; Port = $P.Port; UseLdaps = $P.UseLdaps; BaseDn = $P.BaseDn
                Username = $P.Username; Password = $P.Password; TimeoutSec = $P.TimeoutSec
            } $Ctx
            $sslTxt = if ($root.UseSsl) { 'LDAPS (SSL/TLS)' } else { 'LDAP mit Signing+Sealing' }
            & $Ctx.Log ("Verbunden: {0} | {1}{2}" -f $(if ($root.DnsHost) { $root.DnsHost } else { 'DC-Locator' }), $sslTxt, $(if ($root.IsGc) { ' | Global Catalog' } else { '' }))
            & $Ctx.Log ("Such-Basis (DN): {0}" -f $root.BaseDn)

            # --- Subjekt-Klauseln je nach Modus (ggf. mehrere Chunks) -----------
            $subjectFilters = @('')
            switch ($P.Mode) {
                'all' {
                    & $Ctx.Log 'Modus: ALLE Benutzer.'
                }
                'group' {
                    & $Ctx.Log "Modus: Mitglieder von Gruppe(n) (rekursiv, inkl. primärer Gruppe)."
                    $clauses = @()
                    foreach ($g in $P.Groups) {
                        $info = Resolve-ArgusGroup $root $g $Ctx
                        if ($info) {
                            $clauses += Get-ArgusGroupClause $info
                            & $Ctx.Log ("  Gruppe aufgelöst: {0} -> {1}" -f $g, $info.Dn)
                        }
                        else { & $Ctx.Log "  WARN: Gruppe NICHT gefunden: $g" }
                    }
                    if (-not $clauses) { throw 'Keine der angegebenen Filter-Gruppen wurde gefunden.' }
                    if ($clauses.Count -eq 1)   { $subjectFilters = @($clauses[0]) }
                    elseif ($P.GroupMatchAny)   { $subjectFilters = @('(|' + (-join $clauses) + ')') }   # in MIND. EINER
                    else                        { $subjectFilters = @('(&' + (-join $clauses) + ')') }   # in ALLEN
                }
                'users' {
                    & $Ctx.Log "Modus: Spezifische Benutzer ($(@($P.Users).Count) IDs)."
                    if (-not $P.Users -or @($P.Users).Count -eq 0) { throw 'Keine Benutzer-IDs angegeben.' }
                    # Chunking: sehr lange OR-Filter können Serverlimits reißen
                    $subjectFilters = @()
                    foreach ($chunk in (Split-ArgusList @($P.Users) 60)) {
                        $ors = foreach ($id in $chunk) {
                            $e = Convert-ArgusLdapValue ([string]$id).Trim() $false
                            "(|(sAMAccountName=$e)(userPrincipalName=$e)(mail=$e)(cn=$e))"
                        }
                        $subjectFilters += $(if (@($ors).Count -eq 1) { @($ors)[0] } else { '(|' + (-join $ors) + ')' })
                    }
                    if ($subjectFilters.Count -gt 1) { & $Ctx.Log ("  Aufgeteilt in {0} Teil-Abfragen (à max. 60 IDs)." -f $subjectFilters.Count) }
                }
            }

            # --- Attribut-Filter (UND/ODER) -------------------------------------
            $attrClauses = @()
            foreach ($f in $P.Filters) {
                $c = New-ArgusClause $f.Ldap $f.Op $f.Val $f.Type $(if ($f.UacBit) { $f.UacBit } else { 0 }) $(if ($f.UacInvert) { $f.UacInvert } else { $false })
                $attrClauses += $c
                & $Ctx.Log "  Filter: $c"
            }
            $attrPart = ''
            if ($attrClauses.Count -eq 1) { $attrPart = $attrClauses[0] }
            elseif ($attrClauses.Count -gt 1) {
                if ($P.FilterJoin -eq 'or') { $attrPart = '(|' + (-join $attrClauses) + ')' }
                else                        { $attrPart = -join $attrClauses }
            }

            # --- RAW-Zusatzfilter (Power-User) ----------------------------------
            $rawPart = ''
            if (-not [string]::IsNullOrWhiteSpace($P.RawFilter)) {
                $raw = $P.RawFilter.Trim()
                if (-not $raw.StartsWith('(')) { $raw = "($raw)" }
                $open = ($raw.ToCharArray() | Where-Object { $_ -eq '(' }).Count
                $close = ($raw.ToCharArray() | Where-Object { $_ -eq ')' }).Count
                if ($open -ne $close) { throw "RAW-LDAP-Zusatzfilter: Klammern sind unausgeglichen ($open öffnende, $close schließende)." }
                $rawPart = $raw
                & $Ctx.Log "  RAW-Zusatzfilter: $raw"
            }

            # --- Zu ladende Properties ------------------------------------------
            $loadSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
            $needPrimaryGroup = $false
            foreach ($a in $P.Attributes) {
                [void]$loadSet.Add($a.Ldap)
                if ($a.Type -eq 'memberof') { $needPrimaryGroup = $true }
            }
            foreach ($n in @('objectGUID', 'distinguishedName')) { [void]$loadSet.Add($n) }
            if ($needPrimaryGroup) { [void]$loadSet.Add('primaryGroupID'); [void]$loadSet.Add('objectSid') }
            if ($P.Mode -eq 'users') {
                foreach ($n in @('sAMAccountName', 'userPrincipalName', 'mail', 'cn')) { [void]$loadSet.Add($n) }
            }
            $load = @($loadSet)
            & $Ctx.Log ('Geladene Attribute: ' + ($load -join ', '))

            # --- Benutzerdefinierte Attribute gegen das Schema prüfen -----------
            foreach ($a in $P.Attributes) {
                if ($a.Custom -and -not (Test-ArgusAttributeExists $root $a.Ldap)) {
                    & $Ctx.Log "WARN: Attribut '$($a.Ldap)' existiert nicht im AD-Schema (Tippfehler?). Spalte bleibt leer."
                }
            }

            # --- Gruppen-Spalten vorbereiten (rekursive Mitglieds-Sets) ---------
            $columnSets = @{}
            $columnOrder = @()
            if ($P.IncludeGroupColumns -and @($P.ColumnGroups).Count -gt 0) {
                & $Ctx.Progress 0 'Lese Gruppen-Mitgliedschaften...'
                foreach ($g in $P.ColumnGroups) {
                    if (& $Ctx.IsCancelled) { throw (New-Object System.OperationCanceledException 'Abgebrochen durch Benutzer.') }
                    $info = Resolve-ArgusGroup $root $g $Ctx
                    if (-not $info) { & $Ctx.Log "WARN: Spalten-Gruppe NICHT gefunden: $g"; $columnOrder += , @($g, $null); continue }
                    $set = Get-ArgusGroupMemberGuids $root $info $Ctx
                    $columnSets[$g] = $set
                    $columnOrder += , @($g, $set)
                    & $Ctx.Log ("  Spalte '{0}': {1} Mitglieder (rekursiv)" -f $g, $set.Count)
                }
            }

            # --- DataTable-Struktur ---------------------------------------------
            $table = New-Object System.Data.DataTable 'ArgusExport'
            $usedHeaders = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
            $colDefs = @()
            foreach ($a in $P.Attributes) {
                if (-not $usedHeaders.Add($a.Header)) { & $Ctx.Log "WARN: Doppelte Spalte übersprungen: $($a.Header)"; continue }
                $t = switch ($a.Type) {
                    'date'       { [datetime] }
                    'datedirect' { [datetime] }
                    'uac'        { [bool] }
                    default      { [string] }
                }
                [void]$table.Columns.Add((New-Object System.Data.DataColumn($a.Header, $t)))
                $colDefs += $a
            }
            $groupColHeaders = @{}
            foreach ($pair in $columnOrder) {
                $hdr = 'Mitglied: {0}' -f $pair[0]
                if (-not $usedHeaders.Add($hdr)) { continue }
                [void]$table.Columns.Add((New-Object System.Data.DataColumn($hdr, [bool])))
                $groupColHeaders[$pair[0]] = $hdr
            }

            # --- Hauptsuche (ggf. mehrere Chunks) --------------------------------
            $core = '(&(objectCategory=person)(objectClass=user)'
            $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
            $matchedIds = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
            $pgCache = @{}   # Primärgruppen-SID -> CN
            $firstFilter = $null
            $chunkCount = $subjectFilters.Count
            $chunkIdx = 0

            foreach ($subject in $subjectFilters) {
                if (& $Ctx.IsCancelled) { throw (New-Object System.OperationCanceledException 'Abgebrochen durch Benutzer.') }
                $filter = $core + $subject + $attrPart + $rawPart + ')'
                if (-not $firstFilter) { $firstFilter = $filter }
                if ($chunkCount -gt 1) { & $Ctx.Log ("LDAP-Filter (Teil {0}/{1}): {2}" -f ($chunkIdx + 1), $chunkCount, $filter) }
                else { & $Ctx.Log "LDAP-Filter: $filter" }

                $ds = New-ArgusSearcher $root.Search $filter $load $root.TimeoutSec
                $res = $null
                try {
                    & $Ctx.Progress ([int](100 * $chunkIdx / $chunkCount)) 'Führe LDAP-Suche aus...'
                    $res = $ds.FindAll()
                    $total = $res.Count
                    & $Ctx.Log "Treffer: $total"
                    $i = 0
                    foreach ($r in $res) {
                        $i++
                        if (($i % 25) -eq 0 -or $i -eq $total) {
                            if (& $Ctx.IsCancelled) { throw (New-Object System.OperationCanceledException 'Abgebrochen durch Benutzer.') }
                            $pct = [int](100 * ($chunkIdx + ($i / [Math]::Max(1, $total))) / $chunkCount)
                            & $Ctx.Progress $pct "Verarbeite $i / $total"
                        }

                        $props = $r.Properties
                        $dn    = [string](Get-ArgusPropRaw $props 'distinguishedname')
                        $guid  = Get-ArgusGuidString (Get-ArgusPropRaw $props 'objectguid')
                        $dedupKey = if ($guid) { $guid } elseif ($dn) { $dn } else { $null }
                        if ($dedupKey -and -not $seen.Add($dedupKey)) { continue }

                        if ($P.Mode -eq 'users') {
                            foreach ($n in @('samaccountname', 'userprincipalname', 'mail', 'cn')) {
                                $v = Get-ArgusPropRaw $props $n
                                if ($v) { [void]$matchedIds.Add(([string]$v).Trim()) }
                            }
                        }

                        $row = $table.NewRow()
                        foreach ($a in $colDefs) {
                            $v = $null
                            switch ($a.Type) {
                                'uac' {
                                    $uac = Get-ArgusPropRaw $props 'useraccountcontrol'
                                    if ($null -ne $uac) {
                                        $bitSet = ([int]$uac -band $a.UacBit) -ne 0
                                        $v = if ($a.UacInvert) { -not $bitSet } else { $bitSet }
                                    }
                                }
                                'date'       { $v = Convert-ArgusFileTime (Get-ArgusPropRaw $props $a.Ldap) }
                                'datedirect' {
                                    $d = Get-ArgusPropRaw $props $a.Ldap
                                    if ($d -is [datetime]) { $v = ([datetime]$d).ToLocalTime() } elseif ($null -ne $d) { $v = [string]$d }
                                }
                                'manager'    { $v = Get-ArgusCnFromDn ([string](Get-ArgusPropRaw $props $a.Ldap)) }
                                'memberof'   {
                                    $dns = Get-ArgusMemberOfAll $root $props $dn
                                    $names = New-Object System.Collections.ArrayList
                                    foreach ($gdn in $dns) {
                                        $cn = Get-ArgusCnFromDn ([string]$gdn)
                                        if ($cn) { [void]$names.Add($cn) }
                                    }
                                    # Primäre Gruppe ergänzen (fehlt in memberOf konstruktionsbedingt)
                                    $pgid = Get-ArgusPropRaw $props 'primarygroupid'
                                    $sidB = Get-ArgusPropRaw $props 'objectsid'
                                    if ($null -ne $pgid -and $null -ne $sidB) {
                                        try {
                                            $userSid = New-Object System.Security.Principal.SecurityIdentifier([byte[]]$sidB, 0)
                                            $pgSid = '{0}-{1}' -f $userSid.AccountDomainSid.Value, [int]$pgid
                                            if (-not $pgCache.ContainsKey($pgSid)) {
                                                $pgCache[$pgSid] = $null
                                                $pds = New-ArgusSearcher $root.Domain "(objectSid=$pgSid)" @('cn') $root.TimeoutSec
                                                $pres = $null
                                                try {
                                                    $pds.SizeLimit = 1
                                                    $pres = $pds.FindAll()
                                                    if ($pres.Count -gt 0) { $pgCache[$pgSid] = [string]$pres[0].Properties['cn'][0] }
                                                }
                                                finally {
                                                    if ($pres) { try { $pres.Dispose() } catch { $null = $_ } }
                                                    try { $pds.Dispose() } catch { $null = $_ }
                                                }
                                            }
                                            if ($pgCache[$pgSid] -and -not ($names -contains $pgCache[$pgSid])) { [void]$names.Add($pgCache[$pgSid]) }
                                        }
                                        catch { $null = $_ }
                                    }
                                    if ($names.Count -gt 0) { $v = (($names.ToArray() | Sort-Object) -join '; ') }
                                }
                                'guid'       { $v = Get-ArgusGuidString (Get-ArgusPropRaw $props $a.Ldap) }
                                'sid'        { $v = Get-ArgusSidString (Get-ArgusPropRaw $props $a.Ldap) }
                                default      {
                                    $v = Get-ArgusPropValue $props $a.Ldap
                                    if ($null -ne $v -and $v -isnot [string]) { $v = [string]$v }
                                }
                            }
                            if ($null -eq $v) { $row[$a.Header] = [DBNull]::Value } else { $row[$a.Header] = $v }
                        }

                        foreach ($g in $groupColHeaders.Keys) {
                            $hdr = $groupColHeaders[$g]
                            $set = $columnSets[$g]
                            if ($null -ne $set -and $guid) { $row[$hdr] = $set.Contains($guid) }
                            else { $row[$hdr] = [DBNull]::Value }   # unbekannt statt fälschlich False
                        }

                        $table.Rows.Add($row)
                    }
                }
                finally {
                    if ($res) { try { $res.Dispose() } catch { $null = $_ } }
                    try { $ds.Dispose() } catch { $null = $_ }
                }
                $chunkIdx++
            }

            # --- Abgleich: nicht gefundene Identifikatoren (Modus 'users') ------
            $unmatched = @()
            if ($P.Mode -eq 'users') {
                foreach ($id in $P.Users) {
                    if (-not $matchedIds.Contains(([string]$id).Trim())) { $unmatched += [string]$id }
                }
                if ($unmatched.Count -gt 0) {
                    & $Ctx.Log ("WARN: {0} Identifikator(en) ohne Treffer: {1}" -f $unmatched.Count, ($unmatched -join ', '))
                }
            }

            # --- Sortierung ------------------------------------------------------
            $sortExpr = $null
            switch ($P.SortBy) {
                'sam'  { if ($table.Columns.Contains('SamAccountName')) { $sortExpr = '[SamAccountName] ASC' } }
                'none' { }
                default {
                    $parts = @()
                    if ($table.Columns.Contains('Nachname')) { $parts += '[Nachname] ASC' }
                    if ($table.Columns.Contains('Vorname'))  { $parts += '[Vorname] ASC' }
                    if ($parts.Count) { $sortExpr = $parts -join ', ' }
                }
            }
            if ($sortExpr -and $table.Rows.Count -gt 0) {
                $dv = New-Object System.Data.DataView($table)
                $dv.Sort = $sortExpr
                $table = $dv.ToTable()
            }

            $sw.Stop()
            & $Ctx.Log ("Suche abgeschlossen: {0} Benutzer in {1:n1} s." -f $table.Rows.Count, $sw.Elapsed.TotalSeconds)
            return @{
                Table = $table; Count = $table.Rows.Count; Unmatched = $unmatched
                DurationSec = [Math]::Round($sw.Elapsed.TotalSeconds, 1); Filter = $firstFilter
                Server = $root.DnsHost; BaseDn = $root.BaseDn
            }
        }
        finally {
            Close-ArgusRoot $root
        }
    }

    # ------------------------------------------------------------------------
    # Export: DataTable -> Excel (ImportExcel) und/oder CSV (ohne Abhängigkeiten)
    # $Opt: Folder, FileName, Xlsx, Csv, CsvDelimiter, StrictCsv, AddTimestamp, Meta
    # Liefert @{ XlsxPath; CsvPath }
    # ------------------------------------------------------------------------
    function Get-ArgusExportTargets {
        param([hashtable]$Opt)
        $base = [IO.Path]::GetFileNameWithoutExtension($Opt.FileName)
        if ([string]::IsNullOrWhiteSpace($base)) { $base = 'Argus-Export' }
        if ($Opt.AddTimestamp) { $base += '_' + (Get-Date -Format 'yyyyMMdd-HHmmss') }
        $r = @{ XlsxPath = $null; CsvPath = $null }
        if ($Opt.Xlsx) { $r.XlsxPath = Join-Path $Opt.Folder ($base + '.xlsx') }
        if ($Opt.Csv)  { $r.CsvPath  = Join-Path $Opt.Folder ($base + '.csv') }
        return $r
    }

    function Test-ArgusFileWritable {
        param([string]$Path)
        if (-not (Test-Path -LiteralPath $Path)) { return $true }
        try {
            $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
            $fs.Dispose()
            return $true
        }
        catch { return $false }
    }

    function Export-ArgusData {
        param([System.Data.DataTable]$Table, [hashtable]$Opt, [hashtable]$Ctx)
        if (-not (Test-Path -LiteralPath $Opt.Folder)) { New-Item -Path $Opt.Folder -ItemType Directory -Force | Out-Null }
        # Vom Aufrufer vorab berechnete Zielpfade haben Vorrang (Zeitstempel-Konsistenz
        # zwischen Überschreib-Rückfrage und tatsächlichem Schreiben).
        $targets = if ($Opt.Targets) { $Opt.Targets } else { Get-ArgusExportTargets $Opt }
        $headers = @(foreach ($c in $Table.Columns) { $c.ColumnName })
        $dateCols = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($c in $Table.Columns) { if ($c.DataType -eq [datetime]) { [void]$dateCols.Add($c.ColumnName) } }

        foreach ($p in @($targets.XlsxPath, $targets.CsvPath)) {
            if ($p -and -not (Test-ArgusFileWritable $p)) {
                throw "Die Datei ist gesperrt (vermutlich in Excel geöffnet): $p"
            }
        }

        if ($targets.CsvPath) {
            & $Ctx.Progress 0 'Schreibe CSV...'
            $delim = if ($Opt.CsvDelimiter) { $Opt.CsvDelimiter } else { ';' }
            $enc = New-Object System.Text.UTF8Encoding($true)   # BOM: deutsches Excel braucht ihn für Umlaute
            $writer = New-Object System.IO.StreamWriter($targets.CsvPath, $false, $enc)
            try {
                $quoteChars = [char[]]@($delim[0], '"', "`n", "`r")
                $line = New-Object System.Text.StringBuilder
                $emit = {
                    param($vals)
                    [void]$line.Clear()
                    for ($ci = 0; $ci -lt $vals.Count; $ci++) {
                        if ($ci -gt 0) { [void]$line.Append($delim) }
                        $s = [string]$vals[$ci]
                        if ($s.IndexOfAny($quoteChars) -ge 0) { $s = '"' + $s.Replace('"', '""') + '"' }
                        [void]$line.Append($s)
                    }
                    $writer.WriteLine($line.ToString())
                }
                & $emit $headers
                $ri = 0
                foreach ($row in $Table.Rows) {
                    $ri++
                    if (($ri % 1000) -eq 0) {
                        if (& $Ctx.IsCancelled) { throw (New-Object System.OperationCanceledException 'Abgebrochen durch Benutzer.') }
                        & $Ctx.Progress ([int](100 * $ri / [Math]::Max(1, $Table.Rows.Count))) "CSV: Zeile $ri / $($Table.Rows.Count)"
                    }
                    $vals = foreach ($h in $headers) {
                        $v = $row[$h]
                        if ($v -is [DBNull]) { '' }
                        elseif ($dateCols.Contains($h)) { ([datetime]$v).ToString('yyyy-MM-dd HH:mm:ss') }
                        elseif ($v -is [bool]) { if ($v) { 'True' } else { 'False' } }
                        else { Protect-ArgusCell ([string]$v) 'csv' ([bool]$Opt.StrictCsv) }
                    }
                    & $emit $vals
                }
            }
            finally {
                $writer.Dispose()
            }
            & $Ctx.Log "CSV geschrieben: $($targets.CsvPath)"
        }

        if ($targets.XlsxPath) {
            & $Ctx.Progress 0 'Bereite Excel-Export vor...'
            Import-Module ImportExcel -ErrorAction Stop
            # Alte Datei entfernen: verhindert Reste alter Zeilen / Tabellen-Kollisionen
            if (Test-Path -LiteralPath $targets.XlsxPath) { Remove-Item -LiteralPath $targets.XlsxPath -Force }

            $rowsOut = New-Object System.Collections.ArrayList
            $ri = 0
            foreach ($row in $Table.Rows) {
                $ri++
                if (($ri % 1000) -eq 0) {
                    if (& $Ctx.IsCancelled) { throw (New-Object System.OperationCanceledException 'Abgebrochen durch Benutzer.') }
                    & $Ctx.Progress ([int](80 * $ri / [Math]::Max(1, $Table.Rows.Count))) "Excel: Zeile $ri / $($Table.Rows.Count)"
                }
                $rec = [ordered]@{}
                foreach ($h in $headers) {
                    $v = $row[$h]
                    if ($v -is [DBNull]) { $rec[$h] = $null }
                    elseif ($v -is [string]) { $rec[$h] = Protect-ArgusCell $v 'xlsx' $false }
                    else { $rec[$h] = $v }
                }
                [void]$rowsOut.Add([PSCustomObject]$rec)
            }

            & $Ctx.Progress 85 'Schreibe Excel...'
            $xlParams = @{
                Path          = $targets.XlsxPath
                WorksheetName = 'Export'
                TableName     = 'ArgusExport'
                TableStyle    = 'Medium2'
                FreezeTopRow  = $true
                # Keine Zahlen-Autokonvertierung: führende Nullen (PLZ 01067,
                # EmployeeID) und lange Nummern bleiben unverändert erhalten.
                NoNumberConversion = '*'
            }
            if ($Table.Rows.Count -le 20000) { $xlParams.AutoSize = $true }
            else { & $Ctx.Log 'Hinweis: AutoSize bei >20000 Zeilen deaktiviert (Performance).' }
            $rowsOut | Export-Excel @xlParams

            if ($Opt.Meta) {
                $info = foreach ($k in $Opt.Meta.Keys) { [PSCustomObject]@{ Eigenschaft = $k; Wert = [string]$Opt.Meta[$k] } }
                $info | Export-Excel -Path $targets.XlsxPath -WorksheetName 'Info' -AutoSize
            }
            & $Ctx.Log "Excel geschrieben: $($targets.XlsxPath)"
        }

        & $Ctx.Progress 100 'Export abgeschlossen.'
        return $targets
    }
}

# Engine in die Hauptsession laden (für Headless-Modus, Filter-Vorschau, OU-Picker)
. $script:ArgusEngine

# ============================================================================
#  EINSTELLUNGEN / PROFILE (JSON) - gemeinsam für GUI und Headless
# ============================================================================

function Get-ArgusDefaultSettings {
    @{
        Version    = 2
        Connection = @{
            Server = ''; PortChoice = '389'; CustomPort = 389; UseLdaps = $false
            TimeoutSec = 120; AuthMode = 'current'; Username = ''; BaseDn = ''
        }
        Source     = @{ Mode = 'all'; Groups = @(); GroupMatchAny = $true; Users = @() }
        Filter     = @{ Join = 'and'; Rows = @(); Raw = '' }
        Columns    = @{
            Selected = @($script:AttributeCatalog | Where-Object { $_.Default } | ForEach-Object { $_.Key })
            Custom = @(); GroupColumnsEnabled = $false; GroupColumns = @(); Sort = 'name'
        }
        Export     = @{
            Folder = $script:DefaultExportDir; File = 'Argus-Export'
            Xlsx = $true; Csv = $false; CsvDelimiter = ';'; StrictCsv = $false
            AddTimestamp = $false; WriteLogFile = $true
        }
        Window     = @{ Width = 1100; Height = 800 }
    }
}

# JSON (PSCustomObject) rekursiv in Hashtables/Arrays wandeln (PS-5.1-kompatibel)
function ConvertTo-ArgusHashtable {
    param($InputObject)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($k in $InputObject.Keys) { $h[[string]$k] = ConvertTo-ArgusHashtable $InputObject[$k] }
        return $h
    }
    if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
        $h = @{}
        foreach ($p in $InputObject.PSObject.Properties) { $h[$p.Name] = ConvertTo-ArgusHashtable $p.Value }
        return $h
    }
    if ($InputObject -is [array]) {
        return , @($InputObject | ForEach-Object { ConvertTo-ArgusHashtable $_ })
    }
    return $InputObject
}

# Tief zusammenführen: $Overlay überschreibt $Base (fehlende Schlüssel bleiben)
function Merge-ArgusSettings {
    param([hashtable]$Base, [hashtable]$Overlay)
    $result = @{}
    foreach ($k in $Base.Keys) { $result[$k] = $Base[$k] }
    if ($Overlay) {
        foreach ($k in $Overlay.Keys) {
            if ($result.ContainsKey($k) -and $result[$k] -is [hashtable] -and $Overlay[$k] -is [hashtable]) {
                $result[$k] = Merge-ArgusSettings $result[$k] $Overlay[$k]
            }
            else { $result[$k] = $Overlay[$k] }
        }
    }
    return $result
}

function Read-ArgusProfile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Profil-Datei nicht gefunden: $Path" }
    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $obj = $raw | ConvertFrom-Json
    $loaded = ConvertTo-ArgusHashtable $obj
    if ($loaded -isnot [hashtable]) { throw "Profil-Datei hat ein ungültiges Format: $Path" }
    return Merge-ArgusSettings (Get-ArgusDefaultSettings) $loaded
}

function Save-ArgusProfile {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param([hashtable]$Settings, [string]$Path)
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
    # Kennwörter werden bewusst NIE gespeichert.
    $Settings | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Path -Encoding UTF8
}

# Effektiven Port/SSL aus den Verbindungs-Einstellungen ableiten
function Get-ArgusEffectivePort {
    param([hashtable]$Conn)
    $port = 389; $ssl = $false
    switch ([string]$Conn.PortChoice) {
        '389'    { $port = 389;  $ssl = $false }
        '636'    { $port = 636;  $ssl = $true }
        '3268'   { $port = 3268; $ssl = $false }
        '3269'   { $port = 3269; $ssl = $true }
        'custom' { $port = [int]$Conn.CustomPort; $ssl = [bool]$Conn.UseLdaps }
        default  { $port = 389;  $ssl = [bool]$Conn.UseLdaps }
    }
    return @{ Port = $port; UseSsl = $ssl }
}

# Spaltenauswahl (Katalog-Schlüssel + benutzerdefinierte LDAP-Namen) in die
# Attribut-Definitionsliste für die Engine auflösen. Reihenfolge = Katalog.
function Resolve-ArgusAttributeSelection {
    param([string[]]$SelectedKeys, [string[]]$CustomNames)
    $selected = New-Object System.Collections.Generic.List[object]
    $used = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $usedLdap = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $keySet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($k in @($SelectedKeys)) { if ($k) { [void]$keySet.Add($k) } }

    foreach ($entry in $script:AttributeCatalog) {
        if ($keySet.Contains($entry.Key) -or $keySet.Contains($entry.Ldap)) {
            if (-not $used.Add($entry.Header)) { continue }
            [void]$selected.Add(@{
                Header = $entry.Header; Ldap = $entry.Ldap; Type = $entry.Type
                UacBit = $(if ($entry.ContainsKey('UacBit')) { $entry.UacBit } else { 0 })
                UacInvert = $(if ($entry.ContainsKey('UacInvert')) { [bool]$entry.UacInvert } else { $false })
                Custom = $false
            })
            # uac-Katalogeinträge teilen sich userAccountControl - nicht als
            # "verbraucht" markieren, sonst blockiert 'Aktiviert' den zweiten Eintrag
            if ($entry.Type -ne 'uac') { [void]$usedLdap.Add($entry.Ldap) }
            [void]$keySet.Remove($entry.Key); [void]$keySet.Remove($entry.Ldap)
        }
    }
    # Übrige Schlüssel, die keinem Katalog-Eintrag entsprachen, als LDAP-Namen
    # behandeln. Bereits über den Katalog abgedeckte LDAP-Attribute werden
    # übersprungen (verhindert doppelte Spalten wie 'E-Mail' + custom 'mail').
    $extra = @($keySet) + @($CustomNames | Where-Object { $_ })
    foreach ($c in $extra) {
        $c = ([string]$c).Trim()
        if (-not $c -or $usedLdap.Contains($c) -or -not $used.Add($c)) { continue }
        [void]$usedLdap.Add($c)
        [void]$selected.Add(@{ Header = $c; Ldap = $c; Type = 'string'; UacBit = 0; UacInvert = $false; Custom = $true })
    }
    return , $selected.ToArray()
}

# Filterzeilen aus den Einstellungen (Attr-Header/LDAP-Name, Op-Code, Wert)
# in Engine-Filter auflösen. Liefert @{ Filters; Errors }.
function Resolve-ArgusFilterRows {
    param($Rows)
    $headerToEntry = @{}
    foreach ($e in $script:AttributeCatalog) { $headerToEntry[$e.Header] = $e }
    $filters = New-Object System.Collections.Generic.List[object]
    $errors = New-Object System.Collections.Generic.List[string]
    $i = 0
    foreach ($row in @($Rows)) {
        $i++
        $attr = [string]$row.Attr
        $op   = [string]$row.Op
        $val  = [string]$row.Val
        if ([string]::IsNullOrWhiteSpace($attr)) { continue }
        if (-not $op) { $errors.Add("Filterzeile ${i}: kein Operator gewählt."); continue }
        $needsValue = $op -notin @('present', 'notpresent')
        if ($needsValue -and [string]::IsNullOrWhiteSpace($val)) {
            $errors.Add("Filterzeile ${i} ('$attr'): Wert fehlt - Zeile entfernen oder Wert angeben.")
            continue
        }
        $entry = $headerToEntry[$attr]
        if ($entry) {
            $filters.Add(@{
                Ldap = $entry.Ldap; Op = $op; Val = $val; Type = $entry.Type
                UacBit = $(if ($entry.ContainsKey('UacBit')) { $entry.UacBit } else { 0 })
                UacInvert = $(if ($entry.ContainsKey('UacInvert')) { [bool]$entry.UacInvert } else { $false })
            })
        }
        else {
            $filters.Add(@{ Ldap = $attr.Trim(); Op = $op; Val = $val; Type = 'string'; UacBit = 0; UacInvert = $false })
        }
    }
    return @{ Filters = $filters.ToArray(); Errors = $errors.ToArray() }
}

# Engine-Suchparameter aus einem Settings-Hashtable bauen.
# $Password: SecureString oder $null. Liefert @{ Params; Errors }.
function Get-ArgusSearchParams {
    param([hashtable]$Settings, [securestring]$Password)
    $errors = New-Object System.Collections.Generic.List[string]
    $conn = $Settings.Connection
    $eff = Get-ArgusEffectivePort $conn

    $mode = [string]$Settings.Source.Mode
    if ($mode -notin @('all', 'group', 'users')) { $mode = 'all' }
    $groups = @($Settings.Source.Groups | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $users  = @($Settings.Source.Users  | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($mode -eq 'group' -and $groups.Count -eq 0) { $errors.Add('Quelle "Gruppen": mindestens eine Gruppe angeben.') }
    if ($mode -eq 'users' -and $users.Count -eq 0)  { $errors.Add('Quelle "Spezifische Benutzer": mindestens einen Benutzer angeben.') }

    $resolved = Resolve-ArgusFilterRows $Settings.Filter.Rows
    foreach ($e in $resolved.Errors) { $errors.Add($e) }

    $attrs = Resolve-ArgusAttributeSelection $Settings.Columns.Selected $Settings.Columns.Custom
    $includeGroupCols = [bool]$Settings.Columns.GroupColumnsEnabled
    $columnGroups = @($Settings.Columns.GroupColumns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($includeGroupCols -and $columnGroups.Count -eq 0) { $errors.Add('Gruppen-Spalten sind aktiviert, aber keine Gruppen angegeben.') }
    if ($attrs.Count -eq 0 -and -not $includeGroupCols) { $errors.Add('Mindestens ein Export-Feld oder Gruppen-Spalten wählen.') }

    $username = ''
    if ([string]$conn.AuthMode -eq 'manual') {
        $username = [string]$conn.Username
        if ([string]::IsNullOrWhiteSpace($username)) { $errors.Add('Alternative Anmeldedaten: Benutzername fehlt.') }
        elseif ($null -eq $Password -or ($Password -is [securestring] -and $Password.Length -eq 0)) { $errors.Add('Alternative Anmeldedaten: Kennwort fehlt.') }
    }

    $params = @{
        Server = [string]$conn.Server; Port = $eff.Port; UseLdaps = $eff.UseSsl
        BaseDn = [string]$conn.BaseDn; Username = $username
        Password = $(if ($username) { $Password } else { $null })
        TimeoutSec = [int]$conn.TimeoutSec
        Mode = $mode; Groups = $groups; GroupMatchAny = [bool]$Settings.Source.GroupMatchAny; Users = $users
        Filters = $resolved.Filters; FilterJoin = [string]$Settings.Filter.Join; RawFilter = [string]$Settings.Filter.Raw
        Attributes = $attrs; IncludeGroupColumns = $includeGroupCols; ColumnGroups = $columnGroups
        SortBy = [string]$Settings.Columns.Sort
    }
    return @{ Params = $params; Errors = $errors.ToArray() }
}

function Get-ArgusExportOptions {
    param([hashtable]$Settings, [hashtable]$Meta)
    $exp = $Settings.Export
    @{
        Folder = [string]$exp.Folder; FileName = [string]$exp.File
        Xlsx = [bool]$exp.Xlsx; Csv = [bool]$exp.Csv
        CsvDelimiter = [string]$exp.CsvDelimiter; StrictCsv = [bool]$exp.StrictCsv
        AddTimestamp = [bool]$exp.AddTimestamp; Meta = $Meta
    }
}

function Get-ArgusExportMeta {
    param([hashtable]$Stats)
    $meta = [ordered]@{
        'Tool'          = "$script:ArgusName v$script:ArgusVersion"
        'Exportiert am' = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        'Exportiert von'= ("{0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
        'Computer'      = $env:COMPUTERNAME
    }
    if ($Stats) {
        if ($Stats.Server) { $meta['Domain Controller'] = $Stats.Server }
        if ($Stats.BaseDn) { $meta['Such-Basis (DN)'] = $Stats.BaseDn }
        if ($Stats.Filter) { $meta['LDAP-Filter'] = $Stats.Filter }
        $meta['Treffer'] = $Stats.Count
        $meta['Suchdauer (s)'] = $Stats.DurationSec
        if ($Stats.Unmatched -and @($Stats.Unmatched).Count -gt 0) {
            $meta['Nicht gefundene IDs'] = (@($Stats.Unmatched) -join ', ')
        }
    }
    $meta['Hinweis'] = 'Letzte Anmeldung basiert auf lastLogonTimestamp (Replikationsverzögerung 9-14 Tage).'
    $h = @{}
    foreach ($k in $meta.Keys) { $h[$k] = $meta[$k] }
    return $h
}

# ============================================================================
#  HEADLESS-MODUS (-NoGui) - für Aufgabenplanung / Automatisierung
# ============================================================================
if ($NoGui) {
    $exitCode = 1
    try {
        $settings = if ($ProfilePath) { Read-ArgusProfile $ProfilePath } else { Get-ArgusDefaultSettings }

        # Explizit gesetzte CLI-Parameter überschreiben das Profil
        if ($PSBoundParameters.ContainsKey('Server')) { $settings.Connection.Server = $Server }
        if ($PSBoundParameters.ContainsKey('Port') -and $Port -gt 0) {
            if ($Port -in @(389, 636, 3268, 3269)) { $settings.Connection.PortChoice = [string]$Port }
            else { $settings.Connection.PortChoice = 'custom'; $settings.Connection.CustomPort = $Port }
        }
        if ($PSBoundParameters.ContainsKey('UseLdaps')) { $settings.Connection.UseLdaps = [bool]$UseLdaps; if ($UseLdaps -and $settings.Connection.PortChoice -eq '389') { $settings.Connection.PortChoice = '636' } }
        if ($PSBoundParameters.ContainsKey('BaseDn')) { $settings.Connection.BaseDn = $BaseDn }
        if ($PSBoundParameters.ContainsKey('TimeoutSec')) { $settings.Connection.TimeoutSec = $TimeoutSec }
        if ($Mode) { $settings.Source.Mode = $Mode }
        if ($PSBoundParameters.ContainsKey('Group')) { $settings.Source.Groups = @($Group); if (-not $Mode) { $settings.Source.Mode = 'group' } }
        if ($PSBoundParameters.ContainsKey('GroupMatchAll')) { $settings.Source.GroupMatchAny = -not $GroupMatchAll }
        if ($PSBoundParameters.ContainsKey('User')) { $settings.Source.Users = @($User); if (-not $Mode) { $settings.Source.Mode = 'users' } }
        if ($PSBoundParameters.ContainsKey('Attributes')) { $settings.Columns.Selected = @($Attributes); $settings.Columns.Custom = @() }
        if ($PSBoundParameters.ContainsKey('OutputPath')) { $settings.Export.Folder = $OutputPath }
        if ($PSBoundParameters.ContainsKey('FileName')) { $settings.Export.File = $FileName }
        if ($Format) {
            $settings.Export.Xlsx = ($Format -in @('xlsx', 'both'))
            $settings.Export.Csv  = ($Format -in @('csv', 'both'))
        }
        if ($PSBoundParameters.ContainsKey('AddTimestamp')) { $settings.Export.AddTimestamp = [bool]$AddTimestamp }

        $password = $null
        if ($Credential) {
            $settings.Connection.AuthMode = 'manual'
            $settings.Connection.Username = $Credential.UserName
            $password = $Credential.Password
        }

        $logFile = $script:LogFilePath
        $ctx = @{
            Log = {
                param($m)
                $line = "[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $m
                Write-Host $line
                Write-ArgusFileLog $logFile $line
            }
            Progress    = { param($pct, $status) if ($status) { Write-Progress -Activity 'Argus' -Status $status -PercentComplete ([Math]::Min(100, [Math]::Max(0, $pct))) } }
            IsCancelled = { $false }
        }
        & $ctx.Log "=== $script:ArgusName v$script:ArgusVersion (Headless) ==="

        $built = Get-ArgusSearchParams $settings $password
        if ($built.Errors.Count -gt 0) {
            foreach ($e in $built.Errors) { Write-Error $e }
            exit 1
        }

        if ($settings.Export.Xlsx -and -not (Get-Module -ListAvailable -Name ImportExcel)) {
            throw "Excel-Export benötigt das Modul 'ImportExcel' (Install-Module ImportExcel -Scope CurrentUser). Alternativ -Format csv verwenden."
        }

        # Relative Pfade absolut machen (StreamWriter nutzt sonst das Prozess-CWD)
        try { $settings.Export.Folder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($settings.Export.Folder) } catch { $null = $_ }

        $result = Invoke-ArgusSearch $built.Params $ctx
        Write-Progress -Activity 'Argus' -Completed
        if ($result.Count -eq 0) {
            & $ctx.Log 'Keine Treffer - es wurde keine Datei erzeugt.'
            exit 2
        }

        $meta = Get-ArgusExportMeta $result
        $opt = Get-ArgusExportOptions $settings $meta
        $targets = Get-ArgusExportTargets $opt
        foreach ($p in @($targets.XlsxPath, $targets.CsvPath)) {
            if ($p -and (Test-Path -LiteralPath $p) -and -not $Force) {
                throw "Datei existiert bereits: $p (mit -Force überschreiben)."
            }
        }
        $written = Export-ArgusData $result.Table $opt $ctx
        Write-Progress -Activity 'Argus' -Completed
        & $ctx.Log ("FERTIG: {0} Benutzer exportiert." -f $result.Count)
        if ($written.XlsxPath) { Write-Host "Excel: $($written.XlsxPath)" }
        if ($written.CsvPath)  { Write-Host "CSV:   $($written.CsvPath)" }
        if ($result.Unmatched -and @($result.Unmatched).Count -gt 0) {
            Write-Warning ("{0} Identifikator(en) ohne Treffer: {1}" -f @($result.Unmatched).Count, (@($result.Unmatched) -join ', '))
        }
        $exitCode = 0
    }
    catch {
        Write-Error $_.Exception.Message
        Write-ArgusFileLog $script:LogFilePath ("[{0}] FEHLER: {1}" -f (Get-Date -Format 'HH:mm:ss'), $_.Exception.Message)
        $exitCode = 1
    }
    exit $exitCode
}

# ============================================================================
#  GUI - XAML
# ============================================================================
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Argus - AD Export (LDAP)"
        Height="820" Width="1120"
        WindowStartupLocation="CenterScreen"
        Background="#F8FAFC" FontFamily="Segoe UI" FontSize="13"
        ResizeMode="CanResizeWithGrip" MinWidth="980" MinHeight="700">

  <Window.Resources>
    <Style x:Key="Lbl" TargetType="TextBlock">
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="#334155"/>
      <Setter Property="Margin" Value="0,12,0,4"/>
    </Style>
    <Style x:Key="Hint" TargetType="TextBlock">
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="Foreground" Value="#64748B"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Margin" Value="0,1,0,0"/>
    </Style>
    <Style x:Key="Inp" TargetType="TextBox">
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="BorderBrush" Value="#CBD5E1"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Background" Value="White"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="White"/>
      <Setter Property="BorderBrush" Value="#E2E8F0"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="10"/>
      <Setter Property="Padding" Value="16"/>
      <Setter Property="Margin" Value="0,0,0,14"/>
    </Style>
    <Style x:Key="CardTitle" TargetType="TextBlock">
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="Foreground" Value="#1E293B"/>
    </Style>
    <Style x:Key="Accent" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="#4F46E5" CornerRadius="8" Padding="18,9">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#4338CA"/></Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="#1E293B"/><Setter TargetName="bd" Property="BorderThickness" Value="1"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="bd" Property="Background" Value="#C7D2FE"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Secondary" TargetType="Button">
      <Setter Property="Foreground" Value="#1E293B"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="#E2E8F0" CornerRadius="8" Padding="13,8">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#CBD5E1"/></Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="#1E293B"/><Setter TargetName="bd" Property="BorderThickness" Value="1"/></Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="bd" Property="Background" Value="#F1F5F9"/>
                <Setter Property="Foreground" Value="#94A3B8"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Mini" TargetType="Button" BasedOn="{StaticResource Secondary}">
      <Setter Property="FontSize" Value="11"/>
    </Style>
    <Style x:Key="Danger" TargetType="Button">
      <Setter Property="Foreground" Value="#B91C1C"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="#FEE2E2" CornerRadius="6" Padding="9,6">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#FECACA"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="bd" Property="Background" Value="#FEF2F2"/><Setter Property="Foreground" Value="#FCA5A5"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ComboBox">
      <Setter Property="Padding" Value="6,4"/>
      <Setter Property="Height" Value="32"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style TargetType="TabItem">
      <Setter Property="Padding" Value="16,8"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Menu Grid.Row="0" Background="#F1F5F9" Padding="6,2">
      <MenuItem Header="_Profil">
        <MenuItem x:Name="miProfileLoad" Header="Laden..." InputGestureText="Strg+O"/>
        <MenuItem x:Name="miProfileSave" Header="Speichern..." InputGestureText="Strg+S"/>
        <Separator/>
        <MenuItem x:Name="miProfileDefaults" Header="Standardwerte wiederherstellen"/>
      </MenuItem>
      <MenuItem Header="_Extras">
        <MenuItem x:Name="miTestConnection" Header="Verbindung testen"/>
        <MenuItem x:Name="miOpenLogDir" Header="Protokoll-Ordner öffnen"/>
      </MenuItem>
      <MenuItem Header="_Hilfe">
        <MenuItem x:Name="miAbout" Header="Über Argus..."/>
      </MenuItem>
    </Menu>

    <Border Grid.Row="1" Background="#0F172A" Padding="24,14">
      <StackPanel Orientation="Horizontal">
        <TextBlock Text="Argus" Foreground="White" FontSize="20" FontWeight="Bold" VerticalAlignment="Center"/>
        <Border Background="#312E81" CornerRadius="6" Padding="7,2" Margin="10,0,0,0" VerticalAlignment="Center">
          <TextBlock x:Name="lblVersion" Text="v2.0.0" Foreground="#C7D2FE" FontSize="11" FontWeight="SemiBold"/>
        </Border>
        <TextBlock Text="Active-Directory-Export via LDAP &#8226; kein ADWS/RSAT n&#246;tig &#8226; LDAPS &#8226; Vorschau &#8226; Profile &#8226; Headless"
                   Foreground="#94A3B8" FontSize="12" Margin="16,0,0,0" VerticalAlignment="Center"/>
      </StackPanel>
    </Border>

    <TabControl x:Name="tabMain" Grid.Row="2" Margin="14,12,14,0" Background="Transparent" BorderThickness="0">

      <!-- ================= Tab 1: Verbindung ================= -->
      <TabItem x:Name="tabConnection" Header="1. Verbindung">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="6,10,6,4">
          <StackPanel>
            <Border Style="{StaticResource Card}">
              <StackPanel>
                <TextBlock Text="LDAP-Verbindung" Style="{StaticResource CardTitle}"/>
                <Grid Margin="0,4,0,0">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="2*"/>
                    <ColumnDefinition Width="16"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <StackPanel Grid.Column="0">
                    <TextBlock Text="Domain Controller (Hostname / FQDN / IP)" Style="{StaticResource Lbl}"/>
                    <TextBox x:Name="txtDC" Style="{StaticResource Inp}" Text="" ToolTip="Leer lassen = serverlose Bindung über den DC-Locator."/>
                    <TextBlock Text="Leer lassen = serverlose Bindung. Die Dom&#228;ne wird automatisch &#252;ber den DC-Locator gefunden."
                               Style="{StaticResource Hint}"/>
                  </StackPanel>
                  <StackPanel Grid.Column="2">
                    <TextBlock Text="Port / Protokoll" Style="{StaticResource Lbl}"/>
                    <ComboBox x:Name="cmbPort">
                      <ComboBoxItem Content="389 - LDAP (Signing+Sealing)" IsSelected="True"/>
                      <ComboBoxItem Content="636 - LDAPS (SSL/TLS)"/>
                      <ComboBoxItem Content="3268 - Global Catalog (Forest)"/>
                      <ComboBoxItem Content="3269 - Global Catalog SSL"/>
                      <ComboBoxItem Content="Benutzerdefiniert..."/>
                    </ComboBox>
                    <StackPanel x:Name="pnlCustomPort" Orientation="Horizontal" Visibility="Collapsed" Margin="0,8,0,0">
                      <TextBox x:Name="txtCustomPort" Style="{StaticResource Inp}" Width="80" Text="389"/>
                      <CheckBox x:Name="chkLdaps" Content="SSL/TLS" Margin="10,0,0,0" VerticalAlignment="Center" Foreground="#334155"/>
                    </StackPanel>
                  </StackPanel>
                </Grid>

                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="2*"/>
                    <ColumnDefinition Width="16"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <StackPanel Grid.Column="0">
                    <TextBlock Text="Such-Basis / Base-DN (optional)" Style="{StaticResource Lbl}"/>
                    <Grid>
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <TextBox x:Name="txtBaseDn" Grid.Column="0" Style="{StaticResource Inp}" Text=""/>
                      <Button x:Name="btnPickOu" Grid.Column="1" Content="OU ausw&#228;hlen..." Style="{StaticResource Secondary}" Margin="8,0,0,0"/>
                    </Grid>
                    <TextBlock Text="z. B. OU=Benutzer,DC=firma,DC=local - leer = gesamte Dom&#228;ne (bei Global Catalog: gesamter Forest)."
                               Style="{StaticResource Hint}"/>
                  </StackPanel>
                  <StackPanel Grid.Column="2">
                    <TextBlock Text="Zeitlimit (Sekunden)" Style="{StaticResource Lbl}"/>
                    <TextBox x:Name="txtTimeout" Style="{StaticResource Inp}" Width="80" HorizontalAlignment="Left" Text="120"/>
                    <TextBlock Text="0 = Systemstandard" Style="{StaticResource Hint}"/>
                  </StackPanel>
                </Grid>
              </StackPanel>
            </Border>

            <Border Style="{StaticResource Card}">
              <StackPanel>
                <TextBlock Text="Anmeldung" Style="{StaticResource CardTitle}"/>
                <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                  <RadioButton x:Name="rbAuthCurrent" Content="Aktueller Benutzer" IsChecked="True" GroupName="auth" Margin="0,0,18,0" Foreground="#334155"/>
                  <RadioButton x:Name="rbAuthManual" Content="Andere Anmeldedaten" GroupName="auth" Foreground="#334155"/>
                </StackPanel>
                <Grid x:Name="pnlManualAuth" Visibility="Collapsed" Margin="0,6,0,0">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="16"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <StackPanel Grid.Column="0">
                    <TextBlock Text="Benutzername (DOMAIN\benutzer oder UPN)" Style="{StaticResource Lbl}"/>
                    <TextBox x:Name="txtAuthUser" Style="{StaticResource Inp}"/>
                  </StackPanel>
                  <StackPanel Grid.Column="2">
                    <TextBlock Text="Kennwort" Style="{StaticResource Lbl}"/>
                    <PasswordBox x:Name="pwdAuth" Padding="8,6" BorderBrush="#CBD5E1" BorderThickness="1" Background="White" VerticalContentAlignment="Center"/>
                    <TextBlock Text="Wird nur f&#252;r die LDAP-Bindung verwendet und nie gespeichert." Style="{StaticResource Hint}"/>
                  </StackPanel>
                </Grid>
              </StackPanel>
            </Border>

            <Border Style="{StaticResource Card}">
              <StackPanel>
                <TextBlock Text="Protokollierung" Style="{StaticResource CardTitle}"/>
                <CheckBox x:Name="chkLogFile" Content="Lauf-Protokoll zus&#228;tzlich in Datei schreiben (Audit-Trail)" IsChecked="True"
                          Margin="0,10,0,0" Foreground="#334155"/>
                <TextBlock x:Name="lblLogPath" Text="" Style="{StaticResource Hint}" Margin="22,4,0,0"/>
              </StackPanel>
            </Border>
          </StackPanel>
        </ScrollViewer>
      </TabItem>

      <!-- ================= Tab 2: Quelle und Filter ================= -->
      <TabItem x:Name="tabSource" Header="2. Quelle &amp; Filter">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="6,10,6,4">
          <StackPanel>
            <Border Style="{StaticResource Card}">
              <StackPanel>
                <TextBlock Text="Quelle" Style="{StaticResource CardTitle}"/>
                <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                  <RadioButton x:Name="rbAll"   Content="Alle Benutzer" IsChecked="True" GroupName="src" Margin="0,0,18,0" Foreground="#334155"/>
                  <RadioButton x:Name="rbGroup" Content="Mitglieder von Gruppe(n)" GroupName="src" Margin="0,0,18,0" Foreground="#334155"/>
                  <RadioButton x:Name="rbUsers" Content="Spezifische Benutzer" GroupName="src" Foreground="#334155"/>
                </StackPanel>

                <StackPanel x:Name="pnlGroupSrc" Visibility="Collapsed" Margin="0,10,0,0">
                  <TextBlock Text="Filter-Gruppen (eine pro Zeile - CN, sAMAccountName oder DN)" Style="{StaticResource Lbl}"/>
                  <TextBox x:Name="txtGroupSrc" Style="{StaticResource Inp}" Height="58"
                           AcceptsReturn="True" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto"/>
                  <CheckBox x:Name="chkGroupAny" Content="Mitglied in MINDESTENS einer (statt in allen)" IsChecked="True"
                            Margin="0,8,0,0" Foreground="#334155"/>
                  <TextBlock Text="Rekursiv (verschachtelte Gruppen) inkl. prim&#228;rer Gruppe. Mehrdeutige Namen werden gemeldet."
                             Style="{StaticResource Hint}" Margin="0,4,0,0"/>
                </StackPanel>

                <StackPanel x:Name="pnlUserSrc" Visibility="Collapsed" Margin="0,10,0,0">
                  <TextBlock Text="Benutzer (einer pro Zeile - sAMAccountName, UPN, E-Mail oder CN)" Style="{StaticResource Lbl}"/>
                  <TextBox x:Name="txtUserSrc" Style="{StaticResource Inp}" Height="90"
                           AcceptsReturn="True" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto"/>
                  <TextBlock Text="Nicht gefundene Identifikatoren werden nach der Suche gemeldet. Gro&#223;e Listen werden automatisch in Teil-Abfragen zerlegt."
                             Style="{StaticResource Hint}" Margin="0,4,0,0"/>
                </StackPanel>
              </StackPanel>
            </Border>

            <Border Style="{StaticResource Card}">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <StackPanel Grid.Column="0">
                    <TextBlock Text="Attribut-Filter" Style="{StaticResource CardTitle}"/>
                    <TextBlock Text="Bedingungen auf Attribute. Datumsfelder: 'innerhalb der letzten N Tage' / '&#228;lter als N Tage'. Statusfelder (Aktiviert): wahr/falsch."
                               Style="{StaticResource Hint}" Margin="0,2,0,0"/>
                  </StackPanel>
                  <Button x:Name="btnAddFilter" Grid.Column="1" Content="+ Filter" Style="{StaticResource Mini}" VerticalAlignment="Top"/>
                </Grid>
                <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                  <TextBlock Text="Verkn&#252;pfung:" Foreground="#334155" VerticalAlignment="Center" Margin="0,0,10,0"/>
                  <RadioButton x:Name="rbJoinAnd" Content="ALLE Bedingungen (UND)" IsChecked="True" GroupName="join" Margin="0,0,14,0" Foreground="#334155"/>
                  <RadioButton x:Name="rbJoinOr" Content="MINDESTENS eine (ODER)" GroupName="join" Foreground="#334155"/>
                </StackPanel>
                <StackPanel x:Name="pnlFilters" Margin="0,10,0,0"/>

                <TextBlock Text="Zus&#228;tzlicher LDAP-Filter (RAW, wird mit UND verkn&#252;pft - optional)" Style="{StaticResource Lbl}"/>
                <TextBox x:Name="txtRawFilter" Style="{StaticResource Inp}" Text=""
                         ToolTip="Beispiel: (extensionAttribute5=X*) - wird unver&#228;ndert &#252;bernommen."/>
                <TextBlock Text="F&#252;r Power-User. Wird nicht escaped; Klammer-Balance wird gepr&#252;ft." Style="{StaticResource Hint}"/>

                <TextBlock Text="LDAP-Filter-Vorschau" Style="{StaticResource Lbl}"/>
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <TextBox x:Name="txtFilterPreview" Grid.Column="0" Style="{StaticResource Inp}" IsReadOnly="True"
                           Background="#F8FAFC" FontFamily="Consolas" FontSize="12" TextWrapping="Wrap"/>
                  <Button x:Name="btnCopyFilter" Grid.Column="1" Content="Kopieren" Style="{StaticResource Mini}" Margin="8,0,0,0" VerticalAlignment="Top"/>
                </Grid>
              </StackPanel>
            </Border>
          </StackPanel>
        </ScrollViewer>
      </TabItem>

      <!-- ================= Tab 3: Spalten ================= -->
      <TabItem x:Name="tabColumns" Header="3. Spalten">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="6,10,6,4">
          <StackPanel>
            <Border Style="{StaticResource Card}">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <StackPanel Grid.Column="0">
                    <TextBlock Text="Export-Felder" Style="{StaticResource CardTitle}"/>
                    <TextBlock Text="Spaltenreihenfolge folgt der Katalogreihenfolge." Style="{StaticResource Hint}" Margin="0,2,0,0"/>
                  </StackPanel>
                  <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Top">
                    <Button x:Name="btnSelectAll"     Content="Alle"     Style="{StaticResource Mini}"/>
                    <Button x:Name="btnSelectNone"    Content="Keine"    Style="{StaticResource Mini}" Margin="6,0,0,0"/>
                    <Button x:Name="btnSelectDefault" Content="Standard" Style="{StaticResource Mini}" Margin="6,0,0,0"/>
                  </StackPanel>
                </Grid>
                <TextBox x:Name="txtAttrSearch" Style="{StaticResource Inp}" Margin="0,10,0,0"
                         ToolTip="Tippen, um die Feldliste zu filtern"/>
                <TextBlock Text="Suchfeld: tippen, um die Liste zu filtern." Style="{StaticResource Hint}"/>
                <StackPanel x:Name="pnlAttributes" Margin="0,8,0,0"/>
                <TextBlock Text="Weitere LDAP-Attribute (komma-getrennt, exakte LDAP-Namen)" Style="{StaticResource Lbl}"/>
                <TextBox x:Name="txtCustomAttrs" Style="{StaticResource Inp}" Text=""/>
                <TextBlock Text="Attribute werden gegen das AD-Schema gepr&#252;ft; Tippfehler werden im Protokoll gemeldet. Mehrfachwerte (z. B. proxyAddresses) werden mit ; verbunden."
                           Style="{StaticResource Hint}"/>
                <TextBlock Text="Sortierung des Exports" Style="{StaticResource Lbl}"/>
                <ComboBox x:Name="cmbSort" Width="260" HorizontalAlignment="Left">
                  <ComboBoxItem Content="Nachname, Vorname" IsSelected="True"/>
                  <ComboBoxItem Content="SamAccountName"/>
                  <ComboBoxItem Content="Keine Sortierung"/>
                </ComboBox>
              </StackPanel>
            </Border>

            <Border Style="{StaticResource Card}">
              <StackPanel>
                <TextBlock Text="Gruppen-Mitgliedschaft als Spalten" Style="{StaticResource CardTitle}"/>
                <CheckBox x:Name="chkColumns" Content="Pro Gruppe eine True/False-Spalte anh&#228;ngen"
                          Margin="0,8,0,0" Foreground="#334155"/>
                <TextBlock Text="Gruppen (eine pro Zeile) - unabh&#228;ngig vom Quell-Filter" Style="{StaticResource Lbl}"/>
                <TextBox x:Name="txtColumnGroups" Style="{StaticResource Inp}" Height="58" IsEnabled="False"
                         AcceptsReturn="True" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto"/>
                <TextBlock Text="Erzeugt je Gruppe eine Spalte 'Mitglied: ...'. Rekursive Pr&#252;fung inkl. prim&#228;rer Gruppe."
                           Style="{StaticResource Hint}"/>
              </StackPanel>
            </Border>
          </StackPanel>
        </ScrollViewer>
      </TabItem>

      <!-- ================= Tab 4: Vorschau und Export ================= -->
      <TabItem x:Name="tabPreview" Header="4. Vorschau &amp; Export">
        <Grid Margin="6,10,6,4">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <StackPanel Grid.Row="0" Margin="2,0,2,8">
            <TextBlock x:Name="lblResultInfo" Text="Noch keine Suche ausgef&#252;hrt. Start mit 'Suchen' (F5)."
                       FontWeight="SemiBold" Foreground="#334155"/>
            <TextBlock x:Name="lblUnmatchedInfo" Text="" Foreground="#B45309" TextWrapping="Wrap" Visibility="Collapsed" Margin="0,4,0,0"/>
          </StackPanel>

          <DataGrid x:Name="dgPreview" Grid.Row="1"
                    AutoGenerateColumns="True" IsReadOnly="True" CanUserAddRows="False"
                    EnableRowVirtualization="True" EnableColumnVirtualization="True"
                    GridLinesVisibility="Horizontal" HeadersVisibility="Column"
                    AlternatingRowBackground="#F8FAFC" Background="White"
                    BorderBrush="#E2E8F0" BorderThickness="1"
                    ColumnWidth="Auto" MaxColumnWidth="360"
                    SelectionMode="Extended" SelectionUnit="FullRow"
                    ClipboardCopyMode="IncludeHeader"
                    HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto"/>

          <Border Grid.Row="2" Style="{StaticResource Card}" Margin="0,10,0,0">
            <StackPanel>
              <TextBlock Text="Ausgabe" Style="{StaticResource CardTitle}"/>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="2*"/>
                  <ColumnDefinition Width="16"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                  <TextBlock Text="Export-Ordner" Style="{StaticResource Lbl}"/>
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="txtPath" Grid.Column="0" Style="{StaticResource Inp}" Text=""/>
                    <Button x:Name="btnBrowse" Grid.Column="1" Content="Durchsuchen..." Style="{StaticResource Secondary}" Margin="8,0,0,0"/>
                  </Grid>
                </StackPanel>
                <StackPanel Grid.Column="2">
                  <TextBlock Text="Dateiname (ohne Endung)" Style="{StaticResource Lbl}"/>
                  <TextBox x:Name="txtFile" Style="{StaticResource Inp}" Text="Argus-Export"/>
                </StackPanel>
              </Grid>
              <StackPanel Orientation="Horizontal" Margin="0,12,0,0">
                <CheckBox x:Name="chkXlsx" Content="Excel (.xlsx)" IsChecked="True" Foreground="#334155" VerticalAlignment="Center"/>
                <CheckBox x:Name="chkCsv" Content="CSV (.csv)" Foreground="#334155" Margin="18,0,0,0" VerticalAlignment="Center"/>
                <TextBlock Text="CSV-Trennzeichen:" Foreground="#64748B" Margin="18,0,6,0" VerticalAlignment="Center"/>
                <ComboBox x:Name="cmbDelimiter" Width="130" VerticalAlignment="Center">
                  <ComboBoxItem Content="Semikolon (;)" IsSelected="True"/>
                  <ComboBoxItem Content="Komma (,)"/>
                </ComboBox>
                <CheckBox x:Name="chkTimestamp" Content="Zeitstempel im Dateinamen" Foreground="#334155" Margin="18,0,0,0" VerticalAlignment="Center"/>
              </StackPanel>
              <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
                <CheckBox x:Name="chkStrictCsv" Content="Strenger CSV-Schutz (neutralisiert auch f&#252;hrende +, -, @)" Foreground="#334155" VerticalAlignment="Center"
                          ToolTip="Schutz gegen CSV-Formel-Injektion. Achtung: betrifft dann auch Telefonnummern wie +49..."/>
              </StackPanel>
              <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                <TextBlock x:Name="lblXlsxStatus" Text="" Style="{StaticResource Hint}" VerticalAlignment="Center" Margin="0,0,8,0"/>
                <Button x:Name="btnInstallXlsx" Content="ImportExcel installieren..." Style="{StaticResource Mini}" Visibility="Collapsed"/>
              </StackPanel>
              <StackPanel Orientation="Horizontal" Margin="0,12,0,0">
                <Button x:Name="btnOpenFile"   Content="Datei &#246;ffnen"  Style="{StaticResource Secondary}" IsEnabled="False"/>
                <Button x:Name="btnOpenFolder" Content="Ordner &#246;ffnen" Style="{StaticResource Secondary}" IsEnabled="False" Margin="8,0,0,0"/>
              </StackPanel>
            </StackPanel>
          </Border>
        </Grid>
      </TabItem>
    </TabControl>

    <StackPanel Grid.Row="3" Margin="20,10,20,0">
      <ProgressBar x:Name="pbProgress" Height="16" Minimum="0" Maximum="100"
                   Foreground="#4F46E5" Background="#E2E8F0" BorderThickness="0"/>
      <TextBlock x:Name="lblStatus" Text="Bereit." Foreground="#334155"
                 Margin="0,6,0,0" TextTrimming="CharacterEllipsis"/>
    </StackPanel>

    <Expander x:Name="expLog" Grid.Row="4" Margin="20,8,20,0" IsExpanded="True">
      <Expander.Header>
        <StackPanel Orientation="Horizontal">
          <TextBlock Text="Protokoll" FontWeight="SemiBold" Foreground="#334155" VerticalAlignment="Center"/>
          <Button x:Name="btnCopyLog" Content="Kopieren" Style="{StaticResource Mini}" Margin="14,0,0,0"/>
          <Button x:Name="btnClearLog" Content="Leeren" Style="{StaticResource Mini}" Margin="6,0,0,0"/>
        </StackPanel>
      </Expander.Header>
      <Border CornerRadius="10" Background="#0F172A" Height="130" Margin="0,6,0,0">
        <TextBox x:Name="txtLog" Background="Transparent" Foreground="#A7F3D0"
                 FontFamily="Consolas" FontSize="12" BorderThickness="0" IsReadOnly="True"
                 Padding="12" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap"/>
      </Border>
    </Expander>

    <Grid Grid.Row="5" Margin="20,12,20,16">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <TextBlock x:Name="lblHintBar" Grid.Column="0" Style="{StaticResource Hint}" VerticalAlignment="Center"
                 Text="F5 = Suchen &#8226; Esc = Abbrechen &#8226; Suche zuerst, Export danach aus der Vorschau."/>
      <StackPanel Grid.Column="1" Orientation="Horizontal">
        <Button x:Name="btnTest"   Content="Verbindung testen" Style="{StaticResource Secondary}"/>
        <Button x:Name="btnStart"  Content="Suchen (F5)" Style="{StaticResource Accent}" Margin="10,0,0,0"/>
        <Button x:Name="btnExport" Content="Exportieren" Style="{StaticResource Accent}" IsEnabled="False" Margin="10,0,0,0"/>
        <Button x:Name="btnCancel" Content="Abbrechen" Style="{StaticResource Danger}" IsEnabled="False" Margin="10,0,0,0"/>
      </StackPanel>
    </Grid>

  </Grid>
</Window>
'@

# --- Fenster laden ----------------------------------------------------------
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)
$window.Title = "$script:ArgusName v$script:ArgusVersion - AD Export (LDAP)"

foreach ($name in @(
    'miProfileLoad', 'miProfileSave', 'miProfileDefaults', 'miTestConnection', 'miOpenLogDir', 'miAbout',
    'lblVersion', 'tabMain', 'tabConnection', 'tabSource', 'tabColumns', 'tabPreview',
    'txtDC', 'cmbPort', 'pnlCustomPort', 'txtCustomPort', 'chkLdaps', 'txtBaseDn', 'btnPickOu', 'txtTimeout',
    'rbAuthCurrent', 'rbAuthManual', 'pnlManualAuth', 'txtAuthUser', 'pwdAuth', 'chkLogFile', 'lblLogPath',
    'rbAll', 'rbGroup', 'rbUsers', 'pnlGroupSrc', 'txtGroupSrc', 'chkGroupAny', 'pnlUserSrc', 'txtUserSrc',
    'btnAddFilter', 'rbJoinAnd', 'rbJoinOr', 'pnlFilters', 'txtRawFilter', 'txtFilterPreview', 'btnCopyFilter',
    'txtAttrSearch', 'btnSelectAll', 'btnSelectNone', 'btnSelectDefault', 'pnlAttributes', 'txtCustomAttrs', 'cmbSort',
    'chkColumns', 'txtColumnGroups',
    'lblResultInfo', 'lblUnmatchedInfo', 'dgPreview',
    'txtPath', 'btnBrowse', 'txtFile', 'chkXlsx', 'chkCsv', 'cmbDelimiter', 'chkTimestamp', 'chkStrictCsv',
    'lblXlsxStatus', 'btnInstallXlsx', 'btnOpenFile', 'btnOpenFolder',
    'pbProgress', 'lblStatus', 'expLog', 'txtLog', 'btnCopyLog', 'btnClearLog',
    'btnTest', 'btnStart', 'btnExport', 'btnCancel'
)) { Set-Variable -Name $name -Value $window.FindName($name) }

$lblVersion.Text = "v$script:ArgusVersion"
$lblLogPath.Text = "Protokolldatei: $script:LogFilePath"
$txtPath.Text    = $script:DefaultExportDir

# --- Gemeinsamer (thread-sicherer) Zustand UI <-> Worker --------------------
$sync = [hashtable]::Synchronized(@{})
$sync.Log         = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
$sync.Percent     = 0
$sync.Status      = 'Bereit.'
$sync.Done        = $false
$sync.Success     = $false
$sync.Cancelled   = $false
$sync.ErrorMsg    = $null
$sync.Cancel      = $false
$sync.TestInfo    = $null
$sync.Data        = $null
$sync.Stats       = $null
$sync.ExportPaths = $null

$script:HaveData       = $false
$script:CurrentRunKind = $null
$script:LastStats      = $null
$script:logIndex       = 0

# --- Feld-Checkboxen aus dem Katalog generieren (mit Suchfilter) ------------
$script:AttrCheckBoxes = @{}
$script:AttrGroupPanels = @()
$currentGroup = $null
$currentWrap = $null
foreach ($entry in $script:AttributeCatalog) {
    if ($entry.Group -ne $currentGroup) {
        $currentGroup = $entry.Group
        $hdr = New-Object System.Windows.Controls.TextBlock
        $hdr.Text       = $currentGroup
        $hdr.FontSize   = 11
        $hdr.FontWeight = 'SemiBold'
        $hdr.Foreground = '#4F46E5'
        $hdr.Margin     = '0,8,0,2'
        [void]$pnlAttributes.Children.Add($hdr)
        $currentWrap = New-Object System.Windows.Controls.WrapPanel
        [void]$pnlAttributes.Children.Add($currentWrap)
        $script:AttrGroupPanels += , @($hdr, $currentWrap)
    }
    $cb = New-Object System.Windows.Controls.CheckBox
    $cb.Content    = $entry.Header
    $cb.Tag        = $entry.Key
    $cb.IsChecked  = [bool]$entry.Default
    $cb.Margin     = '8,3,8,3'
    $cb.Width      = 215
    $cb.Foreground = '#334155'
    if ($entry.Tip) { $cb.ToolTip = $entry.Tip }
    [void]$currentWrap.Children.Add($cb)
    $script:AttrCheckBoxes[$entry.Key] = $cb
}

function Update-ArgusAttrSearch {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param()
    $q = $txtAttrSearch.Text.Trim()
    foreach ($entry in $script:AttributeCatalog) {
        $cb = $script:AttrCheckBoxes[$entry.Key]
        $match = (-not $q) -or ($entry.Header -like "*$q*") -or ($entry.Key -like "*$q*") -or ($entry.Ldap -like "*$q*")
        $cb.Visibility = if ($match) { 'Visible' } else { 'Collapsed' }
    }
    foreach ($pair in $script:AttrGroupPanels) {
        $anyVisible = $false
        foreach ($child in $pair[1].Children) { if ($child.Visibility -eq 'Visible') { $anyVisible = $true; break } }
        $vis = if ($anyVisible) { 'Visible' } else { 'Collapsed' }
        $pair[0].Visibility = $vis
        $pair[1].Visibility = $vis
    }
}

# --- Filter-Builder: dynamische Zeilen --------------------------------------
$script:HeaderToEntry = @{}
foreach ($e in $script:AttributeCatalog) { $script:HeaderToEntry[$e.Header] = $e }
$script:FilterAttrHeaders = @($script:AttributeCatalog | ForEach-Object { $_.Header })
$script:OpCodeToLabel = @{}
foreach ($label in $script:FilterOps.Keys) { $script:OpCodeToLabel[$script:FilterOps[$label]] = $label }

function Add-ArgusFilterRow {
    param([string]$AttrHeader = '', [string]$OpCode = 'equals', [string]$Val = '')
    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = '0,0,0,6'
    $c1 = New-Object System.Windows.Controls.ColumnDefinition; $c1.Width = '2*'
    $c2 = New-Object System.Windows.Controls.ColumnDefinition; $c2.Width = '1.6*'
    $c3 = New-Object System.Windows.Controls.ColumnDefinition; $c3.Width = '2*'
    $c4 = New-Object System.Windows.Controls.ColumnDefinition; $c4.Width = 'Auto'
    $grid.ColumnDefinitions.Add($c1); $grid.ColumnDefinitions.Add($c2)
    $grid.ColumnDefinitions.Add($c3); $grid.ColumnDefinitions.Add($c4)

    $cbAttr = New-Object System.Windows.Controls.ComboBox
    $cbAttr.IsEditable = $true   # auch freie LDAP-Attributnamen erlaubt
    $cbAttr.ToolTip = 'Katalog-Feld wählen oder LDAP-Attributnamen direkt eintippen (z. B. extensionAttribute5).'
    foreach ($h in $script:FilterAttrHeaders) { [void]$cbAttr.Items.Add($h) }
    if ($AttrHeader -and $script:FilterAttrHeaders -contains $AttrHeader) { $cbAttr.SelectedItem = $AttrHeader }
    elseif ($AttrHeader) { $cbAttr.Text = $AttrHeader }
    else { $cbAttr.SelectedIndex = 0 }
    $cbAttr.Margin = '0,0,6,0'
    [System.Windows.Controls.Grid]::SetColumn($cbAttr, 0)

    $cbOp = New-Object System.Windows.Controls.ComboBox
    foreach ($opLabel in $script:FilterOps.Keys) { [void]$cbOp.Items.Add($opLabel) }
    $opLbl = $script:OpCodeToLabel[$OpCode]
    if ($opLbl) { $cbOp.SelectedItem = $opLbl } else { $cbOp.SelectedIndex = 0 }
    $cbOp.Margin = '0,0,6,0'
    [System.Windows.Controls.Grid]::SetColumn($cbOp, 1)

    $txtVal = New-Object System.Windows.Controls.TextBox
    $txtVal.Padding = '8,6'
    $txtVal.BorderBrush = '#CBD5E1'
    $txtVal.VerticalContentAlignment = 'Center'
    $txtVal.Margin = '0,0,6,0'
    $txtVal.Text = $Val
    [System.Windows.Controls.Grid]::SetColumn($txtVal, 2)

    $btnDel = New-Object System.Windows.Controls.Button
    $btnDel.Content = 'X'
    $btnDel.ToolTip = 'Filterzeile entfernen'
    $btnDel.Style = $window.FindResource('Danger')
    $btnDel.Tag = $grid
    $btnDel.Add_Click({ $pnlFilters.Children.Remove($this.Tag); Update-ArgusFilterPreview })
    [System.Windows.Controls.Grid]::SetColumn($btnDel, 3)

    [void]$grid.Children.Add($cbAttr)
    [void]$grid.Children.Add($cbOp)
    [void]$grid.Children.Add($txtVal)
    [void]$grid.Children.Add($btnDel)

    $grid.Tag = @{ Attr = $cbAttr; Op = $cbOp; Val = $txtVal }
    [void]$pnlFilters.Children.Add($grid)

    # Live-Vorschau aktuell halten
    $cbAttr.Add_SelectionChanged({ Update-ArgusFilterPreview })
    $cbOp.Add_SelectionChanged({ Update-ArgusFilterPreview })
    $txtVal.Add_TextChanged({ Update-ArgusFilterPreview })
}

# --- Einstellungen <-> UI ----------------------------------------------------
function Get-ArgusUiSettings {
    $portChoice = switch ($cmbPort.SelectedIndex) {
        0 { '389' } 1 { '636' } 2 { '3268' } 3 { '3269' } default { 'custom' }
    }
    $customPort = 389
    $null = [int]::TryParse($txtCustomPort.Text.Trim(), [ref]$customPort)
    if ($customPort -lt 1 -or $customPort -gt 65535) { $customPort = 389 }
    $timeout = 120
    $null = [int]::TryParse($txtTimeout.Text.Trim(), [ref]$timeout)
    if ($timeout -lt 0) { $timeout = 0 }

    $rows = @()
    foreach ($rowGrid in $pnlFilters.Children) {
        $h = $rowGrid.Tag
        if (-not $h) { continue }
        $attrText = [string]$h.Attr.Text
        $opLabel  = [string]$h.Op.SelectedItem
        if (-not $attrText.Trim() -and -not $h.Val.Text.Trim()) { continue }   # komplett leere Zeile
        $rows += @{ Attr = $attrText.Trim(); Op = [string]$script:FilterOps[$opLabel]; Val = [string]$h.Val.Text }
    }

    $selectedKeys = @()
    foreach ($entry in $script:AttributeCatalog) {
        if ($script:AttrCheckBoxes[$entry.Key].IsChecked -eq $true) { $selectedKeys += $entry.Key }
    }

    @{
        Version    = 2
        Connection = @{
            Server = $txtDC.Text.Trim(); PortChoice = $portChoice; CustomPort = $customPort
            UseLdaps = ($chkLdaps.IsChecked -eq $true); TimeoutSec = $timeout
            AuthMode = $(if ($rbAuthManual.IsChecked) { 'manual' } else { 'current' })
            Username = $txtAuthUser.Text.Trim(); BaseDn = $txtBaseDn.Text.Trim()
        }
        Source     = @{
            Mode = $(if ($rbGroup.IsChecked) { 'group' } elseif ($rbUsers.IsChecked) { 'users' } else { 'all' })
            Groups = @($txtGroupSrc.Text -split "\r?\n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            GroupMatchAny = ($chkGroupAny.IsChecked -eq $true)
            Users = @($txtUserSrc.Text -split "\r?\n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        }
        Filter     = @{
            Join = $(if ($rbJoinOr.IsChecked) { 'or' } else { 'and' })
            Rows = $rows
            Raw  = $txtRawFilter.Text.Trim()
        }
        Columns    = @{
            Selected = $selectedKeys
            Custom = @($txtCustomAttrs.Text -split '[,;]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            GroupColumnsEnabled = ($chkColumns.IsChecked -eq $true)
            GroupColumns = @($txtColumnGroups.Text -split "\r?\n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            Sort = $(switch ($cmbSort.SelectedIndex) { 1 { 'sam' } 2 { 'none' } default { 'name' } })
        }
        Export     = @{
            Folder = $txtPath.Text.Trim(); File = $txtFile.Text.Trim()
            Xlsx = ($chkXlsx.IsChecked -eq $true); Csv = ($chkCsv.IsChecked -eq $true)
            CsvDelimiter = $(if ($cmbDelimiter.SelectedIndex -eq 1) { ',' } else { ';' })
            StrictCsv = ($chkStrictCsv.IsChecked -eq $true)
            AddTimestamp = ($chkTimestamp.IsChecked -eq $true)
            WriteLogFile = ($chkLogFile.IsChecked -eq $true)
        }
        Window     = @{ Width = [int]$window.Width; Height = [int]$window.Height }
    }
}

function Set-ArgusUiSettings {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param([hashtable]$S)
    $conn = $S.Connection
    $txtDC.Text = [string]$conn.Server
    $cmbPort.SelectedIndex = switch ([string]$conn.PortChoice) {
        '389' { 0 } '636' { 1 } '3268' { 2 } '3269' { 3 } default { 4 }
    }
    $txtCustomPort.Text = [string]$conn.CustomPort
    $chkLdaps.IsChecked = [bool]$conn.UseLdaps
    $txtTimeout.Text = [string]$conn.TimeoutSec
    $txtBaseDn.Text = [string]$conn.BaseDn
    if ([string]$conn.AuthMode -eq 'manual') { $rbAuthManual.IsChecked = $true } else { $rbAuthCurrent.IsChecked = $true }
    $txtAuthUser.Text = [string]$conn.Username

    switch ([string]$S.Source.Mode) {
        'group' { $rbGroup.IsChecked = $true }
        'users' { $rbUsers.IsChecked = $true }
        default { $rbAll.IsChecked = $true }
    }
    $txtGroupSrc.Text = (@($S.Source.Groups) -join "`r`n")
    $chkGroupAny.IsChecked = [bool]$S.Source.GroupMatchAny
    $txtUserSrc.Text = (@($S.Source.Users) -join "`r`n")

    if ([string]$S.Filter.Join -eq 'or') { $rbJoinOr.IsChecked = $true } else { $rbJoinAnd.IsChecked = $true }
    $pnlFilters.Children.Clear()
    foreach ($row in @($S.Filter.Rows)) {
        Add-ArgusFilterRow ([string]$row.Attr) ([string]$row.Op) ([string]$row.Val)
    }
    $txtRawFilter.Text = [string]$S.Filter.Raw

    $selSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($k in @($S.Columns.Selected)) { if ($k) { [void]$selSet.Add([string]$k) } }
    foreach ($entry in $script:AttributeCatalog) {
        $script:AttrCheckBoxes[$entry.Key].IsChecked = $selSet.Contains($entry.Key)
    }
    $txtCustomAttrs.Text = (@($S.Columns.Custom) -join ', ')
    $chkColumns.IsChecked = [bool]$S.Columns.GroupColumnsEnabled
    $txtColumnGroups.Text = (@($S.Columns.GroupColumns) -join "`r`n")
    $cmbSort.SelectedIndex = switch ([string]$S.Columns.Sort) { 'sam' { 1 } 'none' { 2 } default { 0 } }

    $exp = $S.Export
    if ($exp.Folder) { $txtPath.Text = [string]$exp.Folder }
    if ($exp.File)   { $txtFile.Text = [string]$exp.File }
    $chkXlsx.IsChecked = [bool]$exp.Xlsx
    $chkCsv.IsChecked = [bool]$exp.Csv
    $cmbDelimiter.SelectedIndex = if ([string]$exp.CsvDelimiter -eq ',') { 1 } else { 0 }
    $chkStrictCsv.IsChecked = [bool]$exp.StrictCsv
    $chkTimestamp.IsChecked = [bool]$exp.AddTimestamp
    $chkLogFile.IsChecked = [bool]$exp.WriteLogFile

    if ($S.Window -and $S.Window.Width -ge 900 -and $S.Window.Height -ge 600) {
        $window.Width = [double]$S.Window.Width
        $window.Height = [double]$S.Window.Height
    }
    Update-ArgusFilterPreview
}

# --- Live-Vorschau des LDAP-Filters -----------------------------------------
function Update-ArgusFilterPreview {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param()
    if (-not $txtFilterPreview) { return }
    try {
        $s = Get-ArgusUiSettings
        $resolved = Resolve-ArgusFilterRows $s.Filter.Rows
        $clauses = @()
        foreach ($f in $resolved.Filters) {
            $clauses += New-ArgusClause $f.Ldap $f.Op $f.Val $f.Type $(if ($f.UacBit) { $f.UacBit } else { 0 }) $(if ($f.UacInvert) { $f.UacInvert } else { $false })
        }
        $attrPart = ''
        if ($clauses.Count -eq 1) { $attrPart = $clauses[0] }
        elseif ($clauses.Count -gt 1) {
            if ($s.Filter.Join -eq 'or') { $attrPart = '(|' + (-join $clauses) + ')' } else { $attrPart = -join $clauses }
        }
        $subject = switch ($s.Source.Mode) {
            'group' { '<Gruppen-Klausel: ' + @($s.Source.Groups).Count + ' Gruppe(n), rekursiv>' }
            'users' { '<Benutzer-Klausel: ' + @($s.Source.Users).Count + ' ID(s)>' }
            default { '' }
        }
        $raw = ''
        if ($s.Filter.Raw) { $raw = $s.Filter.Raw }
        $preview = '(&(objectCategory=person)(objectClass=user)' + $subject + $attrPart + $raw + ')'
        if ($resolved.Errors.Count -gt 0) {
            $preview = $preview + '   [Hinweis: ' + ($resolved.Errors -join ' | ') + ']'
        }
        $txtFilterPreview.Text = $preview
        $txtFilterPreview.Foreground = '#334155'
    }
    catch {
        $txtFilterPreview.Text = 'Filter unvollständig: ' + $_.Exception.Message
        $txtFilterPreview.Foreground = '#B45309'
    }
}

# --- Ergebnis-Vorschau (DataGrid) -------------------------------------------
# Manuell erzeugte Spalten mit Index-Bindung "[n]": robust gegen Sonderzeichen
# in Spaltennamen (Klammern/Punkte wären in WPF-Bindungspfaden Syntax).
$script:PreviewCap = 2000
function Set-ArgusPreview {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param([System.Data.DataTable]$Table)
    $dgPreview.ItemsSource = $null
    $dgPreview.Columns.Clear()
    if ($null -eq $Table) { return }

    $previewTable = $Table
    if ($Table.Rows.Count -gt $script:PreviewCap) {
        $previewTable = $Table.Clone()
        for ($i = 0; $i -lt $script:PreviewCap; $i++) { $previewTable.ImportRow($Table.Rows[$i]) }
    }
    for ($ci = 0; $ci -lt $previewTable.Columns.Count; $ci++) {
        $colDef = $previewTable.Columns[$ci]
        $col = New-Object System.Windows.Controls.DataGridTextColumn
        $col.Header = $colDef.ColumnName
        $b = New-Object System.Windows.Data.Binding ("[$ci]")
        if ($colDef.DataType -eq [datetime]) { $b.StringFormat = 'yyyy-MM-dd HH:mm' }
        $col.Binding = $b
        $col.SortMemberPath = $colDef.ColumnName
        [void]$dgPreview.Columns.Add($col)
    }
    $dgPreview.ItemsSource = $previewTable.DefaultView
}

# --- ImportExcel-Status ------------------------------------------------------
function Update-ArgusXlsxStatus {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param()
    $mod = Get-Module -ListAvailable -Name ImportExcel | Sort-Object Version -Descending | Select-Object -First 1
    if ($mod) {
        $lblXlsxStatus.Text = "ImportExcel v$($mod.Version) gefunden - Excel-Export verfügbar."
        $btnInstallXlsx.Visibility = 'Collapsed'
        return $true
    }
    $lblXlsxStatus.Text = "Modul 'ImportExcel' fehlt - Excel-Export nicht möglich (CSV funktioniert immer)."
    $btnInstallXlsx.Visibility = 'Visible'
    return $false
}

# --- OU-Picker ---------------------------------------------------------------
function Show-ArgusOuPicker {
    $s = Get-ArgusUiSettings
    $eff = Get-ArgusEffectivePort $s.Connection
    $connParams = @{
        Server = $s.Connection.Server; Port = $eff.Port; UseLdaps = $eff.UseSsl; BaseDn = ''
        Username = $(if ($s.Connection.AuthMode -eq 'manual') { $s.Connection.Username } else { '' })
        Password = $(if ($s.Connection.AuthMode -eq 'manual') { $pwdAuth.SecurePassword } else { $null })
        TimeoutSec = $s.Connection.TimeoutSec
    }
    $quietCtx = @{ Log = { param($m) $null = $m }; Progress = { param($p, $st) $null = $p; $null = $st }; IsCancelled = { $false } }
    $root = $null
    try { $root = New-ArgusRoot $connParams $quietCtx }
    catch {
        [System.Windows.MessageBox]::Show("Verbindung fehlgeschlagen: $($_.Exception.Message)", 'OU auswählen', 'OK', 'Error') | Out-Null
        return $null
    }

    try {
        $dlg = New-Object System.Windows.Window
        $dlg.Title = 'Such-Basis (OU) auswählen'
        $dlg.Width = 560; $dlg.Height = 520
        $dlg.WindowStartupLocation = 'CenterOwner'
        $dlg.Owner = $window
        $dlg.Background = '#F8FAFC'

        $grid = New-Object System.Windows.Controls.Grid
        $grid.Margin = '14'
        $r1 = New-Object System.Windows.Controls.RowDefinition; $r1.Height = '*'
        $r2 = New-Object System.Windows.Controls.RowDefinition; $r2.Height = 'Auto'
        $grid.RowDefinitions.Add($r1); $grid.RowDefinitions.Add($r2)

        $tree = New-Object System.Windows.Controls.TreeView
        $tree.BorderBrush = '#CBD5E1'
        [System.Windows.Controls.Grid]::SetRow($tree, 0)
        [void]$grid.Children.Add($tree)

        # Lazy-Loading: Dummy-Kind, das beim Aufklappen ersetzt wird
        $newNode = {
            param($HeaderText, $Dn)
            $item = New-Object System.Windows.Controls.TreeViewItem
            $item.Header = $HeaderText
            $item.Tag = $Dn
            [void]$item.Items.Add('...')
            return $item
        }
        $loadChildren = {
            param($Item)
            $Item.Items.Clear()
            $parentEntry = $null; $ds = $null; $res = $null
            try {
                $parentEntry = New-ArgusBoundEntry $root ([string]$Item.Tag)
                $ds = New-ArgusSearcher $parentEntry '(|(objectClass=organizationalUnit)(objectClass=container)(objectClass=builtinDomain))' @('name', 'distinguishedName') $root.TimeoutSec
                $ds.SearchScope = [System.DirectoryServices.SearchScope]::OneLevel
                $res = $ds.FindAll()
                $children = New-Object System.Collections.ArrayList
                foreach ($r in $res) {
                    [void]$children.Add([pscustomobject]@{
                        Name = [string]$r.Properties['name'][0]
                        Dn   = [string]$r.Properties['distinguishedname'][0]
                    })
                }
                foreach ($c in ($children.ToArray() | Sort-Object Name)) {
                    [void]$Item.Items.Add((& $newNode $c.Name $c.Dn))
                }
            }
            catch {
                [void]$Item.Items.Add("Fehler: $($_.Exception.Message)")
            }
            finally {
                if ($res) { try { $res.Dispose() } catch { $null = $_ } }
                if ($ds)  { try { $ds.Dispose() }  catch { $null = $_ } }
                if ($parentEntry) { try { $parentEntry.Dispose() } catch { $null = $_ } }
            }
        }
        $tree.AddHandler(
            [System.Windows.Controls.TreeViewItem]::ExpandedEvent,
            [System.Windows.RoutedEventHandler] {
                param($eventSender, $e)
                $null = $eventSender
                $item = $e.OriginalSource
                if ($item -is [System.Windows.Controls.TreeViewItem] -and $item.Items.Count -eq 1 -and $item.Items[0] -is [string]) {
                    & $loadChildren $item
                }
            })

        $rootItem = & $newNode $root.BaseDn $root.BaseDn
        [void]$tree.Items.Add($rootItem)
        $rootItem.IsExpanded = $true

        $btnPanel = New-Object System.Windows.Controls.StackPanel
        $btnPanel.Orientation = 'Horizontal'
        $btnPanel.HorizontalAlignment = 'Right'
        $btnPanel.Margin = '0,12,0,0'
        [System.Windows.Controls.Grid]::SetRow($btnPanel, 1)
        $btnOk = New-Object System.Windows.Controls.Button
        $btnOk.Content = 'Übernehmen'; $btnOk.Padding = '16,7'; $btnOk.IsDefault = $true
        $btnCancelDlg = New-Object System.Windows.Controls.Button
        $btnCancelDlg.Content = 'Abbrechen'; $btnCancelDlg.Padding = '16,7'; $btnCancelDlg.Margin = '8,0,0,0'; $btnCancelDlg.IsCancel = $true
        [void]$btnPanel.Children.Add($btnOk)
        [void]$btnPanel.Children.Add($btnCancelDlg)
        [void]$grid.Children.Add($btnPanel)
        $dlg.Content = $grid

        $script:OuPickerResult = $null
        $btnOk.Add_Click({
            $sel = $tree.SelectedItem
            if ($sel -is [System.Windows.Controls.TreeViewItem]) { $script:OuPickerResult = [string]$sel.Tag }
            $dlg.DialogResult = $true
        })
        [void]$dlg.ShowDialog()
        return $script:OuPickerResult
    }
    finally {
        Close-ArgusRoot $root
    }
}

# ============================================================================
#  WORKER (separater Runspace) - Engine wird als Text mitinjiziert
# ============================================================================
$script:WorkerBody = {
    function Add-Log {
        param([string]$m)
        $line = "[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $m
        [void]$sync.Log.Add($line)
        if ($LogFile) { Write-ArgusFileLog $LogFile $line }
    }
    $ctx = @{
        Log         = { param($m) Add-Log $m }
        Progress    = { param($pct, $status) $sync.Percent = $pct; if ($status) { $sync.Status = $status } }
        IsCancelled = { [bool]$sync.Cancel }
    }
    try {
        switch ($RunKind) {
            'test' {
                $sync.Status = 'Teste Verbindung...'
                Add-Log 'Verbindungstest gestartet...'
                $root = New-ArgusRoot $ConnParams $ctx
                try {
                    $info = New-Object System.Collections.ArrayList
                    [void]$info.Add(("Domain Controller: {0}" -f $(if ($root.DnsHost) { $root.DnsHost } else { '(serverlose Bindung / DC-Locator)' })))
                    [void]$info.Add(("Verschlüsselung:   {0}" -f $(if ($root.UseSsl) { 'LDAPS (SSL/TLS)' } else { 'LDAP mit Signing+Sealing' })))
                    [void]$info.Add(("Global Catalog:    {0}" -f $(if ($root.IsGc) { 'ja (Forest-weit)' } else { 'nein' })))
                    [void]$info.Add(("Such-Basis (DN):   {0}" -f $root.BaseDn))
                    if ($root.DomainDn -ne $root.BaseDn) { [void]$info.Add(("Domänen-Wurzel:    {0}" -f $root.DomainDn)) }
                    $sync.TestInfo = ($info.ToArray() -join "`n")
                    Add-Log 'Verbindungstest erfolgreich.'
                }
                finally { Close-ArgusRoot $root }
                $sync.Success = $true
                $sync.Percent = 100
                $sync.Status  = 'Verbindung OK.'
            }
            'search' {
                Add-Log ("=== Suche gestartet ({0} v{1}, Benutzer: {2}\{3}) ===" -f 'Argus', $AppVersion, $env:USERDOMAIN, $env:USERNAME)
                $result = Invoke-ArgusSearch $SearchParams $ctx
                $sync.Data  = $result.Table
                $sync.Stats = @{
                    Count = $result.Count; Unmatched = $result.Unmatched
                    DurationSec = $result.DurationSec; Filter = $result.Filter
                    Server = $result.Server; BaseDn = $result.BaseDn
                }
                $sync.Success = $true
                $sync.Percent = 100
                $sync.Status  = "Suche abgeschlossen: $($result.Count) Benutzer."
            }
            'export' {
                Add-Log ("=== Export gestartet ({0} Zeilen) ===" -f $ExportData.Rows.Count)
                $written = Export-ArgusData $ExportData $ExportOptions $ctx
                $sync.ExportPaths = $written
                $sync.Success = $true
                $sync.Percent = 100
                $sync.Status  = 'Export abgeschlossen.'
            }
        }
    }
    catch [System.OperationCanceledException] {
        $sync.Cancelled = $true
        $sync.Status = 'Abgebrochen.'
        Add-Log 'Abgebrochen durch Benutzer.'
    }
    catch {
        $sync.ErrorMsg = $_.Exception.Message
        $sync.Status = "Fehler: $($_.Exception.Message)"
        Add-Log ("FEHLER: {0}" -f $_.Exception.Message)
    }
    finally {
        $sync.Done = $true
    }
}

function Set-ArgusUiBusy {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param([bool]$Busy)
    $tabMain.IsEnabled   = -not $Busy
    $btnStart.IsEnabled  = -not $Busy
    $btnTest.IsEnabled   = -not $Busy
    $btnExport.IsEnabled = (-not $Busy) -and $script:HaveData
    $btnCancel.IsEnabled = $Busy
    foreach ($mi in @($miProfileLoad, $miProfileSave, $miProfileDefaults, $miTestConnection)) { $mi.IsEnabled = -not $Busy }
}

function Start-ArgusWorker {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param([string]$Kind, [hashtable]$Vars)

    $sync.Percent = 0; $sync.Status = 'Starte...'; $sync.Done = $false; $sync.Success = $false
    $sync.Cancelled = $false; $sync.ErrorMsg = $null; $sync.Cancel = $false
    $sync.TestInfo = $null; $sync.ExportPaths = $null
    if ($Kind -eq 'search') {
        $sync.Data = $null; $sync.Stats = $null
        $script:HaveData = $false
        Set-ArgusPreview $null
        $lblResultInfo.Text = 'Suche läuft...'
        $lblUnmatchedInfo.Visibility = 'Collapsed'
        $btnOpenFile.IsEnabled = $false
        $btnOpenFolder.IsEnabled = $false
    }
    if ($Kind -ne 'export') {
        $sync.Log.Clear(); $script:logIndex = 0; $txtLog.Clear()
    }
    $pbProgress.Value = 0
    $script:CurrentRunKind = $Kind

    try {
        $script:runspace = [runspacefactory]::CreateRunspace()
        $script:runspace.ThreadOptions = 'ReuseThread'
        $script:runspace.ApartmentState = 'MTA'
        $script:runspace.Open()
        $script:runspace.SessionStateProxy.SetVariable('sync', $sync)
        $script:runspace.SessionStateProxy.SetVariable('RunKind', $Kind)
        $script:runspace.SessionStateProxy.SetVariable('AppVersion', $script:ArgusVersion)
        $logFileForRun = if ($chkLogFile.IsChecked -eq $true) { $script:LogFilePath } else { '' }
        $script:runspace.SessionStateProxy.SetVariable('LogFile', $logFileForRun)
        foreach ($k in $Vars.Keys) { $script:runspace.SessionStateProxy.SetVariable($k, $Vars[$k]) }

        $script:ps = [powershell]::Create()
        $script:ps.Runspace = $script:runspace
        # Engine + Worker als EIN Skript: so sind die Engine-Funktionen im
        # Runspace definiert, ohne sie doppelt pflegen zu müssen.
        [void]$script:ps.AddScript($script:ArgusEngine.ToString() + "`n" + $script:WorkerBody.ToString())
        $script:handle = $script:ps.BeginInvoke()
    }
    catch {
        try { if ($script:ps) { $script:ps.Dispose(); $script:ps = $null } } catch { $null = $_ }
        try { if ($script:runspace) { $script:runspace.Dispose(); $script:runspace = $null } } catch { $null = $_ }
        [System.Windows.MessageBox]::Show("Der Hintergrund-Vorgang konnte nicht gestartet werden:`n`n$($_.Exception.Message)", 'Argus - Fehler', 'OK', 'Error') | Out-Null
        Set-ArgusUiBusy $false
        return
    }
    $script:RunStart = Get-Date
    Set-ArgusUiBusy $true
    $timer.Start()
}

function Complete-ArgusSearchUi {
    $stats = $sync.Stats
    $script:LastStats = $stats
    $count = if ($stats) { [int]$stats.Count } else { 0 }

    if ($stats -and $stats.Unmatched -and @($stats.Unmatched).Count -gt 0) {
        $u = @($stats.Unmatched)
        $shown = $u | Select-Object -First 30
        $suffix = if ($u.Count -gt 30) { " ... (+$($u.Count - 30) weitere, siehe Protokoll)" } else { '' }
        $lblUnmatchedInfo.Text = "$($u.Count) Identifikator(en) ohne Treffer: $($shown -join ', ')$suffix"
        $lblUnmatchedInfo.Visibility = 'Visible'
    }
    else {
        $lblUnmatchedInfo.Visibility = 'Collapsed'
    }

    if ($count -eq 0) {
        $script:HaveData = $false
        $btnExport.IsEnabled = $false
        Set-ArgusPreview $null
        $lblResultInfo.Text = 'Keine Treffer. Quelle und Filter prüfen (Details im Protokoll).'
        $tabPreview.IsSelected = $true
        [System.Windows.MessageBox]::Show('Die Suche hat keine Benutzer gefunden. Es wurde nichts exportiert.', 'Argus - Keine Treffer', 'OK', 'Information') | Out-Null
        return
    }
    $script:HaveData = $true
    $btnExport.IsEnabled = $true
    Set-ArgusPreview $sync.Data
    $capNote = if ($count -gt $script:PreviewCap) { " Vorschau zeigt die ersten $($script:PreviewCap) Zeilen; der Export enthält alle." } else { '' }
    $lblResultInfo.Text = "$count Benutzer gefunden (Dauer: $($stats.DurationSec) s).$capNote"
    $tabPreview.IsSelected = $true
}

function Complete-ArgusExportUi {
    $written = $sync.ExportPaths
    $files = @()
    if ($written -and $written.XlsxPath) { $files += $written.XlsxPath }
    if ($written -and $written.CsvPath)  { $files += $written.CsvPath }
    $script:LastExportFile = if ($files.Count -gt 0) { $files[0] } else { $null }
    $btnOpenFile.IsEnabled   = ($files.Count -gt 0)
    $btnOpenFolder.IsEnabled = ($files.Count -gt 0)
    [System.Windows.MessageBox]::Show(("Export abgeschlossen:`n`n{0}" -f ($files -join "`n")), 'Argus - Export fertig', 'OK', 'Information') | Out-Null
}

# --- DispatcherTimer --------------------------------------------------------
$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(200)
$timer.Add_Tick({
    $pbProgress.Value = [double]$sync.Percent
    $statusText = [string]$sync.Status
    if ($script:RunStart) {
        $el = (Get-Date) - $script:RunStart
        $statusText = '{0}  [{1:mm\:ss}]' -f $statusText, $el
    }
    $lblStatus.Text = $statusText

    # Fertig? (Done-Flag ZUERST lesen, damit keine Logzeilen verloren gehen)
    $isDone = [bool]$sync.Done
    if (-not $isDone -and $script:ps) {
        # Pipeline-Zusammenbruch erkennen (Worker-Finally lief nie)
        $st = $script:ps.InvocationStateInfo.State
        if ($st -eq 'Completed' -or $st -eq 'Failed' -or $st -eq 'Stopped') {
            $isDone = $true
            if (-not $sync.ErrorMsg) {
                $reason = $script:ps.InvocationStateInfo.Reason
                $sync.ErrorMsg = if ($reason) { $reason.Message } else { 'Der Hintergrund-Vorgang wurde unerwartet beendet.' }
            }
        }
    }

    while ($script:logIndex -lt $sync.Log.Count) {
        $txtLog.AppendText($sync.Log[$script:logIndex] + "`r`n")
        $script:logIndex++
        $txtLog.ScrollToEnd()
    }

    if ($isDone) {
        $timer.Stop()
        $script:RunStart = $null
        try { if ($script:ps -and $script:handle) { $script:ps.EndInvoke($script:handle) } } catch { $null = $_ }
        try { if ($script:ps) { $script:ps.Dispose() } } catch { $null = $_ }
        try { if ($script:runspace) { $script:runspace.Close(); $script:runspace.Dispose() } } catch { $null = $_ }
        $script:ps = $null; $script:runspace = $null; $script:handle = $null

        while ($script:logIndex -lt $sync.Log.Count) {
            $txtLog.AppendText($sync.Log[$script:logIndex] + "`r`n")
            $script:logIndex++
            $txtLog.ScrollToEnd()
        }

        $kind = $script:CurrentRunKind
        Set-ArgusUiBusy $false
        if ($sync.Cancelled) {
            $lblStatus.Text = 'Abgebrochen.'
            $pbProgress.Value = 0
        }
        elseif ($sync.Success) {
            $lblStatus.Text = [string]$sync.Status
            switch ($kind) {
                'test'   { [System.Windows.MessageBox]::Show([string]$sync.TestInfo, 'Argus - Verbindung erfolgreich', 'OK', 'Information') | Out-Null }
                'search' { Complete-ArgusSearchUi }
                'export' { Complete-ArgusExportUi }
            }
        }
        elseif ($sync.ErrorMsg) {
            $lblStatus.Text = "Fehler: $($sync.ErrorMsg)"
            [System.Windows.MessageBox]::Show("Der Vorgang ist fehlgeschlagen:`n`n$($sync.ErrorMsg)`n`nDetails stehen im Protokoll.", 'Argus - Fehler', 'OK', 'Error') | Out-Null
        }
    }
})

# ============================================================================
#  AKTIONEN
# ============================================================================
function Start-ArgusSearchRun {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param()
    $s = Get-ArgusUiSettings
    $pw = if ($s.Connection.AuthMode -eq 'manual') { $pwdAuth.SecurePassword } else { $null }
    $built = Get-ArgusSearchParams $s $pw
    if ($built.Errors.Count -gt 0) {
        [System.Windows.MessageBox]::Show(("Bitte Eingaben prüfen:`n`n- " + ($built.Errors -join "`n- ")), 'Argus - Eingaben unvollständig', 'OK', 'Warning') | Out-Null
        return
    }
    $script:LastSearchSettings = $s
    Start-ArgusWorker 'search' @{ SearchParams = $built.Params }
}

function Start-ArgusTestRun {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param()
    $s = Get-ArgusUiSettings
    if ($s.Connection.AuthMode -eq 'manual' -and (-not $s.Connection.Username -or $pwdAuth.SecurePassword.Length -eq 0)) {
        [System.Windows.MessageBox]::Show('Alternative Anmeldedaten: Benutzername und Kennwort angeben.', 'Argus', 'OK', 'Warning') | Out-Null
        return
    }
    $eff = Get-ArgusEffectivePort $s.Connection
    $connParams = @{
        Server = $s.Connection.Server; Port = $eff.Port; UseLdaps = $eff.UseSsl; BaseDn = $s.Connection.BaseDn
        Username = $(if ($s.Connection.AuthMode -eq 'manual') { $s.Connection.Username } else { '' })
        Password = $(if ($s.Connection.AuthMode -eq 'manual') { $pwdAuth.SecurePassword } else { $null })
        TimeoutSec = $s.Connection.TimeoutSec
    }
    Start-ArgusWorker 'test' @{ ConnParams = $connParams }
}

function Start-ArgusExportRun {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    param()
    if (-not $script:HaveData -or $null -eq $sync.Data) { return }
    $s = Get-ArgusUiSettings
    $errors = @()
    if (-not $s.Export.Folder) { $errors += 'Export-Ordner angeben.' }
    if (-not $s.Export.Xlsx -and -not $s.Export.Csv) { $errors += 'Mindestens ein Format (Excel/CSV) wählen.' }
    if ($s.Export.Xlsx -and -not (Update-ArgusXlsxStatus)) {
        $errors += "Excel-Export benötigt das Modul 'ImportExcel' (Button 'ImportExcel installieren...' im Ausgabe-Bereich) - oder stattdessen CSV wählen."
    }
    if ($errors.Count -gt 0) {
        [System.Windows.MessageBox]::Show(("Bitte Eingaben prüfen:`n`n- " + ($errors -join "`n- ")), 'Argus - Export', 'OK', 'Warning') | Out-Null
        return
    }

    # Relative Pfade absolut machen (StreamWriter nutzt sonst das Prozess-CWD,
    # das vom PowerShell-Arbeitsverzeichnis abweichen kann)
    try { $s.Export.Folder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($s.Export.Folder) } catch { $null = $_ }

    $meta = Get-ArgusExportMeta $script:LastStats
    $opt = Get-ArgusExportOptions $s $meta
    $targets = Get-ArgusExportTargets $opt
    $opt.Targets = $targets   # exakt diese Pfade verwenden (Zeitstempel-Konsistenz)

    $existing = @()
    foreach ($p in @($targets.XlsxPath, $targets.CsvPath)) {
        if ($p -and (Test-Path -LiteralPath $p)) { $existing += $p }
    }
    if ($existing.Count -gt 0) {
        $answer = [System.Windows.MessageBox]::Show(
            ("Folgende Datei(en) existieren bereits und werden überschrieben:`n`n{0}`n`nFortfahren?" -f ($existing -join "`n")),
            'Argus - Überschreiben?', 'YesNo', 'Warning')
        if ($answer -ne 'Yes') { return }
    }
    Start-ArgusWorker 'export' @{ ExportData = $sync.Data; ExportOptions = $opt }
}

# ============================================================================
#  EVENTS
# ============================================================================
$updateSrcPanels = {
    $pnlGroupSrc.Visibility = if ($rbGroup.IsChecked) { 'Visible' } else { 'Collapsed' }
    $pnlUserSrc.Visibility  = if ($rbUsers.IsChecked) { 'Visible' } else { 'Collapsed' }
    Update-ArgusFilterPreview
}
$rbAll.Add_Checked($updateSrcPanels)
$rbGroup.Add_Checked($updateSrcPanels)
$rbUsers.Add_Checked($updateSrcPanels)

$cmbPort.Add_SelectionChanged({
    $pnlCustomPort.Visibility = if ($cmbPort.SelectedIndex -eq 4) { 'Visible' } else { 'Collapsed' }
})

$updateAuthPanel = {
    $pnlManualAuth.Visibility = if ($rbAuthManual.IsChecked) { 'Visible' } else { 'Collapsed' }
}
$rbAuthCurrent.Add_Checked($updateAuthPanel)
$rbAuthManual.Add_Checked($updateAuthPanel)

$chkColumns.Add_Checked({   $txtColumnGroups.IsEnabled = $true })
$chkColumns.Add_Unchecked({ $txtColumnGroups.IsEnabled = $false })

$btnAddFilter.Add_Click({ Add-ArgusFilterRow; Update-ArgusFilterPreview })
$rbJoinAnd.Add_Checked({ Update-ArgusFilterPreview })
$rbJoinOr.Add_Checked({ Update-ArgusFilterPreview })
$txtRawFilter.Add_TextChanged({ Update-ArgusFilterPreview })
$txtGroupSrc.Add_TextChanged({ Update-ArgusFilterPreview })
$txtUserSrc.Add_TextChanged({ Update-ArgusFilterPreview })
$btnCopyFilter.Add_Click({ try { [System.Windows.Clipboard]::SetText($txtFilterPreview.Text) } catch { $null = $_ } })

$txtAttrSearch.Add_TextChanged({ Update-ArgusAttrSearch })
$btnSelectAll.Add_Click({  foreach ($cb in $script:AttrCheckBoxes.Values) { $cb.IsChecked = $true } })
$btnSelectNone.Add_Click({ foreach ($cb in $script:AttrCheckBoxes.Values) { $cb.IsChecked = $false } })
$btnSelectDefault.Add_Click({
    foreach ($entry in $script:AttributeCatalog) { $script:AttrCheckBoxes[$entry.Key].IsChecked = [bool]$entry.Default }
})

$btnPickOu.Add_Click({
    $dn = Show-ArgusOuPicker
    if ($dn) { $txtBaseDn.Text = $dn }
})

$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($txtPath.Text) { $dlg.SelectedPath = $txtPath.Text }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtPath.Text = $dlg.SelectedPath }
})

$btnInstallXlsx.Add_Click({
    $answer = [System.Windows.MessageBox]::Show(
        "Das Modul 'ImportExcel' jetzt für den aktuellen Benutzer installieren?`n`n(Install-Module ImportExcel -Scope CurrentUser)`n`nHinweis: benötigt Internet-/PSGallery-Zugriff. In Offline-Umgebungen das Modul per Softwareverteilung bereitstellen oder CSV exportieren.",
        'ImportExcel installieren', 'YesNo', 'Question')
    if ($answer -ne 'Yes') { return }
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        Install-Module ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
        [System.Windows.MessageBox]::Show('ImportExcel wurde installiert.', 'Argus', 'OK', 'Information') | Out-Null
    }
    catch {
        [System.Windows.MessageBox]::Show("Installation fehlgeschlagen: $($_.Exception.Message)`n`nManuell: Install-Module ImportExcel -Scope CurrentUser", 'Argus - Fehler', 'OK', 'Error') | Out-Null
    }
    finally {
        $window.Cursor = $null
        [void](Update-ArgusXlsxStatus)
    }
})

$btnTest.Add_Click({ Start-ArgusTestRun })
$btnStart.Add_Click({ Start-ArgusSearchRun })
$btnExport.Add_Click({ Start-ArgusExportRun })
$btnCancel.Add_Click({
    $sync.Cancel = $true
    $btnCancel.IsEnabled = $false
    $lblStatus.Text = 'Breche ab... (wartet auf laufende LDAP-Antwort)'
})

$btnOpenFile.Add_Click({
    $written = $sync.ExportPaths
    foreach ($p in @($written.XlsxPath, $written.CsvPath)) {
        if ($p -and (Test-Path -LiteralPath $p)) { Start-Process $p; break }
    }
})
$btnOpenFolder.Add_Click({
    if ($script:LastExportFile -and (Test-Path -LiteralPath $script:LastExportFile)) {
        Start-Process 'explorer.exe' "/select,`"$script:LastExportFile`""
    }
})

$btnCopyLog.Add_Click({ try { [System.Windows.Clipboard]::SetText($txtLog.Text) } catch { $null = $_ } })
$btnClearLog.Add_Click({ $txtLog.Clear() })

# --- Menü -------------------------------------------------------------------
$miProfileLoad.Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'Argus-Profil (*.json)|*.json|Alle Dateien (*.*)|*.*'
    if ($dlg.ShowDialog() -eq $true) {
        try {
            Set-ArgusUiSettings (Read-ArgusProfile $dlg.FileName)
            $lblStatus.Text = "Profil geladen: $($dlg.FileName)"
        }
        catch {
            [System.Windows.MessageBox]::Show("Profil konnte nicht geladen werden:`n$($_.Exception.Message)", 'Argus - Fehler', 'OK', 'Error') | Out-Null
        }
    }
})
$miProfileSave.Add_Click({
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'Argus-Profil (*.json)|*.json'
    $dlg.FileName = 'Argus-Profil.json'
    if ($dlg.ShowDialog() -eq $true) {
        try {
            Save-ArgusProfile (Get-ArgusUiSettings) $dlg.FileName
            $lblStatus.Text = "Profil gespeichert: $($dlg.FileName) (für Headless: Argus.ps1 -NoGui -ProfilePath `"$($dlg.FileName)`")"
        }
        catch {
            [System.Windows.MessageBox]::Show("Profil konnte nicht gespeichert werden:`n$($_.Exception.Message)", 'Argus - Fehler', 'OK', 'Error') | Out-Null
        }
    }
})
$miProfileDefaults.Add_Click({
    Set-ArgusUiSettings (Get-ArgusDefaultSettings)
    if ($pnlFilters.Children.Count -eq 0) { Add-ArgusFilterRow }
    $lblStatus.Text = 'Standardwerte wiederhergestellt.'
})
$miTestConnection.Add_Click({ Start-ArgusTestRun })
$miOpenLogDir.Add_Click({
    try {
        if (-not (Test-Path -LiteralPath $script:LogDir)) { New-Item -Path $script:LogDir -ItemType Directory -Force | Out-Null }
        Start-Process 'explorer.exe' $script:LogDir
    }
    catch { $null = $_ }
})
$miAbout.Add_Click({
    [System.Windows.MessageBox]::Show(
        ("Argus v{0} - Active Directory Export Tool (LDAP Edition)`n`n" -f $script:ArgusVersion) +
        "Reines LDAP (389/636/3268/3269) - kein ADWS, kein RSAT nötig.`n" +
        "Excel-Export via ImportExcel, CSV ohne Abhängigkeiten.`n`n" +
        "Headless-Modus:  Argus.ps1 -NoGui -ProfilePath <profil.json>`n" +
        "Hilfe:  Get-Help .\Argus.ps1 -Full`n`n" +
        "Owner: Semih Danyeri`nKontakt: Microsoft Teams -> 'Semih Danyeri'",
        'Über Argus', 'OK', 'Information') | Out-Null
})

# --- Tastatur ----------------------------------------------------------------
$window.Add_KeyDown({
    param($keySender, $e)
    $null = $keySender
    if ($e.Key -eq 'F5' -and $btnStart.IsEnabled) { Start-ArgusSearchRun; $e.Handled = $true }
    elseif ($e.Key -eq 'Escape' -and $btnCancel.IsEnabled) { $sync.Cancel = $true; $btnCancel.IsEnabled = $false; $e.Handled = $true }
    elseif ($e.Key -eq 'O' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control' -and $miProfileLoad.IsEnabled) { $miProfileLoad.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.MenuItem]::ClickEvent))); $e.Handled = $true }
    elseif ($e.Key -eq 'S' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control' -and $miProfileSave.IsEnabled) { $miProfileSave.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.MenuItem]::ClickEvent))); $e.Handled = $true }
})

# --- Fenster schließen -------------------------------------------------------
$window.Add_Closing({
    try { $sync.Cancel = $true } catch { $null = $_ }
    try { $timer.Stop() } catch { $null = $_ }
    try { if ($script:ps) { [void]$script:ps.BeginStop($null, $null) } } catch { $null = $_ }
    try { Save-ArgusProfile (Get-ArgusUiSettings) $script:LastSettingsPath } catch { $null = $_ }
})

# ============================================================================
#  START
# ============================================================================
$initialSettings = $null
if ($ProfilePath) {
    try { $initialSettings = Read-ArgusProfile $ProfilePath }
    catch {
        [System.Windows.MessageBox]::Show("Profil konnte nicht geladen werden:`n$($_.Exception.Message)", 'Argus - Warnung', 'OK', 'Warning') | Out-Null
    }
}
if (-not $initialSettings -and (Test-Path -LiteralPath $script:LastSettingsPath)) {
    try { $initialSettings = Read-ArgusProfile $script:LastSettingsPath } catch { $initialSettings = $null }
}
if (-not $initialSettings) { $initialSettings = Get-ArgusDefaultSettings }
Set-ArgusUiSettings $initialSettings
if ($pnlFilters.Children.Count -eq 0) { Add-ArgusFilterRow }
[void](Update-ArgusXlsxStatus)
Update-ArgusAttrSearch
Update-ArgusFilterPreview
[void]$txtDC.Focus()

[void]$window.ShowDialog()

# Nach dem Schließen: Restarbeiten aufräumen (Runspace läuft ggf. noch nach)
try { if ($script:ps) { $script:ps.Dispose() } } catch { $null = $_ }
try { if ($script:runspace) { $script:runspace.Dispose() } } catch { $null = $_ }
