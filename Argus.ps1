<#
 .SYNOPSIS
   Argus - Active Directory Export Tool (LDAP Edition)

   Grafisches Tool zum Auslesen und Exportieren von Active-Directory-
   Benutzerobjekten nach Excel - ueber reines LDAP (Port 389).

   Argus benoetigt bewusst KEIN ADWS (Port 9389) und KEIN RSAT /
   ActiveDirectory-PowerShell-Modul. Es laeuft damit auch in Umgebungen,
   in denen ADWS blockiert ist oder RSAT nicht installiert werden darf.
   Einzige Abhaengigkeit: das PowerShell-Modul 'ImportExcel' (wird beim
   Start automatisch angeboten, falls es fehlt).

 .DESCRIPTION
   Funktionsumfang:
     - Quelle waehlbar: alle Benutzer, Mitglieder von Gruppe(n) (rekursiv)
       oder eine Liste spezifischer Benutzer (sAMAccountName/UPN/Mail/CN)
     - Dynamischer Attribut-Filter-Builder (beliebig viele Bedingungen,
       UND-verknuepft, inkl. RAW-LDAP-Modus fuer Power-User)
     - Frei waehlbare Export-Spalten aus einem Attribut-Katalog
       plus beliebige zusaetzliche LDAP-Attribute
     - Optionale True/False-Spalten fuer Gruppenmitgliedschaften
       (rekursive Pruefung ueber verschachtelte Gruppen)
     - Export als formatierte Excel-Tabelle (ImportExcel)
     - Nicht blockierende UI: Die Suche laeuft in einem eigenen Runspace,
       Fortschritt und Log werden live angezeigt.

 .NOTES
   Verifizierte LDAP-Fallstricke, die das ActiveDirectory-Modul sonst fuer
   einen versteckt und die hier explizit behandelt werden:

   1) PageSize-Falle: DirectorySearcher.FindAll() liefert OHNE gesetztes
      PageSize nur die ersten 1000 Treffer zurueck - ohne Fehler.
      -> PageSize = 1000 erzwingt Server-Side-Paging.
   2) FILETIME-"Never"-Falle: accountExpires/pwdLastSet/lastLogonTimestamp
      sind Int64. 0 und 9223372036854775807 bedeuten "nie/ungesetzt";
      FromFileTime() wirft dann bzw. liefert Jahr 1601/30828.
      -> Sentinel-Guard in Convert-FileTime.
   3) Rekursive Gruppenmitgliedschaft: memberOf:1.2.840.113556.1.4.1941:=<DN>
      (LDAP_MATCHING_RULE_IN_CHAIN).
   4) Deaktivierte Konten: userAccountControl:1.2.840.113556.1.4.803:=2
      (Bit-AND-Matching-Rule).
   5) LDAP-Attributnamen != AD-Modul-Property-Namen (z. B. l=City, sn=Surname,
      physicalDeliveryOfficeName=Office, co=Land-Name, c=Land-ISO).

   Owner:   Semih Danyeri
   Kontakt: Microsoft Teams -> "Semih Danyeri"
#>

# --- Assemblies -------------------------------------------------------------
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.DirectoryServices

# --- Voraussetzungen pruefen (vor GUI) -------------------------------------
# Bewusst KEINE RSAT/ActiveDirectory-Modul-Pruefung: reines LDAP braucht das nicht.
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    $answer = [System.Windows.MessageBox]::Show(
        "Das Modul 'ImportExcel' fehlt. Jetzt fuer den aktuellen Benutzer installieren?`n`n(Install-Module ImportExcel -Scope CurrentUser)",
        "ImportExcel benoetigt", 'YesNo', 'Question')
    if ($answer -eq 'Yes') {
        try { Install-Module ImportExcel -Scope CurrentUser -Force -ErrorAction Stop }
        catch {
            [System.Windows.MessageBox]::Show(
                "Installation fehlgeschlagen: $($_.Exception.Message)`n`nManuell: Install-Module ImportExcel -Scope CurrentUser",
                "Fehler", 'OK', 'Error') | Out-Null
            return
        }
    } else { return }
}

<# Vor dem GUI-Start: Mitgliedschaft pruefen (rekursiv, tokenGroups des eigenen Tokens)
$requiredGroup = 'MAAApprover'   # sAMAccountName der Gruppe
$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
if (-not $principal.IsInRole($requiredGroup)) {
    [System.Windows.MessageBox]::Show(
        "Keine Berechtigung. Erforderliche Gruppe: $requiredGroup`nKontakt: Teams -> Semih Danyeri",
        "Argus - Zugriff verweigert", 'OK', 'Warning') | Out-Null
    return
}#>

# --- Attribut-Katalog (Single Source of Truth) ------------------------------
# Type: string | enabled | date | datedirect | manager | guid
#   string     -> Wert direkt aus LDAP-Property (erstes Element)
#   enabled     -> aus userAccountControl abgeleitet (Bit 2 = deaktiviert)
#   date        -> Int64 FILETIME -> DateTime (mit Never-Sentinel-Guard)
#   datedirect  -> bereits DateTime (whenCreated/whenChanged via ADSI)
#   manager     -> DN -> CN extrahieren
$script:AttributeCatalog = @(
    @{ Key='GivenName';      Header='Vorname';          Group='Identitaet';   Ldap='givenName';                 Type='string';     Default=$true  }
    @{ Key='Surname';        Header='Nachname';         Group='Identitaet';   Ldap='sn';                        Type='string';     Default=$true  }
    @{ Key='DisplayName';    Header='Anzeigename';      Group='Identitaet';   Ldap='displayName';               Type='string';     Default=$false }
    @{ Key='CN';             Header='Name (CN)';        Group='Identitaet';   Ldap='cn';                        Type='string';     Default=$false }
    @{ Key='SamAccountName'; Header='SamAccountName';   Group='Identitaet';   Ldap='sAMAccountName';            Type='string';     Default=$true  }
    @{ Key='UPN';            Header='UserPrincipalName';Group='Identitaet';   Ldap='userPrincipalName';         Type='string';     Default=$false }
    @{ Key='Mail';           Header='E-Mail';           Group='Identitaet';   Ldap='mail';                      Type='string';     Default=$true  }
    @{ Key='EmployeeID';     Header='EmployeeID';       Group='Identitaet';   Ldap='employeeID';                Type='string';     Default=$true  }

    @{ Key='Department';     Header='Abteilung';        Group='Organisation'; Ldap='department';                Type='string';     Default=$false }
    @{ Key='Title';          Header='Titel';            Group='Organisation'; Ldap='title';                     Type='string';     Default=$false }
    @{ Key='Company';        Header='Firma';            Group='Organisation'; Ldap='company';                   Type='string';     Default=$false }
    @{ Key='Manager';        Header='Manager (Name)';   Group='Organisation'; Ldap='manager';                   Type='manager';    Default=$false }
    @{ Key='Office';         Header='Buero';            Group='Organisation'; Ldap='physicalDeliveryOfficeName';Type='string';     Default=$false }
    @{ Key='Description';    Header='Beschreibung';     Group='Organisation'; Ldap='description';               Type='string';     Default=$false }

    @{ Key='OfficePhone';    Header='Telefon';          Group='Kontakt';      Ldap='telephoneNumber';           Type='string';     Default=$false }
    @{ Key='MobilePhone';    Header='Mobil';            Group='Kontakt';      Ldap='mobile';                    Type='string';     Default=$false }

    @{ Key='Street';         Header='Strasse';          Group='Adresse';      Ldap='streetAddress';             Type='string';     Default=$true  }
    @{ Key='City';           Header='Stadt';            Group='Adresse';      Ldap='l';                         Type='string';     Default=$true  }
    @{ Key='PostCode';       Header='PLZ';              Group='Adresse';      Ldap='postalCode';                Type='string';     Default=$true  }
    @{ Key='State';          Header='Bundesland';       Group='Adresse';      Ldap='st';                        Type='string';     Default=$false }
    @{ Key='CountryName';    Header='Land';             Group='Adresse';      Ldap='co';                        Type='string';     Default=$true  }
    @{ Key='CountryISO';     Header='Land (ISO)';       Group='Adresse';      Ldap='c';                         Type='string';     Default=$false }

    @{ Key='Enabled';        Header='Aktiviert';        Group='Konto/Status'; Ldap='userAccountControl';        Type='enabled';    Default=$false }
    @{ Key='LastLogonDate';  Header='Letzte Anmeldung'; Group='Konto/Status'; Ldap='lastLogonTimestamp';        Type='date';       Default=$false }
    @{ Key='PwdLastSet';     Header='Passwort gesetzt'; Group='Konto/Status'; Ldap='pwdLastSet';                Type='date';       Default=$false }
    @{ Key='Created';        Header='Erstellt';         Group='Konto/Status'; Ldap='whenCreated';               Type='datedirect'; Default=$false }
    @{ Key='AcctExpires';    Header='Konto laeuft ab';  Group='Konto/Status'; Ldap='accountExpires';            Type='date';       Default=$false }
    @{ Key='MemberOf'; Header='Gruppen (direkt)';       Group='Konto/Status'; Ldap='memberOf';                  Type='memberof';   Default=$false }
)

# Operatoren fuer den Filter-Builder (Label -> interner Code)
$script:FilterOps = [ordered]@{
    'ist gleich'      = 'equals'
    'enthaelt'        = 'contains'
    'beginnt mit'     = 'starts'
    'endet mit'       = 'ends'
    'ist nicht gleich'= 'notequals'
    'ist vorhanden'   = 'present'
    'ist leer'        = 'notpresent'
    'RAW (LDAP)'      = 'raw'
}

# --- Gemeinsamer (thread-sicherer) Zustand UI <-> Worker --------------------
$sync = [hashtable]::Synchronized(@{})
$sync.Log         = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
$sync.Percent     = 0
$sync.Status      = "Bereit."
$sync.Done        = $false
$sync.Success     = $false
$sync.ErrorMsg    = $null
$sync.ResultPath  = $null
$sync.ResultCount = 0

# --- Worker (separater Runspace, beruehrt nur $sync) ------------------------
# WICHTIG: Funktionen aus dem Hauptskript sind im Runspace NICHT verfuegbar.
# Alle Helfer muessen daher INNERHALB dieses Scriptblocks definiert werden.
$worker = {

    function Add-Log([string]$m) {
        [void]$sync.Log.Add(("[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $m))
    }

    # RFC-4515-Escaping. allowWildcard=$false escaped auch '*'.
    function Convert-LdapValue([string]$v, [bool]$allowWildcard) {
        if ($null -eq $v) { return '' }
        $sb = New-Object System.Text.StringBuilder
        foreach ($ch in $v.ToCharArray()) {
            switch ($ch) {
                '\'  { [void]$sb.Append('\5c') }
                '('  { [void]$sb.Append('\28') }
                ')'  { [void]$sb.Append('\29') }
                "`0" { [void]$sb.Append('\00') }
                '*'  { if ($allowWildcard) { [void]$sb.Append('*') } else { [void]$sb.Append('\2a') } }
                default { [void]$sb.Append($ch) }
            }
        }
        $sb.ToString()
    }

    # Baut eine einzelne LDAP-Filterklausel.
    function Build-Clause([string]$ldap, [string]$op, [string]$val, [string]$type) {
        if ($type -eq 'enabled') {
            $truthy    = @('true','wahr','1','yes','ja','aktiv','enabled')
            $isEnabled = $truthy -contains ($val.ToString().Trim().ToLower())
            if ($op -eq 'notequals') { $isEnabled = -not $isEnabled }
            if ($isEnabled) { return '(!(userAccountControl:1.2.840.113556.1.4.803:=2))' }
            else            { return '(userAccountControl:1.2.840.113556.1.4.803:=2)' }
        }
        $e = Convert-LdapValue $val $false
        switch ($op) {
            'equals'     { "($ldap=$e)" }
            'notequals'  { "(!($ldap=$e))" }
            'contains'   { "($ldap=*$e*)" }
            'starts'     { "($ldap=$e*)" }
            'ends'       { "($ldap=*$e)" }
            'present'    { "($ldap=*)" }
            'notpresent' { "(!($ldap=*))" }
            'raw'        { "($ldap=$val)" }   # bewusst NICHT escaped - Power-User
            default      { "($ldap=$e)" }
        }
    }

    function Get-Val($props, [string]$name) {
        if ($props.Contains($name) -and $props[$name].Count -gt 0) { return $props[$name][0] }
        return $null
    }

    # FILETIME Int64 -> lokale DateTime, mit Never-Sentinel-Guard (0 / Int64::MaxValue).
    function Convert-FileTime($raw) {
        if ($null -eq $raw) { return $null }
        try { $i = [int64]$raw } catch { return $null }
        if ($i -le 0 -or $i -ge 9223372036854775807) { return $null }
        try { return [DateTime]::FromFileTime($i) } catch { return $null }
    }

    function Get-CnFromDn([string]$dn) {
        if (-not $dn) { return $null }
        $cn = [regex]::Match($dn, '^CN=(?<cn>(?:[^,\\]|\\.)*)').Groups['cn'].Value
        return ($cn -replace '\\(.)', '$1')
    }

    function Get-GuidString($raw) {
        if ($null -eq $raw) { return $null }
        try { return (New-Object Guid (,[byte[]]$raw)).ToString() } catch { return $null }
    }

    # Erzeugt einen DirectoryEntry-Root fuer die Suche (serverless wenn $Server leer).
    function New-RootEntry([string]$Server, [string]$BaseDn) {
        if ([string]::IsNullOrWhiteSpace($BaseDn)) {
            $rootDsePath = if ($Server) { "LDAP://$Server/RootDSE" } else { "LDAP://RootDSE" }
            $rootDse = New-Object System.DirectoryServices.DirectoryEntry($rootDsePath)
            $BaseDn  = [string]$rootDse.Properties["defaultNamingContext"][0]
            try { $rootDse.Dispose() } catch {}
            if (-not $BaseDn) { throw "defaultNamingContext konnte nicht ermittelt werden (Domain erreichbar?)." }
        }
        $path = if ($Server) { "LDAP://$Server/$BaseDn" } else { "LDAP://$BaseDn" }
        $entry = New-Object System.DirectoryServices.DirectoryEntry($path)
        # Erzwingt eine echte Bindung -> wirft frueh bei Verbindungs-/Auth-Fehlern.
        $null = $entry.NativeObject
        return [pscustomobject]@{ Entry = $entry; BaseDn = $BaseDn }
    }

    function New-Searcher($RootEntry, [string]$Filter, [string[]]$Load) {
        $ds = New-Object System.DirectoryServices.DirectorySearcher($RootEntry)
        $ds.Filter      = $Filter
        $ds.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
        $ds.PageSize    = 1000      # KRITISCH: ohne Paging max. 1000 Treffer
        $ds.SizeLimit   = 0
        $ds.PropertiesToLoad.Clear()
        foreach ($p in $Load) { [void]$ds.PropertiesToLoad.Add($p) }
        return $ds
    }

    function Resolve-GroupDn($RootEntry, [string]$nameOrDn) {
        # Bereits ein DN? (heuristisch)
        if ($nameOrDn -match '^(CN|OU|DC)=.+,(DC|OU|CN)=') { return $nameOrDn }
        $e  = Convert-LdapValue $nameOrDn $false
        $ds = New-Searcher $RootEntry "(&(objectClass=group)(|(cn=$e)(sAMAccountName=$e)(name=$e)))" @('distinguishedName')
        $ds.SizeLimit = 2
        $r = $ds.FindOne()
        try { $ds.Dispose() } catch {}
        if ($r) { return [string]$r.Properties['distinguishedName'][0] }
        return $null
    }

    function Get-RecursiveMemberGuids($RootEntry, [string]$groupDn) {
        $set = New-Object 'System.Collections.Generic.HashSet[string]'
        $f   = "(&(objectCategory=person)(objectClass=user)(memberOf:1.2.840.113556.1.4.1941:=$groupDn))"
        $ds  = New-Searcher $RootEntry $f @('objectGUID')
        $res = $ds.FindAll()
        foreach ($r in $res) {
            $g = Get-GuidString (Get-Val $r.Properties 'objectGUID')
            if ($g) { [void]$set.Add($g) }
        }
        try { $res.Dispose() } catch {}
        try { $ds.Dispose() }  catch {}
        return $set
    }

    try {
        $sync.Status = "Lade ImportExcel..."
        Import-Module ImportExcel -ErrorAction Stop
        Add-Log "ImportExcel geladen."

        # --- Verbindung / Root --------------------------------------------------
        $sync.Status = "Verbinde via LDAP..."
        if ([string]::IsNullOrWhiteSpace($Server)) {
            Add-Log "Kein DC angegeben -> serverlose Bindung (DC-Locator)."
        } else {
            Add-Log "Ziel-DC: $Server"
        }
        $rootInfo  = New-RootEntry $Server $BaseDn
        $rootEntry = $rootInfo.Entry
        Add-Log "Such-Basis (DN): $($rootInfo.BaseDn)"

        # --- Subjekt-Klausel je nach Modus -------------------------------------
        $core    = '(&(objectCategory=person)(objectClass=user)'
        $subject = ''

        switch ($Mode) {
            'all' {
                Add-Log "Modus: ALLE Benutzer."
            }
            'group' {
                Add-Log "Modus: Mitglieder von Gruppe(n) (rekursiv)."
                $dns = @()
                foreach ($g in $GroupFilterNames) {
                    $dn = Resolve-GroupDn $rootEntry $g
                    if ($dn) { $dns += $dn; Add-Log "  Gruppe aufgeloest: $g" }
                    else     { Add-Log "  WARN: Gruppe NICHT gefunden: $g" }
                }
                if (-not $dns) { throw "Keine der angegebenen Filter-Gruppen wurde gefunden." }
                $clauses = @($dns | ForEach-Object { "(memberOf:1.2.840.113556.1.4.1941:=$_)" })
                if ($clauses.Count -eq 1) {
                    $subject = $clauses[0]
                } elseif ($GroupMatchAny) {
                    $subject = '(|' + (-join $clauses) + ')'   # Mitglied in MIND. EINER
                } else {
                    $subject = '(&' + (-join $clauses) + ')'   # Mitglied in ALLEN
                }
            }
            'users' {
                Add-Log "Modus: Spezifische Benutzer ($($Identifiers.Count) IDs)."
                $clauses = @()
                foreach ($id in $Identifiers) {
                    $e = Convert-LdapValue $id $false
                    $clauses += "(|(sAMAccountName=$e)(userPrincipalName=$e)(mail=$e)(cn=$e))"
                }
                if (-not $clauses) { throw "Keine Benutzer-IDs angegeben." }
                $subject = if ($clauses.Count -eq 1) { $clauses[0] } else { '(|' + (-join $clauses) + ')' }
            }
        }

        # --- Attribut-Filter (AND) ---------------------------------------------
        $attrClauses = @()
        foreach ($f in $AttributeFilters) {
            $c = Build-Clause $f.Ldap $f.Op $f.Val $f.Type
            $attrClauses += $c
            Add-Log "  Filter: $c"
        }
        $attrPart = -join $attrClauses

        $filter = $core + $subject + $attrPart + ')'
        Add-Log "LDAP-Filter: $filter"

        # --- Properties zum Laden zusammenstellen ------------------------------
        $loadSet = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($a in $SelectedAttributes) { [void]$loadSet.Add($a.Ldap) }
        foreach ($n in @('objectGUID','distinguishedName','sn','givenName')) { [void]$loadSet.Add($n) }
        $load = @($loadSet)
        Add-Log ("Geladene Attribute: " + ($load -join ', '))

        # --- Gruppen-Spalten vorbereiten (rekursive Mitglieds-Sets) ------------
        $columnSets = @{}
        if ($IncludeColumns -and $ColumnGroups.Count -gt 0) {
            $sync.Status = "Lese Gruppen-Mitgliedschaften..."
            foreach ($g in $ColumnGroups) {
                $dn = Resolve-GroupDn $rootEntry $g
                if (-not $dn) { Add-Log "WARN: Spalten-Gruppe NICHT gefunden: $g"; continue }
                $columnSets[$g] = Get-RecursiveMemberGuids $rootEntry $dn
                Add-Log "  Spalte '$g': $($columnSets[$g].Count) Mitglieder (rekursiv)"
            }
        }

        # --- Hauptsuche --------------------------------------------------------
        $sync.Status = "Fuehre LDAP-Suche aus..."
        $ds  = New-Searcher $rootEntry $filter $load
        $res = $ds.FindAll()
        $total = $res.Count
        Add-Log "Treffer (vor Dedup): $total"
        if ($total -eq 0) { throw "Keine Benutzer im Ergebnis - es wurde keine Datei erzeugt." }

        $seen     = New-Object 'System.Collections.Generic.HashSet[string]'
        $UserList = New-Object System.Collections.ArrayList
        $i = 0

        foreach ($r in $res) {
            $i++
            $sync.Percent = [math]::Round(($i / $total) * 100)
            $sync.Status  = "Verarbeite $i / $total"

            $props = $r.Properties
            $guid  = Get-GuidString (Get-Val $props 'objectGUID')
            if ($guid -and $seen.Contains($guid)) { continue }
            if ($guid) { [void]$seen.Add($guid) }

            $record = [ordered]@{}
            foreach ($a in $SelectedAttributes) {
                switch ($a.Type) {
                    'enabled' {
                        $uac = Get-Val $props 'userAccountControl'
                        $v = if ($null -ne $uac) { -not ([int]$uac -band 2) } else { $null }
                    }
                    'date'       { $v = Convert-FileTime (Get-Val $props $a.Ldap) }
                    'datedirect' {
                        $d = Get-Val $props $a.Ldap
                        $v = if ($d -is [datetime]) { ([datetime]$d).ToLocalTime() } else { $d }
                    }
                    'manager'    { $v = Get-CnFromDn ([string](Get-Val $props $a.Ldap)) }
                    'memberof'   {
                                   $v = $null
                                        if ($props.Contains($a.Ldap) -and $props[$a.Ldap].Count -gt 0) 
                                        {
                                            $names = foreach ($dn in $props[$a.Ldap]) { Get-CnFromDn ([string]$dn) 
                                        }
                                        $v = ($names | Sort-Object) -join '; '
                               }
}
                    default      { $v = Get-Val $props $a.Ldap }
                }
                $record[$a.Header] = $v
            }

            if ($IncludeColumns) {
                foreach ($g in $ColumnGroups) {
                    $col = 'Mitglied von "{0}"' -f $g
                    if ($columnSets.ContainsKey($g)) {
                        $record[$col] = [bool]($guid -and $columnSets[$g].Contains($guid))
                    } else {
                        $record[$col] = $null
                    }
                }
            }

            [void]$UserList.Add([PSCustomObject]$record)
        }

        try { $res.Dispose() } catch {}
        try { $ds.Dispose() }  catch {}

        # --- Sortierung (nur wenn Spalten vorhanden) ---------------------------
        $headers   = @($SelectedAttributes | ForEach-Object { $_.Header })
        $sortProps = @()
        if ($headers -contains 'Nachname') { $sortProps += 'Nachname' }
        if ($headers -contains 'Vorname')  { $sortProps += 'Vorname' }
        $AllUsers = if ($sortProps.Count) { $UserList | Sort-Object $sortProps } else { @($UserList) }

        if (-not $AllUsers -or @($AllUsers).Count -eq 0) {
            throw "Keine Benutzer nach Verarbeitung - es wurde keine Datei erzeugt."
        }

        # --- Export ------------------------------------------------------------
        New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
        $ExcelFullPath = Join-Path $ExportPath $ExcelFile

        $sync.Status = "Schreibe Excel..."
        Add-Log "Exportiere $(@($AllUsers).Count) Benutzer nach Excel..."

        $AllUsers | Export-Excel `
            -Path $ExcelFullPath `
            -WorksheetName "Argus-Export" `
            -TableName "ArgusExport" `
            -TableStyle Medium2 `
            -AutoSize `
            -FreezeTopRow

        try { $rootEntry.Dispose() } catch {}

        $sync.ResultPath  = $ExcelFullPath
        $sync.ResultCount = @($AllUsers).Count
        $sync.Success     = $true
        $sync.Percent     = 100
        $sync.Status      = "Fertig: $(@($AllUsers).Count) Benutzer exportiert."
        Add-Log "FERTIG. Datei: $ExcelFullPath"
    }
    catch {
        $sync.Success  = $false
        $sync.ErrorMsg = $_.Exception.Message
        $sync.Status   = "Fehler: $($_.Exception.Message)"
        Add-Log "FEHLER: $($_.Exception.Message)"
    }
    finally {
        $sync.Done = $true
    }
}

# --- XAML -------------------------------------------------------------------
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Argus - AD Export (LDAP)"
        Height="860" Width="780"
        WindowStartupLocation="CenterScreen"
        Background="#F8FAFC" FontFamily="Segoe UI" FontSize="13"
        ResizeMode="CanResizeWithGrip" MinWidth="720" MinHeight="680">

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
            <Border x:Name="bd" Background="#4F46E5" CornerRadius="8" Padding="18,10">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#4338CA"/></Trigger>
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
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Background="#0F172A" Padding="24,18">
      <StackPanel>
        <TextBlock Text="Argus" Foreground="White" FontSize="20" FontWeight="Bold"/>
        <TextBlock Text="Active-Directory-Export via LDAP (Port 389) &#8226; kein ADWS &#8226; dynamische Filter"
                   Foreground="#94A3B8" FontSize="12" Margin="0,3,0,0"/>
      </StackPanel>
    </Border>

    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Padding="20,16,20,4">
      <StackPanel>

        <!-- Verbindung -->
        <Border Style="{StaticResource Card}">
          <StackPanel>
            <TextBlock Text="Verbindung" Style="{StaticResource CardTitle}"/>
            <TextBlock Text="Domain Controller (Hostname / FQDN / IP)" Style="{StaticResource Lbl}"/>
            <TextBox x:Name="txtDC" Style="{StaticResource Inp}" Text=""/>
            <TextBlock Text="Leer lassen = serverlose Bindung. Die Domaene wird automatisch ueber den DC-Locator gefunden."
                       Style="{StaticResource Hint}"/>
            <TextBlock Text="Such-Basis / Base-DN (optional)" Style="{StaticResource Lbl}"/>
            <TextBox x:Name="txtBaseDn" Style="{StaticResource Inp}" Text=""/>
            <TextBlock Text="z. B. OU=Benutzer,DC=firma,DC=local  -  leer = gesamte Domaene (defaultNamingContext)."
                       Style="{StaticResource Hint}"/>
          </StackPanel>
        </Border>

        <!-- Quelle / Modus -->
        <Border Style="{StaticResource Card}">
          <StackPanel>
            <TextBlock Text="Quelle" Style="{StaticResource CardTitle}"/>
            <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
              <RadioButton x:Name="rbAll"   Content="Alle Benutzer" IsChecked="True" GroupName="src" Margin="0,0,18,0" Foreground="#334155"/>
              <RadioButton x:Name="rbGroup" Content="Mitglieder von Gruppe(n)" GroupName="src" Margin="0,0,18,0" Foreground="#334155"/>
              <RadioButton x:Name="rbUsers" Content="Spezifische Benutzer" GroupName="src" Foreground="#334155"/>
            </StackPanel>

            <!-- Panel: Gruppen-Filter -->
            <StackPanel x:Name="pnlGroupSrc" Visibility="Collapsed" Margin="0,10,0,0">
              <TextBlock Text="Filter-Gruppen (eine pro Zeile - CN, sAMAccountName oder DN)" Style="{StaticResource Lbl}"/>
              <TextBox x:Name="txtGroupSrc" Style="{StaticResource Inp}" Height="58"
                       AcceptsReturn="True" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto"/>
              <CheckBox x:Name="chkGroupAny" Content="Mitglied in MINDESTENS einer (statt in allen)" IsChecked="True"
                        Margin="0,8,0,0" Foreground="#334155"/>
              <TextBlock Text="Rekursiv (verschachtelte Gruppen) via LDAP_MATCHING_RULE_IN_CHAIN."
                         Style="{StaticResource Hint}" Margin="0,4,0,0"/>
            </StackPanel>

            <!-- Panel: Spezifische User -->
            <StackPanel x:Name="pnlUserSrc" Visibility="Collapsed" Margin="0,10,0,0">
              <TextBlock Text="Benutzer (einer pro Zeile - sAMAccountName, UPN, E-Mail oder CN)" Style="{StaticResource Lbl}"/>
              <TextBox x:Name="txtUserSrc" Style="{StaticResource Inp}" Height="80"
                       AcceptsReturn="True" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto"/>
            </StackPanel>
          </StackPanel>
        </Border>

        <!-- Filter-Builder -->
        <Border Style="{StaticResource Card}">
          <StackPanel>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <StackPanel Grid.Column="0">
                <TextBlock Text="Attribut-Filter" Style="{StaticResource CardTitle}"/>
                <TextBlock Text="Beliebig viele Bedingungen, alle UND-verknuepft. Beispiel: PLZ - ist gleich - 44147."
                           Style="{StaticResource Hint}" Margin="0,2,0,0"/>
              </StackPanel>
              <Button x:Name="btnAddFilter" Grid.Column="1" Content="+ Filter" Style="{StaticResource Mini}" VerticalAlignment="Top"/>
            </Grid>
            <StackPanel x:Name="pnlFilters" Margin="0,10,0,0"/>
          </StackPanel>
        </Border>

        <!-- Felder -->
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
                <Button x:Name="btnSelectAll"  Content="Alle"  Style="{StaticResource Mini}"/>
                <Button x:Name="btnSelectNone" Content="Keine" Style="{StaticResource Mini}" Margin="6,0,0,0"/>
              </StackPanel>
            </Grid>
            <StackPanel x:Name="pnlAttributes" Margin="0,8,0,0"/>
            <TextBlock Text="Weitere LDAP-Attribute (komma-getrennt, exakte LDAP-Namen)" Style="{StaticResource Lbl}"/>
            <TextBox x:Name="txtCustomAttrs" Style="{StaticResource Inp}" Text=""/>
            <TextBlock Text="Unbekannte Attribute liefern via LDAP keinen Fehler, sondern bleiben leer."
                       Style="{StaticResource Hint}"/>
          </StackPanel>
        </Border>

        <!-- Gruppen-Spalten -->
        <Border Style="{StaticResource Card}">
          <StackPanel>
            <TextBlock Text="Gruppen-Mitgliedschaft als Spalten" Style="{StaticResource CardTitle}"/>
            <CheckBox x:Name="chkColumns" Content="Pro Gruppe eine True/False-Spalte anhaengen"
                      Margin="0,8,0,0" Foreground="#334155"/>
            <TextBlock Text="Gruppen (eine pro Zeile) - unabhaengig vom Quell-Filter oben" Style="{StaticResource Lbl}"/>
            <TextBox x:Name="txtColumnGroups" Style="{StaticResource Inp}" Height="58"
                     AcceptsReturn="True" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto"/>
            <TextBlock Text="Erzeugt je Gruppe eine Spalte 'Mitglied von &quot;...&quot;'. Rekursive Pruefung."
                       Style="{StaticResource Hint}"/>
          </StackPanel>
        </Border>

        <!-- Ausgabe -->
        <Border Style="{StaticResource Card}">
          <StackPanel>
            <TextBlock Text="Ausgabe" Style="{StaticResource CardTitle}"/>
            <TextBlock Text="Export-Ordner" Style="{StaticResource Lbl}"/>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBox x:Name="txtPath" Grid.Column="0" Style="{StaticResource Inp}" Text="C:\Temp"/>
              <Button x:Name="btnBrowse" Grid.Column="1" Content="Durchsuchen..." Style="{StaticResource Secondary}" Margin="8,0,0,0"/>
            </Grid>
            <TextBlock Text="Dateiname (.xlsx)" Style="{StaticResource Lbl}"/>
            <TextBox x:Name="txtFile" Style="{StaticResource Inp}" Text="Argus-Export.xlsx"/>
          </StackPanel>
        </Border>

        <Button x:Name="btnStart" Content="Export starten"
                Style="{StaticResource Accent}" HorizontalAlignment="Left" Margin="0,2,0,10"/>
      </StackPanel>
    </ScrollViewer>

    <StackPanel Grid.Row="2" Margin="20,4,20,0">
      <ProgressBar x:Name="pbProgress" Height="18" Minimum="0" Maximum="100"
                   Foreground="#4F46E5" Background="#E2E8F0" BorderThickness="0"/>
      <TextBlock x:Name="lblStatus" Text="Bereit." Foreground="#334155"
                 Margin="0,8,0,0" TextTrimming="CharacterEllipsis"/>
    </StackPanel>

    <Border Grid.Row="3" Margin="20,10,20,0" CornerRadius="10" Background="#0F172A" Height="130">
      <TextBox x:Name="txtLog" Background="Transparent" Foreground="#A7F3D0"
               FontFamily="Consolas" FontSize="12" BorderThickness="0" IsReadOnly="True"
               Padding="12" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap"/>
    </Border>

    <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="20,12,20,16">
      <Button x:Name="btnOpenFile"   Content="Excel oeffnen"  Style="{StaticResource Secondary}" IsEnabled="False"/>
      <Button x:Name="btnOpenFolder" Content="Ordner oeffnen" Style="{StaticResource Secondary}" IsEnabled="False" Margin="8,0,0,0"/>
    </StackPanel>

  </Grid>
</Window>
'@

# --- Fenster laden ----------------------------------------------------------
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$txtDC          = $window.FindName('txtDC')
$txtBaseDn      = $window.FindName('txtBaseDn')
$rbAll          = $window.FindName('rbAll')
$rbGroup        = $window.FindName('rbGroup')
$rbUsers        = $window.FindName('rbUsers')
$pnlGroupSrc    = $window.FindName('pnlGroupSrc')
$pnlUserSrc     = $window.FindName('pnlUserSrc')
$txtGroupSrc    = $window.FindName('txtGroupSrc')
$chkGroupAny    = $window.FindName('chkGroupAny')
$txtUserSrc     = $window.FindName('txtUserSrc')
$btnAddFilter   = $window.FindName('btnAddFilter')
$pnlFilters     = $window.FindName('pnlFilters')
$pnlAttributes  = $window.FindName('pnlAttributes')
$txtCustomAttrs = $window.FindName('txtCustomAttrs')
$btnSelectAll   = $window.FindName('btnSelectAll')
$btnSelectNone  = $window.FindName('btnSelectNone')
$chkColumns     = $window.FindName('chkColumns')
$txtColumnGroups= $window.FindName('txtColumnGroups')
$txtPath        = $window.FindName('txtPath')
$btnBrowse      = $window.FindName('btnBrowse')
$txtFile        = $window.FindName('txtFile')
$btnStart       = $window.FindName('btnStart')
$pbProgress     = $window.FindName('pbProgress')
$lblStatus      = $window.FindName('lblStatus')
$txtLog         = $window.FindName('txtLog')
$btnOpenFile    = $window.FindName('btnOpenFile')
$btnOpenFolder  = $window.FindName('btnOpenFolder')

# --- Feld-Checkboxen aus dem Katalog generieren -----------------------------
$script:AttrCheckBoxes = @{}
$currentGroup = $null
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
    }
    $cb = New-Object System.Windows.Controls.CheckBox
    $cb.Content    = $entry.Header
    $cb.Tag        = $entry.Key
    $cb.IsChecked  = [bool]$entry.Default
    $cb.Margin     = '8,2,0,2'
    $cb.Foreground = '#334155'
    [void]$pnlAttributes.Children.Add($cb)
    $script:AttrCheckBoxes[$entry.Key] = $cb
}

# --- Filter-Builder: dynamische Zeilen --------------------------------------
# Attribut-Auswahl: Header -> Katalog-Eintrag (zzgl. RAW-Custom moeglich)
$script:HeaderToEntry = @{}
foreach ($e in $script:AttributeCatalog) { $script:HeaderToEntry[$e.Header] = $e }
$script:FilterAttrHeaders = @($script:AttributeCatalog | ForEach-Object { $_.Header })

function Add-FilterRow {
    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = '0,0,0,6'
    $c1 = New-Object System.Windows.Controls.ColumnDefinition; $c1.Width = '2*'
    $c2 = New-Object System.Windows.Controls.ColumnDefinition; $c2.Width = '1.4*'
    $c3 = New-Object System.Windows.Controls.ColumnDefinition; $c3.Width = '2*'
    $c4 = New-Object System.Windows.Controls.ColumnDefinition; $c4.Width = 'Auto'
    $grid.ColumnDefinitions.Add($c1); $grid.ColumnDefinitions.Add($c2)
    $grid.ColumnDefinitions.Add($c3); $grid.ColumnDefinitions.Add($c4)

    $cbAttr = New-Object System.Windows.Controls.ComboBox
    foreach ($h in $script:FilterAttrHeaders) { [void]$cbAttr.Items.Add($h) }
    $cbAttr.SelectedIndex = 0
    $cbAttr.Margin = '0,0,6,0'
    [System.Windows.Controls.Grid]::SetColumn($cbAttr, 0)

    $cbOp = New-Object System.Windows.Controls.ComboBox
    foreach ($opLabel in $script:FilterOps.Keys) { [void]$cbOp.Items.Add($opLabel) }
    $cbOp.SelectedIndex = 0
    $cbOp.Margin = '0,0,6,0'
    [System.Windows.Controls.Grid]::SetColumn($cbOp, 1)

    $txtVal = New-Object System.Windows.Controls.TextBox
    $txtVal.Padding = '8,6'
    $txtVal.BorderBrush = '#CBD5E1'
    $txtVal.VerticalContentAlignment = 'Center'
    $txtVal.Margin = '0,0,6,0'
    [System.Windows.Controls.Grid]::SetColumn($txtVal, 2)

    $btnDel = New-Object System.Windows.Controls.Button
    $btnDel.Content = 'X'
    $btnDel.Style = $window.FindResource('Danger')
    $btnDel.Tag = $grid
    $btnDel.Add_Click({ $pnlFilters.Children.Remove($this.Tag) })
    [System.Windows.Controls.Grid]::SetColumn($btnDel, 3)

    [void]$grid.Children.Add($cbAttr)
    [void]$grid.Children.Add($cbOp)
    [void]$grid.Children.Add($txtVal)
    [void]$grid.Children.Add($btnDel)

    # Zugriff beim Start: Controls im Grid.Tag ablegen
    $grid.Tag = @{ Attr = $cbAttr; Op = $cbOp; Val = $txtVal }
    [void]$pnlFilters.Children.Add($grid)
}

# --- DispatcherTimer --------------------------------------------------------
$script:logIndex = 0
$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(200)
$timer.Add_Tick({
    $pbProgress.Value = [double]$sync.Percent
    $lblStatus.Text   = [string]$sync.Status
    while ($script:logIndex -lt $sync.Log.Count) {
        $txtLog.AppendText($sync.Log[$script:logIndex] + "`r`n")
        $script:logIndex++
    }
    $txtLog.ScrollToEnd()
    if ($sync.Done) {
        $timer.Stop()
        try { $script:ps.EndInvoke($script:handle) } catch {}
        try { $script:ps.Dispose() }                catch {}
        try { $script:runspace.Close(); $script:runspace.Dispose() } catch {}
        $btnStart.IsEnabled = $true
        if ($sync.Success) {
            $btnOpenFile.IsEnabled   = $true
            $btnOpenFolder.IsEnabled = $true
        }
        elseif ($sync.ErrorMsg) {
            [System.Windows.MessageBox]::Show($sync.ErrorMsg, "Fehler beim Export", 'OK', 'Error') | Out-Null
        }
    }
})

# --- Events -----------------------------------------------------------------
$updateSrcPanels = {
    $pnlGroupSrc.Visibility = if ($rbGroup.IsChecked) { 'Visible' } else { 'Collapsed' }
    $pnlUserSrc.Visibility  = if ($rbUsers.IsChecked) { 'Visible' } else { 'Collapsed' }
}
$rbAll.Add_Checked($updateSrcPanels)
$rbGroup.Add_Checked($updateSrcPanels)
$rbUsers.Add_Checked($updateSrcPanels)

$btnAddFilter.Add_Click({ Add-FilterRow })

$btnSelectAll.Add_Click({  foreach ($cb in $script:AttrCheckBoxes.Values) { $cb.IsChecked = $true } })
$btnSelectNone.Add_Click({ foreach ($cb in $script:AttrCheckBoxes.Values) { $cb.IsChecked = $false } })

$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($txtPath.Text) { $dlg.SelectedPath = $txtPath.Text }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtPath.Text = $dlg.SelectedPath }
})

$btnStart.Add_Click({
    $dc       = $txtDC.Text.Trim()
    $baseDn   = $txtBaseDn.Text.Trim()
    $path     = $txtPath.Text.Trim()
    $file     = $txtFile.Text.Trim()
    $custom   = @($txtCustomAttrs.Text -split '[,;\r\n]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })

    # Modus
    $mode = if ($rbGroup.IsChecked) { 'group' } elseif ($rbUsers.IsChecked) { 'users' } else { 'all' }
    $groupSrc = @($txtGroupSrc.Text -split "\r?\n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $userSrc  = @($txtUserSrc.Text  -split "\r?\n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })

    # Spalten
    $includeCols  = ($chkColumns.IsChecked -eq $true)
    $columnGroups = @($txtColumnGroups.Text -split "\r?\n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })

    # Ausgewaehlte Felder (Katalogreihenfolge) + Custom
    $selected = New-Object System.Collections.Generic.List[object]
    $used     = New-Object System.Collections.Generic.HashSet[string]
    foreach ($entry in $script:AttributeCatalog) {
        if ($script:AttrCheckBoxes[$entry.Key].IsChecked -eq $true) {
            [void]$selected.Add(@{ Header=$entry.Header; Ldap=$entry.Ldap; Type=$entry.Type })
            [void]$used.Add($entry.Header)
        }
    }
    foreach ($c in $custom) {
        if ($used.Contains($c)) { continue }
        [void]$selected.Add(@{ Header=$c; Ldap=$c; Type='string' }); [void]$used.Add($c)
    }
    $selectedArr = $selected.ToArray()

    # Filterzeilen einlesen
    $filters = New-Object System.Collections.Generic.List[object]
    foreach ($row in $pnlFilters.Children) {
        $h = $row.Tag
        if (-not $h) { continue }
        $hdr = [string]$h.Attr.SelectedItem
        $opL = [string]$h.Op.SelectedItem
        $val = [string]$h.Val.Text
        if (-not $hdr -or -not $opL) { continue }
        $opCode = $script:FilterOps[$opL]
        if (($opCode -ne 'present') -and ($opCode -ne 'notpresent') -and ([string]::IsNullOrWhiteSpace($val))) { continue }
        $entry = $script:HeaderToEntry[$hdr]
        $ldap  = if ($entry) { $entry.Ldap } else { $hdr }
        $type  = if ($entry) { $entry.Type } else { 'string' }
        [void]$filters.Add(@{ Ldap=$ldap; Op=$opCode; Val=$val; Type=$type })
    }
    $filterArr = $filters.ToArray()

    # Validierung
    if ($mode -eq 'group' -and -not $groupSrc) {
        [System.Windows.MessageBox]::Show("Bitte mindestens eine Filter-Gruppe angeben.","Eingabe fehlt",'OK','Warning') | Out-Null; return }
    if ($mode -eq 'users' -and -not $userSrc) {
        [System.Windows.MessageBox]::Show("Bitte mindestens einen Benutzer angeben.","Eingabe fehlt",'OK','Warning') | Out-Null; return }
    if (-not $path) {
        [System.Windows.MessageBox]::Show("Bitte einen Export-Ordner angeben.","Eingabe fehlt",'OK','Warning') | Out-Null; return }
    if ($includeCols -and -not $columnGroups) {
        [System.Windows.MessageBox]::Show("Gruppen-Spalten aktiviert, aber keine Gruppen angegeben.","Eingabe fehlt",'OK','Warning') | Out-Null; return }
    if ($selectedArr.Count -eq 0 -and -not $includeCols) {
        [System.Windows.MessageBox]::Show("Bitte mindestens ein Feld oder die Gruppen-Spalten waehlen.","Keine Auswahl",'OK','Warning') | Out-Null; return }
    if (-not $file) { $file = "Argus-Export.xlsx" }
    if ($file -notmatch '\.xlsx$') { $file += '.xlsx' }

    # Zustand zuruecksetzen
    $sync.Percent = 0; $sync.Status = "Starte..."; $sync.Done = $false; $sync.Success = $false
    $sync.ErrorMsg = $null; $sync.ResultPath = $null; $sync.ResultCount = 0
    $sync.Log.Clear(); $script:logIndex = 0
    $txtLog.Clear(); $pbProgress.Value = 0
    $btnStart.IsEnabled = $false; $btnOpenFile.IsEnabled = $false; $btnOpenFolder.IsEnabled = $false

    # Runspace + Worker
    $script:runspace = [runspacefactory]::CreateRunspace()
    $script:runspace.ThreadOptions = 'ReuseThread'
    $script:runspace.ApartmentState = 'MTA'
    $script:runspace.Open()
    $script:runspace.SessionStateProxy.SetVariable('sync',              $sync)
    $script:runspace.SessionStateProxy.SetVariable('Server',            $dc)
    $script:runspace.SessionStateProxy.SetVariable('BaseDn',            $baseDn)
    $script:runspace.SessionStateProxy.SetVariable('Mode',              $mode)
    $script:runspace.SessionStateProxy.SetVariable('GroupFilterNames',  $groupSrc)
    $script:runspace.SessionStateProxy.SetVariable('GroupMatchAny',     ($chkGroupAny.IsChecked -eq $true))
    $script:runspace.SessionStateProxy.SetVariable('Identifiers',       $userSrc)
    $script:runspace.SessionStateProxy.SetVariable('AttributeFilters',  $filterArr)
    $script:runspace.SessionStateProxy.SetVariable('SelectedAttributes',$selectedArr)
    $script:runspace.SessionStateProxy.SetVariable('IncludeColumns',    $includeCols)
    $script:runspace.SessionStateProxy.SetVariable('ColumnGroups',      $columnGroups)
    $script:runspace.SessionStateProxy.SetVariable('ExportPath',        $path)
    $script:runspace.SessionStateProxy.SetVariable('ExcelFile',         $file)

    $script:ps = [powershell]::Create()
    $script:ps.Runspace = $script:runspace
    [void]$script:ps.AddScript($worker.ToString())
    $script:handle = $script:ps.BeginInvoke()
    $timer.Start()
})

$btnOpenFile.Add_Click({   if ($sync.ResultPath -and (Test-Path $sync.ResultPath)) { Start-Process $sync.ResultPath } })
$btnOpenFolder.Add_Click({ if ($sync.ResultPath -and (Test-Path $sync.ResultPath)) { Start-Process explorer.exe "/select,`"$($sync.ResultPath)`"" } })

$window.Add_Closing({
    try { $timer.Stop() } catch {}
    try { if ($script:ps)       { $script:ps.Dispose() } }       catch {}
    try { if ($script:runspace) { $script:runspace.Dispose() } } catch {}
})

# Eine leere Filterzeile als Startpunkt
Add-FilterRow

[void]$window.ShowDialog()