# Funktionale Tests für Argus.ps1 (Engine + Settings) - läuft auch auf Linux pwsh,
# da nur die reinen .NET-Anteile aufgerufen werden (kein System.DirectoryServices).
$ErrorActionPreference = 'Stop'
$file = Join-Path (Split-Path -Parent $PSScriptRoot) 'Argus.ps1'

$tokens = $null; $errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($file, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw "Parse errors: $($errors.Count)" }

# --- Engine-Scriptblock extrahieren und dot-sourcen -------------------------
$engineAssign = $ast.Find({
    $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] -and
    $args[0].Left.Extent.Text -eq '$script:ArgusEngine'
}, $true)
if (-not $engineAssign) { throw 'ArgusEngine assignment not found' }
$sbText = $engineAssign.Right.Extent.Text
$body = $sbText.Substring(1, $sbText.Length - 2)   # äußere {} entfernen
. ([scriptblock]::Create($body))

# --- Kontext-Variablen, die die Settings-Funktionen erwarten ----------------
$script:ArgusVersion = '2.0.0-test'
$script:ArgusName = 'Argus'
$script:DefaultExportDir = [IO.Path]::GetTempPath()

# --- Katalog, FilterOps und Top-Level-Funktionen extrahieren ----------------
foreach ($varName in @('$script:AttributeCatalog', '$script:FilterOps')) {
    $assign = $ast.Find({
        $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] -and
        $args[0].Left.Extent.Text -eq $varName
    }, $true)
    if (-not $assign) { throw "$varName not found" }
    . ([scriptblock]::Create($assign.Extent.Text))
}
$wanted = @('Get-ArgusDefaultSettings', 'ConvertTo-ArgusHashtable', 'Merge-ArgusSettings', 'Read-ArgusProfile',
            'Save-ArgusProfile', 'Get-ArgusEffectivePort', 'Resolve-ArgusAttributeSelection', 'Resolve-ArgusFilterRows',
            'Get-ArgusSearchParams', 'Get-ArgusExportOptions', 'Get-ArgusExportMeta')
$funcs = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
foreach ($name in $wanted) {
    $f = $funcs | Where-Object { $_.Name -eq $name } | Select-Object -First 1
    if (-not $f) { throw "Function $name not found" }
    . ([scriptblock]::Create($f.Extent.Text))
}

# --- Test-Helfer -------------------------------------------------------------
$script:pass = 0; $script:fail = 0
function Assert {
    param([string]$Name, [bool]$Cond, [string]$Detail = '')
    if ($Cond) { $script:pass++; Write-Host "PASS  $Name" }
    else { $script:fail++; Write-Host "FAIL  $Name   $Detail" -ForegroundColor Red }
}
function AssertThrows {
    param([string]$Name, [scriptblock]$Sb)
    try { & $Sb; Assert $Name $false 'kein Fehler geworfen' } catch { Assert $Name $true }
}

# --- RFC-4515-Escaping -------------------------------------------------------
Assert 'LdapValue escapes ( ) \ *' ((Convert-ArgusLdapValue 'a(b)\c*' $false) -eq 'a\28b\29\5cc\2a')
Assert 'LdapValue wildcard erlaubt' ((Convert-ArgusLdapValue 'x*' $true) -eq 'x*')
Assert 'DnForFilter Klammern' ((Convert-ArgusDnForFilter 'CN=G (L),OU=x') -eq 'CN=G \28L\29,OU=x')
Assert 'DnForFilter Backslash zuerst' ((Convert-ArgusDnForFilter 'CN=Smith\, J,OU=x') -eq 'CN=Smith\5c, J,OU=x')
Assert 'DnForPath Slash' ((Convert-ArgusDnForPath 'OU=IT/Serv,DC=x') -eq 'OU=IT\/Serv,DC=x')

# --- DN -> CN ----------------------------------------------------------------
Assert 'CnFromDn einfach' ((Get-ArgusCnFromDn 'CN=Max Mustermann,OU=U,DC=x') -eq 'Max Mustermann')
Assert 'CnFromDn escaptes Komma' ((Get-ArgusCnFromDn 'CN=Smith\, John,OU=U,DC=x') -eq 'Smith, John')
Assert 'CnFromDn case-insensitiv' ((Get-ArgusCnFromDn 'cn=Test,OU=x') -eq 'Test')
Assert 'CnFromDn Hex-Escape' ((Get-ArgusCnFromDn 'CN=A\28B,DC=x') -eq 'A(B')
Assert 'CnFromDn leer -> null' ($null -eq (Get-ArgusCnFromDn ''))

# --- FILETIME-Sentinels ------------------------------------------------------
Assert 'FileTime 0 -> null' ($null -eq (Convert-ArgusFileTime 0))
Assert 'FileTime Max -> null' ($null -eq (Convert-ArgusFileTime 9223372036854775807))
Assert 'FileTime Unsinn -> null' ($null -eq (Convert-ArgusFileTime 'abc'))
$now = Get-Date
$rt = Convert-ArgusFileTime $now.ToFileTime()
Assert 'FileTime Roundtrip' ($rt -and [Math]::Abs(($rt - $now).TotalSeconds) -lt 2)

# --- Klausel-Builder ---------------------------------------------------------
Assert 'Clause equals escaped' ((New-ArgusClause 'mail' 'equals' 'a(b' 'string') -eq '(mail=a\28b)')
Assert 'Clause contains' ((New-ArgusClause 'sn' 'contains' 'Mü' 'string') -eq '(sn=*Mü*)')
Assert 'Clause notpresent' ((New-ArgusClause 'mail' 'notpresent' '' 'string') -eq '(!(mail=*))')
Assert 'Clause ge' ((New-ArgusClause 'employeeID' 'ge' '1000' 'string') -eq '(employeeID>=1000)')
Assert 'Clause raw unescaped' ((New-ArgusClause 'x' 'raw' 'a*b' 'string') -eq '(x=a*b)')
Assert 'UAC Enabled=true -> NOT Bit2' ((New-ArgusClause 'userAccountControl' 'equals' 'wahr' 'uac' 2 $true) -eq '(!(userAccountControl:1.2.840.113556.1.4.803:=2))')
Assert 'UAC Enabled=false -> Bit2' ((New-ArgusClause 'userAccountControl' 'equals' 'falsch' 'uac' 2 $true) -eq '(userAccountControl:1.2.840.113556.1.4.803:=2)')
Assert 'UAC Enabled notequals true' ((New-ArgusClause 'userAccountControl' 'notequals' 'true' 'uac' 2 $true) -eq '(userAccountControl:1.2.840.113556.1.4.803:=2)')
Assert 'UAC PwdNeverExpires=true -> Bit' ((New-ArgusClause 'userAccountControl' 'equals' '1' 'uac' 65536 $false) -eq '(userAccountControl:1.2.840.113556.1.4.803:=65536)')
AssertThrows 'UAC contains verboten' { New-ArgusClause 'userAccountControl' 'contains' 'x' 'uac' 2 $true }
Assert 'Date lastdays FILETIME' ((New-ArgusClause 'lastLogonTimestamp' 'lastdays' '30' 'date') -match '^\(lastLogonTimestamp>=\d{15,}\)$')
Assert 'Date olderdays inkl. nie' ((New-ArgusClause 'pwdLastSet' 'olderdays' '90' 'date') -match '^\(!\(pwdLastSet>=\d{15,}\)\)$')
Assert 'DateDirect GeneralizedTime' ((New-ArgusClause 'whenCreated' 'lastdays' '7' 'datedirect') -match '^\(whenCreated>=\d{14}\.0Z\)$')
AssertThrows 'Date equals verboten' { New-ArgusClause 'pwdLastSet' 'equals' '01.01.2024' 'date' }
AssertThrows 'lastdays auf String verboten' { New-ArgusClause 'mail' 'lastdays' '30' 'string' }
AssertThrows 'lastdays ohne Zahl' { New-ArgusClause 'pwdLastSet' 'lastdays' 'abc' 'date' }
Assert 'Manager equals DN' ((New-ArgusClause 'manager' 'equals' 'CN=Boss (IT),OU=U,DC=x' 'manager') -eq '(manager=CN=Boss \28IT\29,OU=U,DC=x)')
Assert 'Manager present' ((New-ArgusClause 'manager' 'present' '' 'manager') -eq '(manager=*)')
AssertThrows 'Manager contains verboten' { New-ArgusClause 'manager' 'contains' 'Boss' 'manager' }
AssertThrows 'Manager equals ohne DN' { New-ArgusClause 'manager' 'equals' 'Boss' 'manager' }
AssertThrows 'GUID-Filter verboten' { New-ArgusClause 'objectGUID' 'equals' 'x' 'guid' }

# --- Gruppen-Klausel ---------------------------------------------------------
$gc = Get-ArgusGroupClause @{ Dn = 'CN=G (alt),OU=x,DC=y'; Rid = 1105; Name = 'G' }
Assert 'GroupClause chain+primary' ($gc -eq '(|(memberOf:1.2.840.113556.1.4.1941:=CN=G \28alt\29,OU=x,DC=y)(primaryGroupID=1105))')
$gc2 = Get-ArgusGroupClause @{ Dn = 'CN=G,DC=y'; Rid = $null; Name = 'G' }
Assert 'GroupClause ohne RID' ($gc2 -eq '(memberOf:1.2.840.113556.1.4.1941:=CN=G,DC=y)')

# --- Zellen-Schutz -----------------------------------------------------------
Assert 'Xlsx: = neutralisiert' ((Protect-ArgusCell '=SUM(A1)' 'xlsx' $false) -eq "'=SUM(A1)")
Assert 'Xlsx: + bleibt' ((Protect-ArgusCell '+49 231 123' 'xlsx' $false) -eq '+49 231 123')
Assert 'CSV: = neutralisiert' ((Protect-ArgusCell '=1+1' 'csv' $false) -eq "'=1+1")
Assert 'CSV: + bleibt (nicht streng)' ((Protect-ArgusCell '+49 231' 'csv' $false) -eq '+49 231')
Assert 'CSV streng: + neutralisiert' ((Protect-ArgusCell '+49 231' 'csv' $true) -eq "'+49 231")
Assert 'CSV streng: @ neutralisiert' ((Protect-ArgusCell '@cmd' 'csv' $true) -eq "'@cmd")

# --- Chunking / Pipeline-Falle ----------------------------------------------
$chunks = Split-ArgusList @(1..130) 60
Assert 'Split 130/60 -> 3 Chunks' ($chunks.Count -eq 3)
Assert 'Split Chunkgrößen' ($chunks[0].Count -eq 60 -and $chunks[1].Count -eq 60 -and $chunks[2].Count -eq 10)
$one = Split-ArgusList @('a') 60
Assert 'Split 1 Element bleibt Liste' ($one.Count -eq 1 -and $one[0].Count -eq 1)

# --- Attribut-Auswahl --------------------------------------------------------
$sel = Resolve-ArgusAttributeSelection @('Mail', 'UnbekannterKey') @('extensionAttribute5')
Assert 'AttrSelection Katalog+Custom' ($sel.Count -eq 3)
Assert 'AttrSelection Mail via Katalog' ($sel[0].Ldap -eq 'mail' -and -not $sel[0].Custom)
Assert 'AttrSelection Unbekannt als LDAP' (@($sel | Where-Object { $_.Ldap -eq 'UnbekannterKey' -and $_.Custom }).Count -eq 1)
$selDup = Resolve-ArgusAttributeSelection @('Mail', 'mail') @('mail')
Assert 'AttrSelection dedupliziert' (@($selDup).Count -eq 1)
$selUac = Resolve-ArgusAttributeSelection @('Enabled', 'PwdNeverExpires') @()
Assert 'AttrSelection beide UAC-Spalten' (@($selUac).Count -eq 2)

# --- Filterzeilen ------------------------------------------------------------
$rr = Resolve-ArgusFilterRows @(
    @{ Attr = 'PLZ'; Op = 'equals'; Val = '44147' }
    @{ Attr = 'E-Mail'; Op = 'contains'; Val = '' }        # Fehler: Wert fehlt
    @{ Attr = 'E-Mail'; Op = 'present'; Val = '' }          # OK ohne Wert
    @{ Attr = 'extensionAttribute5'; Op = 'equals'; Val = 'X' }  # freier LDAP-Name
)
Assert 'FilterRows: 3 gültig' ($rr.Filters.Count -eq 3)
Assert 'FilterRows: 1 Fehler gemeldet' ($rr.Errors.Count -eq 1)
Assert 'FilterRows: PLZ -> postalCode' ($rr.Filters[0].Ldap -eq 'postalCode')
Assert 'FilterRows: freier Name als string' ($rr.Filters[2].Ldap -eq 'extensionAttribute5' -and $rr.Filters[2].Type -eq 'string')

# --- Settings / Profile ------------------------------------------------------
$def = Get-ArgusDefaultSettings
$built = Get-ArgusSearchParams $def $null
Assert 'Defaults: keine Fehler' ($built.Errors.Count -eq 0) ($built.Errors -join '; ')
Assert 'Defaults: Mode all' ($built.Params.Mode -eq 'all')
Assert 'Defaults: 9 Standard-Spalten' ($built.Params.Attributes.Count -eq 9) "war $($built.Params.Attributes.Count)"

$def.Source.Mode = 'group'
$built2 = Get-ArgusSearchParams $def $null
Assert 'Gruppe ohne Gruppen -> Fehler' ($built2.Errors.Count -ge 1)

$def2 = Get-ArgusDefaultSettings
$def2.Connection.AuthMode = 'manual'
$built3 = Get-ArgusSearchParams $def2 $null
Assert 'Manuelle Auth ohne Daten -> Fehler' ($built3.Errors.Count -ge 1)

Assert 'Port 636 -> SSL' ((Get-ArgusEffectivePort @{ PortChoice = '636' }).UseSsl)
Assert 'Port 3268 -> kein SSL' (-not (Get-ArgusEffectivePort @{ PortChoice = '3268' }).UseSsl)
$effC = Get-ArgusEffectivePort @{ PortChoice = 'custom'; CustomPort = 10389; UseLdaps = $true }
Assert 'Custom Port + LDAPS' ($effC.Port -eq 10389 -and $effC.UseSsl)

$tmpProfile = Join-Path ([IO.Path]::GetTempPath()) "argus-test-$PID.json"
$defSave = Get-ArgusDefaultSettings
$defSave.Source.Groups = @('Gruppe A', 'Gruppe B')
$defSave.Filter.Rows = @(@{ Attr = 'PLZ'; Op = 'equals'; Val = '44147' })
Save-ArgusProfile $defSave $tmpProfile
$reload = Read-ArgusProfile $tmpProfile
Assert 'Profil-Roundtrip Gruppen' (@($reload.Source.Groups).Count -eq 2 -and $reload.Source.Groups[1] -eq 'Gruppe B')
Assert 'Profil-Roundtrip Filterzeile' (@($reload.Filter.Rows).Count -eq 1 -and $reload.Filter.Rows[0].Val -eq '44147')
Assert 'Profil-Merge ergänzt fehlende Keys' ($null -ne $reload.Export.CsvDelimiter)
Remove-Item $tmpProfile -Force

$partial = Merge-ArgusSettings (Get-ArgusDefaultSettings) @{ Export = @{ Csv = $true } }
Assert 'Merge behält Nachbarwerte' ($partial.Export.Csv -and $partial.Export.Xlsx)

# --- Export: Ziele + CSV-Writer (DataTable, ohne ImportExcel) ----------------
$opt = @{ Folder = [IO.Path]::GetTempPath(); FileName = 'unit-test.xlsx'; Xlsx = $true; Csv = $true; AddTimestamp = $false }
$targets = Get-ArgusExportTargets $opt
Assert 'Targets: Endungen ersetzt' ($targets.XlsxPath.EndsWith('unit-test.xlsx') -and $targets.CsvPath.EndsWith('unit-test.csv'))

$table = New-Object System.Data.DataTable 'T'
[void]$table.Columns.Add((New-Object System.Data.DataColumn('Name', [string])))
[void]$table.Columns.Add((New-Object System.Data.DataColumn('Datum', [datetime])))
[void]$table.Columns.Add((New-Object System.Data.DataColumn('Aktiv', [bool])))
$r1 = $table.NewRow(); $r1['Name'] = '=cmd|calc'; $r1['Datum'] = [datetime]'2026-01-02 03:04:05'; $r1['Aktiv'] = $true; $table.Rows.Add($r1)
$r2 = $table.NewRow(); $r2['Name'] = 'Sm;th "Q"'; $r2['Datum'] = [DBNull]::Value; $r2['Aktiv'] = $false; $table.Rows.Add($r2)

$csvDir = Join-Path ([IO.Path]::GetTempPath()) "argus-csv-$PID"
New-Item -Path $csvDir -ItemType Directory -Force | Out-Null
$ctx = @{ Log = { param($m) $null = $m }; Progress = { param($p, $s) $null = $p; $null = $s }; IsCancelled = { $false } }
$csvOpt = @{ Folder = $csvDir; FileName = 'probe'; Xlsx = $false; Csv = $true; CsvDelimiter = ';'; StrictCsv = $false; AddTimestamp = $false }
$written = Export-ArgusData $table $csvOpt $ctx
$lines = Get-Content -LiteralPath $written.CsvPath -Encoding UTF8
Assert 'CSV: 3 Zeilen' ($lines.Count -eq 3)
Assert 'CSV: Header' ($lines[0] -eq 'Name;Datum;Aktiv')
Assert 'CSV: Formel-Schutz' ($lines[1] -eq "'=cmd|calc;2026-01-02 03:04:05;True")
Assert 'CSV: Quoting+DBNull' ($lines[2] -eq '"Sm;th ""Q""";;False')
$bytes = [IO.File]::ReadAllBytes($written.CsvPath)
Assert 'CSV: UTF-8 BOM' ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
Remove-Item $csvDir -Recurse -Force

# --- Export-Meta -------------------------------------------------------------
$meta = Get-ArgusExportMeta @{ Count = 42; DurationSec = 1.5; Filter = '(x=y)'; Server = 'dc1'; BaseDn = 'DC=x'; Unmatched = @('foo') }
Assert 'Meta: Treffer' ($meta['Treffer'] -eq 42)
Assert 'Meta: Unmatched' ($meta['Nicht gefundene IDs'] -eq 'foo')

# --- Worker-Scriptblock: Engine+Worker-Konkatenation parsebar ----------------
$workerAssign = $ast.Find({
    $args[0] -is [System.Management.Automation.Language.AssignmentStatementAst] -and
    $args[0].Left.Extent.Text -eq '$script:WorkerBody'
}, $true)
$wbText = $workerAssign.Right.Extent.Text
$wbBody = $wbText.Substring(1, $wbText.Length - 2)
$combined = $body + "`n" + $wbBody
$e2 = $null; $t2 = $null
[System.Management.Automation.Language.Parser]::ParseInput($combined, [ref]$t2, [ref]$e2) | Out-Null
Assert 'Engine+Worker konkateniert parsebar' ($e2.Count -eq 0)

Write-Host ''
Write-Host ("Ergebnis: {0} PASS, {1} FAIL" -f $script:pass, $script:fail)
if ($script:fail -gt 0) { exit 1 } else { exit 0 }
