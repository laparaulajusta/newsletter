# =============================================================================
#  generar.ps1 - Genera una newsletter HTML a partir d'un fitxer de contingut
#
#  Us normal: arrossega el fitxer .txt de la newsletter sobre generar.bat
#  (o fes doble clic a generar.bat i tria el fitxer).
#
#  Resultat: crea un .html amb el mateix nom que el .txt i obre un correu nou
#  a l'Outlook d'escriptori amb la newsletter a dins, llest per enviar.
#  Si l'Outlook d'escriptori no esta disponible, obre una finestra nova
#  d'Edge amb la newsletter i un correu nou d'Outlook web (copiar i enganxar).
# =============================================================================

param(
    [string]$FitxerContingut,
    [switch]$NoObrir          # nomes genera l'HTML, sense obrir res
)

# --- CONFIGURACIO ------------------------------------------------------------
# On s'obre el correu nou:
#   'escriptori' = Outlook d'escriptori classic (si falla, passa a 'web')
#   'web'        = Outlook web en una finestra nova d'Edge
$ObrirAmb = 'escriptori'

# Adreca d'Outlook web per al mode 'web':
#   Microsoft 365 (compte d'empresa): https://outlook.office.com/mail/deeplink/compose
#   Outlook.com (compte personal):    https://outlook.live.com/mail/0/deeplink/compose
$UrlNouCorreu = 'https://outlook.office.com/mail/deeplink/compose'
# -----------------------------------------------------------------------------

$ErrorActionPreference = 'Stop'
$arrel      = $PSScriptRoot
$dirPlant   = Join-Path $arrel 'plantilla'
$dirBlocs   = Join-Path $dirPlant 'blocs'
$fitxerBase = Join-Path $dirPlant 'base.html'
$avisos     = New-Object System.Collections.ArrayList
$esperar    = $false
$campsGenerals = @('BLOC', 'ASSUMPTE', 'PER', 'CONTINGUT')

function Llegeix([string]$ruta) {
    return [System.IO.File]::ReadAllText($ruta, [System.Text.Encoding]::UTF8)
}

function Avis([string]$text) {
    [void]$avisos.Add($text)
}

# Uneix les linies d'un camp: treu linies buides del principi i del final
# i converteix cada salt de linia en <br />
function Uneix($llista) {
    $l = @($llista)
    $ini = 0
    $fi = $l.Count - 1
    while ($ini -le $fi -and $l[$ini] -eq '') { $ini++ }
    while ($fi -ge $ini -and $l[$fi] -eq '') { $fi-- }
    if ($ini -gt $fi) { return '' }
    return (@($l[$ini..$fi]) -join '<br />')
}

function TeValor($dades, [string]$k) {
    return ($dades.Contains($k) -and ([string]$dades[$k]) -ne '')
}

# Omple una plantilla amb les dades. A les plantilles:
#   {{CAMP}}                      camp obligatori (avisa si falta)
#   {{CAMP?}}                     camp opcional (queda buit si falta)
#   <!--SI:CAMP--> ... <!--/SI:CAMP-->  tros que nomes surt si CAMP te valor
function Omple([string]$plantilla, $dades, [string]$nom) {
    # Camps que admet aquesta plantilla
    $admesos = @([regex]::Matches($plantilla, '\{\{([A-Z0-9_]+)\??\}\}|<!--SI:([A-Z0-9_]+)-->') |
        ForEach-Object { if ($_.Groups[1].Success) { $_.Groups[1].Value } else { $_.Groups[2].Value } })
    foreach ($k in @($dades.Keys)) {
        if (($k -notin $admesos) -and ($k -notin $campsGenerals)) {
            Avis ("{0}: el camp {1} no existeix en aquest bloc (revisa com esta escrit)" -f $nom, $k)
        }
    }

    # Trossos opcionals
    $plantilla = [regex]::Replace($plantilla, '(?s)<!--SI:([A-Z0-9_]+)-->(.*?)<!--/SI:\1-->', {
        param($m)
        if (TeValor $dades $m.Groups[1].Value) { return $m.Groups[2].Value }
        return ''
    })

    # Camps obligatoris buits o que falten
    $obligatoris = [regex]::Matches($plantilla, '\{\{([A-Z0-9_]+)\}\}') |
        ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique
    foreach ($o in $obligatoris) {
        if (-not (TeValor $dades $o)) { Avis ("{0}: falta el camp {1} (ha quedat buit)" -f $nom, $o) }
    }

    foreach ($k in @($dades.Keys)) {
        $v = [string]$dades[$k]
        $plantilla = $plantilla.Replace('{{' + $k + '}}', $v).Replace('{{' + $k + '?}}', $v)
    }
    return [regex]::Replace($plantilla, '\{\{[A-Z0-9_]+\??\}\}', '')
}

function Obre-OutlookEscriptori([string]$html, [string]$assumpte, [string]$per) {
    $outlook = New-Object -ComObject Outlook.Application
    $correu = $outlook.CreateItem(0)       # 0 = correu nou
    $correu.Subject = $assumpte
    if ($per) { $correu.To = $per }
    $correu.BodyFormat = 2                 # 2 = HTML
    $correu.HTMLBody = $html
    $correu.Display()
    # Amaga les linies de quadricula de les taules (nomes es veuen mentre
    # s'edita el correu, pero semblen requadres al voltant dels blocs)
    try { $correu.GetInspector.WordEditor.Windows.Item(1).View.TableGridlines = $false } catch { }
}

function Obre-Edge([string[]]$adreces) {
    try {
        Start-Process -FilePath 'msedge.exe' -ArgumentList (@('--new-window') + $adreces)
    }
    catch {
        foreach ($a in $adreces) { Start-Process $a }
    }
}

try {
    # --- 1. Triar el fitxer de contingut -------------------------------------
    if (-not $FitxerContingut) {
        Add-Type -AssemblyName System.Windows.Forms
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.InitialDirectory = Join-Path $arrel 'newsletters'
        $dlg.Filter = 'Contingut de newsletter (*.txt)|*.txt'
        $dlg.Title = 'Tria la newsletter que vols generar'
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { exit 0 }
        $FitxerContingut = $dlg.FileName
    }
    if (-not (Test-Path -LiteralPath $FitxerContingut)) {
        throw "No es troba el fitxer: $FitxerContingut"
    }
    $FitxerContingut = (Resolve-Path -LiteralPath $FitxerContingut).Path

    # --- 2. Llegir i separar el contingut en seccions ------------------------
    # Primera seccio = dades generals. Cada '---' comenca un bloc nou.
    $linies   = (Llegeix $FitxerContingut) -split "\r?\n"
    $seccions = New-Object System.Collections.ArrayList
    $actual   = [ordered]@{}
    $clau     = $null

    foreach ($linia in $linies) {
        if ($linia.Trim() -eq '---') {
            [void]$seccions.Add($actual)
            $actual = [ordered]@{}
            $clau = $null
            continue
        }
        if ($linia.TrimStart().StartsWith('#')) { continue }

        # Un camp es una paraula en MAJUSCULES seguida de dos punts: TITOL: ...
        if ($linia -cmatch '^([A-Z0-9_]+):\s*(.*)$') {
            $clau = $Matches[1]
            $actual[$clau] = New-Object System.Collections.ArrayList
            [void]$actual[$clau].Add($Matches[2].Trim())
        }
        elseif ($clau) {
            [void]$actual[$clau].Add($linia.Trim())
        }
    }
    [void]$seccions.Add($actual)

    $dadesSeccions = @()
    foreach ($s in $seccions) {
        $d = [ordered]@{}
        foreach ($k in @($s.Keys)) { $d[$k] = Uneix $s[$k] }
        $dadesSeccions += , $d
    }

    $general = $dadesSeccions[0]
    $blocs   = @($dadesSeccions | Select-Object -Skip 1 | Where-Object { $_.Count -gt 0 })

    # --- 3. Construir els blocs ----------------------------------------------
    $disponibles = (Get-ChildItem -Path $dirBlocs -Filter '*.html' |
        ForEach-Object { $_.BaseName }) -join ', '
    $htmlBlocs = New-Object System.Text.StringBuilder
    $n = 0
    foreach ($b in $blocs) {
        $n++
        $tipus = [string]$b['BLOC']
        if (-not $tipus) {
            throw "El bloc numero $n no te la linia BLOC: (tipus disponibles: $disponibles)"
        }
        $rutaBloc = Join-Path $dirBlocs ($tipus.Trim().ToLower() + '.html')
        if (-not (Test-Path -LiteralPath $rutaBloc)) {
            throw "El bloc numero $n es de tipus '$tipus', que no existeix. Tipus disponibles: $disponibles"
        }
        [void]$htmlBlocs.AppendLine((Omple (Llegeix $rutaBloc) $b "Bloc $n ($tipus)"))
    }

    # --- 4. Muntar la newsletter sencera -------------------------------------
    $assumpte = [string]$general['ASSUMPTE']
    if (-not $assumpte) {
        $assumpte = [System.IO.Path]::GetFileNameWithoutExtension($FitxerContingut)
        Avis "Falta ASSUMPTE: a la capcalera; s'usa el nom del fitxer"
        $general['ASSUMPTE'] = $assumpte
    }
    $per = [string]$general['PER']
    $general['CONTINGUT'] = $htmlBlocs.ToString()
    $html = Omple (Llegeix $fitxerBase) $general 'Capcalera'

    $sortida = [System.IO.Path]::ChangeExtension($FitxerContingut, '.html')
    [System.IO.File]::WriteAllText($sortida, $html, (New-Object System.Text.UTF8Encoding($true)))

    Write-Host ''
    Write-Host "Newsletter generada: $sortida" -ForegroundColor Green
    Write-Host "Blocs: $($blocs.Count)"

    # --- 5. Obrir el correu nou ----------------------------------------------
    if (-not $NoObrir) {
        $obert = $false
        if ($ObrirAmb -eq 'escriptori') {
            try {
                Obre-OutlookEscriptori $html $assumpte $per
                $obert = $true
                Write-Host ''
                Write-Host "S'ha obert un correu nou a l'Outlook amb la newsletter."
                Write-Host 'Revisa-la, posa-hi els destinataris i envia-la.'
            }
            catch {
                Write-Host ''
                Write-Host "No s'ha pogut obrir l'Outlook d'escriptori: $($_.Exception.Message)" -ForegroundColor Yellow
                Write-Host "S'obre Outlook web en una finestra nova d'Edge." -ForegroundColor Yellow
                $esperar = $true
            }
        }
        if (-not $obert) {
            $url = $UrlNouCorreu + '?subject=' + [uri]::EscapeDataString($assumpte)
            if ($per) { $url += '&to=' + [uri]::EscapeDataString($per) }
            Obre-Edge @(([uri]$sortida).AbsoluteUri, $url)
            Write-Host ''
            Write-Host 'A la pestanya de la newsletter fes Ctrl+A i Ctrl+C;'
            Write-Host 'a la del correu nou, clic al cos del missatge i Ctrl+V.'
        }
    }

    if ($avisos.Count -gt 0) {
        Write-Host ''
        Write-Host 'AVISOS (revisa el fitxer .txt):' -ForegroundColor Yellow
        foreach ($a in $avisos) { Write-Host "  - $a" -ForegroundColor Yellow }
        $esperar = $true
    }
    if ($esperar -and -not $NoObrir) {
        Write-Host ''
        [void](Read-Host 'Prem Enter per tancar')
    }
}
catch {
    Write-Host ''
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'No s''ha generat la newsletter.'
    Write-Host ''
    if (-not $NoObrir) { [void](Read-Host 'Prem Enter per tancar') }
    exit 1
}
