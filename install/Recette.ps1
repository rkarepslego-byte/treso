<#
.SYNOPSIS
    Recette automatique de l'outil de trésorerie dans Excel (poste Windows).

.DESCRIPTION
    Travaille sur une COPIE du dossier de démonstration (le classeur livré n'est pas modifié) :
      1. ouvre la copie dans Excel (macros activées, sans affichage) ;
      2. exécute la chaîne complète via la macro Recette_Executer : actualisation Power Query
         des relevés de démonstration, import, charges fixes, rapprochement, validation des
         correspondances certaines ;
      3. relance une actualisation (Recette_Reimporter) : aucun doublon ne doit être ajouté ;
      4. compare les soldes et indicateurs obtenus aux valeurs attendues calculées par le
         moteur de référence (recette_attendue.json).
    Code de sortie 0 = recette réussie, 1 = écart détecté.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File Recette.ps1
#>
param(
    [string]$Dossier,
    [switch]$GarderCopie
)

$ErrorActionPreference = "Stop"
$racine = Split-Path -Parent $PSScriptRoot
if (-not $Dossier) { $Dossier = Join-Path $racine "dist" }
$Dossier = (Resolve-Path $Dossier).Path
$attendu = Get-Content (Join-Path $Dossier "recette_attendue.json") -Raw -Encoding UTF8 | ConvertFrom-Json
$inv = [System.Globalization.CultureInfo]::InvariantCulture

$copie = Join-Path $env:TEMP ("treso_recette_" + (Get-Date -Format "yyyyMMdd_HHmmss"))
New-Item -ItemType Directory -Path $copie | Out-Null
Copy-Item (Join-Path $Dossier "Tresorerie.xlsm") $copie
Copy-Item (Join-Path $Dossier "Imports") $copie -Recurse
New-Item -ItemType Directory -Path (Join-Path $copie "Sauvegardes") | Out-Null
Get-ChildItem $copie -Recurse -File | Unblock-File
Write-Host "Recette dans la copie $copie" -ForegroundColor Cyan

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1       # macros autorisées pour cette session d'automatisation
    $wb = $excel.Workbooks.Open((Join-Path $copie "Tresorerie.xlsm"))
    $prefixe = "'" + $wb.Name + "'!"
    Write-Host "Exécution de la chaîne complète (Power Query + VBA)..."
    $resultat = $excel.Run($prefixe + "Recette_Executer", [double]$attendu.date_reference_serie)
    Write-Host "Réimport des mêmes relevés..."
    $resultat2 = $excel.Run($prefixe + "Recette_Reimporter")
    $journal = $excel.Run($prefixe + "Recette_Journal")
    $wb.Close($true)
}
finally {
    if ($excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
}

function Lire-Resultat([string]$texte) {
    $h = @{}
    foreach ($morceau in ($texte -split ";")) {
        $kv = $morceau -split "=", 2
        if ($kv.Count -eq 2) { $h[$kv[0]] = $kv[1] }
    }
    return $h
}

$r = Lire-Resultat $resultat
$r2 = Lire-Resultat $resultat2
$ok = $true
Write-Host ""
if ($r["Erreur"]) { Write-Host "ERREUR pendant l'exécution : $($r['Erreur'])" -ForegroundColor Red; $ok = $false }
"{0,-24} {1,14} {2,14}  {3}" -f "Indicateur", "Attendu", "Obtenu", "Résultat" | Write-Host
foreach ($p in $attendu.valeurs.PSObject.Properties) {
    $valeurAttendue = [double]$p.Value
    $texteObtenu = $r[$p.Name]
    $obtenu = $null
    if ($texteObtenu) { $obtenu = [double]::Parse($texteObtenu, $inv) }
    $conforme = ($null -ne $obtenu) -and ([math]::Abs($obtenu - $valeurAttendue) -lt 0.01)
    if (-not $conforme) { $ok = $false }
    $etat = if ($conforme) { "OK" } else { "ÉCART" }
    $couleur = if ($conforme) { "Green" } else { "Red" }
    Write-Host ("{0,-24} {1,14} {2,14}  {3}" -f $p.Name, $valeurAttendue.ToString("N2"), $texteObtenu, $etat) -ForegroundColor $couleur
}
if ($r2["Erreur"]) { Write-Host "ERREUR au réimport : $($r2['Erreur'])" -ForegroundColor Red; $ok = $false }
if ($r2["Avant"] -ne $r2["Apres"]) {
    Write-Host "Réimport : $($r2['Avant']) opérations avant, $($r2['Apres']) après -> DOUBLONS AJOUTÉS" -ForegroundColor Red
    $ok = $false
} else {
    Write-Host "Réimport des mêmes relevés : aucun doublon ajouté ($($r2['Apres']) opérations)." -ForegroundColor Green
}
$fichierJournal = Join-Path $copie "journal_recette.txt"
Set-Content -Path $fichierJournal -Value ($resultat + "`r`n" + $resultat2 + "`r`n`r`n" + $journal) -Encoding UTF8
Write-Host "Journal détaillé : $fichierJournal"
if (-not $GarderCopie -and $ok) { Remove-Item $copie -Recurse -Force }

Write-Host ""
if ($ok) {
    Write-Host "RECETTE RÉUSSIE : Excel, Power Query et VBA donnent exactement les résultats attendus." -ForegroundColor Green
    exit 0
} else {
    Write-Host "RECETTE EN ÉCHEC : voir les écarts ci-dessus et le journal ($copie)." -ForegroundColor Red
    exit 1
}
