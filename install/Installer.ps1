<#
.SYNOPSIS
    Installation de l'outil de trésorerie sur le poste Windows du client.

.DESCRIPTION
    1. Débloque les fichiers téléchargés (marque « provenant d'Internet » qui bloque les macros).
    2. Crée les dossiers Imports\Banque1, Imports\Banque2 et Sauvegardes.
    3. (option -EmplacementApprouve) Ajoute le dossier aux emplacements approuvés d'Excel.
    4. (option -Reconstruire) Réimporte dans le classeur les modules VBA (src\vba) et les
       requêtes Power Query (src\powerquery). Sert à appliquer une mise à jour du code ou à
       réparer le classeur si Excel a retiré le projet VBA ou les requêtes.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File Installer.ps1 -EmplacementApprouve
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File Installer.ps1 -Dossier "D:\Clients\Dupont\Tresorerie" -Reconstruire
#>
param(
    [string]$Dossier,
    [string]$Classeur = "Tresorerie.xlsm",
    [switch]$EmplacementApprouve,
    [switch]$Reconstruire
)

$ErrorActionPreference = "Stop"
$racine = Split-Path -Parent $PSScriptRoot
if (-not $Dossier) { $Dossier = Join-Path $racine "dist" }
$Dossier = (Resolve-Path $Dossier).Path
$chemin = Join-Path $Dossier $Classeur
if (-not (Test-Path $chemin)) { throw "Classeur introuvable : $chemin" }

Write-Host "Outil de trésorerie - installation dans $Dossier" -ForegroundColor Cyan

Write-Host "1. Déblocage des fichiers..."
Get-ChildItem -Path $Dossier -Recurse -File | Unblock-File

Write-Host "2. Dossiers d'import et de sauvegarde..."
foreach ($d in @("Imports\Banque1", "Imports\Banque2", "Sauvegardes")) {
    New-Item -ItemType Directory -Force -Path (Join-Path $Dossier $d) | Out-Null
}

if ($EmplacementApprouve) {
    Write-Host "3. Ajout du dossier aux emplacements approuvés d'Excel..."
    $cle = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\OutilTresorerie"
    New-Item -Path $cle -Force | Out-Null
    New-ItemProperty -Path $cle -Name "Path" -Value ($Dossier.TrimEnd('\') + '\') -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $cle -Name "AllowSubfolders" -Value 1 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $cle -Name "Description" -Value "Outil de trésorerie" -PropertyType String -Force | Out-Null
}

if ($Reconstruire) {
    Write-Host "4. Réimport du code VBA et des requêtes Power Query..."
    $cleSecurite = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
    if (-not (Test-Path $cleSecurite)) { New-Item -Path $cleSecurite -Force | Out-Null }
    $ancien = (Get-ItemProperty -Path $cleSecurite -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
    Copy-Item $chemin ($chemin + ".avant_reconstruction") -Force
    $excel = $null
    try {
        # Autorise temporairement l'accès au modèle objet du projet VBA
        Set-ItemProperty -Path $cleSecurite -Name AccessVBOM -Value 1 -Type DWord
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AutomationSecurity = 3   # ouverture sans exécuter les macros
        $wb = $excel.Workbooks.Open($chemin)
        $projet = $wb.VBProject

        foreach ($comp in @($projet.VBComponents)) {
            if ($comp.Type -eq 1) { $projet.VBComponents.Remove($comp) }   # 1 = module standard
        }
        $tmp = Join-Path $env:TEMP ("treso_vba_" + [guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $tmp | Out-Null
        $ansi = [System.Text.Encoding]::GetEncoding(1252)
        foreach ($f in Get-ChildItem (Join-Path $racine "src\vba") -Filter *.bas) {
            # Les sources sont en UTF-8 ; l'éditeur VBA importe en Windows-1252
            $texte = [System.IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)
            $texte = ($texte -replace "`r`n", "`n") -replace "`n", "`r`n"
            $cible = Join-Path $tmp $f.Name
            [System.IO.File]::WriteAllText($cible, $texte, $ansi)
            [void]$projet.VBComponents.Import($cible)
            Write-Host "   module $($f.BaseName)"
        }
        $code = [System.IO.File]::ReadAllText((Join-Path $racine "src\vba\ThisWorkbook.cls"), [System.Text.Encoding]::UTF8)
        $module = $projet.VBComponents.Item("ThisWorkbook").CodeModule
        if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
        $module.AddFromString($code)
        Remove-Item $tmp -Recurse -Force

        $tableM = $wb.Worksheets.Item("Param_Import").ListObjects.Item("tbl_CodeM")
        foreach ($f in Get-ChildItem (Join-Path $racine "src\powerquery") -Filter *.pq) {
            $nom = $f.BaseName
            $formule = [System.IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8).Trim()
            $requete = $null
            foreach ($q in $wb.Queries) { if ($q.Name -eq $nom) { $requete = $q } }
            if ($requete) { $requete.Formula = $formule } else { [void]$wb.Queries.Add($nom, $formule) }
            foreach ($ligne in $tableM.ListRows) {
                if ($ligne.Range.Cells.Item(1, 1).Value2 -eq $nom) { $ligne.Range.Cells.Item(1, 4).Value2 = $formule }
            }
            Write-Host "   requête $nom"
        }
        $wb.Save()
        $wb.Close($false)
        Write-Host "   Classeur mis à jour (copie de sécurité : $Classeur.avant_reconstruction)."
    }
    finally {
        if ($excel) {
            $excel.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }
        if ($null -eq $ancien) {
            Remove-ItemProperty -Path $cleSecurite -Name AccessVBOM -ErrorAction SilentlyContinue
        } else {
            Set-ItemProperty -Path $cleSecurite -Name AccessVBOM -Value $ancien -Type DWord
        }
    }
}

Write-Host ""
Write-Host "Installation terminée. Ouvrez $Classeur, activez les macros si Excel le demande," -ForegroundColor Green
Write-Host "puis renseignez l'onglet Paramètres (noms des banques, soldes initiaux et leur date)." -ForegroundColor Green
