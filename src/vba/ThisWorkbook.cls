'===============================================================================
' ThisWorkbook - Événements du classeur
'-------------------------------------------------------------------------------
' Fichier source du module de document « ThisWorkbook » (code uniquement, sans les
' lignes Attribute qui sont ajoutées par le générateur / l'installateur).
'===============================================================================
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    ' La protection « UserInterfaceOnly » n'est pas enregistrée dans le fichier :
    ' elle doit être réappliquée à chaque ouverture.
    modOutils.ProtegerTout
    modAdmin.AssignerBoutons
    modTableauDeBord.AppliquerHorizon
    Me.Worksheets(F_TDB).Activate
    modAdmin.VerifierInstallation
End Sub
