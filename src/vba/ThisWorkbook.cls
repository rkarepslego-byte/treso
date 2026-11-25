'===============================================================================
' ThisWorkbook - Événements du classeur
'-------------------------------------------------------------------------------
' Fichier source du module de document « ThisWorkbook » (code uniquement, sans les
' lignes Attribute qui sont ajoutées par le générateur / l'installateur).
'===============================================================================
Option Explicit

Private Sub Workbook_Open()
    Dim etaitEnregistre As Boolean
    On Error Resume Next
    etaitEnregistre = Me.Saved
    ' La protection « UserInterfaceOnly » n'est pas enregistrée dans le fichier :
    ' elle doit être réappliquée à chaque ouverture.
    modOutils.ProtegerTout
    modAdmin.AssignerBoutons
    modTableauDeBord.AppliquerHorizon
    Me.Worksheets(F_TDB).Activate
    ' Ces réglages techniques ne doivent pas provoquer de demande d'enregistrement à la fermeture
    Me.Saved = etaitEnregistre
    modAdmin.VerifierInstallation
    ' Proposition d'actualisation si de nouveaux relevés ont été déposés (après la fin de l'ouverture)
    Application.OnTime Now + TimeSerial(0, 0, 2), "ProposerActualisation"
End Sub
