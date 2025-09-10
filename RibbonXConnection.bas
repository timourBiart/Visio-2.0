Attribute VB_Name = "RibbonXConnection"
' ======================
' === MODULE STANDARD POUR LE RUBAN ===
' ======================
Option Explicit

' D�claration pour le contr�le du ruban (facultatif)
Public MyRibbon As IRibbonUI

' (facultatif) si tu ajoutes onLoad="Ribbon_OnLoad" dans customUI.xml
Public Sub Ribbon_OnLoad(ByVal ribbon As IRibbonUI)
    Set MyRibbon = ribbon
End Sub

' Callback pour la mise � jour des documents, li�e au bouton du ruban
' customUI.xml : onAction="MettreAJourTout"
Public Sub MettreAJourTout(ByVal control As IRibbonControl)
    On Error GoTo ErrHandler
    ' >>> Appel unique � ta macro centrale (ThisDocument) qui fait :
    ' EnsureListeCablesPage -> ReordonnerPages -> UpdateTOC -> UpdateBOM -> UpdateCableList
    ThisDocument.MettreAJourToutSansAlerte "ribbon"
    MsgBox "Les documents ont �t� mis � jour !", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Erreur pendant la mise � jour : " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' (Optionnel) Wrapper priv� si besoin local, �vite toute ambigu�t� de nom
Private Sub Ruban_MettreAJourToutSansAlerte()
    On Error Resume Next
    ThisDocument.MettreAJourToutSansAlerte "module"
End Sub

' Callback pour l'encodage des valeurs utilisateur, li�e au bouton du ruban
Public Sub EncoderValeursUtilisateur(ByVal control As IRibbonControl)
    frmEncodageDonnees.Show
End Sub

' Callback pour "Convertir en Sch�ma Normalis�"
Public Sub MAJ(ByVal control As IRibbonControl)
    On Error Resume Next
    ConvSchemaNormalise.ConvertirEnSchemaNormalise
End Sub

