Attribute VB_Name = "ConvSchemaNormalise"
Option Explicit

' Vérifie si une forme avec un ID donné existe sur une page
Private Function ShapeExiste(pg As Visio.Page, shapeID As Long) As Boolean
    On Error Resume Next
    Dim tmp As Visio.Shape
    Set tmp = pg.Shapes.ItemFromID(shapeID)
    ShapeExiste = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Public Sub ConvertirEnSchemaNormalise()
    On Error GoTo GestionErreurs

    Dim VisioApp As Visio.Application
    Dim DocSource As Visio.Document
    Dim PageSource As Visio.Page
    Dim PageCible As Visio.Page
    Dim shp As Visio.Shape
    Dim stencil As Visio.Document
    Dim masterFound As Visio.Master
    Dim nouvelleForme As Visio.Shape
    Dim fso As Object, folder As Object, file As Object
    Dim cheminStencils As String
    Dim stencilsOuverts As New Collection
    Dim pageIndexChoisi As Integer
    Dim compoManquant As New Collection
    Dim formesARemplacer As New Collection
    Dim formeData As Object
    Dim dictRemplacement As Object

    ' === Configuration ===
    cheminStencils = "C:\Users\timour.biart\OneDrive - Autographe\Documents\Librairie Autographe\"

    Set VisioApp = Visio.Application
    Set DocSource = VisioApp.ActiveDocument
    If DocSource Is Nothing Then
        MsgBox "Aucun document Visio ouvert.", vbExclamation
        Exit Sub
    End If

    ' --- Sélection de la page ---
    Dim pageIndex As Integer, pageCount As Integer, pageList As String
    pageCount = DocSource.Pages.Count
    If pageCount = 0 Then
        MsgBox "Le document ne contient aucune page.", vbExclamation
        Exit Sub
    End If
    pageList = "Veuillez sélectionner la page à convertir :" & vbCrLf
    For pageIndex = 1 To pageCount
        pageList = pageList & pageIndex & " - " & DocSource.Pages(pageIndex).Name & vbCrLf
    Next pageIndex
    Dim reponse As Variant
    reponse = InputBox(pageList, "Sélection de la page")
    If reponse = "" Then Exit Sub
    If Not IsNumeric(reponse) Then MsgBox "Entrée invalide.", vbExclamation: Exit Sub
    pageIndexChoisi = CInt(reponse)
    If pageIndexChoisi < 1 Or pageIndexChoisi > pageCount Then MsgBox "Numéro de page invalide.", vbExclamation: Exit Sub
    Set PageSource = DocSource.Pages(pageIndexChoisi)

    ' --- Copie de la page + renommage ---
    Dim nomFeuilleSource As String, nomFeuilleCible As String
    nomFeuilleSource = PageSource.Name
    If InStr(1, nomFeuilleSource, "Dessin Electrique", vbTextCompare) > 0 Then
        nomFeuilleCible = Replace(nomFeuilleSource, "Dessin Electrique", "Schéma Electrique Normalisé", 1, -1, vbTextCompare)
    Else
        nomFeuilleCible = nomFeuilleSource & " - Schéma Electrique Normalisé"
    End If
    Set PageCible = PageSource.Duplicate
    PageCible.Name = nomFeuilleCible
    Debug.Print "--- Démarrage du script ---"
    Debug.Print "Page copiée: " & PageCible.Name

    ' ============================================================
    ' === PHASE 1 : Recensement des formes 'A '
    ' ============================================================
    Debug.Print vbCrLf & "=== PHASE 1 : RECENSEMENT ==="

    For Each shp In PageCible.Shapes
        If Not shp.Master Is Nothing Then
            If Left$(shp.Master.Name, 2) = "A " Then
                Set formeData = CreateObject("Scripting.Dictionary")
                formeData.Add "ID", shp.ID
                formeData.Add "MasterName", shp.Master.Name
                formeData.Add "Text", shp.Text
                formeData.Add "PinX", shp.CellsU("PinX").Result("mm")
                formeData.Add "PinY", shp.CellsU("PinY").Result("mm")
                formesARemplacer.Add formeData
                Debug.Print "- Forme 'A' recensée : ID=" & shp.ID & ", Master='" & shp.Master.Name & "', Texte='" & shp.Text & "'"
            End If
        End If
    Next shp
    Debug.Print "Nombre total de formes 'A ' recensées : " & formesARemplacer.Count

    ' ============================================================
    ' === PHASE 2 : Ouverture stencils + remplacement des formes
    ' ============================================================
    Debug.Print vbCrLf & "=== PHASE 2 : REMPLACEMENT DES FORMES ==="
    Set dictRemplacement = CreateObject("Scripting.Dictionary")

    ' Vérification dossier stencils
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set folder = fso.GetFolder(cheminStencils)
    If Err.Number <> 0 Then
        MsgBox "Chemin stencils invalide : " & cheminStencils, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    Dim originalID As Long
    Dim nomMasterNormalise As String
    Dim trouve As Boolean

    For Each formeData In formesARemplacer
        originalID = CLng(formeData("ID"))
        nomMasterNormalise = "B " & Mid$(CStr(formeData("MasterName")), 3)

        trouve = False
        Set masterFound = Nothing

        ' Parcourt des fichiers de stencil
        For Each file In folder.Files
            If UCase$(fso.GetExtensionName(file.Name)) = "VSSX" Or UCase$(fso.GetExtensionName(file.Name)) = "VSS" Then
                On Error Resume Next
                Set stencil = VisioApp.Documents.OpenEx(file.Path, visOpenHidden)
                On Error GoTo 0
                If Not stencil Is Nothing Then
                    stencilsOuverts.Add stencil
                    On Error Resume Next
                    Set masterFound = stencil.Masters.item(nomMasterNormalise)
                    If Err.Number = 0 Then
                        trouve = True
                        On Error GoTo 0
                        Exit For
                    End If
                    Err.Clear
                    On Error GoTo 0
                End If
            End If
        Next file

        If trouve And Not masterFound Is Nothing Then
            Set nouvelleForme = PageCible.Drop(masterFound, formeData("PinX"), formeData("PinY"))
            ' Sécurise la position/texte
            nouvelleForme.CellsU("PinX").Result("mm") = formeData("PinX")
            nouvelleForme.CellsU("PinY").Result("mm") = formeData("PinY")
            nouvelleForme.Text = formeData("Text")
            dictRemplacement.Add CStr(originalID), nouvelleForme.ID
            Debug.Print "- Forme ID=" & originalID & " remplacée par '" & nomMasterNormalise & "' (nouvelle ID=" & nouvelleForme.ID & ")"
        Else
            compoManquant.Add nomMasterNormalise
            Debug.Print "!!! Master '" & nomMasterNormalise & "' non trouvé pour forme ID=" & originalID
        End If
    Next formeData

    Debug.Print "Remplacement terminé. " & dictRemplacement.Count & " forme(s) remplacée(s)."

    ' ============================================================
    ' === PHASE 3 : Suppression des anciennes formes 'A '
    ' ============================================================
    Debug.Print vbCrLf & "=== PHASE 3 : SUPPRESSION DES ANCIENNES FORMES ==="
    Dim nbFormesSuppr As Long: nbFormesSuppr = 0
    For Each formeData In formesARemplacer
        On Error Resume Next
        If ShapeExiste(PageCible, CLng(formeData("ID"))) Then
            PageCible.Shapes.ItemFromID(CLng(formeData("ID"))).Delete
            nbFormesSuppr = nbFormesSuppr + 1
            Debug.Print "- Forme ID=" & formeData("ID") & " supprimée."
        Else
            Debug.Print "- Forme ID=" & formeData("ID") & " déjà supprimée ou introuvable."
        End If
        On Error GoTo 0
    Next formeData

    ' ============================================================
    ' === Bilan final
    ' ============================================================
    Dim msg As String
    msg = "Conversion terminée." & vbCrLf & _
          "Formes 'A ' supprimées : " & nbFormesSuppr & vbCrLf & _
          "Formes remplacées : " & dictRemplacement.Count
    If compoManquant.Count > 0 Then
        Dim s As Variant
        msg = msg & vbCrLf & vbCrLf & "Masters non trouvés :" & vbCrLf
        For Each s In compoManquant
            msg = msg & " - " & CStr(s) & vbCrLf
        Next s
    End If
    MsgBox msg, vbInformation

Fin:
    On Error Resume Next
    Dim sdoc As Variant
    For Each sdoc In stencilsOuverts
        If Not sdoc Is Nothing Then sdoc.Close
    Next sdoc
    On Error GoTo 0
    Exit Sub

GestionErreurs:
    MsgBox "Erreur: " & Err.Description & " (ligne " & Erl & ")", vbCritical
    Resume Fin
End Sub


