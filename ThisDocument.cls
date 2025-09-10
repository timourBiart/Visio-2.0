VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' -----------------------------
' THISDOCUMENT – TEMPLATE PROTÉGÉ COMPLET AVEC ToC
' -----------------------------

' --- Constantes ---
Private Const PAGE_GARDE As String = "Page de garde"
Private Const PAGE_BOM As String = "BOM"
Private Const PAGE_ListeCables As String = "Liste des câbles"
Private Const PAGE_CARTOUCHE As String = "Cartouche"
Private Const MAX_PAGES As Integer = 20
Private Const LEFT_MM As Double = 220
Private Const TOP_MM As Double = 45.6
Private Const COL_NUM_WIDTH_MM As Double = 13.65
Private Const COL_NAME_WIDTH_MM As Double = 180
Private Const LINE_HEIGHT_MM As Double = 7
Private Const FONT_SIZE_PT As Integer = 12

' ======================
' === ÉVÉNEMENTS =======
' ======================

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
    On Error Resume Next
    CreatePageDeGarde
    MettreAJourToutSansAlerte
End Sub

Private Sub Document_PageAdded(ByVal Page As IVPage)
    On Error Resume Next
    With Page.PageSheet
        .CellsU("PageWidth").Result("mm") = 420
        .CellsU("PageHeight").Result("mm") = 297
        .CellsU("PrintPageOrientation").FormulaU = "2"
    End With
    
    Dim cartouchePage As Visio.Page
    Set cartouchePage = GetPageByName(PAGE_CARTOUCHE)
    If LCase$(Page.Name) <> LCase$(PAGE_GARDE) And _
       LCase$(Page.Name) <> LCase$(PAGE_CARTOUCHE) And _
       LCase$(Page.Name) <> LCase$(PAGE_BOM) And _
       LCase$(Page.Name) <> LCase$(PAGE_ListeCables) Then
        If Not cartouchePage Is Nothing Then Page.BackPage = PAGE_CARTOUCHE
    End If
    
    MettreAJourToutSansAlerte
End Sub

Private Sub Document_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    MettreAJourToutSansAlerte
End Sub

Private Sub Document_BeforePrint(ByVal PrintFlags As Long)
    On Error Resume Next
    MettreAJourToutSansAlerte
End Sub

Private Sub Document_PageRenamed(ByVal Page As IVPage, ByVal oldName As String)
    On Error Resume Next
    Dim np As String: np = LCase$(Page.Name)
    If InStr(np, "bom") > 0 And LCase$(oldName) = "bom" Then
        MsgBox "Il est interdit de renommer la page BOM !", vbExclamation
        Page.Name = PAGE_BOM
    ElseIf InStr(np, "page de garde") > 0 And LCase$(oldName) = LCase$(PAGE_GARDE) Then
        MsgBox "Il est interdit de renommer la Page de garde !", vbExclamation
        Page.Name = PAGE_GARDE
    ElseIf InStr(np, "cartouche") > 0 And LCase$(oldName) = LCase$(PAGE_CARTOUCHE) Then
        MsgBox "Il est interdit de renommer la page Cartouche !", vbExclamation
        Page.Name = PAGE_CARTOUCHE
    ElseIf InStr(np, LCase$(PAGE_ListeCables)) > 0 And LCase$(oldName) = LCase$(PAGE_ListeCables) Then
        MsgBox "Il est interdit de renommer la page " & PAGE_ListeCables & " !", vbExclamation
        Page.Name = PAGE_ListeCables
    End If
    MettreAJourToutSansAlerte
End Sub

Private Sub Document_PageDeleted(ByVal Page As IVPage)
    On Error Resume Next
    Dim n As String: n = LCase$(Page.nameU)
    If n = "bom" Then
        MsgBox "Il est interdit de supprimer la page BOM !", vbExclamation
    ElseIf n = LCase$(PAGE_GARDE) Then
        MsgBox "Il est interdit de supprimer la Page de garde !", vbExclamation
    ElseIf n = LCase$(PAGE_CARTOUCHE) Then
        MsgBox "Il est interdit de supprimer la page Cartouche !", vbExclamation
    ElseIf n = "liste des câbles" Or n = "liste des cables" Then
        MsgBox "Il est interdit de supprimer la page " & PAGE_ListeCables & " !", vbExclamation
    End If
    MettreAJourToutSansAlerte
End Sub

Private Function Document_QueryCancelPageDelete(ByVal Page As IVPage) As Boolean
    On Error Resume Next
    If LCase$(Page.Name) = LCase$(PAGE_ListeCables) Or LCase$(Page.Name) = LCase$(PAGE_BOM) Then
        MsgBox "Suppression refusée : la page """ & Page.Name & """ est protégée.", vbExclamation
        Document_QueryCancelPageDelete = True
    End If
End Function

Private Function Document_QueryCancelPageRename(ByVal Page As IVPage, ByVal oldName As String, ByVal newName As String) As Boolean
    On Error Resume Next
    If LCase$(oldName) = LCase$(PAGE_ListeCables) Or LCase$(oldName) = LCase$(PAGE_BOM) Then
        MsgBox "Renommage refusé : la page """ & oldName & """ est protégée.", vbExclamation
        Document_QueryCancelPageRename = True
    End If
End Function

Private Function Document_QueryCancelPageMove(ByVal Page As IVPage, ByVal Before As IVPage) As Boolean
    On Error Resume Next
    If LCase$(Page.Name) = LCase$(PAGE_ListeCables) Or LCase$(Page.Name) = LCase$(PAGE_BOM) Then
        MsgBox "Déplacement refusé : la page """ & Page.Name & """ a une position imposée.", vbExclamation
        Document_QueryCancelPageMove = True
    End If
End Function

' ======================
' === GESTION DES PAGES =====
' ======================

Public Sub ReordonnerPages()
    On Error GoTo Fin
    
    Dim gardePage As Visio.Page, bomPage As Visio.Page, listePage As Visio.Page
    Dim gardeNames As Variant, bomNames As Variant, listeNames As Variant
    gardeNames = Array(PAGE_GARDE, "Page_de_garde", "Garde", "00_Garde")
    bomNames = Array(PAGE_BOM, "Z_BOM", "ZZ_BOM", "Liste des composants")
    listeNames = Array(PAGE_ListeCables, "Liste des cables", "Liste_des_câbles", "Liste_des_cables")
    
    Set gardePage = FindPageByNames(gardeNames)
    Set bomPage = FindPageByNames(bomNames)
    Set listePage = FindPageByNames(listeNames)
    
    Dim n As Integer: n = ThisDocument.Pages.Count
    If n <= 0 Then GoTo Fin
    
    If Not gardePage Is Nothing Then SetIndexSafe gardePage, 1
    If Not bomPage Is Nothing Then
        If Not listePage Is Nothing And n >= 2 Then
            SetIndexSafe bomPage, n - 1        ' BOM en avant-dernière
        Else
            SetIndexSafe bomPage, n
        End If
    End If
    If Not listePage Is Nothing Then SetIndexSafe listePage, n  ' Liste des câbles en dernière
Fin:
End Sub

Private Sub SetIndexSafe(ByVal p As Visio.Page, ByVal target As Integer)
    On Error Resume Next
    Dim n As Integer: n = ThisDocument.Pages.Count
    If target < 1 Then target = 1
    If target > n Then target = n
    If p.Index <> target Then p.Index = target
End Sub

' ======================
' === CREATION PAGE DE GARDE + ToC ===
' ======================

Public Sub CreatePageDeGarde()
    Dim pgGarde As Visio.Page
    Set pgGarde = GetPageByName(PAGE_GARDE)
    If pgGarde Is Nothing Then
        Set pgGarde = ThisDocument.Pages.Add
        pgGarde.nameU = PAGE_GARDE
    End If
End Sub

Public Sub UpdateTOC()
    Dim pgGarde As Visio.Page
    Set pgGarde = GetPageByName(PAGE_GARDE)
    If pgGarde Is Nothing Then Exit Sub
    
    DeleteShapeIfExists pgGarde, "TOC_TABLE"
    DeleteShapeIfExists pgGarde, "TOC_COL_LINE"
    
    Dim i As Integer
    Dim xLeft As Double, yTopPos As Double, xColSepar As Double, currentY As Double
    xLeft = mmToIn(LEFT_MM)
    xColSepar = xLeft + mmToIn(COL_NUM_WIDTH_MM)
    yTopPos = mmToIn(297 - TOP_MM)
    currentY = yTopPos
    
    Dim shpNum As Visio.Shape, shpName As Visio.Shape, p As Visio.Page
    Set shpNum = pgGarde.DrawRectangle(xLeft, currentY, xColSepar, currentY - mmToIn(LINE_HEIGHT_MM))
    shpNum.Text = "Page"
    shpNum.nameU = "TOC_TABLE_TITLE_NUM"
    shpNum.CellsU("Char.Size").FormulaU = FONT_SIZE_PT & " pt"
    shpNum.CellsU("LinePattern").ResultIU = 1
    shpNum.CellsU("FillPattern").ResultIU = 0
    shpNum.CellsU("Para.HorzAlign").ResultIU = 0
    
    Set shpName = pgGarde.DrawRectangle(xColSepar, currentY, xLeft + mmToIn(COL_NUM_WIDTH_MM + COL_NAME_WIDTH_MM), currentY - mmToIn(LINE_HEIGHT_MM))
    shpName.Text = "Page Name"
    shpName.nameU = "TOC_TABLE_TITLE_NAME"
    shpName.CellsU("Char.Size").FormulaU = FONT_SIZE_PT & " pt"
    shpName.CellsU("LinePattern").ResultIU = 1
    shpName.CellsU("FillPattern").ResultIU = 0
    shpName.CellsU("Para.HorzAlign").ResultIU = 0
    
    currentY = currentY - mmToIn(LINE_HEIGHT_MM)
    
    For i = 1 To ThisDocument.Pages.Count
        If i > MAX_PAGES Then Exit For
        Set p = ThisDocument.Pages(i)
        If LCase$(p.Name) <> LCase$(PAGE_CARTOUCHE) Then
            Set shpNum = pgGarde.DrawRectangle(xLeft, currentY, xColSepar, currentY - mmToIn(LINE_HEIGHT_MM))
            shpNum.Text = p.Index
            shpNum.nameU = "TOC_TABLE_" & p.Index & "_NUM"
            shpNum.CellsU("Char.Size").FormulaU = FONT_SIZE_PT & " pt"
            shpNum.CellsU("LinePattern").ResultIU = 1
            shpNum.CellsU("FillPattern").ResultIU = 0
            shpNum.CellsU("Para.HorzAlign").ResultIU = 1
            
            Set shpName = pgGarde.DrawRectangle(xColSepar, currentY, xLeft + mmToIn(COL_NUM_WIDTH_MM + COL_NAME_WIDTH_MM), currentY - mmToIn(LINE_HEIGHT_MM))
            shpName.Text = p.Name
            shpName.nameU = "TOC_TABLE_" & p.Index & "_NAME"
            shpName.CellsU("Char.Size").FormulaU = FONT_SIZE_PT & " pt"
            shpName.CellsU("LinePattern").ResultIU = 1
            shpName.CellsU("FillPattern").ResultIU = 0
            shpName.CellsU("Para.HorzAlign").ResultIU = 0
            
            currentY = currentY - mmToIn(LINE_HEIGHT_MM)
        End If
    Next i
End Sub

' ======================
' === GESTION DE LA BOM ===
' ======================

Public Sub UpdateBOM()
    Dim pgBOM As Visio.Page
    Set pgBOM = GetPageByName(PAGE_BOM)
    If pgBOM Is Nothing Then
        Set pgBOM = ThisDocument.Pages.Add
        pgBOM.nameU = PAGE_BOM
    End If
    
    DeleteShapeIfExists pgBOM, "BOM_TABLE"
    DeleteShapeIfExists pgBOM, "BOM_TITLE"
    
    Dim pageWidth As Double, pageHeight As Double
    pageWidth = mmToIn(420): pageHeight = mmToIn(297)
    Dim leftMarginIn As Double, rightMarginIn As Double
    leftMarginIn = mmToIn(10): rightMarginIn = mmToIn(10)
    Dim availableWidth As Double: availableWidth = pageWidth - leftMarginIn - rightMarginIn
    
    Dim colWeights() As Double, colWidths() As Double, k As Integer, totalWeight As Double
    ReDim colWeights(1 To 5): ReDim colWidths(1 To 5)
    colWeights(1) = 1: colWeights(2) = 1: colWeights(3) = 2: colWeights(4) = 2: colWeights(5) = 1
    For k = 1 To 5: totalWeight = totalWeight + colWeights(k): Next k
    For k = 1 To 5: colWidths(k) = availableWidth * (colWeights(k) / totalWeight): Next k
    
    Dim xLeft As Double, currentY As Double, xPos As Double
    xLeft = leftMarginIn
    currentY = pageHeight - mmToIn(TOP_MM) + mmToIn(5)
    
    Dim headers As Variant: headers = Array("Fournisseur", "Marque", "Nom de la pièce", "Numéro de pièce", "Quantité")
    
    xPos = xLeft
    For k = LBound(headers) To UBound(headers)
        Dim headerShp As Visio.Shape
        Set headerShp = pgBOM.DrawRectangle(xPos, currentY, xPos + colWidths(k + 1), currentY - mmToIn(LINE_HEIGHT_MM))
        headerShp.Text = headers(k)
        headerShp.CellsU("Char.Size").FormulaU = FONT_SIZE_PT & " pt"
        headerShp.CellsU("LinePattern").ResultIU = 1
        headerShp.CellsU("FillPattern").ResultIU = 0
        headerShp.CellsU("Para.HorzAlign").ResultIU = 1
        headerShp.nameU = "BOM_TABLE_HEADER_" & k + 1
        xPos = xPos + colWidths(k + 1)
    Next k
    currentY = currentY - mmToIn(LINE_HEIGHT_MM)
    
    Dim bomItems As Object: Set bomItems = CollectShapeData()
    If bomItems.Count > 0 Then
        Dim itemKey As Variant, i As Integer: i = 1
        For Each itemKey In bomItems.keys
            Dim bomData As Object: Set bomData = bomItems(itemKey)
            xPos = xLeft
            
            Dim s As Visio.Shape
            Set s = pgBOM.DrawRectangle(xPos, currentY, xPos + colWidths(1), currentY - mmToIn(LINE_HEIGHT_MM))
            s.Text = bomData("Prop.Fournisseur"): s.nameU = "BOM_TABLE_ROW_" & i & "_Fournisseur"
            s.CellsU("Char.Size").FormulaU = FONT_SIZE_PT & " pt": s.CellsU("LinePattern").ResultIU = 1: s.CellsU("FillPattern").ResultIU = 0: s.CellsU("Para.HorzAlign").ResultIU = 0
            xPos = xPos + colWidths(1)
            
            Set s = pgBOM.DrawRectangle(xPos, currentY, xPos + colWidths(2), currentY - mmToIn(LINE_HEIGHT_MM))
            s.Text = bomData("Prop.Marque"): s.nameU = "BOM_TABLE_ROW_" & i & "_Marque"
            s.CellsU("Char.Size").FormulaU = FONT_SIZE_PT & " pt": s.CellsU("LinePattern").ResultIU = 1: s.CellsU("FillPattern").ResultIU = 0: s.CellsU("Para.HorzAlign").ResultIU = 0
            xPos = xPos + colWidths(2)
            
            Set s = pgBOM.DrawRectangle(xPos, currentY, xPos + colWidths(3), currentY - mmToIn(LINE_HEIGHT_MM))
            s.Text = bomData("Prop.NomPiece"): s.nameU = "BOM_TABLE_ROW_" & i & "_NomPiece"
            s.CellsU("Char.Size").FormulaU = FONT_SIZE_PT & " pt": s.CellsU("LinePattern").ResultIU = 1: s.CellsU("FillPattern").ResultIU = 0: s.CellsU("Para.HorzAlign").ResultIU = 0
            xPos = xPos + colWidths(3)
            
            Set s = pgBOM.DrawRectangle(xPos, currentY, xPos + colWidths(4), currentY - mmToIn(LINE_HEIGHT_MM))
            s.Text = bomData("Prop.NumPiece"): s.nameU = "BOM_TABLE_ROW_" & i & "_NumPiece"
            s.CellsU("Char.Size").FormulaU = FONT_SIZE_PT & " pt": s.CellsU("LinePattern").ResultIU = 1: s.CellsU("FillPattern").ResultIU = 0: s.CellsU("Para.HorzAlign").ResultIU = 0
            xPos = xPos + colWidths(4)
            
            Set s = pgBOM.DrawRectangle(xPos, currentY, xPos + colWidths(5), currentY - mmToIn(LINE_HEIGHT_MM))
            s.Text = bomData("Quantite"): s.nameU = "BOM_TABLE_ROW_" & i & "_Quantite"
            s.CellsU("Char.Size").FormulaU = FONT_SIZE_PT & " pt": s.CellsU("LinePattern").ResultIU = 1: s.CellsU("FillPattern").ResultIU = 0: s.CellsU("Para.HorzAlign").ResultIU = 1
            
            currentY = currentY - mmToIn(LINE_HEIGHT_MM)
            i = i + 1
        Next itemKey
    End If
End Sub

' ======================
' === LISTE DES CÂBLES (groupée par CouleurFil) ===
' ======================

Public Sub UpdateCableList()
    Dim pgC As Visio.Page
    Set pgC = GetPageByName(PAGE_ListeCables)
    If pgC Is Nothing Then
        Set pgC = ThisDocument.Pages.Add
        pgC.nameU = PAGE_ListeCables
    End If
    
    DeleteShapeIfExists pgC, "CABLE_TABLE"
    DeleteShapeIfExists pgC, "CABLE_TITLE"
    
    Dim pageWidth As Double, pageHeight As Double
    pageWidth = mmToIn(420): pageHeight = mmToIn(297)
    Dim leftMarginIn As Double, rightMarginIn As Double
    leftMarginIn = mmToIn(10): rightMarginIn = mmToIn(10)
    Dim availableWidth As Double: availableWidth = pageWidth - leftMarginIn - rightMarginIn
    
    Dim xLeft As Double, currentY As Double
    xLeft = leftMarginIn
    currentY = pageHeight - mmToIn(TOP_MM) + mmToIn(5)
    
    Dim headers As Variant
    headers = Array( _
        "CableID", "RepereFil", "Longueur", "Couleurs", _
        "FromShape", "FromPin", "FromConnType", _
        "ToShape", "ToPin", "ToConnType" _
    )
    
    Dim colWeights() As Double, colWidths() As Double, k As Integer, totalW As Double
    ReDim colWeights(1 To 10): ReDim colWidths(1 To 10)
    colWeights(1) = 1: colWeights(2) = 1: colWeights(3) = 1: colWeights(4) = 1.6
    colWeights(5) = 1.6: colWeights(6) = 1: colWeights(7) = 1.4
    colWeights(8) = 1.6: colWeights(9) = 1: colWeights(10) = 1.4
    For k = 1 To 10: totalW = totalW + colWeights(k): Next k
    For k = 1 To 10: colWidths(k) = availableWidth * (colWeights(k) / totalW): Next k
    
    Dim xPos As Double: xPos = xLeft
    For k = LBound(headers) To UBound(headers)
        Dim h As Visio.Shape
        Set h = pgC.DrawRectangle(xPos, currentY, xPos + colWidths(k + 1), currentY - mmToIn(LINE_HEIGHT_MM))
        h.Text = headers(k)
        h.CellsU("Char.Size").FormulaU = FONT_SIZE_PT & " pt"
        h.CellsU("LinePattern").ResultIU = 1
        h.CellsU("FillPattern").ResultIU = 0
        h.CellsU("Para.HorzAlign").ResultIU = 1
        h.nameU = "CABLE_TABLE_HEADER_" & (k + 1)
        xPos = xPos + colWidths(k + 1)
    Next k
    currentY = currentY - mmToIn(LINE_HEIGHT_MM)
    
    Dim rows As Collection
    Set rows = CollectCableData() ' groupé par CouleurFil
    
    If rows.Count > 0 Then
        Dim i As Long: i = 1
        Dim r As Variant
        For Each r In rows
            Dim vals() As String
            vals = r  ' 0..9
            xPos = xLeft
            For k = 0 To 9
                Dim cellShape As Visio.Shape
                Set cellShape = pgC.DrawRectangle(xPos, currentY, xPos + colWidths(k + 1), currentY - mmToIn(LINE_HEIGHT_MM))
                cellShape.Text = vals(k)
                cellShape.nameU = "CABLE_TABLE_ROW_" & i & "_" & (k + 1)
                cellShape.CellsU("Char.Size").FormulaU = FONT_SIZE_PT & " pt"
                cellShape.CellsU("LinePattern").ResultIU = 1
                cellShape.CellsU("FillPattern").ResultIU = 0
                If k = 5 Or k = 6 Or k = 9 Or k = 8 Then
                    cellShape.CellsU("Para.HorzAlign").ResultIU = 1
                Else
                    cellShape.CellsU("Para.HorzAlign").ResultIU = 0
                End If
                xPos = xPos + colWidths(k + 1)
            Next k
            currentY = currentY - mmToIn(LINE_HEIGHT_MM)
            i = i + 1
        Next r
    End If
End Sub

' ======================
' === MACROS BOUTONS (optionnelles) ===
' ======================

' S’assurer que la page "Liste des câbles" existe
Private Sub EnsureListeCablesPage()
    On Error Resume Next
    Dim p As Visio.Page
    Set p = GetPageByName(PAGE_ListeCables)
    If p Is Nothing Then
        Set p = ThisDocument.Pages.Add
        p.nameU = PAGE_ListeCables
    End If
End Sub

' (Optionnel) Bouton : mise à jour complète
Public Sub Bouton_MiseAJour_Documents(Optional ByVal context As String)
    On Error Resume Next
    MettreAJourToutSansAlerte context
End Sub

' (Optionnel) Bouton : uniquement Liste des câbles (+ ToC/ordre)
Public Sub Bouton_MiseAJour_ListeCables(Optional ByVal context As String)
    On Error Resume Next
    EnsureListeCablesPage
    UpdateCableList
    ReordonnerPages
    UpdateTOC
End Sub

' (Optionnel) Compat : ancien nom
Public Sub Bouton_MiseAJour(Optional ByVal context As String)
    On Error Resume Next
    MettreAJourToutSansAlerte context
End Sub

' ======================
' === MACRO CENTRALE POUR LE BOUTON ===
' ======================

Public Sub MettreAJourToutSansAlerte(Optional ByVal context As String)
    On Error Resume Next
    EnsureListeCablesPage          ' garantit la page
    ReordonnerPages                ' Liste des câbles dernière, BOM avant-dernière
    UpdateTOC
    UpdateBOM
    UpdateCableList
End Sub

' ======================
' === UTILITAIRES GÉNÉRAUX ======
' ======================

Private Function CollectShapeData() As Object
    Dim bomItems As Object: Set bomItems = CreateObject("Scripting.Dictionary")
    Dim p As Visio.Page, shp As Visio.Shape
    For Each p In ThisDocument.Pages
        If LCase$(p.Name) <> LCase$(PAGE_GARDE) And _
           LCase$(p.Name) <> LCase$(PAGE_CARTOUCHE) And _
           LCase$(p.Name) <> LCase$(PAGE_BOM) And _
           LCase$(p.Name) <> LCase$(PAGE_ListeCables) Then
            For Each shp In p.Shapes
                On Error Resume Next
                Dim nomPiece As String, numPiece As String, fournisseur As String, marque As String
                If shp.CellExistsU("Prop.Fournisseur", visExistsAnywhere) Then fournisseur = shp.CellsU("Prop.Fournisseur").ResultStr("")
                If shp.CellExistsU("Prop.NomPiece", visExistsAnywhere) Then nomPiece = shp.CellsU("Prop.NomPiece").ResultStr("")
                If shp.CellExistsU("Prop.NumPiece", visExistsAnywhere) Then numPiece = shp.CellsU("Prop.NumPiece").ResultStr("")
                If shp.CellExistsU("Prop.Marque", visExistsAnywhere) Then marque = shp.CellsU("Prop.Marque").ResultStr("")
                On Error GoTo 0
                If fournisseur <> "" Or nomPiece <> "" Or numPiece <> "" Or marque <> "" Then
                    Dim key As String: key = fournisseur & "|" & marque & "|" & nomPiece & "|" & numPiece
                    If bomItems.Exists(key) Then
                        Dim cur As Object: Set cur = bomItems(key)
                        cur("Quantite") = cur("Quantite") + 1
                    Else
                        Dim it As Object: Set it = CreateObject("Scripting.Dictionary")
                        it.Add "Prop.NomPiece", nomPiece
                        it.Add "Prop.NumPiece", numPiece
                        it.Add "Prop.Fournisseur", fournisseur
                        it.Add "Prop.Marque", marque
                        it.Add "Quantite", 1
                        bomItems.Add key, it
                    End If
                End If
            Next shp
        End If
    Next p
    Set CollectShapeData = bomItems
End Function

Private Function GetPageByName(ByVal nameU As String) As Visio.Page
    Dim pg As Visio.Page
    For Each pg In ThisDocument.Pages
        If LCase$(pg.nameU) = LCase$(nameU) Then
            Set GetPageByName = pg
            Exit Function
        End If
    Next
End Function

Private Sub DeleteShapeIfExists(pg As Visio.Page, ByVal baseName As String)
    Dim sh As Visio.Shape, i As Long
    For i = pg.Shapes.Count To 1 Step -1
        Set sh = pg.Shapes(i)
        If Left$(LCase$(sh.nameU), Len(baseName)) = LCase$(baseName) Then sh.Delete
    Next i
End Sub

Private Function mmToIn(ByVal mm As Double) As Double
    mmToIn = mm / 25.4
End Function

Private Function FindPageByNames(names As Variant) As Visio.Page
    Dim p As Visio.Page, i As Long
    For Each p In ThisDocument.Pages
        For i = LBound(names) To UBound(names)
            If LCase$(Trim$(p.nameU)) = LCase$(Trim$(names(i))) Or _
               LCase$(Trim$(p.Name)) = LCase$(Trim$(names(i))) Then
                Set FindPageByNames = p
                Exit Function
            End If
        Next i
    Next p
End Function

' ======================
' === UTILITAIRES CÂBLES & COULEURS ===
' ======================

' Valeur affichée (évaluée) d’une Prop.* sur la forme courante
Private Function GetPropStrHere(ByVal shp As Visio.Shape, ByVal cellU As String) As String
    On Error Resume Next
    If shp.CellExistsU(cellU, visExistsAnywhere) Then
        GetPropStrHere = shp.CellsU(cellU).ResultStr("")
    End If
End Function

' Trouve la 1D "porteuse" (tronc) à l’intérieur d’un groupe
Private Function FindConnector1D(ByVal shp As Visio.Shape) As Visio.Shape
    On Error Resume Next
    If shp.OneD = 1 Then
        Set FindConnector1D = shp
        Exit Function
    End If
    If shp.Type = visTypeGroup Then
        Dim s As Visio.Shape, nm As String
        For Each s In shp.Shapes
            If s.OneD = 1 Then
                nm = LCase$(s.nameU)
                If InStr(nm, "tronc") > 0 Or InStr(nm, "trunk") > 0 Or InStr(nm, "main") > 0 Then
                    Set FindConnector1D = s
                    Exit Function
                End If
            End If
        Next s
        For Each s In shp.Shapes
            If s.OneD = 1 Then Set FindConnector1D = s: Exit Function
        Next s
        For Each s In shp.Shapes
            Dim cand As Visio.Shape
            Set cand = FindConnector1D(s)
            If Not cand Is Nothing Then Set FindConnector1D = cand: Exit Function
        Next s
    End If
End Function

' Lit les extrémités d'un connecteur 1D (From/To) — (non utilisée directement mais utile)
Private Sub InferConnectorEndpoints(ByVal host As Visio.Shape, _
                                    ByRef FromShape As String, ByRef FromPin As String, _
                                    ByRef ToShape As String, ByRef ToPin As String)
    On Error Resume Next
    FromShape = "": FromPin = "": ToShape = "": ToPin = ""
    
    Dim i As Integer
    For i = 1 To host.Connects.Count
        Dim c As Visio.Connect
        Set c = host.Connects(i)
        If c.FromSheet Is host Then
            Dim src As String: src = ""
            If Not c.FromCell Is Nothing Then src = LCase$(c.FromCell.nameU)
            Dim tShape As Visio.Shape: Set tShape = c.ToSheet
            Dim tPin As String: tPin = "": If Not c.ToCell Is Nothing Then tPin = c.ToCell.nameU
            
            If InStr(src, "beginx") > 0 Then
                FromShape = tShape.Name: FromPin = tPin
            ElseIf InStr(src, "endx") > 0 Then
                ToShape = tShape.Name: ToPin = tPin
            Else
                If FromShape = "" Then
                    FromShape = tShape.Name: FromPin = tPin
                ElseIf ToShape = "" Then
                    ToShape = tShape.Name: ToPin = tPin
                End If
            End If
        End If
    Next i
    
    If FromShape = "" And ToShape <> "" Then
        FromShape = ToShape: FromPin = ToPin
        ToShape = "": ToPin = ""
    End If
    On Error GoTo 0
End Sub

' Conteneur direct si la forme est dans un groupe
Private Function GetContainingShape(ByVal shp As Visio.Shape) As Visio.Shape
    On Error Resume Next
    Dim parentShapes As Object
    Set parentShapes = shp.Parent
    If Not parentShapes Is Nothing Then
        Dim owner As Object
        Set owner = parentShapes.Parent
        If TypeName(owner) = "IVShape" Or TypeName(owner) = "Shape" Then
            Set GetContainingShape = owner
        End If
    End If
End Function

' Remonte jusqu’au conteneur racine (groupe le plus haut)
Private Function TopContainerShape(ByVal shp As Visio.Shape) As Visio.Shape
    On Error Resume Next
    Dim cur As Visio.Shape: Set cur = shp
    Dim nxt As Visio.Shape
    Do
        Set nxt = GetContainingShape(cur)
        If nxt Is Nothing Then Exit Do
        Set cur = nxt
    Loop
    Set TopContainerShape = cur
End Function

' Vrai si deux formes appartiennent au même conteneur racine
Private Function SameTopContainer(ByVal a As Visio.Shape, ByVal b As Visio.Shape) As Boolean
    On Error Resume Next
    Dim ta As Visio.Shape, tb As Visio.Shape
    Set ta = TopContainerShape(a)
    Set tb = TopContainerShape(b)
    If (ta Is Nothing) Or (tb Is Nothing) Then
        SameTopContainer = False
    Else
        SameTopContainer = (ta.UniqueID(visGetOrMakeGUID) = tb.UniqueID(visGetOrMakeGUID))
    End If
End Function

' Extrémités "externes" d’une 1D (ignore connexions vers sous-formes du même master)
Private Function ExtractExternalEndpoints(ByVal host As Visio.Shape, ByVal ownerTop As Visio.Shape) As Collection
    On Error Resume Next
    Dim eps As New Collection
    Dim i As Integer
    For i = 1 To host.Connects.Count
        Dim c As Visio.Connect: Set c = host.Connects(i)
        If c.FromSheet Is host Then
            Dim tShape As Visio.Shape: Set tShape = c.ToSheet
            If Not tShape Is Nothing Then
                If Not SameTopContainer(ownerTop, tShape) Then
                    Dim pinName As String: pinName = ""
                    If Not c.ToCell Is Nothing Then pinName = c.ToCell.nameU
                    Dim arr(0 To 1) As String
                    arr(0) = tShape.Name
                    arr(1) = pinName
                    eps.Add arr
                End If
            End If
        End If
    Next i
    Set ExtractExternalEndpoints = eps
End Function

' Place un endpoint dans From* si vide, sinon dans To* si vide
Private Sub MergeEndpointIntoItem(ByRef item As Object, ByVal shName As String, ByVal pinName As String, _
                                  ByVal fromType As String, ByVal toType As String)
    On Error Resume Next
    If item("FromShape") = "" Then
        item("FromShape") = shName
        item("FromPin") = pinName
        If item("FromConnType") = "" Then
            item("FromConnType") = IIf(fromType <> "", fromType, toType)
        End If
    ElseIf item("ToShape") = "" Then
        item("ToShape") = shName
        item("ToPin") = pinName
        If item("ToConnType") = "" Then
            item("ToConnType") = IIf(toType <> "", toType, fromType)
        End If
    End If
End Sub

' ======================
' === COLLECTE LISTE CÂBLES : regroupement par CouleurFil ===
' ======================

' Parcours récursif : traite toute sous-forme portant Prop.CouleurFil
Private Sub CollectColorItemsFromShape(ByVal shp As Visio.Shape, ByRef colorMap As Object)
    On Error Resume Next
    
    Dim couleur As String
    couleur = GetPropStrHere(shp, "Prop.CouleurFil")
    
    If couleur <> "" Then
        Dim cableID As String, repere As String, longueur As String
        Dim fromType As String, toType As String
        cableID = GetPropStrHere(shp, "Prop.CableID")
        repere = GetPropStrHere(shp, "Prop.RepereFil")
        longueur = GetPropStrHere(shp, "Prop.Longueur")
        fromType = GetPropStrHere(shp, "Prop.FromConnType")
        toType = GetPropStrHere(shp, "Prop.ToConnType")
        
        Dim host As Visio.Shape, ownerTop As Visio.Shape
        Set host = Nothing: Set ownerTop = Nothing
        If shp.OneD = 1 Then
            Set host = shp
            Set ownerTop = TopContainerShape(shp)
        ElseIf shp.Type = visTypeGroup Then
            Set host = FindConnector1D(shp)
            Set ownerTop = TopContainerShape(shp)
        End If
        
        Dim eps As Collection
        If Not host Is Nothing Then
            Set eps = ExtractExternalEndpoints(host, ownerTop)
        End If
        
        If Not colorMap.Exists(couleur) Then
            Dim bucket As Object
            Set bucket = CreateObject("Scripting.Dictionary")
            Dim col As New Collection
            bucket.Add "Items", col
            colorMap.Add couleur, bucket
        End If
        
        Dim bucketRef As Object
        Set bucketRef = colorMap(couleur)
        
        Dim items As Collection
        Set items = bucketRef("Items")
        
        Dim item As Object
        If items.Count > 0 Then
            Set item = items(items.Count)
            If item("FromShape") <> "" And item("ToShape") <> "" Then
                Set item = Nothing
            End If
        End If
        
        If item Is Nothing Then
            Set item = CreateObject("Scripting.Dictionary")
            item.Add "CableID", ""
            item.Add "RepereFil", ""
            item.Add "Longueur", ""
            item.Add "Couleur", couleur
            item.Add "FromShape", "": item.Add "FromPin", "": item.Add "FromConnType", ""
            item.Add "ToShape", "": item.Add "ToPin", "": item.Add "ToConnType", ""
            items.Add item
        End If
        
        If item("CableID") = "" And cableID <> "" Then item("CableID") = cableID
        If item("RepereFil") = "" And repere <> "" Then item("RepereFil") = repere
        If item("Longueur") = "" And longueur <> "" Then item("Longueur") = longueur
        
        If Not eps Is Nothing Then
            Dim e As Variant
            For Each e In eps
                MergeEndpointIntoItem item, e(0), e(1), fromType, toType
            Next e
        End If
    End If
    
    If shp.Type = visTypeGroup Then
        Dim s As Visio.Shape
        For Each s In shp.Shapes
            CollectColorItemsFromShape s, colorMap
        Next s
    End If
End Sub

' Fonction principale : retourne lignes regroupées par couleur et par paires From/To
Private Function CollectCableData() As Collection
    On Error Resume Next
    
    Dim colorMap As Object
    Set colorMap = CreateObject("Scripting.Dictionary") ' key = CouleurFil ; value = bucket dict ("Items" Collection)
    
    Dim p As Visio.Page, shp As Visio.Shape
    For Each p In ThisDocument.Pages
        If LCase$(p.Name) <> LCase$(PAGE_GARDE) And _
           LCase$(p.Name) <> LCase$(PAGE_CARTOUCHE) And _
           LCase$(p.Name) <> LCase$(PAGE_BOM) And _
           LCase$(p.Name) <> LCase$(PAGE_ListeCables) Then
            For Each shp In p.Shapes
                CollectColorItemsFromShape shp, colorMap
            Next shp
        End If
    Next p
    
    Dim rows As New Collection
    Dim keys As Variant, idx As Long, keyStr As String
    Dim bucket As Object, items As Collection, i As Long
    keys = colorMap.keys
    If Not IsEmpty(keys) Then
        For idx = LBound(keys) To UBound(keys)
            keyStr = CStr(keys(idx))
            Set bucket = colorMap(keyStr)
            Set items = bucket("Items")
            For i = 1 To items.Count
                Dim it As Object
                Set it = items(i)
                
                Dim arr(0 To 9) As String
                arr(0) = it("CableID")
                arr(1) = it("RepereFil")
                arr(2) = it("Longueur")
                arr(3) = it("Couleur")
                arr(4) = it("FromShape")
                arr(5) = it("FromPin")
                arr(6) = it("FromConnType")
                arr(7) = it("ToShape")
                arr(8) = it("ToPin")
                arr(9) = it("ToConnType")
                rows.Add arr
            Next i
        Next idx
    End If
    
    Set CollectCableData = rows
End Function


