VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEncodageDonnees 
   Caption         =   "Encodage des données du document"
   ClientHeight    =   8745.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21540
   OleObjectBlob   =   "frmEncodageDonnees.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEncodageDonnees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnnuler_Click()
    ' Ferme le formulaire sans enregistrer les données
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next

    Dim docSheet As Visio.Shape
    Set docSheet = ThisDocument.DocumentSheet

    ' Écrit les valeurs des TextBox dans les cellules utilisateur du document
    docSheet.CellsU("User.NomProjet").FormulaForce = Chr(34) & Me.TextNomProjet.Value & Chr(34)
    docSheet.CellsU("User.NumeroProjet").FormulaForce = Chr(34) & Me.TextNumeroProjet.Value & Chr(34)
    docSheet.CellsU("User.NumeroDessin").FormulaForce = Chr(34) & Me.TextNumeroDessin.Value & Chr(34)
    docSheet.CellsU("User.Client").FormulaForce = Chr(34) & Me.TextClient.Value & Chr(34)
    docSheet.CellsU("User.Departement").FormulaForce = Chr(34) & Me.TextDepartement.Value & Chr(34)
    docSheet.CellsU("User.CreationDate").FormulaForce = Chr(34) & Me.TextDateCreation.Value & Chr(34)
    docSheet.CellsU("User.Dessinateur").FormulaForce = Chr(34) & Me.TextDessinateur.Value & Chr(34)
    docSheet.CellsU("User.Verificateur").FormulaForce = Chr(34) & Me.TextVerificateur.Value & Chr(34)
    docSheet.CellsU("User.Rev1Nom").FormulaForce = Chr(34) & Me.Text1Nom.Value & Chr(34)
    docSheet.CellsU("User.Rev2Nom").FormulaForce = Chr(34) & Me.Text2Nom.Value & Chr(34)
    docSheet.CellsU("User.Rev3Nom").FormulaForce = Chr(34) & Me.Text3Nom.Value & Chr(34)
    docSheet.CellsU("User.Rev1Mod").FormulaForce = Chr(34) & Me.Text1Mod.Value & Chr(34)
    docSheet.CellsU("User.Rev2Mod").FormulaForce = Chr(34) & Me.Text2Mod.Value & Chr(34)
    docSheet.CellsU("User.Rev3Mod").FormulaForce = Chr(34) & Me.Text3Mod.Value & Chr(34)
    docSheet.CellsU("User.Rev1Date").FormulaForce = Chr(34) & Me.Text1Date.Value & Chr(34)
    docSheet.CellsU("User.Rev2Date").FormulaForce = Chr(34) & Me.Text2Date.Value & Chr(34)
    docSheet.CellsU("User.Rev3Date").FormulaForce = Chr(34) & Me.Text3Date.Value & Chr(34)

    ' Ferme le formulaire
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next

    Dim docSheet As Visio.Shape
    Set docSheet = ThisDocument.DocumentSheet

    ' Charge les valeurs existantes dans les TextBox pour que l'utilisateur puisse les modifier
    Me.TextNomProjet.Value = docSheet.CellsU("User.NomProjet").ResultStr(visUCMDSafe)
    Me.TextNumeroProjet.Value = docSheet.CellsU("User.NumeroProjet").ResultStr(visUCMDSafe)
    Me.TextNumeroDessin.Value = docSheet.CellsU("User.NumeroDessin").ResultStr(visUCMDSafe)
    Me.TextClient.Value = docSheet.CellsU("User.Client").ResultStr(visUCMDSafe)
    Me.TextDepartement.Value = docSheet.CellsU("User.Departement").ResultStr(visUCMDSafe)
    Me.TextDateCreation.Value = docSheet.CellsU("User.CreationDate").ResultStr(visUCMDSafe)
    Me.TextDessinateur.Value = docSheet.CellsU("User.Dessinateur").ResultStr(visUCMDSafe)
    Me.TextVerificateur.Value = docSheet.CellsU("User.Verificateur").ResultStr(visUCMDSafe)
    Me.Text1Nom.Value = docSheet.CellsU("User.Rev1Nom").ResultStr(visUCMDSafe)
    Me.Text2Nom.Value = docSheet.CellsU("User.Rev2Nom").ResultStr(visUCMDSafe)
    Me.Text3Nom.Value = docSheet.CellsU("User.Rev3Nom").ResultStr(visUCMDSafe)
    Me.Text1Mod.Value = docSheet.CellsU("User.Rev1Mod").ResultStr(visUCMDSafe)
    Me.Text2Mod.Value = docSheet.CellsU("User.Rev2Mod").ResultStr(visUCMDSafe)
    Me.Text3Mod.Value = docSheet.CellsU("User.Rev3Mod").ResultStr(visUCMDSafe)
    Me.Text1Date.Value = docSheet.CellsU("User.Rev1Date").ResultStr(visUCMDSafe)
    Me.Text2Date.Value = docSheet.CellsU("User.Rev2Date").ResultStr(visUCMDSafe)
    Me.Text3Date.Value = docSheet.CellsU("User.Rev3Date").ResultStr(visUCMDSafe)
    
End Sub
