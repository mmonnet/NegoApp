'Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Drawing.Printing
Imports Microsoft.Office.Interop.Excel
Imports System.Configuration

Public Class Form1
    Inherits System.Windows.Forms.Form

    Private Sub OuvrirToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim path, str As String
        Dim iDeb, i, J As Integer

        On Error GoTo errorFind

        path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)
        ' recherche l'occurence partir de la droite
        iDeb = InStrRev(path, "")
        ' recherche la longueur de a chaine path
        J = Len(path)
        ' recherche l'occurence
        ' retourne la partie gauche du chemin d'accès au programme
        path = Strings.Left(path, iDeb - 1)
        i = InStrRev(path, "\")
        '' recherche la longueur de a chaine path
        'J = Len(path)
        'path = Strings.Left(path, i - 1)
        '' recherche l'occurence du premier caractère \ à partir de la droite
        'i = InStrRev(path, "\")
        '' recherche la longueur de a chaine path
        'J = Len(path)
        'str = Strings.Left(path, i - 1)
        ' Lancement du fichier d'aide
        nomFichier = path & "\Reference\Fichier Vide.xlsx"
        nomFichierFinal = path & "\Tests\Fichier Vide.xlsx"

        ' ouvrir l'application excel ; Ok
        xlApp = CreateObject("Excel.Application")

        xlApp.Visible = True  ' Rendre visible l'application

        ' Ouvrir le fichier Excel et affectation; Erreur 0x080010105
        wkbObj = xlApp.Workbooks.Open(nomFichier)

        wkbObj.SaveAs(nomFichierFinal)
        wkbObj.Close(SaveChanges:=True, Filename:=nomFichierFinal)
        'wkbObj.Close(nomFichier)
        nomFichierFinal = ""
        wkbObj = Nothing
        xlApp.Visible = False
        xlApp.Quit()
        xlApp = Nothing
        MsgBox("Création de fichier réussie")
        Exit Sub

errorFind:

        MsgBox(Err.Description & " Saving Excel file failed")
        '        set XlApp.Quit()
        ' Reseting.
        wkbObj.Close(nomFichierFinal)
        'wkbObj.Close(nomFichier)
        wkbObj = Nothing
        xlApp.Visible = False
        xlApp = Nothing

    End Sub

    Private Sub DonnéesDeBaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuCarac.Click
        FrmCarac.Show()
    End Sub

    Private Sub ActionDeNégoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuAction.Click

    End Sub

    Private Sub FichierToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuFichier.Click

    End Sub

    Private Sub MenuStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub LireFichier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LireFichierNego.Click
        With LireFichier
            .Title = "Sélection du fichier Excel"
            ' Le fait d'annuler génère une erreur ce qui permet de quitter la procédure:
            .ReadOnlyChecked = True       ' Masque la boîte 'Ouvrir en lecture seule'
            .Filter = "Fichiers Excel (*.xlsx)|*.xlsx"      ' Règle le filtre de fichiers
            .ShowDialog()                           ' Ouvre une boîte de type Enregistrement
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            nomFichier = .FileName
        End With

        lireDansFichierExcel(nomFichier)

    End Sub

    Private Sub EnregistrerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EnregistrerFichierNego.Click
        EcrireDansFichierExcel(nomFichier)

    End Sub
    Private Sub QuitterToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Quitter.Click
        End
    End Sub

    Private Sub AProposToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuAPropos.Click
        FrmAPropos.Show()

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ChxStrat = 0    'définition de la stratégie de base
    End Sub

    Private Sub ObjectifsQuantitatifs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub StatiqueToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StatiqueToolStripMenuItem.Click
        FrmNego.Show()
    End Sub

    Private Sub DynamiqueToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DynamiqueToolStripMenuItem.Click
        FrmNegoDyn.Show()
    End Sub
End Class
