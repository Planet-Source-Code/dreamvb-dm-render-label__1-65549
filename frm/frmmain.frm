VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM Render Label"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.RenderLabel RenderLabel1 
      Height          =   6525
      Left            =   120
      TabIndex        =   5
      Top             =   135
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   11509
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   9210
      TabIndex        =   3
      Top             =   7305
      Width           =   9270
      Begin VB.Label lblUrl 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   75
         TabIndex        =   4
         Top             =   60
         Width           =   45
      End
   End
   Begin VB.TextBox txtFilename 
      Height          =   300
      Left            =   855
      TabIndex        =   2
      Top             =   6825
      Width           =   6960
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00808080&
      Caption         =   "&GO"
      Height          =   300
      Left            =   7875
      TabIndex        =   0
      Top             =   6810
      Width           =   600
   End
   Begin VB.Label lblPage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filename:"
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   6885
      Width           =   675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub TimesTables()
Dim xRnd As Integer, x As Integer, sLine As String
    xRnd = Int(12 * Rnd) + 1
    
    sLine = "<P>" + vbCrLf
    sLine = sLine & "<U>A Random Times Table Lister</U>" + vbCrLf
    sLine = sLine & "<br>" + vbCrLf
    sLine = sLine & "<br>" + vbCrLf
    
    
    For x = 1 To 12
        sLine = sLine & "<b>" & x & "&nbsp;</b> x " & xRnd & " =&nbsp;" & "<b>" & (x * xRnd) & "</b><br>" + vbCrLf
    Next
    
    sLine = sLine & "<br>" + vbCrLf
    sLine = sLine & "</a><a href=""onClick:(%27TABLES%27)"">Show new Random table</a>" + vbCrLf
    sLine = sLine & "<br>" + vbCrLf
    sLine = sLine & "<br>" + vbCrLf
    sLine = sLine & "<a href=""about.htm"">Back</a>"
    
    sLine = sLine & "</p>"
    RenderLabel1.DocRenderHtml sLine
    sLine = ""
    
End Sub

Function FixPath(lPath As String) As String
    If Right(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

Private Sub cmdGo_Click()
    RenderLabel1.HtmlDocNav txtFilename.Text
End Sub

Private Sub Form_Load()
    txtFilename.Text = FixPath(App.Path) & "index.htm"
End Sub

Private Sub RenderLabel1_HyperLinkClick(Key As String, URL As String, Text As String)
    Select Case Key
        Case "TEST"
            MsgBox "You clicked the TEST link"
        Case "EXIT"
            If MsgBox("Do you want to exit now.", vbYesNo Or vbQuestion) = vbNo Then Exit Sub
            Unload Form1
        Case "ABOUT"
            MsgBox "DM Render Label." & vbCrLf & vbTab & "By DreamVB" & vbCrLf & "Please Vote..", vbInformation, "About"
        Case "TABLES"
            Call TimesTables
    End Select
    
End Sub
