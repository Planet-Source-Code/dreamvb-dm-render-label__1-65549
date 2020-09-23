VERSION 5.00
Begin VB.UserControl RenderLabel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   HasDC           =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7380
   Begin VB.PictureBox WebDC 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   0
      MouseIcon       =   "RenderLabel.ctx":0000
      ScaleHeight     =   281
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   472
      TabIndex        =   0
      Top             =   0
      Width           =   7080
   End
End
Attribute VB_Name = "RenderLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'GUI Api calls
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, _
ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
'
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

'Sound Flag Consts
Private Const SND_NODEFAULT = &H2
Private Const SND_RESOURCE = &H40004
Private Const SND_SYNC = &H0

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Main HTML Data
Private sHtmlData As String
'Hold all the HTML Elements Tags, HTML
Private Type HtmlElements
    IsTag As Boolean
    TAG As String
    StrHTML As String
End Type

'Used to hold a collection of HTML Elements
Private HtmlElement() As HtmlElements
'HTML Element Counter
Private ElementCount As Integer

'HTML Document Stuff
Private Type HtmlDocument
    TextColor As OLE_COLOR
    TextFont As String
    BgColor As OLE_COLOR
    LeftMargin As Integer
    TopMargin As Integer
    LinkColor As OLE_COLOR
End Type

'Store HTML Document Info
Private HtmlDoc As HtmlDocument
'Current Text been Rendered
Private PrintHTML As String
'Lets us know when we need to skip a line
Private Skip As Boolean
'Lets us know when we need a line break
Private LineBreak As Boolean
'Current Index of current processing HTML Element
Private CurTagPos As Long
'Current TextHeight size
Private m_TextHeight As Single
'
Private BulletOn As Boolean, bOrder As Boolean
Private BulletX As Integer, bOrderIdx As Integer
'
'Used to hold information about our Hyperlinks
Private Type HtmlLinks
    y As Integer
    x As Integer
    Text As String
    URL As String
End Type
'Hyperlink count
Private LinkCount As Integer
'Collection of hyperlinks info
Private HyperLinks(100) As HtmlLinks
'Current Index of found hyperlink in collection
Private HyperLinkIdx As Integer
Dim ReDirect As Boolean, TempStr As String

Event HyperLinkClick(Key As String, URL As String, Text As String)

Private Sub DocHighlight(sColor As OLE_COLOR)
Dim hBrush As Long, rc As RECT, dc As Long, iWidth As Long
    
    With WebDC
        iWidth = .TextWidth(HtmlElement(CurTagPos + 1).StrHTML)
        dc = MakeDc(iWidth, CLng(m_TextHeight))
        hBrush = CreateSolidBrush(sColor)
        SetRect rc, 0, 0, iWidth, m_TextHeight
        FillRect dc, rc, hBrush
        DeleteObject hBrush
        BitBlt .hdc, .CurrentX + 1, .CurrentY, iWidth, m_TextHeight, dc, 0, 0, vbSrcCopy
    End With
    
End Sub

Public Sub PlayMouseClick()
On Error Resume Next
    Const sFlags = SND_RESOURCE Or SND_SYNC Or SND_NODEFAULT
        
    If (waveOutGetNumDevs >= 1) Then
        'If sound card found play wav
        PlaySound "CLICK", ByVal 0&, sFlags
        Exit Sub
    End If

End Sub

Sub DocRedirect()
Dim e_pos As Integer, TimeVal As Integer, sUrl As String
    'Allows you to redirect to a page
    If Len(TempStr) <> 0 Then e_pos = InStr(1, TempStr, ";", vbBinaryCompare)
    
    If (e_pos > 0) Then
        TimeVal = Val(Left(TempStr, e_pos - 1))
        TempStr = Trim(Right(TempStr, Len(TempStr) - e_pos))
        If UCase(Left(TempStr, 4) = "URL=") Then
            sUrl = Trim(Right(TempStr, Len(TempStr) - 4))
            If IsFileHere(sUrl) = False Then
                Exit Sub
            Else
                Call Sleep(TimeVal * 500)
                Call HtmlDocNav(sUrl)
                sUrl = ""
                TempStr = ""
            End If
        End If
    End If
    
End Sub

Private Function MakeDc(w As Long, h As Long) As Long
Dim myDc As Long, hBmp As Long
    'Function to Crate a New Dc
    myDc = CreateCompatibleDC(GetDC(0))
    hBmp = CreateCompatibleBitmap(GetDC(0), w, h)
    DeleteObject SelectObject(myDc, hBmp)
    MakeDc = myDc
End Function

Private Sub OpenURL(sUrl As String, hwnd As Long)
    ShellExecute hwnd, "open", sUrl, vbNullString, vbNullString, 1
End Sub

Private Function URL_Decode(TUrl As String) As String
' This is used for decodeing HTTP URLS
Dim Xpos As Integer
Dim CGI_Str As String

    While (InStr(TUrl, "%") <> 0)
        Xpos = InStr(TUrl, "%")
        CGI_Str = Mid(TUrl, Xpos + 1, 2)
        TUrl = Replace(TUrl, "%" & CGI_Str, Chr("&H" & CGI_Str))
    Wend
    
    URL_Decode = Replace(TUrl, "+", " ")
End Function

Private Sub SetBackGround(WebDC As PictureBox, sPicFile As String)
Dim iPic As IPictureDisp
Dim hBrush As Long, rc As RECT
    'This fucntion is used to set the background image of the HTML Render control
    If Not IsFileHere(sPicFile) Then Exit Sub
    
    Set iPic = LoadPicture(sPicFile)
    
    With WebDC
        hBrush = CreatePatternBrush(iPic)
        SetRect rc, 0, 0, .ScaleWidth, .ScaleHeight
        FillRect .hdc, rc, hBrush
        DeleteObject hBrush
        Set iPic = Nothing
    End With
    
End Sub

Private Function CountIF(lpStr As String, nChar As String) As Long
Dim x As Integer, idx As Long
    'Return the number of nChar within the string lpStr
    'ex CountIF("hello","l") returns 2
    Do While (x < Len(lpStr))
        x = x + 1
        If Mid(lpStr, x, 1) = nChar Then idx = idx + 1
    Loop
    CountIF = idx
End Function

Private Function Hex2Lng(sHex As String) As Long
Dim r, g, b As Integer
    'Converts a WebHex color to a Long
    If Left(sHex, 1) = "#" Then sHex = Mid(sHex, 2, Len(sHex))
    If (6 - Len(sHex)) > 0 Then sHex = sHex & String(6 - Len(sHex), "0")
    r = Val("&H" & Mid(sHex, 1, 2))
    g = Val("&H" & Mid(sHex, 3, 2))
    b = Val("&H" & Mid(sHex, 5, 2))
    
    Hex2Lng = RGB(r, g, b)
    r = 0: g = 0: b = 0
End Function

Private Function IsFileHere(lzFileName As String) As Boolean
    If Dir(lzFileName) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Public Sub HtmlDocNav(lzFile As String)
    Call PlayMouseClick
    Call HtmlDocReset
    sHtmlData = OpenFile(lzFile)
    Call FormatSpecialChars
    Call PhaseHTML
    Call RenderHTML
    
    'Redirection found
    If ReDirect Then
        ReDirect = False
        DocRedirect
    End If
    
End Sub

Public Sub DocRenderHtml(sHtml As String)
    Call HtmlDocReset
    sHtmlData = sHtml
    Call FormatSpecialChars
    Call PhaseHTML
    Call RenderHTML
    
End Sub

Private Sub DoURL(sHyperLink As String, Index As Integer)
Dim s_pos As Integer, e_pos As Integer, sType As String
Dim Temp As String

    s_pos = InStr(1, sHyperLink, ":", vbBinaryCompare)
    
    If (s_pos > 0) Then sType = Trim(UCase(Left(sHyperLink, s_pos - 1)))
    
    Select Case sType
        Case "HTTP", "MAILTO"
            Call OpenURL(sHyperLink, UserControl.hwnd)
        Case "ONCLICK"
            Temp = URL_Decode(sHyperLink)
            Temp = Trim(Mid(Temp, s_pos + 1, Len(Temp)))
            s_pos = InStr(Temp, "('")
            e_pos = InStr(s_pos, Temp, "')")
            
            If (s_pos > 0) And (e_pos > 0) Then
                Temp = Trim(Mid(Temp, s_pos + 2, e_pos - s_pos - 2))
                RaiseEvent HyperLinkClick(Temp, HyperLinks(Index).Text, HyperLinks(Index).URL)
            End If
            
        Case Else
            Call HtmlDocNav(sHyperLink)
    End Select
    
End Sub

Private Sub HtmlDocReset()
Dim x As Integer

    'Reset all variables
    For x = 0 To UBound(HyperLinks)
        With HyperLinks(x)
            .Text = ""
            .x = -1
            .y = -1
        End With
    Next x
    
    m_TextHeight = WebDC.TextHeight("Qq")
    bOrder = False
    BulletOn = False
    bOrderIdx = 0
    PrintHTML = ""
    sHtmlData = ""
    HyperLinkIdx = 0
    CurTagPos = 0
    ElementCount = 0
    LinkCount = 0
    Erase HtmlElement
End Sub

Private Sub AddHyperLink(sHyper As HtmlLinks, Index As Integer)
On Error Resume Next
    'Add a new hyperlink to the hyperlink collection
    With HyperLinks(Index)
        .Text = sHyper.Text
        .URL = sHyper.URL
        .x = sHyper.x
        .y = sHyper.y
    End With
End Sub

Private Sub AddElement(sHtmlDoc As HtmlElements)
    'Resize HTML Elements collection
    ReDim Preserve HtmlElement(ElementCount) As HtmlElements
    'Add a new HTML Element
    With HtmlElement(ElementCount)
        .IsTag = sHtmlDoc.IsTag
        .StrHTML = sHtmlDoc.StrHTML
        .TAG = sHtmlDoc.TAG
    End With
    'Update HTML Element counter
    ElementCount = ElementCount + 1
    
End Sub

Private Sub AddImage(lzFile As String)
Dim bmp As BITMAP
Dim dc As Long
Dim iPic As IPictureDisp
    'This function we use to convert a VB Picture handle to a DC
    If IsFileHere(lzFile) = False Then Exit Sub
    
    Set iPic = LoadPicture(lzFile)

    GetObjectAPI iPic.Handle, Len(bmp), bmp
    
    dc = MakeDc(bmp.bmWidth, bmp.bmHeight)
    SelectObject dc, iPic.Handle
    
    With WebDC
        .CurrentY = .CurrentY + ImgH
        TransparentBlt .hdc, .CurrentX, .CurrentY, bmp.bmWidth, bmp.bmHeight, _
        dc, 0, 0, bmp.bmWidth, bmp.bmHeight, RGB(255, 0, 255)
        .CurrentY = .CurrentY + bmp.bmHeight
        
        lzFile = ""
        DeleteObject dc
        Set iPic = Nothing
    End With
End Sub

Private Sub FormatText(lpStr As String, sWebDc As PictureBox)
Dim e_pos As Integer, d_pos As Integer, n_Pos As Integer
Dim sValue As String, sName As String
Dim HyperLinkA As HtmlLinks
On Error Resume Next

    e_pos = 1
    x = -1

    Do While (x < CountIF(lpStr, "="))
        x = x + 1
        d_pos = InStr(e_pos, lpStr, "=")    'Get next equals sign
        If d_pos = 0 Then Exit Do           'No Sign so we exit

        sName = Mid$(lpStr, e_pos, d_pos - e_pos)  'Extract tag Name
        n_Pos = InStr(d_pos, lpStr, " ")

        If n_Pos = 0 Then n_Pos = Len(lpStr) + 1
        'Extract Tag
        sValue = Trim$(Mid$(lpStr, d_pos + 1, n_Pos - d_pos - 1)) 'Extract tag value
        'Strip away quotes
        sValue = Replace(sValue, Chr(34), "")
        
        'Process HTML Tags
        Select Case UCase(sName)
            Case "HIGHLIGHT"
                Call DocHighlight(Hex2Lng(sValue))
            Case "HTTP-EQUIV"
                If UCase(sValue) = "REFRESH" Then
                    ReDirect = True
                End If
            Case "CONTENT"
                If (ReDirect) Then TempStr = sValue
            Case "COLOR" 'Font Color
                sWebDc.ForeColor = Hex2Lng(sValue)
            Case "TEXT"
                HtmlDoc.TextColor = Hex2Lng(sValue)
            Case "BGCOLOR" ' HTML Background Color
                sWebDc.BackColor = Hex2Lng(sValue)
            Case "LEFTMARGIN"
                HtmlDoc.LeftMargin = Val(sValue)
            Case "BACKGROUND" 'Html Background image
                Call SetBackGround(WebDC, sValue)
            Case "LINK" 'Hyperlink Color
                HtmlDoc.LinkColor = Hex2Lng(sValue)
            Case "HREF" 'Hyperlinks
                HyperLinkA.Text = HtmlElement(CurTagPos + 1).StrHTML
                HyperLinkA.URL = sValue
                HyperLinkA.x = WebDC.CurrentX
                HyperLinkA.y = WebDC.CurrentY
                Call AddHyperLink(HyperLinkA, LinkCount)
                LinkCount = LinkCount + 1
            Case "SRC" 'Images
                Call AddImage(sValue)
        End Select
        
        e_pos = n_Pos + 1
Loop

    e_pos = 0: d_pos = 0: n_Pos = 0
    sName = "": sValue = ""
    
End Sub

Private Sub FormatTags(sTag As String)
Dim sTagStart As Integer, sNextTag As String, Temp As String
    
    'Set Skip and LineBreak to false
    Skip = False
    LineBreak = False
    
    'Apply the render to the picturebox
    With WebDC
        Select Case UCase(sTag)
            Case "<UL>"
                BulletOn = True
                bOrder = False
                .CurrentX = .CurrentX + 10
                BulletX = .CurrentX
            Case "<OL>"
                bOrderIdx = 0
                bOrder = True
                BulletOn = True
                .CurrentX = .CurrentX + 10
                BulletX = .CurrentX
            Case "</UL>", "</OL>"
                .CurrentX = HtmlDoc.LeftMargin
                BulletX = 0
                BulletOn = False
                bOrder = False
                .CurrentY = .CurrentY + m_TextHeight
            Case "<LI>"
                .CurrentY = .CurrentY + m_TextHeight
                .CurrentX = BulletX
            Case "<CENTER>"
                'Center Text
                .CurrentX = (.ScaleWidth - .TextWidth(HtmlElement(CurTagPos + 1).StrHTML)) \ 2 - HtmlDoc.LeftMargin
            Case "<RIGHT>"
                .CurrentX = (.ScaleWidth - .TextWidth(HtmlElement(CurTagPos + 1).StrHTML)) - HtmlDoc.LeftMargin
            Case "<LEFT>"
                .CurrentX = HtmlDoc.LeftMargin
            Case "</A>"
                Exit Sub
            Case "<HR>"
                'Horizontal Line
                LineBreak = True
                WebDC.Line (HtmlDoc.LeftMargin, .CurrentY + 2)-(.ScaleWidth - HtmlDoc.LeftMargin, .CurrentY + 2), &H808080
                WebDC.Line (HtmlDoc.LeftMargin, .CurrentY + 1)-(.ScaleWidth - HtmlDoc.LeftMargin, .CurrentY + 1), vbWhite
            Case "<HTML>", "</HTML>", "<BODY>", "</BODY>", _
                "<HEAD>", "</HEAD>", "<TITLE>", "</TITLE>"
                Skip = True
            Case "<B>", "<STRONG>"
                .FontBold = True
            Case "</B>", "</I>", "</U>", "</STRONG>", "</S>"
                .FontBold = False
                .FontItalic = False
                .FontUnderline = False
                .FontStrikethru = False
            Case "<I>"
                .FontItalic = True
            Case "<U>"
                .FontUnderline = True
            Case "<S>"
                .FontStrikethru = True
            
            Case "<BR>", "</BR>"
                LineBreak = True
            Case "<P>", "</P>"
                LineBreak = True
            Case "</FONT>"
                .FontName = HtmlDoc.TextFont
            Case Else
                sTagStart = InStr(1, sTag, " ")
                If (sTagStart > 0) Then
                    Temp = Trim(Mid(sTag, sTagStart + 1, Len(sTag) - sTagStart - 1))
                    
                    
                    sNextTag = UCase(Trim(Mid(sTag, 1, sTagStart - 1)))
                    Select Case sNextTag
                        Case "<META"
                            Call FormatText(Temp, WebDC)
                        Case "<FONT"
                            Call FormatText(Temp, WebDC)
                        Case "<A"
                            Call FormatText(Temp, WebDC)
                            .ForeColor = HtmlDoc.LinkColor
                            .FontUnderline = True
                        Case "<BODY"
                            Call FormatText(Temp, WebDC)
                        Case "<IMG"
                            Call FormatText(Temp, WebDC)
                    End Select
                    
                    'sNextTag = UCase(Left(sTag, sTagStart))
                    'MsgBox sNextTag
                End If
        End Select
        
        m_TextHeight = .TextHeight("Qq")
    End With
    
End Sub

Private Sub RenderHTML()
    'This is the main part that does all teh rendering
    With WebDC
        .Cls
        'Apply Margins
        .CurrentY = HtmlDoc.LeftMargin
        .CurrentX = HtmlDoc.TopMargin
        .BackColor = HtmlDoc.BgColor
        
        .ForeColor = HtmlDoc.TextColor
        .FontName = HtmlDoc.TextFont
        
        Do While (CurTagPos < ElementCount)
            
            If HtmlElement(CurTagPos).IsTag Then
                Call FormatTags(HtmlElement(CurTagPos).TAG)
            Else
                PrintHTML = Trim(HtmlElement(CurTagPos).StrHTML)

                If (Skip) Then PrintHTML = ""
                
                If (LineBreak) Then
                    .CurrentX = HtmlDoc.LeftMargin
                    .CurrentY = .CurrentY + m_TextHeight
                End If
                
                If Len(PrintHTML) <> 0 Then
                    PrintHTML = Replace(PrintHTML, "&nbsp;", " ")
                    PrintHTML = Replace(PrintHTML, "&lt;", "<")
                    PrintHTML = Replace(PrintHTML, "&gt;", ">")
                    
                    If (Not bOrder And BulletOn) Then
                        WebDC.Print Chr(149) & " " & PrintHTML;
                    ElseIf (bOrder And BulletOn) Then
                        bOrderIdx = bOrderIdx + 1
                        WebDC.Print bOrderIdx & " " & PrintHTML;
                    Else
                        WebDC.Print PrintHTML;
                    End If
                End If
                
                'Defaults
                .ForeColor = HtmlDoc.TextColor
                .FontName = HtmlDoc.TextFont
                .FontUnderline = False
                .FontBold = False
                .FontItalic = False
                .FontStrikethru = False
            End If
            CurTagPos = CurTagPos + 1
        Loop
       .Refresh
    End With
    
End Sub

Private Sub PhaseHTML()
On Error Resume Next
Dim Temp As String
Dim iLen As Long, x As Long
Dim e_Start As Integer, e_End As Integer
Dim ch As String * 1
Dim Temp_Elment As HtmlElements

    'This is the main part that scans a HTML string and stores it tokens.
    
    e_Start = 1
    e_End = 1
    
    iLen = Len(sHtmlData)
    Temp = ""
    
    For e_End = 0 To iLen
        With Temp_Elment
        
            ch = Mid(sHtmlData, e_End, 1)
        
            'Locate Opening tag
            If ch = "<" Then e_Start = e_End
            'Locate Closeing tag
            If e_Start > 0 Then e_End = InStr(e_Start, sHtmlData, ">")
        
            'Get the HTML tags ie <B></B> etc
            If e_End > 0 Then
                Temp = Trim(Mid(sHtmlData, e_Start, e_End - e_Start + 1))
                If Len(Temp) > 0 Then
                    .IsTag = True
                    .StrHTML = ""
                    .TAG = Temp
                    AddElement Temp_Elment
                    e_End = e_End + 1
                End If
                e_Start = e_End
            End If
        
            If (e_Start > 0) Then
                'Find opening tag
                e_End = InStr(e_Start, sHtmlData, "<")
            End If
        
            'Get the Text between the Tag ie <B>Hello</B>
            If e_End > 0 And e_End - e_Start > 0 Then
                Temp = Mid(sHtmlData, e_Start, e_End - e_Start)
                If Len(Temp) > 0 Then
                    .StrHTML = Temp
                    .IsTag = False
                    AddElement Temp_Elment
                End If
                e_Start = e_End
            Else
                'Opps no closeing tag so we just run to the length of the html
                If e_End = 0 Then e_End = iLen
            End If
            Temp = ""
        End With
        DoEvents
    Next
    
End Sub

Private Sub FormatSpecialChars()
Dim x As Integer
    'Format special chars
    sHtmlData = Replace(sHtmlData, vbCrLf, " ")
    sHtmlData = Replace(sHtmlData, vbTab, " ")
    sHtmlData = Replace(sHtmlData, "&copy;", "©")
    sHtmlData = Replace(sHtmlData, "&reg;", "®")
    sHtmlData = Replace(sHtmlData, "&pound;", "£")
    
    sHtmlData = Replace(sHtmlData, "&yen;", "¥")
    sHtmlData = Replace(sHtmlData, "&euro;", "€")

    sHtmlData = Replace(sHtmlData, "&amp;", "&")
    sHtmlData = Replace(sHtmlData, "&para;", "¶")
    
    sHtmlData = Replace(sHtmlData, "&laquo;", "«")
    sHtmlData = Replace(sHtmlData, "&raquo;", "»")
    sHtmlData = Replace(sHtmlData, "&plusmn;", "±")
    
    sHtmlData = Replace(sHtmlData, "&iexcl;", "¡")
    sHtmlData = Replace(sHtmlData, "&deg;", "°")
    
    'This replaces all string like &#65; to it's char value A
    For x = 0 To 255
        sHtmlData = Replace(sHtmlData, "&#" & CStr(x) & ";", Chr(x))
    Next
    x = 0
End Sub

Private Function OpenFile(lFile As String) As String
Dim fp As Long, sTemp As String, vLst As Variant, sLine As String
Dim vData() As Byte, x As Long

    'Opens a given filename and returns the contents
    fp = FreeFile
    Open lFile For Binary As #fp
        If LOF(fp) = 0 Then Exit Function
        ReDim vData(0 To LOF(fp))
        Get #fp, , vData()
    Close #fp
    
    sTemp = StrConv(vData, vbUnicode)
    Erase vData
    vLst = Split(sTemp, vbCrLf)
    sTemp = ""
    
    For x = 0 To UBound(vLst)
        sLine = Trim(vLst(x))
        If Len(sLine) <> 0 Then
            sTemp = sTemp & sLine & vbCrLf
            sLine = ""
        End If
    Next x
    
    OpenFile = sTemp
    Erase vLst
End Function

Private Function IsHyperLink(YPos As Integer, Xpos As Integer) As Integer
Dim x  As Integer, idx As Integer
Dim isLink As Boolean

    isLink = False
    On Error Resume Next
    
    For x = 0 To LinkCount - 1
        If (HyperLinks(x).y \ m_TextHeight = YPos - 1) Then
            idx = x
            isLink = True
            
            If (Xpos - HyperLinks(x).x) > WebDC.TextWidth(HyperLinks(x).Text) Then
                isLink = False
                x = LinkCount
            End If
            
            If (Xpos < HyperLinks(x).x) Then
                isLink = False
                x = LinkCount
            End If
            
            x = LinkCount
        End If
    Next
    
    If (Not isLink) Then idx = -1
    IsHyperLink = idx
    
End Function

Private Sub UserControl_Initialize()
    HtmlDoc.TextColor = vbBlack
    HtmlDoc.TextFont = WebDC.FontName
    HtmlDoc.LinkColor = vbBlue
    HtmlDoc.BgColor = UserControl.BackColor
    '
    HtmlDoc.LeftMargin = 8
    HtmlDoc.TopMargin = 8
End Sub

Private Sub UserControl_Show()
    HtmlDoc.TextFont = UserControl.Font.Name
End Sub

Private Sub WebDC_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo ErrFlag:
    
    HyperLinkIdx = IsHyperLink(y \ m_TextHeight, CInt(x))
    
    If (HyperLinkIdx <> -1) Then
        WebDC.MousePointer = vbCustom
    Else
        WebDC.MousePointer = vbDefault
    End If
    
    Exit Sub
ErrFlag:
    WebDC.MousePointer = vbDefault
End Sub

Private Sub WebDC_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (HyperLinkIdx <> -1) And (Button = vbLeftButton) Then
        Call DoURL(HyperLinks(HyperLinkIdx).URL, HyperLinkIdx)
    End If
End Sub

Private Sub UserControl_Resize()
    WebDC.Width = UserControl.ScaleWidth
    WebDC.Height = UserControl.ScaleHeight
End Sub

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
End Sub



