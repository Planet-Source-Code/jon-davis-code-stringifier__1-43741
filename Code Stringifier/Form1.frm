VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Code Stringifier"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer WBTimeoutTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8280
      Top             =   480
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   615
      Left            =   5640
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   735
      ExtentX         =   1296
      ExtentY         =   1085
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8280
      Top             =   0
   End
   Begin VB.CheckBox chkAddBreaks 
      Caption         =   "Add real line breaks"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.ComboBox cboLanguage 
      Height          =   315
      ItemData        =   "Form1.frx":0E42
      Left            =   1800
      List            =   "Form1.frx":0E5B
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   2895
   End
   Begin RichTextLib.RichTextBox txtCodeString 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5318
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0EA6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtContent 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5318
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0F26
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Stringify to language:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblStringifiedOutput 
      Caption         =   "JavaScript String"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Content"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private wbinit As Boolean

Private Sub cboLanguage_Change()
    lblStringifiedOutput.Caption = cboLanguage.Text & " String"
    Timer1_Timer
End Sub

Private Sub cboLanguage_Click()
    cboLanguage_Change
End Sub

Private Sub chkAddBreaks_Click()
    Timer1_Timer
End Sub

Private Sub Form_Load()
    cboLanguage.ListIndex = 0
    Dim lang As String
    lang = GetSetting(App.Title, "Settings", "LastLanguage")
    On Error Resume Next
    cboLanguage.Text = lang
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtContent.Width = Form1.ScaleWidth - (txtContent.Left * 2)
    txtCodeString.Width = Form1.ScaleWidth - (txtCodeString.Left * 2)
    Form1.Height = 8130 ' .. too lazy to add height resizing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Settings", "LastLanguage", cboLanguage.Text
End Sub

Private Sub txtContent_Change()
    Timer1.Enabled = True
End Sub

Private Sub txtContent_KeyDown(KeyCode As Integer, Shift As Integer)
    Timer1.Enabled = True
End Sub

Private Sub txtContent_KeyPress(KeyAscii As Integer)
    Timer1.Enabled = True
End Sub

Private Sub txtContent_KeyUp(KeyCode As Integer, Shift As Integer)
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    txtCodeString.Text = JSFix(txtContent.Text)
    Timer1.Enabled = False
End Sub

Private Function JSFix(str As String)
    Select Case cboLanguage.Text
    Case "C#"
        str = Replace(str, "\", "\\")
        str = Replace(str, vbCr, "\r")
        str = Replace(str, """", "\""")
        If chkAddBreaks.Value = 0 Then
            str = Replace(str, vbLf, "\n")
        Else
            str = Replace(str, vbLf, "\n""" & vbCrLf & vbTab & "+ """)
        End If
        str = "string str = """ & str & """;" & vbCrLf
    Case "JavaScript/JScript"
        str = Replace(str, vbCr, "")
        str = Replace(str, "\", "\\")
        str = Replace(str, """", "\""")
        If chkAddBreaks.Value = 0 Then
            str = Replace(str, vbLf, "\n")
        Else
            str = Replace(str, vbLf, "\n""" & vbCrLf & vbTab & "+ """)
        End If
        str = "var str = """ & str & """;" & vbCrLf
    Case "VBScript"
        str = Replace(str, """", """""")
        str = Replace(str, vbCr, "")
        If chkAddBreaks.Value = 0 Then
            str = Replace(str, vbLf, """ & vbCrLf & """)
        Else
            str = Replace(str, vbLf, """ & vbCrLf & _" & vbCrLf & vbTab & """")
        End If
        str = "Dim str" & vbCrLf & "str = """ & str & """" & vbCrLf
    Case "VB6"
        str = Replace(str, """", """""")
        str = Replace(str, vbCr, "")
        If chkAddBreaks.Value = 0 Then
            str = Replace(str, vbLf, """ & vbCrLf & """)
        Else
            str = Replace(str, vbLf, """ & vbCrLf & _" & vbCrLf & vbTab & """")
        End If
        str = "Dim str As String" & vbCrLf & "str = """ & str & """" & vbCrLf
    Case "VB.Net"
        str = Replace(str, """", """""")
        str = Replace(str, vbCr, "")
        If chkAddBreaks.Value = 0 Then
            str = Replace(str, vbLf, """ & vbCrLf & """)
            str = "Dim str As String = " & _
              """" & str & """" & vbCrLf
        Else
            str = Replace(str, vbLf, """ & vbCrLf _" & vbCrLf & vbTab & "& """)
            str = "Dim str As String" & vbCrLf & "str = """ & str & """" & vbCrLf
        End If
    Case "HTML (mini)"
        str = Replace(str, "&", "&amp;")
        str = Replace(str, """", "&quot;")
        str = Replace(str, vbCr, "")
        str = Replace(str, "<", "&lt;")
        str = Replace(str, ">", "&gt;")
        If chkAddBreaks.Value = 0 Then
            str = Replace(str, vbLf, "<br>")
        Else
            str = Replace(str, vbLf, "<br>" & vbCrLf)
        End If
    Case "HTML (IE)"
        str = TEXT2IEHTML(str)
    End Select
    JSFix = str
End Function

Function TEXT2IEHTML(str As String)
    'On Error Resume Next
    If Not wbinit Then
        WebBrowser1.Navigate "about:<html><head></head><body>.</body></html>"
        WBTimeoutTimer.Enabled = True
        Do Until WebBrowser1.Busy Or Not WBTimeoutTimer.Enabled
            DoEvents
        Loop
        Do Until Not WebBrowser1.Busy
            DoEvents
        Loop
        wbinit = True
    End If
    
    WebBrowser1.Document.body.innerText = txtContent.Text
    str = WebBrowser1.Document.body.innerhtml
    If chkAddBreaks.Value > 0 Then
        str = Replace(str, "<BR>", "<BR>" & vbCrLf)
        str = Replace(str, "<P>", vbCrLf & "<P>")
    End If
    TEXT2IEHTML = str
End Function

Private Sub WBTimeoutTimer_Timer()
    WBTimeoutTimer.Enabled = False
End Sub
