VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "音乐解析"
   ClientHeight    =   4470
   ClientLeft      =   5115
   ClientTop       =   6315
   ClientWidth     =   7215
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   7215
   Begin VB.CheckBox qx 
      BackColor       =   &H00FFFFFF&
      Caption         =   "全选"
      Height          =   180
      Left            =   0
      TabIndex        =   11
      Top             =   540
      Width           =   975
   End
   Begin VB.ListBox List4 
      Height          =   2040
      Left            =   480
      TabIndex        =   10
      Top             =   7440
      Width           =   7215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List3 
      Height          =   2040
      Left            =   600
      TabIndex        =   7
      Top             =   4560
      Width           =   7215
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   4215
      Left            =   600
      TabIndex        =   5
      Top             =   7080
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7435
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0352
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      Locked          =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   4
      Top             =   360
      Width           =   7215
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Text            =   "2848433"
      Top             =   0
      Width           =   7215
   End
   Begin VB.ListBox List2 
      Height          =   1140
      Left            =   0
      MultiSelect     =   1  'Simple
      TabIndex        =   2
      Top             =   720
      Width           =   7215
   End
   Begin VB.ListBox List1 
      Height          =   2040
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   7215
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4215
      Left            =   11640
      TabIndex        =   0
      Top             =   1920
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7435
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":03EF
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2520
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label dl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  下载  "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   5880
      TabIndex        =   9
      Top             =   4080
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label chis 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "选择文件夹..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   3600
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   2100
   End
   Begin WMPLibCtl.WindowsMediaPlayer W 
      Height          =   135
      Left            =   7680
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   2566
      _cy             =   238
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sai() As String
Dim 狈 As String
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Const BIF_RETURNONLYFSDIRS = 1                                        '从这里开始为API声明
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "Shell32" _
      (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "Shell32" _
      (ByVal pidList As Long, _
      ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
      (ByVal lpString1 As String, ByVal _
      lpString2 As String) As Long

Private Type BrowseInfo
     hWndOwner As Long
     pIDLRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags    As Long
     lpfnCallback     As Long
     lParam     As Long
     iImage     As Long
End Type

Dim a As New Shell
Dim pat As String



Private Sub chis_Click()
dl.Enabled = False
Dim b As Folder
Set b = a.BrowseForFolder(0, "选择音乐存储路径", 0)
a.Open b
On Error GoTo bbzlp
pat = b.Items.Item.Path
Debug.Print b.Items.Item.Path
dl.Enabled = True
bbzlp:
'dl.Enabled = False
Exit Sub
End Sub

Private Sub dl_Click()
Dim r
Dim i As Integer
For i = 0 To List3.ListCount - 1
'< > / \ | : " * ?
Dim gfwjm As String
gfwjm = Right(List4.List(i) + ".mp3", Len(List4.List(i) + ".mp3") - InStr(List4.List(i) + ".mp3", "]"))
gfwjm = Replace(gfwjm, "<", " - ")
gfwjm = Replace(gfwjm, ">", " - ")
gfwjm = Replace(gfwjm, "/", " - ")
gfwjm = Replace(gfwjm, "\", " - ")
gfwjm = Replace(gfwjm, ":", " - ")
gfwjm = Replace(gfwjm, Chr(34), " - ")
gfwjm = Replace(gfwjm, "|", " - ")
gfwjm = Replace(gfwjm, "*", " - ")
gfwjm = Replace(gfwjm, "?", " - ")
r = URLDownloadToFile(0, List3.List(i), pat + "\" + gfwjm, 0, 0)
Next i
dl.Enabled = False
End Sub

Private Sub Form_Load()
Me.CommonDialog1.InitDir = ""
W.settings.volume = 2
Dim i As Long
Dim b() As Byte
b = Inet1.OpenURL("https://music.163.com/#/album?id=2848433", icByteArray)

RichTextBox1.Text = Utf8ToUnicode(b) 'myHTML
RichTextBox2.Text = Utf8ToUnicode(b) 'myHTML
狈 = RichTextBox2.Text
Sai = Split(RichTextBox1.Text, "<li><a href=" + Chr(34) + "/song?")
For i = 1 To UBound(Sai)
List1.AddItem "http://music.163.com/song/media/outer/url?" + Left(Sai(i), InStr(Sai(i), Chr(34)) - 1)
List2.AddItem "[" + Replace(Str(i), " ", "") + "]" + Right(Left(Sai(i), InStr(Sai(i), "</a>") - 1), Len(Left(Sai(i), InStr(Sai(i), "</a>") - 1)) - 1 - InStr(Left(Sai(i), InStr(Sai(i), "</a>") - 1), Chr(34) + ">"))
Next i
Text2.Text = Mid(狈, InStr(狈, "<title>") + 7, InStr(狈, "</title>") - InStr(狈, "<title>") - 7) '</title>
End Sub

Private Sub List1_Click()
Debug.Print List1.List(List1.ListIndex)
End Sub

Private Sub List2_Click()
Dim i As Integer
List3.Clear
List4.Clear
W.URL = List1.List(List2.ListIndex)
W.Controls.play
List1.Selected(List2.ListIndex) = True
For i = 0 To List2.ListCount - 1
If List2.Selected(i) = True Then
List3.AddItem List1.List(i)
List4.AddItem List2.List(i)
End If
Next i
If List3.ListCount <> List2.ListCount Then
qx.Value = 0
Else
qx.Value = 1
End If

If List3.ListCount <> 0 Then
chis.Visible = True
dl.Visible = True
dl.Enabled = False
Else
dl.Enabled = False
chis.Visible = False
dl.Visible = False
End If
End Sub

Private Sub qx_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
If qx.Value = 0 Then
For i = 0 To List2.ListCount - 1
List2.Selected(i) = False
Next i
List2.Selected(0) = False
ElseIf qx.Value = 1 Then
For i = 0 To List2.ListCount - 1
List2.Selected(i) = True
Next i
List2.Selected(0) = True
End If
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) = 0 Then
RichTextBox1.Text = ""
RichTextBox2.Text = ""
List1.Clear
List2.Clear
Text2.Text = ""
W.Controls.stop
End If
End Sub

Private Sub Text1_GotFocus()
If Text1.Text = "在此输入网易云音乐对应歌单或专辑链接末尾的ID号" Then
Text1.Text = ""
Text1.ForeColor = 0
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Len(Text1.Text) <> 0 Then
RichTextBox1.Text = ""
RichTextBox2.Text = ""
List1.Clear
List2.Clear
Dim i As Long
Dim c As String
Dim b() As Byte
c = Str((Text1.Text))
c = Replace(c, " ", "")
b = Inet1.OpenURL("https://music.163.com/#/album?id=" + c, icByteArray)
RichTextBox1.Text = Utf8ToUnicode(b) 'myHTML
RichTextBox2.Text = Utf8ToUnicode(b) 'myHTML
狈 = RichTextBox2.Text
Sai = Split(RichTextBox1.Text, "<li><a href=" + Chr(34) + "/song?")
For i = 1 To UBound(Sai)
List1.AddItem "http://music.163.com/song/media/outer/url?" + Left(Sai(i), InStr(Sai(i), Chr(34)) - 1)
List2.AddItem "[" + Replace(Str(i), " ", "") + "]" + Right(Left(Sai(i), InStr(Sai(i), "</a>") - 1), Len(Left(Sai(i), InStr(Sai(i), "</a>") - 1)) - 1 - InStr(Left(Sai(i), InStr(Sai(i), "</a>") - 1), Chr(34) + ">"))
Next i
Text2.Text = Mid(狈, InStr(狈, "<title>") + 7, InStr(狈, "</title>") - InStr(狈, "<title>") - 7) '</title>

RichTextBox1.Text = ""
RichTextBox2.Text = ""

c = Str((Text1.Text))
c = Replace(c, " ", "")
b = Inet1.OpenURL("https://music.163.com/#/playlist?id=" + c, icByteArray)
RichTextBox1.Text = Utf8ToUnicode(b) 'myHTML
RichTextBox2.Text = Utf8ToUnicode(b) 'myHTML
狈 = RichTextBox2.Text
Sai = Split(RichTextBox1.Text, "<li><a href=" + Chr(34) + "/song?")
For i = 1 To UBound(Sai)
List1.AddItem "http://music.163.com/song/media/outer/url?" + Left(Sai(i), InStr(Sai(i), Chr(34)) - 1)
List2.AddItem "[" + Replace(Str(i), " ", "") + "]" + Right(Left(Sai(i), InStr(Sai(i), "</a>") - 1), Len(Left(Sai(i), InStr(Sai(i), "</a>") - 1)) - 1 - InStr(Left(Sai(i), InStr(Sai(i), "</a>") - 1), Chr(34) + ">"))
Next i

RichTextBox1.Text = ""
RichTextBox2.Text = ""

c = Str((Text1.Text))
c = Replace(c, " ", "")
b = Inet1.OpenURL("https://music.163.com/#/artist?id=" + c, icByteArray)
RichTextBox1.Text = Utf8ToUnicode(b) 'myHTML
RichTextBox2.Text = Utf8ToUnicode(b) 'myHTML
狈 = RichTextBox2.Text
Sai = Split(RichTextBox1.Text, "<li><a href=" + Chr(34) + "/song?")
For i = 1 To UBound(Sai)
List1.AddItem "http://music.163.com/song/media/outer/url?" + Left(Sai(i), InStr(Sai(i), Chr(34)) - 1)
List2.AddItem "[" + Replace(Str(i), " ", "") + "]" + Right(Left(Sai(i), InStr(Sai(i), "</a>") - 1), Len(Left(Sai(i), InStr(Sai(i), "</a>") - 1)) - 1 - InStr(Left(Sai(i), InStr(Sai(i), "</a>") - 1), Chr(34) + ">"))
Next i

If Text2.Text = "网易云音乐" Then
Text2.Text = Mid(狈, InStr(狈, "<title>") + 7, InStr(狈, "</title>") - InStr(狈, "<title>") - 7) '</title>
End If
If List1.ListCount = 0 Then
RichTextBox1.Text = ""
RichTextBox2.Text = ""
Text1.Text = "2848433"
List1.Clear
List2.Clear
b = Inet1.OpenURL("https://music.163.com/#/album?id=2848433", icByteArray)

RichTextBox1.Text = Utf8ToUnicode(b) 'myHTML
RichTextBox2.Text = Utf8ToUnicode(b) 'myHTML
狈 = RichTextBox2.Text
Sai = Split(RichTextBox1.Text, "<li><a href=" + Chr(34) + "/song?")
For i = 1 To UBound(Sai)
List1.AddItem "http://music.163.com/song/media/outer/url?" + Left(Sai(i), InStr(Sai(i), Chr(34)) - 1)
List2.AddItem "[" + Replace(Str(i), " ", "") + "]" + Right(Left(Sai(i), InStr(Sai(i), "</a>") - 1), Len(Left(Sai(i), InStr(Sai(i), "</a>") - 1)) - 1 - InStr(Left(Sai(i), InStr(Sai(i), "</a>") - 1), Chr(34) + ">"))
Next i
Text2.Text = Mid(狈, InStr(狈, "<title>") + 7, InStr(狈, "</title>") - InStr(狈, "<title>") - 7) '</title>
End If
End If
If List2.ListCount = 0 Then
qx.Enabled = False
Else
qx.Enabled = True
End If
End If
End Sub

Private Sub Text1_LostFocus()
If Text1.Text = "" Then
Text1.Text = "在此输入网易云音乐对应歌单或专辑链接末尾的ID号"
Text1.ForeColor = &H808080
End If
End Sub
