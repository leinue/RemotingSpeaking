VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "爪机控制计算机"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6960
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   6960
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "DelCookies"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   3855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   840
      Width           =   6615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   2775
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   4080
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Sever IP:"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   290
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UserIP As String
Dim i As Integer
Private Sub Command1_Click()
Winsock1(0).Close
End Sub
Private Sub Form_Load()
i = 0
Text1.Text = Winsock1(0).LocalIP + ":1234"
Winsock1(0).LocalPort = 1234
Winsock1(0).Listen
End Sub
Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
PrintDe ("A customer enters")
If Winsock1(0).State <> sckClosed Then
Winsock1(0).Close
End If
Winsock1(0).Accept requestID
End Sub
Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim PostData As String
Dim a, TextLine As String
Dim posSpeaking As Integer
Winsock1(0).GetData PostData
PrintDe (PostData)
If InStr(PostData, "codeSendmessage") <> 0 Then
MsgBox ("已接受到手机端指令-SendMessage")
End If
If InStr(PostData, "codeRunCMD") <> 0 Then
Shell "cmd.exe"
End If
posSpeaking = InStr(PostData, "SpeakingText")
If posSpeaking <> 0 Then
Dim con() As String, con0 As String, con0len As Integer, conf As String
Dim ii As Integer, ItContainInput As Integer
con = Split(PostData, Chr(13) + Chr(10))
For ii = 0 To UBound(con)
ItContainInput = InStr(con(ii), "input=")
If ItContainInput <> 0 Then
con0 = Mid(con(ii), Len("input=") + 1)
Exit For
End If
Next
con0len = Len(con0) - Len("&SpeakingText=Iput")
conf = Replace(con0, "&SpeakingText=Iput", "")
Dim resRepalced As String, resSplited() As String, j As Integer, ResFinal As String
ResFinal = ""
resRepalced = Replace(URLdecode(conf), "&#", "")
resSplited = Split(resRepalced, ";")
For j = 0 To UBound(resSplited) - 1
ResFinal = ResFinal + CStr(ChrW$(resSplited(j)))
Next
CreateObject("SAPI.SpVoice").Speak (ResFinal)
End If
Open App.Path & "\index.html" For Input As #1
Do While Not EOF(1)
   Line Input #1, TextLine
   a = a & TextLine
   Winsock1(0).SendData TextLine
Loop
Close #1
End Sub
Private Sub PrintDe(ByVal data As String)
Text2.Text = Text2.Text + data + Chr(13) + Chr(10)
End Sub
Private Sub Winsock1_SendComplete(Index As Integer)
'Load Winsock1(Index)
i = i - 1
Winsock1(0).Close
Winsock1(0).Listen
End Sub
Function str2asc(strstr)
        str2asc = Hex(Asc(strstr))
End Function
Function asc2str(ascasc)
        asc2str = Chr(ascasc)
End Function
Public Function URLdecode(ByRef Text As String) As String
    Const Hex = "0123456789ABCDEF"
    Dim lngA As Long, lngB As Long, lngChar As Long, lngChar2 As Long
    URLdecode = Text
    lngB = 1
    For lngA = 1 To LenB(Text) - 1 Step 2
        lngChar = Asc(MidB$(URLdecode, lngA, 2))
        Select Case lngChar
            Case 37
                lngChar = InStr(Hex, MidB$(Text, lngA + 2, 2)) - 1
                If lngChar >= 0 Then
                    lngChar2 = InStr(Hex, MidB$(Text, lngA + 4, 2)) - 1
                    If lngChar2 >= 0 Then
                        MidB$(URLdecode, lngB, 2) = Chr$((lngChar * &H10&) Or lngChar2)
                        lngA = lngA + 4
                    Else
                        If lngB < lngA Then MidB$(URLdecode, lngB, 2) = MidB$(Text, lngA, 2)
                    End If
                Else
                    If lngB < lngA Then MidB$(URLdecode, lngB, 2) = MidB$(Text, lngA, 2)
                End If
            Case 43
                MidB$(URLdecode, lngB, 2) = " "
            Case Else
                If lngB < lngA Then MidB$(URLdecode, lngB, 2) = MidB$(Text, lngA, 2)
        End Select
        lngB = lngB + 2
    Next lngA
    URLdecode = LeftB$(URLdecode, lngB - 1)
End Function
