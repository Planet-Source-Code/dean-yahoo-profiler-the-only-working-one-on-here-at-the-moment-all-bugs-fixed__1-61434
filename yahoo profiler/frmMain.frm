VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWinSck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "YAHOO PROFILER THE ONLY ONE THAT WORKS"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   5295
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text28 
      Height          =   285
      Left            =   8760
      TabIndex        =   47
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox Text25 
      Height          =   285
      Left            =   1200
      TabIndex        =   46
      Top             =   7320
      Width           =   3855
   End
   Begin VB.TextBox Text27 
      Height          =   285
      Left            =   1200
      TabIndex        =   37
      Top             =   8040
      Width           =   3855
   End
   Begin VB.TextBox Text26 
      Height          =   285
      Left            =   1200
      TabIndex        =   36
      Top             =   7680
      Width           =   3855
   End
   Begin VB.TextBox Text24 
      Height          =   285
      Left            =   1200
      TabIndex        =   35
      Top             =   6960
      Width           =   3855
   End
   Begin VB.TextBox Text23 
      Height          =   285
      Left            =   1200
      TabIndex        =   34
      Top             =   6600
      Width           =   3855
   End
   Begin VB.TextBox Text22 
      Height          =   285
      Left            =   1200
      TabIndex        =   33
      Top             =   6240
      Width           =   3855
   End
   Begin VB.TextBox Text21 
      Height          =   285
      Left            =   1200
      TabIndex        =   32
      Top             =   5880
      Width           =   3855
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   1200
      TabIndex        =   31
      Top             =   5520
      Width           =   3855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FF00&
      Caption         =   "Grab Details"
      Height          =   2895
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtGetSource 
      Height          =   285
      Left            =   8760
      TabIndex        =   28
      Text            =   "http://profiles.yahoo.com/"
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   8760
      TabIndex        =   27
      Text            =   "_12px_1.gif"" "
      Top             =   7080
      Width           =   1695
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   8760
      TabIndex        =   26
      Text            =   "src=""http://us.i1.yimg.com/us.yimg.com/i/us/msg/6/gr/"
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   8760
      TabIndex        =   25
      Text            =   "</dd>"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   8760
      TabIndex        =   24
      Text            =   "Occupation:</dt> <dd>"
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   8760
      TabIndex        =   23
      Text            =   "</dd>"
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   8760
      TabIndex        =   22
      Text            =   "Sex:</dt> <dd>"
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   8760
      TabIndex        =   21
      Text            =   "</dd>"
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   8760
      TabIndex        =   20
      Text            =   "Age:</dt> <dd>"
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   8760
      TabIndex        =   19
      Text            =   "</dd>"
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   8760
      TabIndex        =   18
      Text            =   "Location:</dt> <dd>"
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   8760
      TabIndex        =   17
      Text            =   "</dd>"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   8760
      TabIndex        =   16
      Text            =   "Nickname:</dt> <dd>"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   8760
      TabIndex        =   15
      Text            =   "http://us.f2.yahoofs.com/users/"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   8760
      TabIndex        =   14
      Text            =   "</dd>"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   8760
      TabIndex        =   13
      Text            =   "Status:</dt> <dd>"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF0000&
      Caption         =   "Save Picture"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5040
      Width           =   5055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "d/l"
      Height          =   495
      Left            =   7320
      TabIndex        =   11
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   120
      ScaleHeight     =   4035
      ScaleWidth      =   4875
      TabIndex        =   10
      Top             =   840
      Width           =   4935
      Begin VB.Image Image1 
         Height          =   1620
         Left            =   -120
         Picture         =   "frmMain.frx":0000
         Top             =   1200
         Visible         =   0   'False
         Width           =   5100
      End
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   8760
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   7800
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   8760
      TabIndex        =   8
      Text            =   "</dd>"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   8760
      TabIndex        =   7
      Text            =   "Name:</dt> <dd>"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Yahoo Profiler"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1080
         TabIndex        =   29
         Text            =   "-dean-"
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton cmdGetSource 
         BackColor       =   &H0000FF00&
         Caption         =   "GO"
         Height          =   285
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Yahoo Id :"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   9960
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Window"
      Height          =   3015
      Left            =   11400
      TabIndex        =   3
      Top             =   120
      Width           =   3375
      Begin VB.TextBox txtData 
         Height          =   2655
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Label Label10 
      Caption         =   "On Line:"
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   8040
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Occupation:"
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Sex:"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Marital Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Age:"
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Location:"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Nickname:"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Real Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Idle..."
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   8400
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   8400
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
    End Type


Private Declare Function OleLoadPicturePath Lib "oleaut32.dll" _
    (ByVal szURLorPath As Long, ByVal punkCaller As Long, _
    ByVal dwReserved As Long, ByVal clrReserved As OLE_COLOR, _
    ByRef riid As TGUID, ByRef ppvRet As IPicture) As Long
    Dim filename As String
    '-------------------------------------------------------------------------------------
    'this winsock source code belongs to kyro and many others which ive learnt from over the years,
    'im sure this could be made cleaner but i cannot be arsed ,
    'just a yahoo profiler which parses the source code
    'yahoo have changed there web pages now so all old profilers wont work


'-----------------------------------------
'Format URL as create a connection request
'-----------------------------------------
Private Sub cmdGetSource_Click()
Image1.Visible = False
Text20 = ""
Text21 = ""
Text22 = ""
Text23 = ""
Text24 = ""
Text25 = ""
Text26 = ""
Text27 = ""
Text28 = ""
txtData.Text = ""

    Dim CurrentServer As String
    
    CurrentServer = GetServer(txtGetSource & Text4)   'get server from URL
    
    'if the URL is blank or there is not server then do not continue
    If CurrentServer = "" Then Exit Sub
    
    'Setup the winsock and connect
    With Winsock1
    
        .Close 'if a connection if being established
               'or is already established close it
        
        'if there is a proxy in the "txtProxy" text box then use it
        
            
        
        
            .RemoteHost = CurrentServer 'The server (Ex: kyro-genics.com)
            .RemotePort = 80 '80 is standard port for ALL HTML requests
            
        
        
        .Connect 'connect with the server and port
  
    End With
    

End Sub



Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
filename = "http://us." & Text28
Picture1.Picture = LoadPicture(filename)


End Sub

Private Sub Command4_Click()
SavePicture Picture1.Picture, App.Path & "\" + Text4.Text + ".jpg"
End Sub

Private Sub Command5_Click()
txtData.Text = Convert2SingleLine(txtData.Text)
Pause (0.06)
ParseIt1 txtData.Text, Text1, Text2, Text20
ParseIt1 txtData.Text, Text8, Text9, Text21
ParseIt1 txtData.Text, Text10, Text11, Text22
ParseIt1 txtData.Text, Text12, Text13, Text23
ParseIt1 txtData.Text, Text5, Text6, Text24
ParseIt1 txtData.Text, Text14, Text15, Text25
ParseIt1 txtData.Text, Text16, Text17, Text26
ParseIt1 txtData.Text, Text18, Text19, Text27
ParseIt1 txtData.Text, "> <a href=http://us.", "><img", Text28
If Text20 = "&nbsp;" Then
Text20 = "NO Details"
End If
If Text21 = "&nbsp;" Then
Text21 = "NO Details"
End If
If Text22 = "&nbsp;" Then
Text22 = "NO Details"
End If
If Text23 = "&nbsp;" Then
Text23 = "NO Details"
End If
If Text24 = "&nbsp;" Then
Text24 = "NO Details"
End If
If Text25 = "&nbsp;" Then
Text25 = "NO Details"
End If
If Text26 = "&nbsp;" Then
Text26 = "NO Details"
End If
If Text27 = "&nbsp;" Then
Text27 = "NO Details"
End If
filename = "http://us." & Text28
Picture1.Picture = LoadPicture(filename)

If Text28 = "" Then
Image1.Visible = True
Else
Image1.Visible = False
End If
End Sub

'--------------------------------------------
'Kill active Winsock connections upon exiting
'--------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    
    Cancel = True
    
    Winsock1.Close
    
    End
    
End Sub





'---------------------------
'Send a request for the HTML
'---------------------------
Private Sub Winsock1_Connect()

    'status message (not important)
    lblStatus = "Connecting to " & Winsock1.RemoteHost & " on Port 80)"
    
    txtData = Empty 'clear the Data Window
    
    Winsock1.SendData SendHeader(txtGetSource & Text4) 'Send request
    
End Sub

'----------------------------------------------
'Get and Add the Data (HTML) to the Data Window
'----------------------------------------------
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    'status message (not important)
    lblStatus = "Receiving data from " & Winsock1.RemoteHost
    
    On Error Resume Next 'on error go to next line of code
    
    Dim ReturnedHTML As String
    
    Winsock1.GetData ReturnedHTML 'Data (HTML) returned by the server
    
    txtData = txtData & ReturnedHTML 'Add the data to the Data Window
   
    txtData = FixFeed(txtData) 'fix the invalid line feed characters in the HTML
Call Command5_Click
End Sub

'-----------------------------------------------
'If an Error occurs display error in data window
'-----------------------------------------------
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    'status message (not important)
    lblStatus = "Idle..."

    Winsock1.Close 'Kill the connection
    
    txtData = "Error : " & Description & vbCrLf 'description of the error
    txtData = txtData & "Error ID : " & Scode 'Error id
    
End Sub

Public Function LoadPicture(ByVal filename As String) As Picture
    Dim myTGUID As TGUID
    myTGUID.Data1 = &H7BF80980
    myTGUID.Data2 = &HBF32
    myTGUID.Data3 = &H101A
    myTGUID.Data4(0) = &H8B
    myTGUID.Data4(1) = &HBB
    myTGUID.Data4(2) = &H0
    myTGUID.Data4(3) = &HAA
    myTGUID.Data4(4) = &H0
    myTGUID.Data4(5) = &H30
    myTGUID.Data4(6) = &HC
    myTGUID.Data4(7) = &HAB
    On Error GoTo LblError
    OleLoadPicturePath StrPtr(filename), 0, 0, 0, myTGUID, LoadPicture
    Exit Function
LblError:
    Set LoadPicture = VB.LoadPicture(filename)
End Function
Sub Pause(interval)
Dim Current
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub
Function ParseIt1(Expression As String, DelimiterA As String, DelimiterB As String, LstBx As TextBox)
On Error Resume Next
Dim A As Long, B As Long
A = 1
While InStr(A, Expression, DelimiterA) > 0
  A = InStr(A, Expression, DelimiterA) + Len(DelimiterA)
  B = InStr(A, Expression, DelimiterB)
  LstBx.Text = Mid$(Expression, A, B - A)
Wend
End Function
