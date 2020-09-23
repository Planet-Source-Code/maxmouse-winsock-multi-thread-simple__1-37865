VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox Cons 
      Height          =   2790
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.ListBox Threads 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   2400
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Q 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MasterIndex As Integer

Private Sub Form_Load()
Load Socket(1)

With Socket(1)
.LocalPort = 10
.Listen
End With

Form1.Caption = Socket(0).LocalIP
End Sub

Private Sub Socket_Close(Index As Integer)
Threads.AddItem Index
Unload Socket(Index)
End Sub

Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
With Socket(Index)
.Close
.Accept requestID
End With

If Index > MasterIndex Then MasterIndex = Index

Socket(Index).SendData "Connected! Your IP is: " & Socket(Index).RemoteHostIP & " Connected on YOUR Localport: " & Socket(Index).RemotePort
Cons.AddItem "IP: " & Socket(Index).RemoteHostIP & " Index: " & Index

If Threads.ListCount = 0 Then
Load Socket(MasterIndex + 1)

With Socket(MasterIndex + 1)
.LocalPort = 10
.Listen
End With

Q.Caption = "Socket " & MasterIndex + 1 & " is Next"

Else

Load Socket(Threads.List(0))

With Socket(Threads.List(0))
.LocalPort = 10
.Listen
End With

Q.Caption = "Socket " & Threads.List(0) & " is Next"
Threads.RemoveItem (0)

End If
End Sub
