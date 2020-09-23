VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Win Dll Error Descripcion"
   ClientHeight    =   2445
   ClientLeft      =   2490
   ClientTop       =   2520
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6150
   Begin VB.TextBox Text2 
      Height          =   1740
      Left            =   135
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   5925
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      Default         =   -1  'True
      Height          =   405
      Left            =   1440
      TabIndex        =   1
      Top             =   135
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const NERR_BASE = 2100
Private Const MAX_NERR = NERR_BASE + 899

Private Const LOAD_LIBRARY_AS_DATAFILE = &H2

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Any) As Long
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long




Public Function DisplayError(ByVal lCode As Long) As String

       Dim sMsg As String
       Dim sRtrnCode As String
       Dim lFlags As Long
       Dim hModule As Long
       Dim lRet As Long
       hModule = 0
        sRtrnCode = Space$(256)
        lFlags = FORMAT_MESSAGE_FROM_SYSTEM
         ' if lRet is in the network range, load the message source
         If (lCode >= NERR_BASE And lCode <= MAX_NERR) Then
            hModule = LoadLibraryEx("netmsg.dll", 0&, LOAD_LIBRARY_AS_DATAFILE)
            If (hModule <> 0) Then
                lFlags = lFlags Or FORMAT_MESSAGE_FROM_HMODULE
            End If
         End If
        ' Call FormatMessage() to allow for message text to be acquired
        ' from the system or the supplied module handle.
        '
        lRet = FormatMessage(lFlags, hModule, lCode, 0&, sRtrnCode, 256&, 0&)
        If lRet = 0 Then
           sRtrnCode = "The system cannot find message text for message number " & lCode
        End If
        ' if you loaded a message source, unload it.        '
        If (hModule <> 0) Then
            FreeLibrary (hModule)
        End If
     '//... now display this string
     sMsg = sRtrnCode
     DisplayError = sMsg
End Function






Private Sub Command1_Click()
If Trim(Text1) <> "" Then
    Text2 = DisplayError(Text1)
End If
End Sub


