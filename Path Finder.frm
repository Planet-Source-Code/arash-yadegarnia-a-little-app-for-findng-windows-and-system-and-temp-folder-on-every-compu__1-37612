VERSION 5.00
Begin VB.Form FRMmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PATH FINDER"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5835
   Icon            =   "Path Finder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.Label LBLwin 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   5
      Left            =   75
      TabIndex        =   5
      Top             =   1950
      Width           =   5640
   End
   Begin VB.Label LBLwin 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   4
      Left            =   75
      TabIndex        =   4
      Top             =   1125
      Width           =   5640
   End
   Begin VB.Label LBLwin 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   3
      Left            =   75
      TabIndex        =   3
      Top             =   375
      Width           =   5640
   End
   Begin VB.Label LBLwin 
      Caption         =   "System Directory :"
      Height          =   240
      Index           =   2
      Left            =   75
      TabIndex        =   2
      Top             =   825
      Width           =   1815
   End
   Begin VB.Label LBLwin 
      Caption         =   "Temp Directory :"
      Height          =   240
      Index           =   1
      Left            =   75
      TabIndex        =   1
      Top             =   1650
      Width           =   1815
   End
   Begin VB.Label LBLwin 
      Caption         =   "Windows Directory :"
      Height          =   240
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   1815
   End
End
Attribute VB_Name = "FRMmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------------------------------
'PATH  FINDER  VER 1.0.0   BY   ARASH YADEGARNIA
'THIS IS THE EASIEST AND FASTEST CODE TO GET WINDOWS,TEMP,SYSTEM DIRECTORY ON EVERY COMPUTER..!!!
'.I'm Iranian...so I apolagize for any possible wrong words or grammer useing.
'------------------------------------------------------------------------------------------------------------------------------
'************************Declareing three API Function for getting pathes on every computer************************************
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Sub Getpath_WINDOWS()
Dim WindirS As String * 255         'declares a full lenght string for DIR name(for getting the path)
                                    
Dim TEMP                            'a temporarry variable that holds LENGHT OF THE FINAL PATH STRING
Dim Result                          'a variable for holding the the output of the function
TEMP = GetWindowsDirectory(WindirS, 255)     'holds the FUUL(include unneccessary charecters)Path
Result = Left(WindirS, TEMP)                 'holds final path
LBLwin(3).Caption = Result
End Sub
Private Sub Getpath_SYSTEM()
Dim WindirS As String * 255         'declares a full lenght string for DIR name(for getting the path)
                                        
Dim TEMP                            'a temporarry variable that holds LENGHT OF THE FINAL PATH STRING!
Dim Result                          'a variable for holding the the output of the function
TEMP = GetSystemDirectory(WindirS, 255)      'holds the FUUL(include unneccessary charecters)Path
Result = Left(WindirS, TEMP)                 'holds final path
LBLwin(4).Caption = Result
End Sub
Private Sub Getpath_TEMP()
'this API(TEMP) is different from others(in placing arguments)

Dim WindirS As String * 255         'declares a full lenght string for DIR name(for getting the path)
                                        
Dim TEMP                            'a temporarry variable that holds LENGHT OF THE FINAL PATH STRING!
Dim Result                          'a variable for holding the the output of the function
TEMP = GetTempPath(255, WindirS)            'holds the FUUL(include unneccessary charecters)Path
Result = Left(WindirS, TEMP)                'holds final path
LBLwin(5).Caption = Result
End Sub
Private Sub Form_Load()
Getpath_WINDOWS                     'Calls the getpath_WINDOWS sub
Getpath_SYSTEM                      'Calls the getpath_SYSTEM sub
Getpath_TEMP                        'Calls the getpath_TEMP sub
End Sub
