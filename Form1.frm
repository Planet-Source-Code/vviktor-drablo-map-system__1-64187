VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drablo map engine"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   335
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   334
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timWalk 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3360
      Top             =   3720
   End
   Begin VB.PictureBox picMChar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picTiles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picSight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FontTransparent =   0   'False
      Height          =   4800
      Left            =   120
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   316
      TabIndex        =   1
      Top             =   120
      Width           =   4800
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Wiktor Toporek
'Contact:
'mail: witek1@konto.pl
'   or wtoporek@gmail.com


Dim Sector() As Integer
Dim MapName As String
Dim PlayerX As Single, PlayerY As Single
Const TCols As Integer = 8
Const SecSize As Integer = 32
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim col(0 To 255) As Boolean
Dim ChangeLeg As Boolean
Dim PlayerSide As Integer
Dim LastKey As Integer


Private Sub Form_Load()
    picTiles.Picture = LoadPicture(App.Path & "\s02tiles.bmp")
    
    picChar.Picture = LoadPicture(App.Path & "\char.bmp")
    picMChar.Picture = LoadPicture(App.Path & "\charmask.gif") 'Maska postaci
    
    LoadCollisionSet "Collision.dat"
    
    LoadMap "Las1"

    timWalk_Timer
End Sub


Public Sub LoadCollisionSet(File As String)
    Dim FN As Integer
    Dim Linia As String
    Dim Z As Integer
    
    FN = FreeFile
    Open App.Path & "\" & File For Input As #FN
        Line Input #FN, Linia
        For Z = 1 To Len(Linia)
            col(Z - 1) = CBool(Mid(Linia, Z, 1))
        Next
    Close FN
    
    
End Sub



Public Sub RefreshCamera()

    On Error Resume Next
    Dim X As Integer, Y As Integer
    Dim SecW As Integer, SecH As Integer
    Dim CX As Integer, CY As Integer
    Dim PX As Integer, PY As Integer
    
    SecW = CInt(picSight.ScaleWidth / SecSize)
    SecH = CInt(picSight.ScaleHeight / SecSize)
    picSight.Cls
    

    For Y = Int(PlayerY / SecSize) - Int(SecH / 2) To Int(PlayerY / SecSize) + Int(SecH / 2)
        For X = Int(PlayerX / SecSize) - Int(SecW / 2) To Int(PlayerX / SecSize) + Int(SecW / 2)
            If X > -1 And Y > -1 And X <= UBound(Sector, 1) And Y <= UBound(Sector, 2) Then
                CX = 1
                CY = 0
                CY = Int(Sector(X, Y) / TCols)
                CX = Sector(X, Y) - (Fix(Sector(X, Y) / TCols) * TCols)
                If Not (CX = 1 And CY = 0) Then
                    PX = Int(SecW / 2) * SecSize - PlayerX + X * SecSize
                    PY = Int(SecH / 2) * SecSize - PlayerY + Y * SecSize
                    BitBlt picSight.hDC, PX, PY, SecSize, SecSize, picTiles.hDC, CX * SecSize, CY * SecSize, vbSrcCopy
                End If
            End If
        Next
    Next
    

    BitBlt picSight.hDC, picSight.ScaleWidth / 2 - 16, picSight.ScaleHeight / 2 - 16, 32, 32, picMChar.hDC, 32 * Abs(ChangeLeg), PlayerSide * 32, vbMergePaint
    BitBlt picSight.hDC, picSight.ScaleWidth / 2 - 16, picSight.ScaleHeight / 2 - 16, 32, 32, picChar.hDC, 32 * Abs(ChangeLeg), PlayerSide * 32, vbSrcAnd


End Sub

Public Sub LoadMap(File As String)
    Dim Linia As String
    Dim arg As Variant
    Dim X As Integer, Y As Integer
    Dim FN As Integer
    
    

    Erase Sector
    
    FN = FreeFile
    

    Open App.Path & "\" & File & ".map" For Input As FN
        

        Line Input #FN, Linia
        arg = Split(Linia, "||")
        MapName = CStr(arg(0))
        ReDim Sector(0 To CInt(arg(1)), 0 To CInt(arg(2)))
        
        'Pozycja startowa gracza:
        PlayerX = CSng(arg(3) * 32 + 16)
        PlayerY = CSng(arg(4) * 32 + 16)
        
        Do While Not EOF(FN)
            Line Input #FN, Linia
            If Linia <> "" Then
                For X = 0 To Len(Linia) - 1
                    Sector(X, Y) = CInt(255 - Asc(Mid(Linia, X + 1, 1))) 'Po sektorze do tablicy
                Next
                Y = Y + 1
            End If
        Loop
    Close FN

End Sub

Private Sub picSight_KeyDown(KeyCode As Integer, Shift As Integer)
    LastKey = KeyCode
    timWalk.Enabled = True
End Sub


Private Sub picSight_KeyUp(KeyCode As Integer, Shift As Integer)
    timWalk.Enabled = False
End Sub

Private Sub timWalk_Timer()
    On Error Resume Next
    Dim NewX As Single, NewY As Single
    
    Select Case LastKey
        Case vbKeyLeft
            NewX = PlayerX - 8
            
            If Not col(Sector(CInt((NewX - SecSize / 2) / SecSize), Int(PlayerY / SecSize))) Then
                PlayerX = NewX
                PlayerSide = 3
                ChangeLeg = True - ChangeLeg
            End If

        Case vbKeyUp
            NewY = PlayerY - 8
            
            If Not col(Sector(CInt((PlayerX - SecSize / 2) / SecSize), Int(NewY / SecSize))) Then
                PlayerY = NewY
                PlayerSide = 0
                ChangeLeg = True - ChangeLeg
            End If
        Case vbKeyRight
            NewX = PlayerX + 8
            
            If Not col(Sector(CInt((NewX - SecSize / 2) / SecSize), Int(PlayerY / SecSize))) Then
                PlayerX = NewX
                PlayerSide = 1
                ChangeLeg = True - ChangeLeg
            End If
        Case vbKeyDown
            NewY = PlayerY + 8
            
            If Not col(Sector(CInt((PlayerX - SecSize / 2) / SecSize), Int(NewY / SecSize))) Then
                PlayerY = NewY
                PlayerSide = 2
                ChangeLeg = True - ChangeLeg
            End If
    End Select

    RefreshCamera
    
    
    picSight.ForeColor = vbRed
    picSight.Print "Map name: " & MapName
    picSight.Print "Player position: (" & PlayerX & ", " & PlayerY & ")"
    picSight.Print "Player sector: (" & Int(PlayerX / 32) & ", " & Int(PlayerY / 32) & ")"
    
End Sub
