VERSION 5.00
Begin VB.Form OGLWin 
   BorderStyle     =   0  'None
   Caption         =   "OpenGL with VB: Solar System"
   ClientHeight    =   2265
   ClientLeft      =   1155
   ClientTop       =   1005
   ClientWidth     =   4200
   Icon            =   "OGLWin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.PictureBox glView2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   137
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "OGLWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'     This File contains code for Rendering, Animating and picking OpenGL Window      '
'                                       + VB Stuff                                    '
'         (You have)CopyRight © 2003 Saadat Ali Shah, shahji_2000@yahoo.com           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Done As Boolean      'Done=True => Closing App.
Dim Rev!(2 To 5), Rot!(2 To 4), Rt!, Rt2! 'Rev():Revolution angle around Sun (or Earth), Rot(): Rotation angle around your own Centre
Dim Mode As Integer         'Rendering Mode or Mouse Picking Mode
Public Selected As Byte     'Identifies object thats under mouse pointer
Dim Tmr As New VbTimer      'Our very own HiRes Timer
Dim Year(2 To 4) As Double  'Time Scale: 1 Year in secs.
Dim Days_per_year(2 To 4) As Double  'implies how may Rots a planet does in a year
Dim Description(Sun To Moon, 0 To 11) As String, Tip(0 To Moon) As String
Dim TexReady As Boolean     'Waoh! this provides a way around a little Bug,see glView2_paint()
Const Rdn_to_Degree = 3.1415926535 / 180 'Radians to Degree


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
 Dim i As Integer
 'Initialize...
 Mode = GL_RENDER: Selected = Earth
 Done = False: TexReady = False
 hDC2 = glView2.hDC 'Save Hdc of OpenGL Drawing Windows
 Year(Earth) = 20  '1 earth year equals ... secs
 Year(Mercury) = Year(Earth) * (88 / 365) 'Extract Mercury's year length from Earth's
 Year(Venus) = Year(Earth) * (225 / 365)
 Days_per_year(Earth) = 10  'Unfortunately putting real value '365' will make ur head spin (..with our year lenght)
 Days_per_year(Mercury) = 88 / 58: Days_per_year(Venus) = 225 / 243
  
 gHW = Me.hwnd      'Save handle to the form.
 Hook               'Begin subclassing.
 If InitGL = False Then Unload Me: Exit Sub
 
 Me.Show
 Tmr.Paused = False ' start Timer, can Also use: Tmr.Start
 glView2_Paint
 
 'Take Control...
 Do
  DoEvents 'Let VB do its stuff
  If Not Done Then
    'Update Animating Vars...
    Tmr.UpdateTimer 'Get Elapsed Time
    If Tmr.ElapsedSeconds > 0 Then  'If Timer not Paused
      For i = Mercury To Earth
        If Not (Selected = i Or i = Earth) Then ' Only inc Rev if the Object is not selected, Exclude earth for now
          Rev(i) = Rev(i) + 360! / (Year(i) / Tmr.ElapsedSeconds)
          If (Rev(i) >= 360!) Then Rev(i) = 0!
        End If
        Rot(i) = Rot(i) + 360! / ((Year(i) / Days_per_year(i)) / Tmr.ElapsedSeconds)
        If (Rot(i) >= 360!) Then Rot(i) = 0!
      Next i
      'Other Rotation Vars...
      Rt = Rt + 360! / (20 / Tmr.ElapsedSeconds): If (Rt >= 360!) Then Rt = 0!
      Rt2 = Rt2 + 360! / (40 / Tmr.ElapsedSeconds): If (Rt2 >= 360!) Then Rt2 = 0!
    End If
    
    If Selected > 0 Then glView2_Paint  'Don't Draw glView2 if nothing is selected
  End If
 Loop Until Done = True
End Sub

Private Sub Form_Paint()
    Dim V2w As Single, V2h As Single, h As Single, w As Single, i As Byte
    wglMakeCurrent hDC2, hglRC2
    'even though i made ScaleMode = pixels these guies r in twips??!
    w = glView2.ScaleX(glView2.Width, vbTwips, vbPixels): h = glView2.ScaleX(glView2.Height, vbTwips, vbPixels)
    glViewport 0, 0, w, h
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    gluPerspective 60!, w / h, 1!, 50! 'calculate aspect ratio of window
    glMatrixMode GL_MODELVIEW
    glLoadIdentity
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unhook  'Stop subclassing.
 Done = True
  
 If hglRC2 <> 0 Then
  wglMakeCurrent 0, 0
  wglDeleteContext hglRC2
 End If
End Sub




Private Sub glView2_Paint()
  wglMakeCurrent hDC2, hglRC2
  glClear clrColorBufferBit Or clrDepthBufferBit
  glLoadIdentity
  glTranslatef 0!, 0!, -10! 'setup camera
  'Draw Which ever Object was Selected
 If Selected = Earth Then
    glColor4f 1!, 1!, 1!, 1!
    glBindTexture GL_TEXTURE_2D, TArray2(Earth)
    glPushMatrix
    glRotatef Rot(Earth), 0, 1, 0
    glRotatef -90, 1!, 0!, 0!: gluSphere QObj, 3.5, 30, 30
    glBindTexture GL_TEXTURE_2D, TArray2(6)
    glPopMatrix
    glEnable GL_BLEND
    glColor4f 1!, 1!, 1!, 0.3!
    glRotatef -Rot(Earth), 0, 1, 0
    glRotatef -90, 1!, 0!, 0!: gluSphere QObj, 3.7, 30, 30
    glDisable GL_BLEND
  End If
  
  SwapBuffers hDC2
End Sub

