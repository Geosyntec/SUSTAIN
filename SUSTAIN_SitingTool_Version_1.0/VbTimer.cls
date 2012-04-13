VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VbTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                 This Class adds support for HiResolution Timmer                     '
'        (You have)CopyRight © 2003 Saadat Ali Shah, shahji_2000@yahoo.com            '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

'local variable(s) to hold property value(s)
Private m_ElapsedSeconds As Double, m_FPS As Double
Private m_Paused As Boolean   'While paused m_elapsedseconds will return 0, use either start again or set paused=False to get time
Private m_Freq As Currency, m_TStart As Currency, m_TLast As Currency, m_TCurrent As Currency

Private Sub Class_Initialize()
  'Get clock frequency
  If (QueryPerformanceFrequency(m_Freq)) = 0 Then 'Waving Bye Bye!
   MsgBox "Your System doesn't support High Resolutiom Timer!", vbCritical
   'For those of u, who want to use this class: instead of following line use : End
   OGLWin.Done = True ' Here is the signal, Program will terminate!
  End If
End Sub

Public Function UpdateTimer(Optional ElapsedFrames As Integer = 1) As Double
  If Not Paused Then
   QueryPerformanceCounter m_TCurrent
   m_ElapsedSeconds = (m_TCurrent - m_TLast) / m_Freq
   m_TLast = m_TCurrent 'Update
   m_FPS = ElapsedFrames / m_ElapsedSeconds ' a Quick fix for FPS
  End If
  UpdateTimer = m_ElapsedSeconds 'Resturn Elapsed Seconds
End Function

Public Sub Start()
  'Go
  m_Paused = False
  QueryPerformanceCounter m_TStart
  m_TLast = m_TStart: m_TCurrent = 0
End Sub
Public Property Get ElapsedSeconds() As Currency
  ElapsedSeconds = m_ElapsedSeconds
End Property

Public Property Get Frequency() As Currency
  Frequency = m_Freq
End Property

Public Property Get FPS() As Currency
  FPS = m_FPS
End Property

Public Property Get Paused() As Boolean
  Paused = m_Paused
End Property

Public Property Get TotalTime() As Currency
  TotalTime = (m_TCurrent - m_TStart) / m_Freq
End Property

Public Property Let Paused(ByVal NewValue As Boolean)
  m_Paused = NewValue
  If m_Paused Then
    m_ElapsedSeconds = 0
  Else
    Start
  End If
End Property
