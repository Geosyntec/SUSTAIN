Attribute VB_Name = "ModuleSWMMRunDLL"
Public Declare Function swmm_open Lib "swmm5.dll" (ByVal F1 As String, ByVal F2 As String, ByVal F3 As String) As Long
Public Declare Function swmm_start Lib "swmm5.dll" (ByVal saveFlag As Long) As Long
Public Declare Function swmm_step Lib "swmm5.dll" (elapsedTime As Double) As Long
Public Declare Function swmm_end Lib "swmm5.dll" () As Long
Public Declare Function swmm_close Lib "swmm5.dll" () As Long

Private Declare Function ShowWindow& Lib "USER32" (ByVal hWnd As Long, ByVal nCmdShow As Long)
Public Const SW_SHOWNORMAL& = 1


' SWMM5_IFACE.BAS
'
' Example code for interfacing SWMM 5
' with Visual Basic Applications
'
' Remember to add SWMM5.BAS to the application

Public Function RunSWMMDll(inpFile As String, rptFile As String, outFile As String, pTotalSimulationTime As Long) As Long
'------------------------------------------------------------------------------
'  Input:   inpFile = name of SWMM 5 input file
'           rptFile = name of status report file
'           outFile = name of binary output file
'  Output:  returns a SWMM 5 error code or 0 if there are no errors
'  Purpose: runs the dynamic link library version of SWMM 5.
'------------------------------------------------------------------------------
Dim err As Long
Dim elapsedTime As Double

'Open the status option
Dim lres As Variant
lres = ShowWindow(FrmSWMMSimulationStatus.hWnd, SW_SHOWNORMAL)
FrmSWMMSimulationStatus.ProgressBar.Min = 0
FrmSWMMSimulationStatus.ProgressBar.Max = 100

Dim pStep
Dim iCounter As Integer
Dim p50thSimulationTime As Long
p50thSimulationTime = CLng(pTotalSimulationTime / 50)

Dim pStartTime, pFinishTime
pStartTime = Now
Dim pTotalSeconds, pSecond, pMinute, pHour

' --- open a SWMM project
err = swmm_open(inpFile, rptFile, outFile)

If err = 0 Then

  ' --- initialize all processing systems
  err = swmm_start(1)
    
  If err = 0 Then

    ' --- step through the simulation
    Do
      ' --- allow Windows to process any pending events
      DoEvents

      ' --- extend the simulation until the next reporting time
      err = swmm_step(elapsedTime)  ' days

      '//////////////////////////////////////////
      ' call a progress reporting function here,
      iCounter = iCounter + 1
      If (iCounter > p50thSimulationTime) Then
        iCounter = 1
        pStep = (elapsedTime / pTotalSimulationTime) * 100
        If (pStep > 100) Then
            pStep = 100
        End If
        pFinishTime = Now
        FrmSWMMSimulationStatus.Caption = "SWMM Simulation " & "[" & Format(pStep, "#.#") & " %]"

        pTotalSeconds = DateDiff("s", pStartTime, pFinishTime)
        pHour = pTotalSeconds \ 3600
        pMinute = Format((pTotalSeconds - (pHour * 3600)) \ 60, "00")
        pSecond = Format((pTotalSeconds - (pMinute * 60)), "00")
        
                
        FrmSWMMSimulationStatus.lblStatus.Caption = "Time Elapsed: " & pHour & ":" & pMinute & ":" & pSecond
        FrmSWMMSimulationStatus.ProgressBar.value = CInt(pStep)
        FrmSWMMSimulationStatus.ProgressBar.Refresh
      End If
      ' using elapsedTime as an argument
      '//////////////////////////////////////////
      
     Loop While elapsedTime > 0 And err = 0
    
  End If

  ' --- close all processing systems
  swmm_end
End If

' --- close the project
swmm_close

' --- return the error code
RunSWMMDll = err

'** close the simulation form
Unload FrmSWMMSimulationStatus

End Function



