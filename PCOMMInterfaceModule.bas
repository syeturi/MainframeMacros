Attribute VB_Name = "PCOMMInterfaceModule"
' Author. Srinivas Yeturi
' 2005

Global PCOM_ConnList As Object
Global PCOM_Ps As Object
Global PCOM_Oia As Object
Global PCOM_Session As Object
Global PCOM_Field As Object

Function PCOMM_SendKeys(Handle As String, Cmdtxt As String)

Set PCOM_ConnList = CreateObject("pcomm.autECLConnlist")
Set PCOM_Oia = CreateObject("pcomm.autECLOIA")
Set PCOM_Ps = CreateObject("pcomm.autECLPS")
Set PCOM_Session = CreateObject("pcomm.autECLSession")
Set PCOM_Metrics = CreateObject("pcomm.autECLWinMetrics")

PCOM_Session.SetConnectionByName (Handle)
PCOM_Oia.SetConnectionByName (Handle)
PCOM_Ps.SetConnectionByName (Handle)

Call Check_PCOMM_Status(Handle)
PCOM_Ps.SendKeys Cmdtxt
End Function

Function PCOMM_Cursor_CurrentX(Handle As String) As Integer

Set PCOM_ConnList = CreateObject("pcomm.autECLConnlist")
Set PCOM_Oia = CreateObject("pcomm.autECLOIA")
Set PCOM_Ps = CreateObject("pcomm.autECLPS")
Set PCOM_Session = CreateObject("pcomm.autECLSession")

PCOM_Session.SetConnectionByName (Handle)
PCOM_Oia.SetConnectionByName (Handle)
PCOM_Ps.SetConnectionByName (Handle)

Call Check_PCOMM_Status(Handle)

PCOMM_Cursor_CurrentX = PCOM_Ps.CursorPosRow
End Function


Function PCOMM_Cursor_CurrentY(Handle As String) As Integer

Set PCOM_ConnList = CreateObject("pcomm.autECLConnlist")
Set PCOM_Oia = CreateObject("pcomm.autECLOIA")
Set PCOM_Ps = CreateObject("pcomm.autECLPS")
Set PCOM_Session = CreateObject("pcomm.autECLSession")

PCOM_Session.SetConnectionByName (Handle)
PCOM_Oia.SetConnectionByName (Handle)
PCOM_Ps.SetConnectionByName (Handle)
Call Check_PCOMM_Status

PCOMM_Cursor_CurrentY = PCOM_Ps.CursorPoscol
End Function


Function PCOMM_Write_Screen(Handle As String, FieldX As Integer, FieldY As Integer, FieldValue As String)

Set PCOM_ConnList = CreateObject("pcomm.autECLConnlist")
Set PCOM_Oia = CreateObject("pcomm.autECLOIA")
Set PCOM_Ps = CreateObject("pcomm.autECLPS")
Set PCOM_Session = CreateObject("pcomm.autECLSession")

PCOM_Session.SetConnectionByName (Handle)
PCOM_Oia.SetConnectionByName (Handle)
PCOM_Ps.SetConnectionByName (Handle)
Call Check_PCOMM_Status(Handle)
Call PCOM_Ps.SetText(FieldValue, FieldX, FieldY)

End Function

Function PCOMM_Read_Screen(Handle As String, Vx As Integer, vy As Integer, vl As Integer) As String

Set PCOM_ConnList = CreateObject("pcomm.autECLConnlist")
Set PCOM_Oia = CreateObject("pcomm.autECLOIA")
Set PCOM_Ps = CreateObject("pcomm.autECLPS")
Set PCOM_Session = CreateObject("pcomm.autECLSession")

PCOM_Session.SetConnectionByName (Handle)
PCOM_Oia.SetConnectionByName (Handle)
PCOM_Ps.SetConnectionByName (Handle)
Call Check_PCOMM_Status(Handle)

PCOMM_Read_Screen = PCOM_Ps.GetTextRect(Vx, vy, Vx, vy + vl - 1)

End Function

    
Function Check_PCOMM_Status(Handle As String)

Set PCOM_ConnList = CreateObject("pcomm.autECLConnlist")

Set PCOM_Oia = CreateObject("pcomm.autECLOIA")
Set PCOM_Ps = CreateObject("pcomm.autECLPS")
Set PCOM_Session = CreateObject("pcomm.autECLSession")

PCOM_Session.SetConnectionByName (Handle)
PCOM_Oia.SetConnectionByName (Handle)
PCOM_Ps.SetConnectionByName (Handle)

If PCOM_Oia.WaitForAppAvailable(5000) = False Then
MsgBox " PCOMM Screen Not Responding .. Cannot continue"
End
End If

If PCOM_Oia.WaitForInputReady(5000) = False Then
MsgBox " PCOMM Screen Not Responding .. Cannot continue"
End
End If

End Function

