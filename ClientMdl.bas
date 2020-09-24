Attribute VB_Name = "ClientMdl"
Private Type Id_
YourName As String
YourPass As String
CurName As String
YourClient As String
YourKey As String
End Type
Global Id As Id_
Global PM(1 To 100) As New frmPM
Public Function MkMsg(strType As String, strValue As String, Optional strFlag As String, Optional strMore As String) As String

MkMsg = strType & "|@|" & strValue & "|@|" & strFlag & "|@|" & strMore

End Function
Public Sub Send(Text As String)
DoEvents
frmMain.ws.SendData Text & "|%|"
DoEvents
End Sub
Public Sub UHOH(Text As String)
Static frmU As New frmUHOH
frmU.RichTextBox1 = Text
frmU.Visible = False
frmU.Show vbModal, frmMain
Unload frmU
Exit Sub
End Sub


