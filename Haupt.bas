Attribute VB_Name = "Haupt"
Option Explicit
Public UN$, AUP$, UP$, TMExeV$, TMNotV$, TMNotPr$
Public EigDatDirekt$, DownDirekt$, ReadODirekt$, GeraldDirekt$, TMServerDirekt$, DokumenteDirekt$
Public FSO As New FileSystemObject
Public FPos&, FNr&
Public plzVz$ ' pVerz$, vVerz$,
'Public Const uVerz$ = "c:\"
Public Const vNS$ = vbNullString
Public DBCn As New ADODB.Connection
Dim altAusgabe As New CString


Public Function Ausgeb(Text$, Optional obDauer%)
 Dim aktText As New CString
 If Not FürIcon.Visible Then
'  Debug.Print Text
 Else
  aktText = Text
  aktText.Append vbCrLf
  aktText.Append altAusgabe
  aktText.Cut 3000
  FürIcon.Ausgabe = aktText
  If obDauer <> 0 Then
   altAusgabe = aktText
  End If
  If InStrB(Text, "READ-COMMITTED") <> 0 Then
   MsgBox "Beinahe-Stop in Ausgeb:" & vbCrLf & "instrb(text, 'READ-COMMITTED') <> 0" & vbCrLf & "Text: " & Text
  End If
  DoEvents
 End If
End Function ' Ausgeb


Public Function ProgEnde()
 End
End Function ' ProgEnde
