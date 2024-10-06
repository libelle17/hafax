VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FürIcon 
   Caption         =   "Sammelfax an Hausärzte"
   ClientHeight    =   11325
   ClientLeft      =   570
   ClientTop       =   450
   ClientWidth     =   17355
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11325
   ScaleWidth      =   17355
   Begin VB.TextBox Ausgabe 
      Height          =   2655
      Left            =   7200
      TabIndex        =   29
      Top             =   8400
      Width           =   2415
   End
   Begin VB.CommandButton FaxeEinstellen 
      Caption         =   "&Faxe einstellen"
      Height          =   375
      Left            =   7200
      TabIndex        =   8
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox DateiDemo 
      Height          =   10935
      Left            =   9720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   28
      Top             =   240
      Width           =   7575
   End
   Begin VB.CommandButton DAuswahl 
      Caption         =   "Ad&ressen auswählen"
      Height          =   375
      Index           =   1
      Left            =   7680
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox DateiName 
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   4
      Top             =   720
      Width           =   6855
   End
   Begin VB.CommandButton Faxen 
      Caption         =   "&Faxen"
      Height          =   255
      Left            =   7200
      TabIndex        =   7
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Nephrologie"
      Height          =   255
      Index           =   14
      Left            =   7200
      TabIndex        =   17
      Top             =   4560
      Width           =   2575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Nervenheilkunde"
      Height          =   255
      Index           =   13
      Left            =   7200
      TabIndex        =   19
      Top             =   5040
      Width           =   2575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Labor"
      Height          =   255
      Index           =   12
      Left            =   7200
      TabIndex        =   27
      Top             =   6720
      Width           =   2575
   End
   Begin VB.CheckBox Check2 
      Caption         =   "alle"
      Height          =   255
      Left            =   7200
      TabIndex        =   26
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Anästhesie"
      Height          =   255
      Index           =   11
      Left            =   7200
      TabIndex        =   25
      Top             =   6480
      Width           =   2575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Chirurgie"
      Height          =   255
      Index           =   10
      Left            =   7200
      TabIndex        =   24
      Top             =   6240
      Width           =   2575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Dermatologie"
      Height          =   255
      Index           =   9
      Left            =   7200
      TabIndex        =   23
      Top             =   6000
      Width           =   2575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Orthopädie"
      Height          =   255
      Index           =   8
      Left            =   7200
      TabIndex        =   22
      Top             =   5760
      Width           =   2575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "plastische Chirurgie"
      Height          =   255
      Index           =   7
      Left            =   7200
      TabIndex        =   21
      Top             =   5520
      Width           =   2575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "HNO"
      Height          =   255
      Index           =   6
      Left            =   7200
      TabIndex        =   20
      Top             =   5280
      Width           =   2575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Kardiologie"
      Height          =   255
      Index           =   5
      Left            =   7200
      TabIndex        =   18
      Top             =   4800
      Width           =   2575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Kinder- und Jugendmedizin"
      Height          =   255
      Index           =   4
      Left            =   7200
      TabIndex        =   16
      Top             =   4320
      Width           =   2575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Urologie"
      Height          =   255
      Index           =   3
      Left            =   7200
      TabIndex        =   15
      Top             =   4080
      Width           =   2575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Gynäkologie"
      Height          =   255
      Index           =   2
      Left            =   7200
      TabIndex        =   14
      Top             =   3840
      Width           =   2575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Innere Medizin"
      Height          =   255
      Index           =   1
      Left            =   7200
      TabIndex        =   13
      Top             =   3600
      Width           =   2575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Prakt.Arzt / Allgemeinmedizin"
      Height          =   255
      Index           =   0
      Left            =   7200
      TabIndex        =   12
      Top             =   3360
      Width           =   2575
   End
   Begin VB.CommandButton DatenbankWählen 
      Caption         =   "&Datenbank wählen"
      Height          =   375
      Left            =   7200
      TabIndex        =   10
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Fortschritt 
      Height          =   10095
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   11
      Top             =   1200
      Width           =   6375
   End
   Begin VB.CommandButton Abbruch 
      Caption         =   "Abbru&ch"
      Height          =   255
      Left            =   7200
      TabIndex        =   9
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton DAuswahl 
      Caption         =   "&Datei auswählen"
      Height          =   375
      Index           =   0
      Left            =   7680
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox DateiName 
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   6855
   End
   Begin VB.CommandButton Start 
      Caption         =   "&Vorauswahl"
      Height          =   255
      Left            =   7200
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label DateiBez 
      Caption         =   "&Adressen:"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.Label DateiBez 
      Caption         =   "D&atei:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "FürIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindowsDirectory& Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer$, ByVal nSize&)
Private Declare Function sndPlaySound32& Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName$, ByVal uFlags&)
Const XPFehler$ = "Windows XP-Fehler.wav"
Const XPHinweis$ = "Windows XP-Hinweis.wav"
Const ringout$ = "ringout.wav"
Const recycle$ = "recycle.wav"
Dim WinDir$
Dim ErrNumber&, ErrDescription$, ErrLastDllError&, ErrSource
'Const EigDat = "\\linux\daten\eigene Dateien"  '"\\linux\Gemein\eigene Dateien" ' "\\anmeld2\u"
Dim EigDat$
Const RegStelle$ = "Software\GSProducts\HAFax"
Const HCU = &H80000001

Public WithEvents dbv As DBVerb
Attribute dbv.VB_VarHelpID = -1
Function WD()
 WinDir = Space(144)
 Call GetWindowsDirectory(WinDir, 144)
 WinDir = REPLACE(Trim(WinDir), Chr(0), "")
 WD = WinDir
End Function ' WD
Function Sound(Pfad$)
    Call sndPlaySound32(Pfad, 1)
End Function
Public Function Tön(Datei$)
 Call WD
 Call Sound(WinDir + "\media\" + Datei)
End Function
Public Function XPH()
 Call Tön(XPHinweis)
End Function

Public Function GetFileToOpen(Index%)
 Dim fileflags As FileOpenConstants
 Dim filefilter$
 On Error GoTo fehler
 'Set the text in the dialog title bar
 
 Select Case Index
  Case 0
   CommonDialog1.DialogTitle = "Zu sendende Datei auswählen:"
   CommonDialog1.InitDir = EigDat
  Case 1
   CommonDialog1.DialogTitle = "Adressdatei auswählen:"
   CommonDialog1.InitDir = "u:\"
 End Select
 'Set the default file name and filter
 CommonDialog1.FileName = vNS
 filefilter = "Verzeichnisse (*.*)|*.*|Alle Dateien (*.*)|*.*"
 CommonDialog1.Filter = filefilter
 CommonDialog1.FilterIndex = 0
 'Verify that the file exists
 fileflags = cdlOFNFileMustExist + cdlOFNHideReadOnly
 CommonDialog1.Flags = fileflags
 'Show the Open common dialog box
 CommonDialog1.ShowOpen
 'Return the path and file name selected or
 'Return an empty string if the user cancels the dialog
 Me.DateiName(Index) = CommonDialog1.FileName
 Exit Function
fehler:
 ErrNumber = Err.Number
 ErrDescription = Err.Description
 ErrLastDllError = Err.LastDllError
 ErrSource = Err.source
 Call XPH
 Select Case MsgBox("FNr: " + CStr(ErrNumber) + vbCrLf + "LastDLLError: " + CStr(ErrLastDllError) + vbCrLf + "Source: " + IIf(IsNull(ErrSource), "", CStr(ErrSource)) + vbCrLf + "Description: " + ErrDescription + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in Wandel/" + App.Path)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' GetFileToOpen

Private Sub Abbruch_Click()
 Unload Me
 ProgEnde
End Sub ' Abbruch_Click()

Private Sub DatenbankWählen_Click()
 Call dbv.Auswahl("kvaerzte", "hae", "Hausärzte")
End Sub ' DatenbankWählen_Click()

Private Sub DAuswahl_Click(Index As Integer)
 Call GetFileToOpen(Index)
 If Index = 1 And Me.DateiName(1) <> vNS Then
  Call demo
 End If
End Sub ' DAuswahl_Click(index As Integer)

Private Sub demo()
 Dim Fdat As New CString, zeile$, zlenge&
 Close #401
 Open Me.DateiName(1) For Input As #401
 Fdat = Input(LOF(401), #401)
 Me.DateiDemo = Fdat.Value
 Close #401
End Sub ' demo()

Private Sub FaxeEinstellen_Click()
 Dim Spli$(), zwi$, i&
 Open Me.DateiName(1) For Input As #402
 i = 1
 Do While Not EOF(402)
  Line Input #402, zwi
  Spli = Split(zwi, "|")
  If UBound(Spli) > 2 Then
   Spli(2) = REPLACE$(REPLACE$(REPLACE$(Spli(2), "0(", ""), ")", ""), " ", "")
   If Not IsNumeric(Trim(Spli(2))) Then Spli(2) = REPLACE(Spli(2), "ehem.", "")
   If IsNumeric(Trim(Spli(2))) Then
    Me.Fortschritt = i & ": " & Trim(Spli(1)) & Trim(Spli(4)) & vbCrLf & Me.Fortschritt
    DoEvents
    Call doFaxEinstellen(Trim(Spli(1)), Trim(Spli(2)), Me.DateiName(0))
    DoEvents
   Else
    Call dbv.Ausgeb(Spli(1) + " nicht faxbar, keine Nummer", True)
   End If
  ElseIf InStrB(zwi, "Zeilen") <> 0 Then
  Else
   MsgBox "Zeile nicht wohlgeformt: " & vbCrLf & zwi
   Close #402
   Exit Sub
  End If
  i = i + 1
 Loop
 Close #402
 Exit Sub
End Sub ' FaxeEinstellen_Click()

Private Sub doFaxEinstellen(HAName$, nr$, uName$)
 Dim uReinName$
 uReinName = uName
 If (Left(uReinName, 2) <> "\\") Then
 Dim Pos1%, pos0%
 Do
  Pos1 = pos0
  pos0 = InStr(pos0 + 1, uReinName, "\")
 Loop Until pos0 = 0
 uReinName = Mid(uReinName, Pos1 + 1)
 uReinName = Left(uReinName, Len(uReinName) - 4)
 uReinName = "p:\zufaxen\" + uReinName + " an " + HAName + " an Fax " + nr + Right(uName, 4)
 End If
 FileCopy uName, uReinName
End Sub ' U:\QZ\30.6.14 Einladung.doc

Private Sub Faxen_Click()
 Dim Spli$(), zwi$, i&
 Open Me.DateiName(1) For Input As #402
 i = 1
 Do While Not EOF(402)
  Line Input #402, zwi
  Spli = Split(zwi, "|")
  If UBound(Spli) > 2 Then
   If Not IsNumeric(Trim(Spli(2))) Then Spli(2) = REPLACE(Spli(2), "ehem.", "")
   If IsNumeric(Trim(Spli(2))) Then
    Me.Fortschritt = i & ": " & Trim(Spli(1)) & Trim(Spli(4)) & vbCrLf & Me.Fortschritt
    DoEvents
    Call FaxSend(RecName:=Trim(Spli(1)), RecNum:=Trim(Spli(2)), DateiName:=Me.DateiName(0), obstreng:=True)
    DoEvents
   Else
    Stop
   End If
  ElseIf InStrB(zwi, "Zeilen") <> 0 Then
  Else
   Stop
  End If
  i = i + 1
 Loop
 Close #402
 Exit Sub
#If False Then
 rs.MoveFirst
 Do While Not rs.EOF
  If Not IsNull(rs!Fax) And rs!Fax <> "" And rs!Fax <> altfax Then
'   Select Case rs!nachname
'    Case "Rößler", "Winkler-Huber", "Frowein", "Folkerts", "Kirchhoff", "Papadopoulos", "Bühler", "Wasmer", "Mialkowskyj", "Kiener", "Eder", "Klier-Deißenböck", "Metzler", "Gross", "Mohr", "Zimmermann", "Peschers", "Hösler", "Geiger", "Düll", "Cappeller", "Frank", "Früchte", "Grahamer", "Kessler", "Klein", "Franke-Wirsching", "Klinzing-Eidens", "Kreie", "Laube", "Lechner", "Kobras", "Moser", "Niezel", "Kollmann", "Past", "Pfeuffer", "Pöhlmann", "Kaltenegger", "Pöschl", "Proß", "Räpple", "Reitmeier", "Ressel", "Rester", "Ringmaier", "Roß", "Ross", "Senner", "Slavin", "Skoruppa", "Szymkowiak", "Linenauer", "Schneider", "Schorten", "Schuff", "Stolzki", "Stöwer", "Tomahogh", "Turba-Schillinger", "Krombholz", "Hübner-Krombholz", "Walter", "Wainryb", "Wegele", "Zankl", "Käufl", "Hernas", "Goß", "Peller", "Garnerus", "Magoley", "Bringmann", "Ziegelhöfer", "Schmucker"
   FaxName = ""
   FaxName = "Dr." + rs!vorname + " " + rs!nachname
   Dim altFortschritt$
   altFortschritt = Me.Fortschritt
   Me.Fortschritt = i & ": " & FaxName & ", " & rs!ort & " " & rs!zulg & " " & rs!Fax & " ..." & vbCrLf & Me.Fortschritt
'   If rs!Fax = "08131616381" Then
    DoEvents
    Call FaxSend(FaxName, rs!Fax, Me.DateiName)
    DoEvents
'   End If
   Me.Fortschritt = i & ": " & FaxName & ", " & rs!ort & " " & rs!zulg & " " & rs!Fax & vbCrLf & altFortschritt
'   End Select
   i = i + 1
  End If
  altfax = rs!Fax
  rs.Move 1
 Loop
#End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
  Unload Me
  ProgEnde
 End If ' KeyAscii = 27
End Sub ' Form_KeyPress(KeyAscii As Integer)

Private Sub Form_Load()
 Set dbv = New DBVerb
 Me.DateiName(0) = getReg(1, RegStelle, "DateiName")
 Me.DateiName(1) = getReg(1, RegStelle, "Adressen")
 If LenB(Me.DateiName(1)) = 0 Then Me.DateiName(1) = "\\linux1\daten\eigene Dateien\testha.txt"
 demo
 EigDat = getReg(1, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal") ' eigene Dateien
End Sub ' Form_Load()

Private Sub Form_Resize()
 On Error Resume Next
 Me.Fortschritt.Height = Me.Height - Me.Fortschritt.Top - 500
 Me.DateiDemo.Height = Me.Height - Me.DateiDemo.Top - 500
 Me.DateiDemo.Width = Me.Width - Me.DateiDemo.Left - 150
End Sub ' Form_Resize()

Private Sub Form_Unload(Cancel As Integer)
 Call fStSpei(HCU, RegStelle, "DateiName", Me.DateiName(0))
 Call fStSpei(HCU, RegStelle, "Adressen", Me.DateiName(1))
End Sub

Private Sub Start_Click()
 Dim FaxName$, i&
#Const DAO = 0
#If DAO Then
 Dim db As DAO.Database, db0 As DAO.Database
 Dim rs As DAO.Recordset, rs0 As DAO.Recordset, rs1 As DAO.Recordset, i&
 Set db = DAO.OpenDatabase("u:\anamnese\quelle.mdb")
 Set db0 = DAO.OpenDatabase("u:\anamnese\KV-Ärzte 20060806.mdb")
 Set rs = db.OpenRecordset("select distinct telefax as recnum from hausärzte where not telefax in (""08136 9809"",""089 1238544"",""089 32100130"",""089 3211070"",""08131 505111"",""08131 594100"",""0721/9573420"",""08131 353556"",""08131 57120"",""08131 597160"",""08131 669833"",""08131 79302"",""08131 9069029"")")
 Set rs0 = db0.OpenRecordset("HAE", dbOpenTable)
 rs0.Index = "fax1k"
 Set rs1 = db.OpenRecordset("Hausärzte", dbOpenDynaset)
 Call fxset(EigDat)
 Do While Not rs.EOF
  If Len(rs!RecNum) > 1 Then
   FaxName = ""
   rs0.Seek "=", REPLACE(REPLACE(REPLACE(rs!RecNum, " ", ""), "/", ""), "-", "")
   If rs0.NoMatch Then
    Call rs1.FindFirst("telefax = '" + rs!RecNum + "'")
    If Not rs1.NoMatch Then
     FaxName = "Dr." + rs1!vorname + " " + rs1!nachname
    End If
   Else
    FaxName = "Dr." + rs0!vorname + " " + rs0!nachname
   End If
   If InStr(FaxName, "Kellerer") > 0 Or InStr(FaxName, "Ranft") > 0 Then
    Stop
   End If
   Debug.Print i, rs!RecNum, FaxName
   Call FaxSend(FaxName, rs!RecNum, Me.DateiName)
   i = i + 1
  End If
  rs.Move 1
 Loop
#Else
 Const tue = "update kvaerzte.hae, quelle.hausaerzte set hausaerzte.gelöscht = hae.gelöscht where hausaerzte.kvnr = hae.kvnu and hausaerzte.nachname = hae.nachname;"
' Const sql = "SELECT Anrede,KVnu as KVNr,Fax1k as Fax,Email,zulg, Ort,Vorname,Nachname,gelöscht FROM KVAerzte.hae where bstelle like ""%Dachau%"" and zulg in (""Allgemeinmedizin"",""Nervenheilkunde"",""Innere Medizin"", ""Chirurgie"", ""Praktischer Arzt"", ""Haut- u.Geschlechtskrankheiten"",""Orthopädie"",""Neurologie"",""Arzt"",""Radiologie"",""Diagnostische Radiologie"",""Plastische und Ästhetische Chi"") and not gelöscht union SELECT if(geschlecht=""w"",""Frau"",""Herr"") as Anrede,KVNr, replace(Telefax,"" "","""") as Fax,E_mail as Email,zulassungsgebiet as zulg, Ort,Vorname,Nachname,gelöscht FROM quelle.hausaerzte where zulassungsgebiet not in (""Kinder- und Jugendmedizin"") and not nichtmehr and not gelöscht order by kvnr;"
' das gleiche ohne die meisten Fachärzte
 Dim CT$(14), sql$, anfg%
 CT(0) = "'%Allg%' or zulg like '%prakt%'"
 CT(1) = "'%Innere%' or zulg like '%Intern%'"
 CT(2) = "'%gyn%' or zulg like '%Frauen%'"
 CT(3) = "'%urol%'"
 CT(4) = "'%kind%' or zulg like '%pädi%'"
 CT(5) = "'%nephr%' or zulg like '%nier%'"
 CT(6) = "'%kard%'"
 CT(7) = "'%nerv%' or zulg like '%neur%' or zulg like '%psych%'"
 CT(8) = "'%hals%' or zulg like '%hno%'"
 CT(9) = "'%plast%'"
 CT(10) = "'%orth%'"
 CT(11) = "'%derm%'"
 CT(12) = "'%chir%'" ' umfaßt dann auch plast. Chir
 CT(13) = "'%anä%'"
 CT(14) = "'%labor%'"
 sql = "SELECT distinct * from (select Anrede,KVnu as KVNr,Fax1k as Fax,Email,zulg, Ort,Vorname,Nachname,straße,gelöscht FROM kvaerzte.hae where bstelle like ""%Dachau%"" and not gelöscht union SELECT if(geschlecht=""w"",""Frau"",""Herr"") as Anrede,KVNr, replace(Telefax,"" "","""") as Fax,E_mail as Email,zulassungsgebiet as zulg, Ort,Vorname,Nachname,straße,gelöscht FROM quelle.hausaerzte where not nichtmehr and not gelöscht) as innen "
 sql = "SELECT distinct * from (select Anrede,KVnu as KVNr,Fax1k as Fax,Email,zulg, Ort,Vorname,Nachname,straße,gelöscht FROM kvaerzte.hae where bstelle like ""%Dachau%"" and not gelöscht union SELECT (select distinct anrede from kvaerzte.hae where vorname = vorname and anrede <> """" limit 1) as Anrede,KVNr, replace(fax,"" "","""") as Fax,"""" as Email,fachgruppe as zulg, Ort,Vorname,name as Nachname,strasse as straße,"""" as gelöscht FROM quelle.listenausgabeuew) as innen where fax <> """" "
 anfg = 0
 If Me.Check2 = 1 Then
  
 Else
  For i = Me.Check1.LBound To Me.Check1.UBound
   If Me.Check1(i) = 1 Then
    If Not anfg Then
     sql = sql & " and (zulg like "
     anfg = -1
    Else
     sql = sql & " zulg like "
    End If
    sql = sql & CT(i) & " or"
   End If
  Next i
 End If
 If anfg Then
  sql = Left(sql, Len(sql) - 3) & ")"
 Else
'  sql = sql & " where fax <> """" "
 End If
 sql = sql & " order by fax, kvnr, nachname, vorname;"
' Const CStrMy$ = "DRIVER={MySQL ODBC 3.51 Driver};server=linux;user=praxis;pwd=...;database=quelle;option=3"
 Dim acn As New ADODB.Connection
 acn.Open dbv.cnVorb("kvaerzte", "hae", "Hausärzte")
 Dim rs As New ADODB.Recordset
 Set rs = acn.Execute(tue)
 Set rs = Nothing
 Call rs.Open(sql, acn, adOpenDynamic, adLockReadOnly)
 Call fxset(EigDat)
 Dim altfax$
 i = 1
 Open Me.DateiName(1) For Output As #333
 Do While Not rs.EOF
  If rs!Fax <> altfax Then
'   Print #333, rs!kvnr, rs!anrede, rs!nachname, rs!vorname, rs!Fax, rs!Email, rs!zulg, rs!ort, rs!gelöscht
   Print #333, Right(Space(7) & rs!kvnr, 7) & " | " & Left("Dr." + rs!vorname + " " + rs!nachname + Space(35), 35) & " | " & Right(Space(20) & rs!Fax, 20) & " | " & Left(rs!ort & ", " & rs!strasse & ", " & rs!zulg & Space(45), 45) & " | " & IIf(IsNull(rs!gelöscht), " ", rs!gelöscht)
   i = i + 1
  End If
  altfax = rs!Fax
  rs.Move 1
 Loop
 Print #333, i & " Zeilen"
 Close #333
 Call Shell(Environ("windir") & "\system32\notepad.exe u:\testha.txt", vbMaximizedFocus)
 Exit Sub
#End If
End Sub

