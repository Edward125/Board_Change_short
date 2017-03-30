VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Do not test a pad-resistor in shorts"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15885
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   15885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   5520
      TabIndex        =   20
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load.."
      Height          =   315
      Left            =   6480
      TabIndex        =   19
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton cmdAddReplace 
      Caption         =   "Add a replace node"
      Height          =   315
      Left            =   5460
      TabIndex        =   18
      Top             =   3960
      Width           =   1995
   End
   Begin VB.TextBox txtAddReplaceNode 
      Height          =   285
      Left            =   1920
      TabIndex        =   17
      Text            =   "NodeA->NodeB"
      Top             =   3960
      Width           =   3495
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Del"
      Height          =   315
      Left            =   1080
      TabIndex        =   8
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdAddDevice 
      Caption         =   "Add..."
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdOutput 
      Caption         =   "Output Board File"
      Height          =   375
      Left            =   7440
      TabIndex        =   6
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Node"
      Height          =   315
      Left            =   7500
      TabIndex        =   5
      Top             =   3960
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtBoard 
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Text            =   "c:\board"
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdLoadBoard 
      Caption         =   "Load board information"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.ListBox lstNet 
      Height          =   2985
      Left            =   0
      TabIndex        =   1
      Top             =   4320
      Width           =   15735
   End
   Begin VB.ListBox lstDevice 
      Height          =   2790
      ItemData        =   "Form1.frx":08CA
      Left            =   120
      List            =   "Form1.frx":08CC
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   2880
      TabIndex        =   9
      Top             =   720
      Width           =   6615
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   4335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   4335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5040
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddAll 
         Caption         =   "Add All"
         Height          =   375
         Left            =   5040
         TabIndex        =   10
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Detail"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Select target name, all node will be target name"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   3975
      End
   End
   Begin VB.Image Image1 
      Height          =   4065
      Left            =   9600
      Picture         =   "Form1.frx":08CE
      Top             =   120
      Width           =   6240
   End
   Begin VB.Label labBoard 
      Height          =   375
      Left            =   9120
      TabIndex        =   16
      Top             =   120
      Width           =   1995
   End
   Begin VB.Label Label2 
      Caption         =   "net change list"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   1635
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private allPin() As Pin
Private allNetChange As NetChange

Private Sub cmdAdd_Click()

Dim i As Integer, j As Integer
Dim i2 As Integer, j2 As Integer
Dim ar() As String
Dim a As String
Dim b As String
Dim aFind As Boolean, bFind As Boolean, abFind As Boolean, baFind As Boolean

a = UCase(Me.Option2.Caption)
b = UCase(Me.Option1.Caption)
If a = b Then Exit Sub

For i = 0 To Me.lstNet.ListCount - 1
    ar = Split(Me.lstNet.List(i), "->")
    For j = 0 To UBound(ar)
        If ar(j) = a Then
            aFind = True
        End If
        If ar(j) = b Then
            bFind = True
        End If
        
    Next
    If aFind And bFind Then 'already in list for both nodes
      Me.lstNet.ListIndex = i
      Exit Sub
    End If
    If aFind Or bFind Then
        Me.lstNet.ListIndex = i
        If aFind Then
            For i2 = i + 1 To Me.lstNet.ListCount - 1
                ar = Split(Me.lstNet.List(i2), "->")
                For j2 = 0 To UBound(ar)
                    If ar(j2) = b Then abFind = True
                Next
                If abFind = True Then Exit For
            Next
            If abFind = False Then
            Me.lstNet.List(i) = Me.lstNet.List(i) & "->" & b
            Else
            Me.lstNet.List(i) = Me.lstNet.List(i) & "->" & Me.lstNet.List(i2)
            Me.lstNet.RemoveItem (i2)
            
            End If
            
        End If
        If bFind Then
            For i2 = i + 1 To Me.lstNet.ListCount - 1
                ar = Split(Me.lstNet.List(i2), "->")
                For j2 = 0 To UBound(ar)
                    If ar(j2) = a Then baFind = True
                Next
                If baFind = True Then Exit For
            Next
            If baFind = False Then
            Me.lstNet.List(i) = Me.lstNet.List(i) & "->" & a
            Else
            Me.lstNet.List(i) = Me.lstNet.List(i) & "->" & Me.lstNet.List(i2)
            Me.lstNet.RemoveItem (i2)
            
            End If
            

        End If

       Exit For
    End If
    
Next

If aFind Or bFind Then
'    If aFind = True And bFind = False Then
'        Me.lstNet.List(i) = Me.lstNet.List(i) & "->" & b
'    End If
'    If bFind = True And aFind = False Then
'        Me.lstNet.List(i) = Me.lstNet.List(i) & "->" & a
'    End If
Else
    If Me.Option1.Value Then
        Me.lstNet.AddItem Me.Option2.Caption & "->" & Me.Option1.Caption
    ElseIf Me.Option2.Value Then
        Me.lstNet.AddItem Me.Option1.Caption & "->" & Me.Option2.Caption
    End If
End If
End Sub

Private Sub cmdAddAll_Click()
Dim i As Integer

Me.Option1.Value = True

For i = 0 To Me.lstDevice.ListCount - 1
    Me.lstDevice.ListIndex = i
    If cmdAdd.Enabled = True Then cmdAdd_Click
    DoEvents
Next
MsgBox "over"
End Sub

Private Sub cmdAddDevice_Click()
Dim f As New FileSystemObject
Dim tr As TextStream
On Error GoTo errH

Me.CommonDialog1.ShowOpen
Set tr = f.OpenTextFile(Me.CommonDialog1.FileName)
Do Until tr.AtEndOfStream
    Me.lstDevice.AddItem tr.ReadLine

Loop

errH:
  If Err Then MsgBox Err.Description


End Sub

Private Sub cmdAddReplace_Click()
Dim ar() As String
Me.txtAddReplaceNode.Text = UCase(Me.txtAddReplaceNode.Text)

ar = Split(Me.txtAddReplaceNode, "->")

Me.Option1.Caption = ar(0)
Me.Option2.Caption = ar(1)
Call cmdAdd_Click

End Sub

Private Sub cmdDel_Click()
Dim i As Integer
i = Me.lstDevice.ListIndex
If i = -1 Then Exit Sub
Me.lstDevice.RemoveItem i
If i < Me.lstDevice.ListCount Then
   Me.lstDevice.ListIndex = i
Else
    Me.lstDevice.ListIndex = Me.lstDevice.ListCount - 1
End If
End Sub

Private Sub cmdDelete_Click()
If Me.lstNet.ListIndex = -1 Then Exit Sub

If Right(Me.lstNet.List(Me.lstNet.ListIndex), 7) <> "->(XXX)" Then
    Me.lstNet.List(Me.lstNet.ListIndex) = Me.lstNet.List(Me.lstNet.ListIndex) & "->(XXX)"
Else
    Me.lstNet.List(Me.lstNet.ListIndex) = Replace(Me.lstNet.List(Me.lstNet.ListIndex), "->(XXX)", "")
End If
 

End Sub

Private Sub cmdLoad_Click()

Me.CommonDialog1.CancelError = True
On Error GoTo errH
Me.CommonDialog1.ShowOpen

Dim f As New FileSystemObject
Dim tr As TextStream
Dim s As String
Set tr = f.OpenTextFile(Me.CommonDialog1.FileName, ForReading)
Me.lstNet.Clear
Do Until tr.AtEndOfStream
   s = Trim(tr.ReadLine)
   If s <> "" Then Me.lstNet.AddItem s
Loop

errH:
   If Err Then MsgBox Err.Description
   
End Sub

Private Sub cmdLoadBoard_Click()
Dim f As New FileSystemObject
Dim tr As TextStream
Set tr = f.OpenTextFile(Me.txtBoard.Text)
Dim Section As Integer   'if reading in [DEVICES] section=1 ,in [CONNECT] section=2
Dim s As String
Dim Head As String 'if in device,this is device name
                   'if in connections this is net name
                   
Dim blnPanelBoard As Boolean 'this is a panel board or not
Dim blnThisBoard As Boolean   'if in a panel file .you need select a board.

Dim iPin As Long
ReDim allPin(0)


Section = 0
Do Until tr.AtEndOfStream
    s = Trim(tr.ReadLine)
    
    If InStr(s, "BOARD ") = 1 Then
        blnPanelBoard = True
        If MsgBox("Do you want to deal this board?" & vbCrLf & s, vbQuestion Or vbYesNo, "select a board") = vbYes Then
           blnThisBoard = True
           Me.labBoard.Caption = s
        End If
    End If
    If s = "CAPACITOR" And blnPanelBoard = False Then blnThisBoard = True ' this is not a panel. let it begin
    
    If blnThisBoard = True Then
            If (s = "DEVICES" Or s = "CONNECTIONS") And Section <> 0 Then
                Exit Do 'finish read.
            End If
            If s = "END BOARD" Then Exit Do
            
            If s = "DEVICES" Or s = "CONNECTIONS" Then
                If s = "DEVICES" Then
                    Section = 1
                End If
                If s = "CONNECTIONS" Then
                    Section = 2
                End If
            Else
            
                If Section = 1 Then 'read in devices
                     
                End If
                
                If Section = 2 Then 'read in connections
                    If s <> "" Then
                        If Head = "" And s <> "" Then
                            Head = s
                        ElseIf Head <> "" Then
                            'CRTBD2.6
                            ReDim Preserve allPin(iPin)
                            allPin(iPin).net = Head
                            allPin(iPin).DeviceName = getLeft(s)
                            allPin(iPin).pinName = getRight(s)
                            iPin = iPin + 1
                            If Right(s, 1) = ";" Then
                                Head = ""
                            End If
                            
                        End If
                    End If
                End If
                
            End If
    End If 'blnthisboard=true
Loop

tr.Close
Set tr = Nothing
Set f = Nothing
If UBound(allPin) = 0 Then
    MsgBox "No device pin found, May be you need save board as netlist format in [board constant->IPG global-->list format]", vbCritical
Else
    MsgBox "Load OK ,Total Find device pin:" & UBound(allPin), vbInformation
End If

End Sub
Private Function getLeft(s As String) As String
    Dim i As Integer
    i = InStr(s, ".")
    getLeft = Left(s, i - 1)
    
End Function
Private Function getRight(s As String) As String
    Dim i As Integer
    
    
    i = InStr(s, ".")
    getRight = Replace(Mid(s, i + 1), ";", "")
    
End Function

Private Sub cmdOutput_Click()
Dim f As New FileSystemObject
Dim trIn As TextStream
Dim trOut As TextStream
Dim ar() As String, ar2() As String

Dim i As Integer, ti As Integer
Dim j As Integer
Dim Section As Integer 'section=1 'DEVICES
                       'section=2 'CONNECTIONS

Dim blnPanelBoard As Boolean 'this is a panel board or not
Dim blnThisBoard As Boolean   'if in a panel file .you need select a board.

Dim s As String
Dim s2 As String
Me.cmdOutput.Caption = "working..."
Me.cmdOutput.Enabled = False

Set trIn = f.OpenTextFile(Me.txtBoard)
Set trOut = f.OpenTextFile(Me.txtBoard & ".2", ForWriting, True)

Do Until trIn.AtEndOfStream
    s = Trim(trIn.ReadLine)
    
    If blnThisBoard = False Then 'not deal a board
    If Me.labBoard.Caption <> "" Then  'this is a panel board file.
        If s = Me.labBoard.Caption Then
            blnThisBoard = True
        End If
    Else                        ' this is not a panel. let it begin
        blnThisBoard = True
    End If
    End If
    
    If blnThisBoard Then
                If s = "DEVICES" Then
                    Section = 1
                ElseIf s = "CONNECTIONS" Then
                    Section = 2
                End If
                    
                If Section = 2 Then 'CONNECTIONS
                    For i = 0 To Me.lstNet.ListCount - 1
                        If Right(Me.lstNet.List(i), 7) <> "->(XXX)" Then
                            ar = Split(Me.lstNet.List(i), "->")
                            For j = 0 To UBound(ar) - 1
                                If UCase(s) = UCase(ar(j)) Then
                                    s = ar(UBound(ar))
                                End If
                            Next
                        End If
                    Next
                End If
                If Section = 1 Then 'DEVICES
                    For i = 0 To Me.lstNet.ListCount - 1
                        If Right(Me.lstNet.List(i), 7) <> "->(XXX)" Then
                            ar = Split(Me.lstNet.List(i), "->")
                             
                '            If Right(s, Len(ar(0)) + 1) = "." & ar(0) Then
                '                ti = InStr(s, ".")
                '                s = Left(s, ti) & ar(1)
                '            ElseIf Right(s, Len(ar(0)) + 2) = "." & ar(0) & ";" Then
                '                ti = InStr(s, ".")
                '                s = Left(s, ti) & ar(1) & ";"
                '            End If
                            For j = 0 To UBound(ar) - 1
                                If UCase(Right(s, Len(ar(j)) + 1)) = "." & UCase(ar(j)) Then
                                    ti = InStr(s, ".")
                                    s = Left(s, ti) & ar(UBound(ar))
                                ElseIf UCase(Right(s, Len(ar(j)) + 2)) = "." & UCase(ar(j)) & ";" Then
                                    ti = InStr(s, ".")
                                    s = Left(s, ti) & ar(UBound(ar)) & ";"
                                End If
                            Next
                        End If
                    Next
                End If
    End If 'deal this board
    
    trOut.WriteLine s
Loop
trIn.Close
trOut.Close
'board_xy


Set trIn = f.OpenTextFile(Me.txtBoard & "_xy")
Set trOut = f.OpenTextFile(Me.txtBoard & "_xy.2", ForWriting, True)

blnThisBoard = False

Do Until trIn.AtEndOfStream
    s = Trim(trIn.ReadLine)
    
    If blnThisBoard = False Then 'not deal a board
    If Me.labBoard.Caption <> "" Then  'this is a panel board file.
        If s = Me.labBoard.Caption Then
            blnThisBoard = True
        End If
    Else                        ' this is not a panel. let it begin
        blnThisBoard = True
    End If
    End If
        
    If blnThisBoard = True Then
            For i = 0 To Me.lstNet.ListCount - 1
                If Right(Me.lstNet.List(i), 7) <> "->(XXX)" Then
                    ar = Split(Me.lstNet.List(i), "->")
                    If Left(s, 4) = "NODE" Then
                        ar2() = Split(s, " ")
                        For j = 0 To UBound(ar) - 1
                            If UCase(ar(j)) = UCase(ar2(1)) Then
                                ar2(1) = ar(UBound(ar))
                                If UBound(ar2) >= 2 Then
                                    Debug.Print s
                                    If UCase(ar2(2)) = "NO_ACCESS;" Then s = ""
                                Else
                                    s = Join(ar2, " ")
                                End If
                            ElseIf UCase(ar(UBound(ar))) = UCase(ar2(1)) Then 'target is no_probe
                                If UBound(ar2) >= 2 Then
                                    Debug.Print s, "target no probe"
                                    If ar2(2) = "NO_ACCESS;" Then s = ""
                                End If
                            End If
                        Next
                    End If
                End If
            Next
    End If 'thisboard=true'
    trOut.WriteLine s
    
Loop

trIn.Close
trOut.Close
Set trOut = f.OpenTextFile(Me.txtBoard & "_short.txt", ForWriting, True)
For j = 0 To Me.lstNet.ListCount
    trOut.WriteLine Me.lstNet.List(j)
Next
trOut.Close


MsgBox "write to file " & Me.txtBoard & ".2 OK!" & vbCrLf & "short readme to " & Me.txtBoard & "_short.txt"

Me.cmdOutput.Caption = "Output"
Me.cmdOutput.Enabled = True

End Sub

Private Sub cmdSave_Click()
Dim f As New FileSystemObject
Dim trOut As TextStream
Dim j As Integer

Set trOut = f.OpenTextFile(Me.txtBoard & "_short.txt", ForWriting, True)
For j = 0 To Me.lstNet.ListCount
    trOut.WriteLine Me.lstNet.List(j)
Next
trOut.Close
MsgBox "ok  " & Me.txtBoard & "_short.txt", vbInformation
End Sub

Private Sub Form_Load()
   'ReDim Preserve allPin(3)

End Sub
  
 
Private Sub Form_Resize()
On Error Resume Next
Me.lstNet.Width = Me.Width - 100

End Sub

Private Sub lstDevice_Click()
Dim i As Long
On Error GoTo errH

Me.Label1.Caption = "Detail of " & Me.lstDevice.Text

Dim Anode As Integer
Dim Bnode As Integer

Anode = -1
Bnode = -1
 
Me.Option1.Caption = ""
Me.Option2.Caption = ""

Me.cmdAdd.Enabled = False

For i = 0 To UBound(allPin)
    If UCase(allPin(i).DeviceName) = UCase(lstDevice.Text) Then
         If Anode = -1 Then
            Anode = i
         ElseIf Bnode = -1 Then
            Bnode = i
         Else
            MsgBox "wrong, more than 2 pin find .", vbCritical
            Exit Sub
         End If
    End If
Next

Me.Option1.Caption = allPin(Anode).net
Me.Option2.Caption = allPin(Bnode).net

If Anode >= 0 And Bnode >= 0 Then Me.cmdAdd.Enabled = True
errH:
 If Err.Number <> 9 And Err.Number <> 0 Then MsgBox Err.Description
End Sub

Private Sub lstNet_DblClick()
Dim ar() As String
Dim i As Integer
Dim j As Integer
Dim s As String
If Right(Me.lstNet.List(Me.lstNet.ListIndex), 7) = "->(XXX)" Then
    Exit Sub
End If

i = Me.lstNet.ListIndex
ar = Split(Me.lstNet.List(i), "->")
s = ar(UBound(ar))

For j = 0 To UBound(ar) - 1
    s = s & "->" & ar(j)
Next

Me.lstNet.List(i) = s

End Sub

Private Sub txtBoard_DblClick()
Me.CommonDialog1.CancelError = True
On Error GoTo errH
Me.CommonDialog1.ShowOpen

Me.txtBoard.Text = Me.CommonDialog1.FileName
errH:
   If Err Then MsgBox Err.Description
   

End Sub
