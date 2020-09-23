VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TabindeXer"
   ClientHeight    =   4395
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Down"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Up"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   1800
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Private rst As ADODB.Recordset

Option Explicit

Private Sub CancelButton_Click()
    Connect.Hide
End Sub

Private Sub RefreshList()

    Dim c As VBComponent
    Dim p As VBProject
    Dim vbc As VBControl
    Dim vbf As VBForm
    Dim sc As String
    Dim sp As String
    Dim tempIndex As Long
   
    Set rst = New ADODB.Recordset
    rst.Fields.Append "ctlname", adVarChar, 255
    rst.Fields.Append "index", adInteger
    rst.Fields.Append "tabindex", adInteger
    rst.Open
   
    Screen.MousePointer = vbHourglass
   
    sp = VBInstance.ActiveVBProject.Name
    sc = VBInstance.SelectedVBComponent.Name
    
    If sc <> "" And sp <> "" Then
        Set p = VBInstance.VBProjects.Item(sp)
        Set c = p.VBComponents.Item(sc)
        If c.Type = vbext_ct_VBForm Then
            c.Activate
            Set vbf = c.Designer
            For Each vbc In vbf.VBControls
                Select Case UCase(vbc.ClassName)
                ' conrols who do not need to be added (like labels etc) go here
                Case "LABEL", "FRAME", "LINE", "PROGRESSBAR", "LET ME KNOW IF YOU KNOW ANY OTHERS THAT SHOULD BE HERE ;)"
                Case Else
                    On Error Resume Next
                    ' prevents errors for controls wo don't have tabindex (like timer)
                    Err.Clear
                    rst.AddNew
                    rst("ctlname") = vbc.Properties("name")
                    rst("index") = vbc.Properties("index")
                    rst("tabindex") = vbc.Properties("tabindex")
                    If Err.Number = 0 Then
                        rst.Update
                    Else
                       rst.CancelUpdate
                    End If
                End Select
            Next vbc
        Else
            MsgBox "No Form Selected"
        End If
    
        ' add to list
        List1.Clear
        rst.Sort = "tabindex asc"
        If Not (rst.EOF And rst.BOF) Then rst.MoveFirst
        
        Do Until rst.EOF
            Dim strToAdd As String
            strToAdd = rst("ctlname")
            If rst("index") <> -1 Then strToAdd = strToAdd & "(" & rst("index") & ")"
            List1.AddItem strToAdd
            rst.MoveNext
        Loop
    
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub Command1_Click()

    If List1.ListIndex > 0 Then
        Dim strLine As String
        strLine = List1.List(List1.ListIndex)
        List1.List(List1.ListIndex) = List1.List(List1.ListIndex - 1)
        List1.List(List1.ListIndex - 1) = strLine
        List1.ListIndex = List1.ListIndex - 1
    End If

End Sub

Private Sub Command2_Click()

    If List1.ListIndex < (List1.ListCount - 1) Then
        Dim strLine As String
        strLine = List1.List(List1.ListIndex)
        List1.List(List1.ListIndex) = List1.List(List1.ListIndex + 1)
        List1.List(List1.ListIndex + 1) = strLine
        List1.ListIndex = List1.ListIndex + 1
    End If

End Sub

Private Sub Form_Load()

    RefreshList

End Sub

Private Sub OKButton_Click()

   Dim c As VBComponent
   Dim p As VBProject
   Dim vbc As VBControl
   Dim vbf As VBForm
   Dim sc As String
   Dim sp As String
   Dim t As Integer
   
    Dim ctlIndex As Long
    Dim ctlName As String
    Dim iPos As Integer
    Dim vbcc As VBControl
   
   Screen.MousePointer = vbHourglass
   
    sp = VBInstance.ActiveVBProject.Name
    sc = VBInstance.SelectedVBComponent.Name
    
    If sc <> "" And sp <> "" Then
        Set p = VBInstance.VBProjects.Item(sp)
        Set c = p.VBComponents.Item(sc)
        If c.Type = vbext_ct_VBForm Then
            c.Activate
            Set vbf = c.Designer
            For t = 0 To List1.ListCount - 1
                
                iPos = InStr(List1.List(t), "(")
                If iPos > 0 Then
                    ctlIndex = Mid(List1.List(t), iPos + 1, Len(List1.List(t)) - (iPos + 1))
                    ctlName = Left(List1.List(t), iPos - 1)
                    For Each vbcc In vbf.VBControls
                        If vbcc.Properties("index") = ctlIndex And vbcc.Properties("name") = ctlName Then
                            Set vbc = vbcc
                            Exit For
                        End If
                    Next vbcc
                Else
                    Set vbc = vbf.VBControls(List1.List(t))
                End If
                vbc.Properties("tabindex") = t
            Next t
        End If
    
    
    End If
    Screen.MousePointer = vbDefault
    
    Unload Me

End Sub
