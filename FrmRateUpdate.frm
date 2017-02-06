VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRateUpdate 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Rate Update"
   ClientHeight    =   2955
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "FrmRateUpdate.frx":0000
   ScaleHeight     =   2955
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   2460
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   873
      SimpleText      =   "s"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Generate"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "FrmRateUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mgroup As String, mstore As String
    Dim mitemcode As String, i As Double
    Dim mbranch As String, myearcode As String
    Dim mrow As Double, mrowsel As Integer
    Dim mstdate As String, meddate As String
    Public cnt As New ADODB.Connection
    Public stSQL As String
    Public oExcel As New Application
    Public oBook As Workbook
    Public oSheet As Worksheet
    Public rnStart As Range
    Public stADO As String, AddonRate As String

Private Sub Command1_Click()
    Dim rsstock As New ADODB.Recordset, rspurch As New ADODB.Recordset, rsitem As New ADODB.Recordset
    Dim mcode As String, mname As String, mrate As Double, mVal1 As Double, mval As Double, mqty As Double, mtotval As Double, mtotval1 As Double
    Dim msqlstr As String, mtotqty As Double
    
    If findvalue("select name from sysobjects where name='tempstock'") <> "" Then cnnData.Execute "drop table tempstock"
    If findvalue("select name from sysobjects where name='temppurch'") <> "" Then cnnData.Execute "drop table temppurch"
        
    rsstock.CursorLocation = adUseClient
    rsstock.Open "Select Code,name from Itemmaster Where Flag = 'L'", cnnData, adOpenDynamic, adLockReadOnly
    
    If rsstock.EOF <> True Then
        Do While rsstock.EOF <> True
            StatusBar1.SimpleText = "Update Rate for " + rsstock!Code + "    " + rsstock!Name + ""
            mcode = rsstock!Code
            mname = findvalue("Select Code From Itemdetail Where Code='" + mcode + "'")
            If mname <> "" Then
                cnnData.Execute "Update ItemDetail Set PurchRate = Isnull((Select Top 1 Rate From Stock Where Code = ItemDetail.Code And Subtype = 'RP' Order By Docdate Desc),0),CostRate = Isnull((Select Top 1 Rate From Stock Where Code = ItemDetail.Code And Subtype = 'RP' Order By Docdate Desc),0) Where Code = '" + mcode + "'"
                cnnData.Execute "Update ItemDetail set LastPurchRate = Isnull((Select Top 1 Rate From Stock Where Code = ItemDetail.Code And Subtype = 'RP' Order By Docdate Desc),0),LastCostRate = Isnull((Select Top 1 Rate From Stock Where Code = ItemDetail.Code And Subtype = 'RP' Order By Docdate Desc),0)  Where Code = '" + mcode + "'"
                cnnData.Execute "Update ItemDetail Set LASTVALUE  = Isnull((Select Top 1 Rate From Stock Where Code = ItemDetail.Code And Subtype in ('RS','XS') Order By Docdate Desc),0)   Where Code = '" + mcode + "'"
            Else
                Dim MRitesh As New ADODB.Recordset
                Dim mLastpurch As String, mLastSales As String
                MRitesh.Open "Select * from ItemDetail Where 1 =2", cnnData, adOpenDynamic, adLockOptimistic
                With MRitesh
                    .AddNew
                    NullValueChange MRitesh
                    !Code = mcode
                    mLastpurch = findvalue("Select Top 1 Isnull(Rate,0) From Stock Where Code = '" + mcode + "' And Subtype = 'RP' Order By Docdate Desc")
                    mLastSales = findvalue("Select Top 1 isnull(Rate,0) From Stock Where Code = '" + mcode + "' And Subtype = 'RS' Order By Docdate Desc")
                    !PurchRate = Val(mLastpurch)
                    !Costrate = Val(mLastpurch)
                    !LastPurchRate = Val(mLastpurch)
                    !LastCostRate = Val(mLastpurch)
                    !Lastvalue = Val(mLastSales)
                    .Update
                    .Close
                End With
            End If
            rsstock.MoveNext
        Loop
    End If
    rsstock.Close
    Unload Me
End Sub

Public Function NullValueChange(xdatabase As ADODB.Recordset)
    Dim i As Integer, mTouchValue As Double
    'mTouchValue = GetTouchValue
    On Error Resume Next
    For i = 0 To xdatabase.Fields.Count - 1
        If xdatabase.Fields(i).Type = 200 Then 'varchar
            xdatabase.Fields(i) = ""
        ElseIf xdatabase.Fields(i).Type = 6 Then 'money
            xdatabase.Fields(i) = 0
        ElseIf xdatabase.Fields(i).Type = 5 Then 'float
            xdatabase.Fields(i) = 0
        ElseIf xdatabase.Fields(i).Type = 3 Then 'int
            xdatabase.Fields(i) = 0
        ElseIf xdatabase.Fields(i).Type = 135 Then 'datetime
            xdatabase.Fields(i) = "01/01/1900"
        ElseIf xdatabase.Fields(i).Type = 131 Then 'numeric
            xdatabase.Fields(i) = 0
        ElseIf xdatabase.Fields(i).Type = 201 Then 'text
            xdatabase.Fields(i) = ""
        ElseIf xdatabase.Fields(i).Type = 11 Then 'bit
            xdatabase.Fields(i) = 0
        ElseIf xdatabase.Fields(i).Type = 2 Then 'smallint
            xdatabase.Fields(i) = 0
        ElseIf xdatabase.Fields(i).Type = 17 Then 'tinyint
            xdatabase.Fields(i) = 0
        End If
        
'        If UCase(xdatabase.Fields(i).Name) = "HWSERIAL" Then xdatabase.Fields(i).Value = gHWSerial
'        If UCase(xdatabase.Fields(i).Name) = "HWSERIAL2" Then xdatabase.Fields(i).Value = gHWSerial
'        If UCase(xdatabase.Fields(i).Name) = "AUTHIDS" Then xdatabase.Fields(i).Value = gUser
'        If UCase(xdatabase.Fields(i).Name) = "AUTHORISE" Then xdatabase.Fields(i).Value = "A00"
'        If UCase(xdatabase.Fields(i).Name) = "ENTEREDBY" Then xdatabase.Fields(i).Value = gUser
'        If UCase(xdatabase.Fields(i).Name) = "BRANCH" Then xdatabase.Fields(i).Value = gBranch
'        If UCase(xdatabase.Fields(i).Name) = "TOUCHVALUE" Then xdatabase.Fields(i).Value = mTouchValue
    Next
End Function

