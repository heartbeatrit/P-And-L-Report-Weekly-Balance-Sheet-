VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{E644C91B-78F7-4D86-8316-181A41A236A9}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmStockVal 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   9705
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":0ECA
   ScaleHeight     =   8445
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtSearch 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2400
      TabIndex        =   21
      Top             =   6480
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.OptionButton optInAnyWord 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Any Where"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "MM"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton optInWholeWord 
      BackColor       =   &H00E0E0E0&
      Caption         =   "By First Character"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "MM"
      Top             =   6480
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox CmbOption 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "Form1.frx":101CA
      Left            =   1440
      List            =   "Form1.frx":101D7
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   8640
      Width           =   1215
   End
   Begin VB.TextBox TxtDataBase 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8520
      TabIndex        =   14
      Top             =   -1200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox TxtServer 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8520
      TabIndex        =   13
      Top             =   -1560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin ProjetXPButton.XPButton cmdgroupselect 
      Height          =   375
      Left            =   7920
      TabIndex        =   11
      Top             =   6960
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Select All"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ProjetXPButton.XPButton CmdReport 
      Height          =   495
      Left            =   7080
      TabIndex        =   10
      Top             =   7680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "&Generate"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridgroup 
      Height          =   3375
      Left            =   1560
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5953
      _Version        =   393216
      RowHeightMin    =   350
      BackColorBkg    =   -2147483634
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker StartDate 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   53805057
      CurrentDate     =   40157
   End
   Begin VB.ComboBox CmbBranch 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1920
      Width           =   4215
   End
   Begin MSComCtl2.DTPicker EndDate 
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   53805057
      CurrentDate     =   40157
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridstore 
      Height          =   1335
      Left            =   5760
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2355
      _Version        =   393216
      BackColorBkg    =   -2147483634
      AllowUserResizing=   3
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin ProjetXPButton.XPButton cmdstoreselect 
      Height          =   375
      Left            =   5880
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Select All"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblmonthly 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   1560
      TabIndex        =   22
      Top             =   6480
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Option"
      Height          =   255
      Left            =   600
      TabIndex        =   18
      Top             =   8640
      Width           =   615
   End
   Begin VB.Image ImageSelect 
      Height          =   480
      Left            =   600
      Picture         =   "Form1.frx":101FE
      Top             =   8580
      Width           =   480
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Server"
      Height          =   255
      Left            =   7560
      TabIndex        =   16
      Top             =   -1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "DataBase"
      Height          =   255
      Left            =   7560
      TabIndex        =   15
      Top             =   -1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Script "
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   -360
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Store"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   -3120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Branch"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1920
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   3360
      Picture         =   "Form1.frx":11040
      Top             =   240
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1110
      Left            =   0
      Picture         =   "Form1.frx":155F9
      Top             =   240
      Width           =   3360
   End
End
Attribute VB_Name = "frmStockVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mGroup As String, mstore As String
    Dim mitemcode As String, i As Double
    Dim mbranch As String, myearcode As String
    Dim mrow As Double, mrowsel As Integer
    Dim mstdate As String, meddate As String
    'Dim oExcel As Object
    'Dim oBook As Object
   
    Public cnt As New ADODB.Connection
    Public stSQL As String
    Public oExcel As New Application
    Public oBook As Workbook
    Public oSheet As Worksheet, Osheet2 As Worksheet
    Public rnStart As Range
    Public stADO As String, AddonRate As String
Public Function gridselectunselect(xgrid As MSHFlexGrid, xcmdbutton As XPButton, xcol As Integer, xpicture As Image)
    Dim i As Integer
    For i = xgrid.FixedRows To xgrid.Rows - 1
        xgrid.Col = xcol
        xgrid.Row = i
        If xcmdbutton.Caption = "&Select All" Then
            Set xgrid.CellPicture = xpicture.Picture
        Else
            Set xgrid.CellPicture = LoadPicture()
        End If
    Next
End Function

Private Sub CmbBranch_LostFocus()
    Dim rsstore As New ADODB.Recordset, mb As String
    rsstore.CursorLocation = adUseClient
    mb = Mid(CmbBranch.Text, 1, InStr(1, CmbBranch.Text, "-") - 1)
    
    rsstore.Open "select code,name from stores Where branch = '" + mb + "' group by code,name", cnnData, adOpenStatic, adLockReadOnly, adCmdText
    If rsstore.RecordCount > 0 Then
        Set gridstore.DataSource = rsstore
        gridstore.ColWidth(0) = 500
        gridstore.ColWidth(1) = 0
        gridstore.ColWidth(2) = 4000
    End If
    rsstore.Close
    cmdstoreselect_Click
End Sub

Private Sub cmdgroupselect_Click()
    gridselectunselect gridgroup, cmdgroupselect, 0, ImageSelect
    If cmdgroupselect.Caption = "&Select All" Then
        cmdgroupselect.Caption = "&Unselect All"
    Else
        cmdgroupselect.Caption = "&Select All"
    End If
End Sub

Private Sub CmdReport_Click()
    Dim ActiveSheet As Object
    Dim rsstock As New ADODB.Recordset, rspurch As New ADODB.Recordset, rsitem As New ADODB.Recordset
    Dim mcode As String, mname As String, mrate As Double, mval As Double, mqty As Double, mtotval As Double
    Dim msqlstr As String, mtotqty As Double, mxqty As Double, mday As Integer, mdaydiff As Integer
    Dim mtot(7) As Currency
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Visible = True
    'Set oBook = oExcel.Workbooks.Add
    'Add data to cells of the first worksheet in the new workbook
    Set oBook = oExcel.Workbooks.Add
    Set oSheet = oBook.Worksheets(1)
    'mstdate = Format(StartDate.Value, "mm/dd/yyyy")
    'meddate = Format(EndDate.Value, "mm/dd/yyyy")
    mstdate = Format(StartDate.Value, "dd-mmm-yyyy")
    meddate = Format(EndDate.Value, "dd-mmm-yyyy")
    'Excel.Sheets(1).Delete
    If DateValue(StartDate.Value) > DateValue(EndDate.Value) Then
        MsgBox "Starting date not accept greater then ending date", vbOKOnly + vbInformation
        Exit Sub
    End If
    If CmbBranch.Text = "" Then
        MsgBox "No branch selected", vbOKOnly + vbInformation
        Exit Sub
    End If
    mbranch = Mid(CmbBranch.Text, 1, InStr(1, CmbBranch.Text, "-") - 1)
    groupcollect
    If mGroup = "" Then
        MsgBox "Dear User, No Product Group Selected", vbOKOnly + vbInformation
        Exit Sub
    End If
    storecollect
    If mstore = "" Then
        MsgBox "Dear User, No Store Selected", vbOKOnly + vbInformation
        Exit Sub
    End If
    reportgenerate
End Sub

Private Sub cmdstoreselect_Click()
    gridselectunselect gridstore, cmdstoreselect, 0, ImageSelect
    If cmdstoreselect.Caption = "&Select All" Then
        cmdstoreselect.Caption = "&Unselect All"
    Else
        cmdstoreselect.Caption = "&Select All"
    End If
End Sub

Private Sub gridgroup_dblClick()
    gridgroup.Row = gridgroup.RowSel
    gridgroup.Col = 0
    If gridgroup.CellPicture.Type = 0 Then
        Set gridgroup.CellPicture = ImageSelect.Picture
    Else
        Set gridgroup.CellPicture = LoadPicture()
    End If
End Sub

Private Sub gridstore_Click()
    gridstore.Row = gridstore.RowSel
    gridstore.Col = 0
    If gridstore.CellPicture.Type = 0 Then
        Set gridstore.CellPicture = ImageSelect.Picture
    Else
        Set gridstore.CellPicture = LoadPicture()
    End If
End Sub

Private Sub Form_Load()
    'Call Main
    'TxtServer.Text = sname
    'TxtDataBase.Text = Nature + "001"
    'On Error GoTo err:
    'cnnData.Open "Provider=SQLOLEDB.1;User Id=sa;Password=;server=" & TxtServer.Text & ";Initial Catalog=" & TxtDataBase.Text & ""
    'cnnData.Open "Provider=SQLOLEDB.1;User Id=sa;Password=;server=" & gSQLServer & ";Initial Catalog=" & gNature001 & ""
    'cnnData.CommandTimeout = 1000
    
    'ActiveSheet.Rows.Clear
    frmStockVal.Caption = UCase(gNature001)
    StartDate.Value = Date
    EndDate.Value = Date
    
    'Me.Caption = "+ gDataBaseGen +"
    
    Dim rsgroup As New ADODB.Recordset
    rsgroup.CursorLocation = adUseClient
    Dim mGrp As String
    mGrp = findvalue("Select Code From Itemmaster Where COde = '100061'")
    
    
    If UCase(frmStockVal.Caption) = "TFATDEMO" Then
        rsgroup.Open "select code,name from itemmaster where flag='L' And GRP in ('100062','100019','100065','100052','100056','100053','100064') order by name", cnnData, adOpenStatic, adLockReadOnly, adCmdText
    Else
        rsgroup.Open "select code,name from itemmaster where flag='L' And Code Not in (Select Code from Itemmaster i where GRP in  ('" + mGrp + "')) order by name", cnnData, adOpenStatic, adLockReadOnly, adCmdText
    End If
    If rsgroup.RecordCount > 0 Then
        Set gridgroup.DataSource = rsgroup
        gridgroup.ColWidth(0) = 500
        gridgroup.ColWidth(1) = 2000
        gridgroup.ColWidth(2) = 4500
    End If
    rsgroup.Close
    
     
'    rsgroup.Open "select code,name from Stores order by name", cnnData, adOpenStatic, adLockReadOnly, adCmdText
'    If rsgroup.RecordCount > 0 Then
'        Set gridstore.DataSource = rsgroup
'        gridstore.ColWidth(0) = 500
'        gridstore.ColWidth(1) = 2000
'        gridstore.ColWidth(2) = 4500
'    End If
'    rsgroup.Close
    
    Dim rscnn As New ADODB.Recordset
    rscnn.CursorLocation = adUseClient
    rscnn.Open "Select code+'-'+name as name from " + gDataBaseGen + ".dbo.tfatbranch where compcode = '001' order by code", cnnData, adOpenStatic, adLockReadOnly, adCmdText
    If rscnn.RecordCount > 0 Then
        CmbBranch.Clear
        Do Until rscnn.EOF
            CmbBranch.AddItem rscnn!Name
            rscnn.MoveNext
        Loop
    End If
    rscnn.Close
CmbBranch.ListIndex = 0
CmbBranch.Text = "HO0000-"
'CmbBranch_LostFocus

'cmdstoreselect_Click
cmdgroupselect_Click
End Sub
Private Function groupcollect()
    mGroup = ""
    For i = 1 To gridgroup.Rows - 1
        gridgroup.Col = 0
        gridgroup.Row = i
        If gridgroup.CellPicture.Type <> 0 Then
           mGroup = mGroup + "'" + gridgroup.TextMatrix(i, 1) + "',"
        End If
    Next
    If mGroup <> "" Then mGroup = Mid(mGroup, 1, Len(mGroup) - 1)
End Function

Private Function storecollect()
    mstore = ""
    For i = 1 To gridstore.Rows - 1
        gridstore.Col = 0
        gridstore.Row = i
        If gridstore.CellPicture.Type <> 0 Then
            mstore = mstore + "'" + gridstore.TextMatrix(i, 1) + "',"
        End If
    Next
    If mstore <> "" Then mstore = Mid(mstore, 1, Len(mstore) - 1)
End Function

Private Function reportgenerate()
    Dim rsstock As New ADODB.Recordset, rspurch As New ADODB.Recordset, rsitem As New ADODB.Recordset
    Dim mcode As String, mname As String, mrate As Double, mVal1 As Double, mval As Double, mqty As Double, mtotval As Double, mtotval1 As Double
    Dim msqlstr As String, mtotqty As Double
    Dim mtotalPurchase As Double, mtotalIPX As Double
    Dim mGrandTotal As Double
    Dim mAbcGroup As Double, mOthergroup As Double
    
    If findvalue("select name from sysobjects where name='tempstock'") <> "" Then cnnData.Execute "drop table tempstock"
    If findvalue("select name from sysobjects where name='temppurch'") <> "" Then cnnData.Execute "drop table temppurch"
    If findvalue("select name from sysobjects where name='PStock'") <> "" Then cnnData.Execute "drop table PStock"
    
   
    
    If UCase(frmStockVal.Caption) = "TFATDEMO" Then
        msqlstr = "select * into tempstock from (Select Itemmaster.GRP,Itemmaster.Code,Isnull((select Sum(OpQty) from ItOpen Where Branch = '" + mbranch + "'  And Code =itemmaster.Code And Notinstock = 0 And Left(authorise,1)='A'),0) + Isnull(sum(qty),0) as stock,ISNULL((Select Fiforate from itemdetail where branch ='" + mbranch + "' And Code = itemmaster.Code),0) as Fiforate  " _
                & " from Itemmaster Left Outer Join stock on Stock.Code = itemmaster.code And stock.Branch = '" + mbranch + "' and docdate<='" + meddate + "' and  Stock.store in (" + mstore + ") and notinstock=0 and left(stock.authorise,1)='A'" _
                & " Where itemmaster.Code in (Select Code from itemmaster i where grp in ('100067','100066','100052','100062','100019','100065','100052','100056','100053','100064','100069')) " _
                & " group by Stock.Branch,Itemmaster.Code,Itemmaster.GRP Having Isnull((select Sum(OpQty) from ItOpen Where Branch = '" + mbranch + "'  And Code =itemmaster.Code And Notinstock = 0 And Left(authorise,1)='A'),0) + Isnull(sum(qty),0) >0) as a"
    Else
        msqlstr = "select * into tempstock from (select Itemmaster.GRP,Itemmaster.Code,Isnull((select Sum(OpQty) from ItOpen Where Branch = '" + mbranch + "'  And Code =itemmaster.Code And Notinstock = 0 And Left(authorise,1)='A'),0) + Isnull(sum(qty),0) as stock,ISNULL((Select Fiforate from itemdetail where branch ='" + mbranch + "' And Code = itemmaster.Code),0) as Fiforate  " _
                & " from Itemmaster Left Outer Join stock on Stock.Code = itemmaster.code And stock.Branch = '" + mbranch + "' and docdate<='" + meddate + "' and  Stock.store in (" + mstore + ") and notinstock=0 and left(stock.authorise,1)='A'" _
                & " Where itemmaster.Code in (" + mGroup + ") " _
                & " group by Stock.Branch,Itemmaster.Code,Itemmaster.GRP  Having Isnull((select Sum(OpQty) from ItOpen Where Branch = '" + mbranch + "'  And Code =itemmaster.Code And Notinstock = 0 And Left(authorise,1)='A'),0) + Isnull(sum(qty),0) >0) as a"
    End If
    
    cnnData.CommandTimeout = 3000
    cnnData.Execute msqlstr
    
    
    mrow = 5
    
    If UCase(frmStockVal.Caption) = "TFATDEMO" Then
        msqlstr = "Select * into temppurch from (select type,prefix,srl,sno,subtype,code,qty,rate,NewRate,amt,docdate,touchvalue,TaxAmt,Addtax,Disc,DiscAmt from stock where  stock.Branch = '" + mbranch + "' and code in (select Code From Itemmaster i Where Grp in ('100067','100066','100052','100062','100019','100065','100052','100056','100053','100064')) and subtype in ('RP','IC','IA') And docdate<='" + meddate + "' and left(authorise,1)='A' and qty>0 and notinstock=0) as a"
    Else
        msqlstr = "Select * into temppurch from (select type,prefix,srl,sno,subtype,code,qty,rate,NewRate,amt,docdate,touchvalue,TaxAmt,Addtax,Disc,DiscAmt from stock where  stock.Branch = '" + mbranch + "' and code in (" + mGroup + ") and subtype in ('RP','IC','IA') And docdate<='" + meddate + "' and left(authorise,1)='A' and qty>0 and notinstock=0) as a"
    End If
    cnnData.CommandTimeout = 3000
    cnnData.Execute msqlstr
    
    cnnData.Execute "Create Index TPurch_0 On TempPurch (type,prefix,sno,srl)"
    
    
    cnnData.CommandTimeout = 3000
    If UCase(frmStockVal.Caption) = "TFATDEMO" Then
        cnnData.Execute "Select type,prefix,srl,sno,subtype,code,qty,rate,amt,docdate,touchvalue,Left(chlnnumber,3) as CType,substring(chlnnumber,4,8) as CPrefix,substring(chlnnumber,12,5) as CSno,Right(chlnnumber,6) as CSrl Into PStock From Stock Where  stock.Branch = '" + mbranch + "' And stock.docdate<='" + meddate + "'  and Code in (select code from itemmaster  i where grp in ('100067','100066','100052','100062','100019','100065','100052','100056','100053','100064'))"
    Else
        cnnData.Execute "Select type,prefix,srl,sno,subtype,code,qty,rate,amt,docdate,touchvalue,Left(chlnnumber,3) as CType,substring(chlnnumber,4,8) as CPrefix,substring(chlnnumber,12,5) as CSno,Right(chlnnumber,6) as CSrl Into PStock From Stock Where  stock.Branch = '" + mbranch + "' And stock.docdate<='" + meddate + "'  and Code in (" + mGroup + ")"
    End If
    cnnData.Execute "Create Index PStock_0 On PStock (Ctype,Cprefix,Csno,Csrl)"
    
    cnnData.Execute "Alter Table Temppurch add PurRate varchar(1)"
    cnnData.Execute "Update TempPurch set PurRate=0"
   
    Dim mIRate As String, mQtyIPX As Double, IPXRate As Double, PurChaseRate As Double, mAddonRate As Double
    Dim OldIpxrate As Double, OldPurchaserate As Double, OldAddonRate As Double
    
    rsstock.CursorLocation = adUseClient
    
    
    rsstock.Open "Select * from Tempstock Where stock>0 Order by (select grp from Itemmaster i Where code = tempstock.code) ,Code", cnnData, adOpenStatic, adLockReadOnly, adCmdText
     If findvalue("select name from sysobjects where name='TmpStkVal'") <> "" Then
        cnnData.Execute "drop table TmpStkVal"
    End If
        
    msqlstr = "Create Table TmpStkVal (Grp varchar(254),Code varchar(100),Name Varchar(254),Stock Money,FifoRate MOney,FifoValue Money)    "
    cnnData.Execute msqlstr
   
    If rsstock.RecordCount > 0 Then
        rsstock.MoveFirst
        Do Until rsstock.EOF
            If rsstock!Stock <> 0 Then
                mcode = rsstock!Code
                mname = findvalue("Select name from Itemmaster Where code = '" + rsstock!Code + "'")
                mqty = 0
                mQtyIPX = 0
                mval = 0
                mVal1 = 0
                PurChaseRate = 0
                OldPurchaserate = 0
                OldIpxrate = 0
                IPXRate = 0
                AddonRate = 0
                OldAddonRate = 0
                rspurch.CursorLocation = adUseClient
'                If mcode = "500440" Then
'                MsgBox "ok"
'                End If
                rspurch.Open "Select code,qty,rate,newrate,DiscAmt,amt,purrate,Subtype from Temppurch where code='" + mcode + "' order by  docdate desc,Touchvalue,Srl Desc,sno DESC", cnnData, adOpenStatic, adLockReadOnly, adCmdText
             'AddonRate = Val(findvalue("Select IsNull(aVG(NewRate),0) From ItOpen Where Branch = '" + mbranch + "' And code='" + mcode + "' And Notinstock =0 And Left(authorise,1) ='A'"))
            If rspurch.RecordCount > 0 Then
                    rspurch.MoveFirst
                    Do Until rspurch.EOF
                    'Modify by ritesh on 06-Feb-2013 for Care office
                    If (mQtyIPX + mqty + rspurch!Qty) <= rsstock!Stock Then
                            If Val(rspurch!Rate) <> 0 Then
                                    If CmbOption.Text = "" Then
                                        mval = mval + Round(rspurch!Qty * rspurch!Rate, 2)
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        mval = mval + Round(rspurch!Qty * rspurch!NewRate, 2)
                                    Else
                                        mval = mval + Round(rspurch!Amt - rspurch!DiscAmt, 2)
                                    End If
                                mqty = mqty + rspurch!Qty
                                If Val(PurChaseRate) = 0 And rspurch!Rate <> 0 Then
                                    If CmbOption.Text = "" Then
                                        PurChaseRate = rspurch!Rate
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        PurChaseRate = rspurch!NewRate
                                    Else
                                        PurChaseRate = Round((rspurch!Qty * rspurch!Rate) - rspurch!DiscAmt, 3)
                                    End If
                                    OldPurchaserate = PurChaseRate
                                ElseIf Val(PurChaseRate) <> 0 And Val(OldPurchaserate) <> Val(PurChaseRate) And rspurch!Rate <> 0 Then
                                    If CmbOption.Text = "" Then
                                        PurChaseRate = rspurch!Rate
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        PurChaseRate = rspurch!NewRate
                                    Else
                                        PurChaseRate = Round(rspurch!Rate - rspurch!DiscAmt, 3)
                                    End If
                                    OldPurchaserate = PurChaseRate
                                ElseIf Val(PurChaseRate) <> 0 And Val(OldPurchaserate) = Val(PurChaseRate) And rspurch!Rate <> 0 Then
                                    If CmbOption.Text = "" Then
                                        PurChaseRate = rspurch!Rate
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        PurChaseRate = rspurch!NewRate
                                    Else
                                        PurChaseRate = Round(rspurch!Rate - rspurch!DiscAmt, 3)
                                    End If
                                    OldPurchaserate = PurChaseRate
                                End If
                            End If
                            mIRate = rspurch!Purrate
                        ElseIf (mQtyIPX + mqty + rspurch!Qty) > rsstock!Stock Then
                            If Val(rspurch!Rate) <> 0 Then
                                    If CmbOption.Text = "" Then
                                        mval = mval + Round((rsstock!Stock - (mqty + mQtyIPX)) * rspurch!Rate, 2)
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        mval = mval + Round((rsstock!Stock - (mqty + mQtyIPX)) * rspurch!NewRate, 2)
                                    Else
                                        mval = mval + Round(((rsstock!Stock - (mqty + mQtyIPX)) * rspurch!Rate) - rspurch!DiscAmt, 2)
                                    End If
                                
                                mqty = rsstock!Stock - mQtyIPX
                                If Val(PurChaseRate) = 0 And Val(rspurch!Rate) <> 0 Then
                                    If CmbOption.Text = "" Then
                                        PurChaseRate = rspurch!Rate
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        PurChaseRate = rspurch!NewRate
                                    Else
                                        PurChaseRate = rspurch!Rate - rspurch!DiscAmt
                                    End If
                                    OldPurchaserate = PurChaseRate
                                ElseIf Val(PurChaseRate) <> 0 And Val(OldPurchaserate) <> Val(PurChaseRate) Then
                                    If CmbOption.Text = "" Then
                                        PurChaseRate = rspurch!Rate
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        PurChaseRate = rspurch!NewRate
                                    Else
                                        PurChaseRate = rspurch!Rate - rspurch!DiscAmt
                                    End If
                                    'PurChaseRate = rspurch!Rate
                                    OldPurchaserate = PurChaseRate
                                ElseIf Val(PurChaseRate) <> 0 And Val(OldPurchaserate) = Val(PurChaseRate) And rspurch!Rate <> 0 Then
                                    If CmbOption.Text = "" Then
                                        PurChaseRate = rspurch!Rate
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        PurChaseRate = rspurch!NewRate
                                    Else
                                        PurChaseRate = rspurch!Rate - rspurch!DiscAmt
                                    End If
                                    'PurChaseRate = rspurch!Rate
                                    OldPurchaserate = PurChaseRate
                                End If
                            End If
                            mIRate = rspurch!Purrate
                        End If
                        If mqty + mQtyIPX >= rsstock!Stock Then rspurch.MoveLast
                        rspurch.MoveNext
                    Loop
            Else
                Dim Rspurch1 As New ADODB.Recordset
                Rspurch1.CursorLocation = adUseClient
                Rspurch1.Open "Select code,OpQty as qty,oprate as rate,newrate,opvalue as amt,0 as purrate,Subtype from ItOpen where Branch = '" + mbranch + "' And code='" + mcode + "' And OpQty<>0  order by  docdate desc,touchvalue desc", cnnData, adOpenStatic, adLockReadOnly, adCmdText
                mqty = 0
                mQtyIPX = 0
                mval = 0
                mVal1 = 0
                PurChaseRate = 0
                OldPurchaserate = 0
                OldIpxrate = 0
                IPXRate = 0
                AddonRate = 0
                OldAddonRate = 0
                If Rspurch1.RecordCount > 0 Then
                    Rspurch1.MoveFirst
                    Do Until Rspurch1.EOF
                    If (mQtyIPX + mqty + Rspurch1!Qty) <= rsstock!Stock Then
                            If Val(Rspurch1!Rate) <> 0 Then
                                    If CmbOption.Text = "" Then
                                        mval = mval + Round(Rspurch1!Qty * Rspurch1!Rate, 2)
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        mval = mval + Round(Rspurch1!Qty * Rspurch1!NewRate, 2)
                                    Else
                                        mval = mval + Round(Rspurch1!Rate - Rspurch1!DiscAmt, 2)
                                    End If
                                mqty = mqty + Rspurch1!Qty
                                If Val(PurChaseRate) = 0 And Rspurch1!Rate <> 0 Then
                                    If CmbOption.Text = "" Then
                                        PurChaseRate = Rspurch1!Rate
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        PurChaseRate = Rspurch1!NewRate
                                    Else
                                        PurChaseRate = Round((Rspurch1!Qty * Rspurch1!Rate) - Rspurch1!DiscAmt, 3)
                                    End If
                                    OldPurchaserate = PurChaseRate
                                ElseIf Val(PurChaseRate) <> 0 And Val(OldPurchaserate) <> Val(PurChaseRate) And Rspurch1!Rate <> 0 Then
                                    If CmbOption.Text = "" Then
                                        PurChaseRate = Rspurch1!Rate
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        PurChaseRate = Rspurch1!NewRate
                                    Else
                                        PurChaseRate = Round(Rspurch1!Rate - Rspurch1!DiscAmt, 3)
                                    End If
                                    OldPurchaserate = PurChaseRate
                                ElseIf Val(PurChaseRate) <> 0 And Val(OldPurchaserate) = Val(PurChaseRate) And Rspurch1!Rate <> 0 Then
                                    If CmbOption.Text = "" Then
                                        PurChaseRate = Rspurch1!Rate
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        PurChaseRate = Rspurch1!NewRate
                                    Else
                                        PurChaseRate = Round(Rspurch1!Rate - Rspurch1!DiscAmt, 3)
                                    End If
                                    OldPurchaserate = PurChaseRate
                                End If
                            End If
                            mIRate = Rspurch1!Purrate
                        ElseIf (mQtyIPX + mqty + Rspurch1!Qty) > rsstock!Stock Then
                            If Val(Rspurch1!Rate) <> 0 Then
                                    If CmbOption.Text = "" Then
                                        mval = mval + Round((rsstock!Stock - (mqty + mQtyIPX)) * Rspurch1!Rate, 2)
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        mval = mval + Round((rsstock!Stock - (mqty + mQtyIPX)) * Rspurch1!NewRate, 2)
                                    Else
                                        mval = mval + Round(((rsstock!Stock - (mqty + mQtyIPX)) * Rspurch1!Rate) - Rspurch1!DiscAmt, 2)
                                    End If
                                
                                mqty = rsstock!Stock - mQtyIPX
                                If Val(PurChaseRate) = 0 And Val(Rspurch1!Rate) <> 0 Then
                                    If CmbOption.Text = "" Then
                                        PurChaseRate = Rspurch1!Rate
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        PurChaseRate = Rspurch1!NewRate
                                    Else
                                        PurChaseRate = Rspurch1!Rate - Rspurch1!DiscAmt
                                    End If
                                    OldPurchaserate = PurChaseRate
                                ElseIf Val(PurChaseRate) <> 0 And Val(OldPurchaserate) <> Val(PurChaseRate) Then
                                    If CmbOption.Text = "" Then
                                        PurChaseRate = Rspurch1!Rate
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        PurChaseRate = Rspurch1!NewRate
                                    Else
                                        PurChaseRate = Rspurch1!Rate - Rspurch1!DiscAmt
                                    End If
                                    'PurChaseRate = rsPurch1!Rate
                                    OldPurchaserate = PurChaseRate
                                ElseIf Val(PurChaseRate) <> 0 And Val(OldPurchaserate) = Val(PurChaseRate) And Rspurch1!Rate <> 0 Then
                                    If CmbOption.Text = "" Then
                                        PurChaseRate = Rspurch1!Rate
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        PurChaseRate = Rspurch1!NewRate
                                    Else
                                        PurChaseRate = Rspurch1!Rate - Rspurch1!DiscAmt
                                    End If
                                    'PurChaseRate = rsPurch1!Rate
                                    OldPurchaserate = PurChaseRate
                                End If
                            End If
                            mIRate = Rspurch1!Purrate
                        End If
                        If mqty + mQtyIPX >= rsstock!Stock Then Rspurch1.MoveLast
                        Rspurch1.MoveNext
                    Loop
                End If
                Rspurch1.Close
            End If
            rspurch.Close
            If mqty < rsstock!Stock Then
                Rspurch1.CursorLocation = adUseClient
                Rspurch1.Open "Select code,OpQty as qty,oprate as rate,newrate,opvalue as amt,0 as purrate,Subtype from ItOpen where Branch = '" + mbranch + "' And code='" + mcode + "' And OpQty<>0  order by  docdate desc,touchvalue", cnnData, adOpenStatic, adLockReadOnly, adCmdText
                If Rspurch1.RecordCount > 0 Then
                    Rspurch1.MoveFirst
                    Do Until Rspurch1.EOF
                    If (mQtyIPX + mqty + Rspurch1!Qty) <= rsstock!Stock Then
                            If Val(Rspurch1!Rate) <> 0 Then
                                    If CmbOption.Text = "" Then
                                        mval = mval + Round(Rspurch1!Qty * Rspurch1!Rate, 2)
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        mval = mval + Round(Rspurch1!Qty * Rspurch1!NewRate, 2)
                                    Else
                                        mval = mval + Round(Rspurch1!Amt - Rspurch1!DiscAmt, 2)
                                    End If
                                mqty = mqty + Rspurch1!Qty
                                If Val(PurChaseRate) = 0 And Rspurch1!Rate <> 0 Then
                                    If CmbOption.Text = "" Then
                                        PurChaseRate = Rspurch1!Rate
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        PurChaseRate = Rspurch1!NewRate
                                    Else
                                        PurChaseRate = Round((Rspurch1!Qty * Rspurch1!Rate) - Rspurch1!DiscAmt, 3)
                                    End If
                                    OldPurchaserate = PurChaseRate
                                ElseIf Val(PurChaseRate) <> 0 And Val(OldPurchaserate) <> Val(PurChaseRate) And Rspurch1!Rate <> 0 Then
                                    If CmbOption.Text = "" Then
                                        PurChaseRate = Rspurch1!Rate
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        PurChaseRate = Rspurch1!NewRate
                                    Else
                                        PurChaseRate = Round(Rspurch1!Rate - Rspurch1!DiscAmt, 3)
                                    End If
                                    OldPurchaserate = PurChaseRate
                                ElseIf Val(PurChaseRate) <> 0 And Val(OldPurchaserate) = Val(PurChaseRate) And Rspurch1!Rate <> 0 Then
                                    If CmbOption.Text = "" Then
                                        PurChaseRate = Rspurch1!Rate
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        PurChaseRate = Rspurch1!NewRate
                                    Else
                                        PurChaseRate = Round(Rspurch1!Rate - Rspurch1!DiscAmt, 3)
                                    End If
                                    OldPurchaserate = PurChaseRate
                                End If
                            End If
                            mIRate = Rspurch1!Purrate
                        ElseIf (mQtyIPX + mqty + Rspurch1!Qty) > rsstock!Stock Then
                            If Val(Rspurch1!Rate) <> 0 Then
                                    If CmbOption.Text = "" Then
                                        mval = mval + Round((rsstock!Stock - (mqty + mQtyIPX)) * Rspurch1!Rate, 2)
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        mval = mval + Round((rsstock!Stock - (mqty + mQtyIPX)) * Rspurch1!NewRate, 2)
                                    Else
                                        mval = mval + Round(((rsstock!Stock - (mqty + mQtyIPX)) * Rspurch1!Amt) - Rspurch1!DiscAmt, 2)
                                    End If
                                mqty = rsstock!Stock - mQtyIPX
                                If Val(PurChaseRate) = 0 And Val(Rspurch1!Rate) <> 0 Then
                                    If CmbOption.Text = "" Then
                                        PurChaseRate = Rspurch1!Rate
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        PurChaseRate = Rspurch1!NewRate
                                    Else
                                        PurChaseRate = Rspurch1!Rate - Rspurch1!DiscAmt
                                    End If
                                    OldPurchaserate = PurChaseRate
                                ElseIf Val(PurChaseRate) <> 0 And Val(OldPurchaserate) <> Val(PurChaseRate) Then
                                    If CmbOption.Text = "" Then
                                        PurChaseRate = Rspurch1!Rate
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        PurChaseRate = Rspurch1!NewRate
                                    Else
                                        PurChaseRate = Rspurch1!Rate - Rspurch1!DiscAmt
                                    End If
                                    'PurChaseRate = rsPurch1!Rate
                                    OldPurchaserate = PurChaseRate
                                ElseIf Val(PurChaseRate) <> 0 And Val(OldPurchaserate) = Val(PurChaseRate) And Rspurch1!Rate <> 0 Then
                                    If CmbOption.Text = "" Then
                                        PurChaseRate = Rspurch1!Rate
                                    ElseIf CmbOption.Text = "Cost Rate" Then
                                        PurChaseRate = Rspurch1!NewRate
                                    Else
                                        PurChaseRate = Rspurch1!Rate - Rspurch1!DiscAmt
                                    End If
                                    'PurChaseRate = rsPurch1!Rate
                                    OldPurchaserate = PurChaseRate
                                End If
                            End If
                            mIRate = Rspurch1!Purrate
                        End If
                        If mqty + mQtyIPX >= rsstock!Stock Then Rspurch1.MoveLast
                        Rspurch1.MoveNext
                    Loop
                End If
             Rspurch1.Close
            End If
       End If
           Dim mGrp As String
           mGrp = CStr(findvalue("Select Grp From Itemmaster Where Code = '" + mcode + "'"))
           
           If mval <> 0 And mqty <> 0 Then PurChaseRate = Round(mval / mqty, 2)
 
            msqlstr = "Insert into TmpStkVal  values ('" + mGrp + "','" + mcode + "','" + mname + "'," _
                            & "" & Round(mqty, 2) & "," & Round(PurChaseRate, 2) & "," & Round(mval, 2) & ")"
    
            cnnData.Execute msqlstr
       
         rsstock.MoveNext
        
        Loop
    End If
    rsstock.Close
'    Dim mIRate As String, mQtyIPX As Double, IPXRate As Double, PurChaseRate As Double, mAddonRate As Double
'    Dim OldIpxrate As Double, OldPurchaserate As Double, OldAddonRate As Double
'
    rsstock.CursorLocation = adUseClient
    rsstock.Open "Select * from TmpStkVal", cnnData, adOpenStatic, adLockReadOnly, adCmdText
    If rsstock.RecordCount > 0 Then
        rsstock.MoveFirst
        Do Until rsstock.EOF
            mGrandTotal = mGrandTotal + Val(rsstock!FifoValue)
            rsstock.MoveNext
        Loop
    End If
    rsstock.Close
    
    
    'A and A br group
    If findvalue("select name from sysobjects where name='tempstock1'") <> "" Then cnnData.Execute "drop table tempstock1"
    If findvalue("select name from sysobjects where name='temppurch1'") <> "" Then cnnData.Execute "drop table temppurch1"
    If findvalue("select name from sysobjects where name='PStock1'") <> "" Then cnnData.Execute "drop table PStock1"
    'mGroup = "'100062','100058'"
    If UCase(frmStockVal.Caption) = "TFATDEMO" Then
            msqlstr = "Select * into tempstock1 from (select * From TempStock Where Grp in ('100062','100069') as a"
    ElseIf UCase(frmStockVal.Caption) = "TFATSILLP" Then
            msqlstr = "select * into tempstock1 from (select * From TempStock where Grp = '100020') as a"
    ElseIf UCase(frmStockVal.Caption) = "TFATDSPL001" Then
            msqlstr = "select * into tempstock1 from (select * from TempStock where Grp = '100015') as a"
    ElseIf UCase(frmStockVal.Caption) = "TFATNMJ001" Then
            msqlstr = "Select * into tempstock1 from (select * From TempStock where Grp = '100003') as a "
    End If
   
    cnnData.CommandTimeout = 3000
    cnnData.Execute msqlstr
    
    mrow = 5
    rsstock.CursorLocation = adUseClient
    If UCase(frmStockVal.Caption) = "TFATDEMO" Then
        rsstock.Open "Select * from TmpStkVal Where Grp   in ('100062','100069')", cnnData, adOpenStatic, adLockReadOnly, adCmdText
    ElseIf UCase(frmStockVal.Caption) = "TFATSILLP" Then
        rsstock.Open "Select * from TmpStkVal Where Grp  ='100020'", cnnData, adOpenStatic, adLockReadOnly, adCmdText
    ElseIf UCase(frmStockVal.Caption) = "TFATDSPL001" Then
        rsstock.Open "Select * from TmpStkVal Where Grp  ='100015'", cnnData, adOpenStatic, adLockReadOnly, adCmdText
    ElseIf UCase(frmStockVal.Caption) = "TFATNMJ001" Then
        rsstock.Open "Select * from TmpStkVal Where Grp  ='100003'", cnnData, adOpenStatic, adLockReadOnly, adCmdText
    End If
    If rsstock.RecordCount > 0 Then
        rsstock.MoveFirst
        Do Until rsstock.EOF
            mAbcGroup = mAbcGroup + Val(rsstock!FifoValue)
            rsstock.MoveNext
        Loop
    End If
    rsstock.Close
    
    
    If UCase(frmStockVal.Caption) = "TFATDEMO" Then
        msqlstr = "Select * From TmpStkVal Where Grp in ('100062','100069') And Stock > 0 Order BY GRP,Code"
    ElseIf UCase(frmStockVal.Caption) = "TFATSILLP" Then
        msqlstr = "Select * From TmpStkVal Where  Grp = '100020' And  Stock > 0 Order BY GRP,Code"
    ElseIf UCase(frmStockVal.Caption) = "TFATDSPL001" Then
        msqlstr = "Select * From TmpStkVal Where  Grp = '100015' And Stock > 0 Order BY GRP,Code"
    ElseIf UCase(frmStockVal.Caption) = "TFATNMJ001" Then
        msqlstr = "Select * From TmpStkVal Where  Grp = '100003' And Stock > 0 Order BY GRP,Code"
    End If
    
    rsstock.Open msqlstr, cnnData, adOpenDynamic, adLockReadOnly
    If rsstock.EOF <> True Then
        Set oSheet = oBook.Worksheets(2)
        Dim nrow As Integer
        nrow = 1
        Do While rsstock.EOF <> True
            oSheet.Cells(nrow, 1) = findvalue("Select Name From Itemmaster i where code = '" + rsstock!grp + "'")
            oSheet.Cells(nrow, 2) = rsstock!Code
            oSheet.Cells(nrow, 3) = rsstock!Stock
            oSheet.Cells(nrow, 4) = rsstock!fiforate
            oSheet.Cells(nrow, 5) = Val(rsstock!FifoValue)
            nrow = nrow + 1
            rsstock.MoveNext
        Loop
    End If
    rsstock.Close
    
    Set oSheet = oBook.Worksheets(1)
    
    If findvalue("select name from sysobjects where name='tempstock2'") <> "" Then cnnData.Execute "drop table tempstock2"
    If findvalue("select name from sysobjects where name='temppurch2'") <> "" Then cnnData.Execute "drop table temppurch2"
    If findvalue("select name from sysobjects where name='PStock2'") <> "" Then cnnData.Execute "drop table PStock2"
    
    'mGroup = "'100057','100056','100059','100060'"
    
    If UCase(frmStockVal.Caption) = "TFATDEMO" Then
            msqlstr = "select * into tempstock2 from (select * From TempStock Where Grp in ('100067','100066','100052','100019','100065','100056','100064')) as a"
    ElseIf UCase(frmStockVal.Caption) = "TFATSILLP" Then
            msqlstr = "select * into tempstock2 from (select * From TempStock where Grp In ('100002','100023','100009') )  as a "
    ElseIf UCase(frmStockVal.Caption) = "TFATDSPL001" Then
        msqlstr = "select * into tempstock2 from (select * From TempStock where Grp IN ('100018')) as a "
    ElseIf UCase(frmStockVal.Caption) = "TFATNMJ001" Then
            msqlstr = "select * into tempstock2 from (select * FROM TempStock where Grp = '') as a "
    End If
   
    cnnData.CommandTimeout = 3000
    cnnData.Execute msqlstr
    
    
    
    mrow = 5
    
    rsstock.CursorLocation = adUseClient
    If UCase(frmStockVal.Caption) = "TFATDEMO" Then
        rsstock.Open "Select * from TmpStkVal Where Grp  in ('100067','100066','100052','100019','100065','100056','100064')", cnnData, adOpenStatic, adLockReadOnly, adCmdText
    ElseIf UCase(frmStockVal.Caption) = "TFATSILLP" Then
        rsstock.Open "Select * from TmpStkVal Where Grp  In ('100002','100023','100009')", cnnData, adOpenStatic, adLockReadOnly, adCmdText
    ElseIf UCase(frmStockVal.Caption) = "TFATDSPL001" Then
        rsstock.Open "Select * from TmpStkVal Where Grp  ='100018'", cnnData, adOpenStatic, adLockReadOnly, adCmdText
    ElseIf UCase(frmStockVal.Caption) = "TFATNMJ001" Then
        rsstock.Open "Select * from TmpStkVal Where Grp  =''", cnnData, adOpenStatic, adLockReadOnly, adCmdText
    End If
    If rsstock.RecordCount > 0 Then
        rsstock.MoveFirst
        Do Until rsstock.EOF
            mOthergroup = mOthergroup + IIf(IsNull(rsstock!FifoValue) = True, 0, rsstock!FifoValue)
            rsstock.MoveNext
        Loop
    End If
    rsstock.Close
    
    If UCase(frmStockVal.Caption) = "TFATDEMO" Then
        msqlstr = "Select * From TmpStkVal where Grp  in ('100067','100066','100052','100019','100065','100056','100064') And Stock > 0 Order BY GRP,Code"
    ElseIf UCase(frmStockVal.Caption) = "TFATSILLP" Then
        msqlstr = "Select * From TmpStkVal where Grp  in ('100002','100023','100009') And Stock > 0 Order BY GRP,Code"
    ElseIf UCase(frmStockVal.Caption) = "TFATDSPL001" Then
        msqlstr = "Select * From TmpStkVal where Grp  in ('100018') And Stock > 0 Order BY GRP,Code"
    ElseIf UCase(frmStockVal.Caption) = "TFATNMJ001" Then
        msqlstr = "Select * From TmpStkVal where Grp  in ('') And Stock > 0 Order BY GRP,Code"
    End If
    
    rsstock.Open msqlstr, cnnData, adOpenDynamic, adLockReadOnly
    If rsstock.EOF <> True Then
        'Set oSheet = oBook.Worksheets(2)
        'oExcel.Workbooks.Add
        Set oSheet = oBook.Worksheets(3)
        Dim xrow As Integer
        xrow = 1
        Do While rsstock.EOF <> True
            oSheet.Cells(xrow, 1) = findvalue("Select Name From Itemmaster i where code = '" + rsstock!grp + "'")
            oSheet.Cells(xrow, 2) = rsstock!Code
            oSheet.Cells(xrow, 3) = rsstock!Stock
            oSheet.Cells(xrow, 4) = rsstock!fiforate
            oSheet.Cells(xrow, 5) = Val(rsstock!FifoValue)
            xrow = xrow + 1
            rsstock.MoveNext
        Loop
    End If
    rsstock.Close
    
    Set oSheet = oBook.Worksheets(1)
    
    
    mrow = 1
    
    If UCase(frmStockVal.Caption) = "TFATDEMO" Then
        oSheet.Cells(mrow, 1) = "SURAJ ENTERPRISE PVT LTD"
    ElseIf UCase(frmStockVal.Caption) = "TFATSILLP" Then
        oSheet.Cells(mrow, 1) = "SURAJ IMPEX LLP"
    ElseIf UCase(frmStockVal.Caption) = "TFATDSPL001" Then
        oSheet.Cells(mrow, 1) = "DINESH STAINLESS PVT LTD"
    ElseIf UCase(frmStockVal.Caption) = "TFATNMJ001" Then
        oSheet.Cells(mrow, 1) = "NISHA MAHESH JAIN"
    End If
    mrow = mrow + 1
    oSheet.Cells(mrow, 6) = "DATE :"
    oSheet.Cells(mrow, 7) = Format(EndDate.Value, "DD-MMM-YYYY")
    oSheet.Rows.Cells(mrow, 7).HorizontalAlignment = xlLeft
    oSheet.Range("A1:G1").Merge
'    Selection.Merge
    oSheet.Rows.Cells(1, 1).HorizontalAlignment = xlCenter
    oSheet.Rows.Cells(1, 1).VerticalAlignment = xlCenter
    oSheet.Cells.Font.Bold = True
    oSheet.Cells.Font.Size = 11
    mrow = mrow + 2
    oSheet.Cells(mrow, 4) = "SHARE"
    mrow = mrow + 2
    oSheet.Cells(mrow, 1) = "Stock - FIFO"
    
    oSheet.Cells(mrow, 3) = Round(Val(mGrandTotal), 0)
    oSheet.Cells(mrow, 5) = "Opening "
    
    Dim mOp As String
    mOp = Val(findvalue("Select isnull(Opening,0) From Master Where Name = 'Opening Stock  - Share'"))
    
    'mOp = mOp + findvalue("Select isnull(Opening,0) From Master Where Name = 'Opening Stock - Metal'")
    'mOp = mOp + findvalue("Select isnull(sum(Debit-Credit),0) From Ledger Where Code in (Select Code From Master Where name = 'Opening Stock - Metal') And left(Authorise,1) ='A'")
    
    mOp = Val(mOp) + findvalue("Select isnull(sum(Debit-Credit),0) From Ledger Where Code in (Select Code From Master Where name = 'Opening Stock  - Share') And left(Authorise,1) ='A'")
    
    Dim mPurch As String
    mPurch = findvalue("Select isnull(Opening,0) From Master Where Name = 'Trading Purchase'")
    mPurch = Val(mPurch) + Val(findvalue("Select isnull(Opening,0) From Master Where Name = 'Purchase W/O Stt'"))
    
    mPurch = mPurch + findvalue("Select isnull(sum(Debit-Credit),0) From Ledger Where Code in (Select Code From Master Where name = 'Trading Purchase') And left(Authorise,1) ='A' And Docdate Between '" + Format(StartDate.Value, "DD-MMM-YYYY") + "' And '" + Format(EndDate.Value, "DD-MMM-YYYY") + "' And Type+Prefix+Srl In (Select Type+Prefix+Srl From Stock Where Docdate Between '" + Format(StartDate.Value, "DD-MMM-YYYY") + "' And '" + Format(EndDate.Value, "DD-MMM-YYYY") + "'   And Code not in (Select Code From Itemmaster Where grp In ('100061')) And Subtype = 'RP')")
    mPurch = mPurch + findvalue("Select isnull(sum(Debit-Credit),0) From Ledger Where Code in (Select Code From Master Where name = 'Purchase W/O Stt') And left(Authorise,1) ='A' And Docdate Between '" + Format(StartDate.Value, "DD-MMM-YYYY") + "' And '" + Format(EndDate.Value, "DD-MMM-YYYY") + "' And Type+Prefix+Srl In (Select Type+Prefix+Srl From Stock Where Docdate Between '" + Format(StartDate.Value, "DD-MMM-YYYY") + "' And '" + Format(EndDate.Value, "DD-MMM-YYYY") + "' And  Code not in (Select Code From Itemmaster Where grp In ('100061')) And Subtype = 'RP')")
    
    
    Dim mSales As String
    mSales = findvalue("Select isnull(Opening,0) From Master Where Name = 'Trading Sales'")
    mSales = mSales + findvalue("Select isnull(sum(Debit-Credit),0) From Ledger Where Code in (Select Code From Master Where name = 'Trading Sales') And left(Authorise,1) ='A' And Docdate Between '" + Format(StartDate.Value, "DD-MMM-YYYY") + "' And '" + Format(EndDate.Value, "DD-MMM-YYYY") + "'   And Type+Prefix+Srl In (Select Type+Prefix+Srl From Stock Where Docdate Between '" + Format(StartDate.Value, "DD-MMM-YYYY") + "' And '" + Format(EndDate.Value, "DD-MMM-YYYY") + "' And  Code not in (Select Code From Itemmaster Where grp In ('100061')) And Subtype = 'RS')")
    
    mSales = mSales + findvalue("Select isnull(Opening,0) From Master Where Name = 'Sales W/O S T T'")
    mSales = mSales + findvalue("Select isnull(sum(Debit-Credit),0) From Ledger Where Code in (Select Code From Master Where name = 'Sales W/O S T T') And left(Authorise,1) ='A' And Docdate Between '" + Format(StartDate.Value, "DD-MMM-YYYY") + "' And '" + Format(EndDate.Value, "DD-MMM-YYYY") + "' And  Type+Prefix+Srl In (Select Type+Prefix+Srl From Stock Where Docdate Between '" + Format(StartDate.Value, "DD-MMM-YYYY") + "' And '" + Format(EndDate.Value, "DD-MMM-YYYY") + "' And  Code not in (Select Code From Itemmaster Where grp In ('100061')) And Subtype = 'RS')")
    
    Dim mDivi As String
    
    mDivi = findvalue("Select isnull(Opening,0) From Master Where Name = 'Dividend On Share'")
    mDivi = mDivi + findvalue("Select isnull(sum(Debit-Credit),0) From Ledger Where Code in (Select Code From Master Where name = 'Dividend On Share') And left(Authorise,1) ='A' And Docdate Between '" + Format(StartDate.Value, "DD-MMM-YYYY") + "' And '" + Format(EndDate.Value, "DD-MMM-YYYY") + "'")
    
    oSheet.Cells(mrow, 7) = Round(Val(mOp), 0)
    mrow = mrow + 1
    'oSheet.Cells(mrow, 6) = "(+)"
    
    mrow = mrow + 1
    'oSheet.Cells(mrow, 1) = "(-)"
    
    oSheet.Cells(mrow, 5) = "Trading Purchase (+) "
    oSheet.Cells(mrow, 7) = Round(Val(mPurch), 0)
    mrow = mrow + 1
    
    mrow = mrow + 1
    oSheet.Cells(mrow, 6) = "Total"
    
    oSheet.Cells(mrow, 7) = Round(Val(mOp), 0) + Round(Val(mPurch), 0)
    oSheet.Cells(mrow, 1) = "Net Value (-)"
    
    oSheet.Cells(mrow, 3) = Val(Val(Round(mOp, 0)) + Val(Round(mPurch, 0))) - Val(Abs(Round(mSales)))
    
    mrow = mrow + 1
    'oSheet.Cells(mrow, 6) = "(-)"
    
    mrow = mrow + 1
    oSheet.Cells(mrow, 5) = "Trading Sales (-)"
    
    Dim mNet As String
    mNet = Val(Val(Round(mOp, 0)) + Val(Round(mPurch, 0))) - Val(Abs(Round(mSales, 0)))
    
    oSheet.Cells(mrow, 7) = Abs(Val(Round(mSales, 0)))
    oSheet.Cells(mrow, 1) = "Profit"
    oSheet.Cells(mrow, 3) = Val(Round(mGrandTotal, 0)) - Val(Round(mNet, 0))
    mrow = mrow + 1
    oSheet.Cells(mrow, 6) = "Net Value "
    oSheet.Cells(mrow, 7) = Round(Val(mNet), 0)
    
    
    mrow = mrow + 1
    'oSheet.Cells(mrow, 1) = "(+)"
     
    mrow = mrow + 1
    
    oSheet.Cells(mrow, 1) = "Dividend (+)"
    
    oSheet.Cells(mrow, 3) = Round(Val(Abs(mDivi)), 0)
    mrow = mrow + 1
    

    mrow = mrow + 1
    oSheet.Cells(mrow, 1) = "Gross Profit : "
    
    oSheet.Cells(mrow, 3) = Val(Round(Val(mGrandTotal), 0) - Round(Val(mNet), 0)) + Val(Abs(Round(mDivi, 0)))
    
    mrow = mrow + 2
    oSheet.Cells(mrow, 1) = "A1 Share  "
    oSheet.Cells(mrow, 3) = Val(Round(Val(mAbcGroup), 0))
    
    mrow = mrow + 2
    oSheet.Cells(mrow, 1) = "Other Group "
    
    oSheet.Cells(mrow, 3) = Val(Round(Val(mOthergroup), 0))
    
    mrow = mrow + 2
    oSheet.Cells(mrow, 1) = "Total"
    oSheet.Cells(mrow, 3) = Val(Round(Val(Val(mAbcGroup) + Val(mOthergroup)), 0))
    
    oSheet.Range("C2:C100").ColumnWidth = 14.86
    oSheet.Range("G2:G100").ColumnWidth = 14.86
    oSheet.Range("B2:B100").ColumnWidth = 17
    oSheet.Protect
    
    MsgBox "Report generated", vbOKOnly + vbInformation
End Function


Private Sub TxtSearch_Change()
       Dim mCol As Integer, mrow As Integer, i As Integer
    mCol = gridgroup.ColSel
    mrow = gridgroup.Row
    For i = gridgroup.FixedRows To gridgroup.Rows - 1
        If optInWholeWord.Value = True Then
            If Mid(UCase(gridgroup.TextMatrix(i, mCol)), 1, Len(TxtSearch.Text)) = UCase(TxtSearch.Text) Then
                gridgroup.Row = i
                gridgroup.Col = mCol
                gridgroup.TopRow = i
                gridgroup.SetFocus
                Exit For
            End If
        Else
            If InStr(UCase(gridgroup.TextMatrix(i, mCol)), UCase(TxtSearch.Text)) <> 0 Then
                gridgroup.Row = i
                gridgroup.Col = mCol
                gridgroup.TopRow = i
                gridgroup.SetFocus
                Exit For
            End If
        End If
    Next
    TxtSearch.SetFocus

End Sub
