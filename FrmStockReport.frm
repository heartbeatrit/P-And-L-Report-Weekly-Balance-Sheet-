VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{E644C91B-78F7-4D86-8316-181A41A236A9}#1.0#0"; "XPButton.ocx"
Begin VB.Form FrmStockReport 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Stock Report"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CmbBranch 
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1320
      Width           =   4215
   End
   Begin MSComCtl2.DTPicker StartDate 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1800
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
      Format          =   91488257
      CurrentDate     =   40157
   End
   Begin ProjetXPButton.XPButton CmdReport 
      Height          =   495
      Left            =   9600
      TabIndex        =   2
      Top             =   6240
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridstore 
      Height          =   3015
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5318
      _Version        =   393216
      BackColorBkg    =   -2147483634
      AllowUserResizing=   3
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin ProjetXPButton.XPButton cmdstoreselect 
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
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
   Begin ProjetXPButton.XPButton cmdgroupselect 
      Height          =   255
      Left            =   10320
      TabIndex        =   8
      Top             =   5760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridgroup 
      Height          =   3015
      Left            =   5760
      TabIndex        =   9
      Top             =   2640
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5318
      _Version        =   393216
      BackColorBkg    =   -2147483634
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Group"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Branch"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image ImageSelect 
      Height          =   480
      Left            =   0
      Picture         =   "FrmStockReport.frx":0000
      Top             =   6960
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   1110
      Left            =   0
      Picture         =   "FrmStockReport.frx":0E42
      Top             =   0
      Width           =   3360
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   3360
      Picture         =   "FrmStockReport.frx":819B
      Top             =   600
      Width           =   1800
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Store Selection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "As On Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
End
Attribute VB_Name = "FrmStockReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gbranch As String, mstore As String, mstdate As String, mGroup As String

Private Sub cmdgroupselect_Click()
    gridselectunselect gridstore, cmdstoreselect, 0, ImageSelect
    If cmdstoreselect.Caption = "&Select All" Then
        cmdstoreselect.Caption = "&Unselect All"
    Else
        cmdstoreselect.Caption = "&Select All"
    End If
End Sub

Private Sub CmdReport_Click()
    Dim oExcel As Object
    Dim oBook As Object
    Dim ActiveSheet As Object
    Dim rsstock As New ADODB.Recordset, rspurch As New ADODB.Recordset, rsitem As New ADODB.Recordset
    Dim mcode As String, mname As String, mrate As Double, mval As Double, mqty As Double, mtotval As Double
    Dim msqlstr As String, mtotqty As Double, mxqty As Double, mday As Integer, mdaydiff As Integer
    Dim mtot(7) As Currency
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Visible = True
    Set oBook = oExcel.Workbooks.Add
    
    'Add data to cells of the first worksheet in the new workbook
    Set ActiveSheet = oBook.Worksheets(1)
    
    'mstdate = Format(StartDate.Value, "mm/dd/yyyy")
    'meddate = Format(EndDate.Value, "mm/dd/yyyy")
    mstdate = Format(StartDate.Value, "dd-mmm-yyyy")
    If CmbBranch.Text = "" Then
        MsgBox "No branch selected", vbOKOnly + vbInformation
        Exit Sub
    End If
    
    gbranch = Mid(CmbBranch.Text, 1, InStr(1, CmbBranch.Text, "-") - 1)
    storecollect
    If mstore = "" Then
        MsgBox "Dear User, No Store Selected", vbOKOnly + vbInformation
        Exit Sub
    End If
    
    groupcollect
    If mGroup = "" Then
        MsgBox "Dear User, No Product Group Selected", vbOKOnly + vbInformation
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

Private Sub Form_Load()
    StartDate.Value = Date
    Dim rsstore As New ADODB.Recordset
    rsstore.CursorLocation = adUseClient
    rsstore.Open "select code,name from stores Where Branch = 'HEADOF' group by code,name", cnnData, adOpenStatic, adLockReadOnly, adCmdText
    If rsstore.RecordCount > 0 Then
        Set gridstore.DataSource = rsstore
        gridstore.ColWidth(0) = 500
        gridstore.ColWidth(1) = 0
        gridstore.ColWidth(2) = 4000
    End If
    rsstore.Close
    Dim rscnn As New ADODB.Recordset
    rscnn.CursorLocation = adUseClient
    rscnn.Open "Select code+'-'+name as name from " + gnTfatSet + ".dbo.tfatbranch order by code", cnnData, adOpenStatic, adLockReadOnly, adCmdText
    If rscnn.RecordCount > 0 Then
        CmbBranch.Clear
        Do Until rscnn.EOF
            CmbBranch.AddItem rscnn!Name
            rscnn.MoveNext
        Loop
    End If
    rscnn.Close
    
    Dim rsgroup As New ADODB.Recordset
    rsgroup.CursorLocation = adUseClient
    rsgroup.Open "select code,name from itemmaster where flag='G' order by name", cnnData, adOpenStatic, adLockReadOnly, adCmdText
    If rsgroup.RecordCount > 0 Then
        Set gridgroup.DataSource = rsgroup
        gridgroup.ColWidth(0) = 500
        gridgroup.ColWidth(1) = 0
        gridgroup.ColWidth(2) = 4000
    End If
    rsgroup.Close
    
    
End Sub
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
    ActiveSheet.Cells(3, 2) = "Stock Report"
    ActiveSheet.Cells(4, 2) = "Item Code"
    ActiveSheet.Cells(4, 3) = "Item Name"
    ActiveSheet.Cells(4, 4) = "Store Name"
    ActiveSheet.Cells(4, 5) = "Cl Stock"
    ActiveSheet.Cells(4, 6) = "Qty"
    ActiveSheet.Cells(4, 7) = "Serial Number"
    ActiveSheet.Cells(4, 8) = "Product Group"
    
    msqlstr = "Select (Select Name From Itemmaster i Where Code in (Select GRP from itemmaster where code=Stock.code)) as " _
              & " PGrp,Code,(Select name from itemmaster where code=Stock.code) as Productname, " _
              & "  (Select Top 1 Name From Stores Where Code = Stock.Store And Branch = Stock.Branch) as StoreName, " _
              & "  SUM(Qty) as ClStock,Stock.Code+'/'+Stock.Store as SerialNumber From Stock " _
              & "  Where Branch ='" + gbranch + "'  And Stock.Store in (" + mstore + ") And docdate<='" + mstdate + "' and notinstock=0 and left(authorise,1)='A' And Stock.Code not in (Select Code From Itemmaster a Where GRP in ('100023','100473')) And  Stock.code in (select code from itemmaster a where grp in (" + mGroup + "))  " _
              & "  Group by Code,Stock.store,Stock.Branch Having SUM(Qty)<>0"
    cnnData.CommandTimeout = 3000
    rsstock.Open msqlstr, cnnData, adOpenDynamic, adLockReadOnly
    Dim i As Integer
    mrow = 5
    If rsstock.EOF <> True Then
        Do While rsstock.EOF <> True
            For i = 0 To Val(rsstock.Fields("ClStock") - 1)
                ActiveSheet.Cells(mrow, 2) = CStr(rsstock!Code)
                ActiveSheet.Cells(mrow, 3) = rsstock!ProductName
                ActiveSheet.Cells(mrow, 4) = rsstock!StoreName
                ActiveSheet.Cells(mrow, 5) = rsstock!ClStock
                ActiveSheet.Cells(mrow, 6) = 1
                ActiveSheet.Cells(mrow, 7) = rsstock!SerialNumber + "/" + CStr(i + 1)
                ActiveSheet.Cells(mrow, 8) = rsstock!pGrp
                mrow = mrow + 1
            Next
            rsstock.MoveNext
        Loop
    End If
    rsstock.Close
    Unload Me
End Function

Private Sub gridgroup_Click()
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


