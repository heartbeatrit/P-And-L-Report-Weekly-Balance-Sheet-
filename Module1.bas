Attribute VB_Name = "Module1"
    
    Public cnnData As New ADODB.Connection
    Public rst As New ADODB.Recordset
    Public rsl As New ADODB.Recordset
    Public gNature001 As String, gDataBaseGen As String, gDataBaseGenStr As String, gnTfatSet As String, gSQLServer As String
    
Sub Main()
    Dim mdatabase As String, mPassword As String
    Dim msqlserver As String, mPwd As String
    cnnData.Open "Provider=Microsoft.jet.oledb.4.0;" & "Data Source=" + App.Path + "\GeneralData.mdb"
    Dim rsmain As New ADODB.Recordset
    rsmain.CursorLocation = adUseClient
    rsmain.Open "Select * from sqlserver", cnnData, adOpenKeyset, adLockPessimistic
    If rsmain.EOF = False Then
        mdatabase = rsmain.Fields("databasename")
        msqlserver = rsmain.Fields("servername")
        If gYearCode <> "" Then
            YearStDate = "01-Apr-" + Right(gYearCode, 4)
            YearEdDate = "31-Mar-" + Right(CStr(Val(gYearCode) + 1), 4)
        End If
        gNature001 = rsmain.Fields("nature001")
        gnTfatSet = rsmain.Fields("ntfatset")
        gDataBaseGen = gnTfatSet
        gDataBaseGenStr = gnTfatSet + ".dbo."
        mPassword = IIf(IsNull(rsmain.Fields("Password")), "", rsmain.Fields("Password"))
    End If
    cnnData.Close
    
    cnnData.Open "Provider=SQLOLEDB.1;User Id=sa;Password='" + mPassword + "';server=" & msqlserver & ";Initial Catalog=" & mdatabase & ""
    cnnData.CommandTimeout = 1000
    'frmStockVal.Caption = gNature001
    frmStockVal.Show
    
    'FrmRateUpdate.Show
    'FrmStockReport.Show
End Sub

Private Sub PrintAreaWithpageBreaks()
    Dim pages As Integer
    Dim pageBegin As String
    Dim PrArea As String
    Dim i As Integer
    Dim q As Integer
    Dim nRows As Integer, nPagebreaks As Integer
    Dim R As Range
    Set R = ActiveSheet.UsedRange
    'add pagebreak every 40 rows
    nRows = R.Rows.Count
    If nRows > 40 Then
        nPagebreaks = Int(nRows / 40)
        For i = 1 To nPagebreaks
           ActiveWindow.SelectedSheets.HPageBreaks.Add before:=R.Cells(20 * i + 1, 1)
        Next i
    End If
    'can be used in a separate macro, as I Start counting the number of pagebreaks
    pages = ActiveSheet.HPageBreaks.Count
    pageBegin = "$A$1"
    For i = 1 To pages
      If i > 1 Then pageBegin = ActiveSheet.HPageBreaks(i - 1).Location.Address
      q = ActiveSheet.HPageBreaks(i).Location.Row - 1
      PrArea = pageBegin & ":" & "$W$" & Trim$(Str$(q))
      ActiveSheet.PageSetup.PrintArea = PrArea
      ' the cell in column 1 and in the row immediately below the pagebreak
      ' contains text for the footer
      ActiveSheet.PageSetup.CenterFooter = Cells(q, 1)
    '  ActiveSheet.PrintOut copies:=1
    Next i
End Sub


Public Function findvalue(ByVal xsqlstr As String)
    findvalue = ""
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open xsqlstr, cnnData, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then findvalue = rs.Fields(0)
    rs.Close
End Function

