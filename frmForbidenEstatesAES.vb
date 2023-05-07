Public Class frmForbidenEstatesAES

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        cmdOK.Select()
        If CStr(Me.Tag) = "Add" Then AddEstate() Else EditEstate()
    End Sub

    Public Sub AddEstate()
        If CheckValues() = False Then Exit Sub
        Dim lngSectionsCode As Long
        Dim strEstateNumber As String = ReturnEstateTableNumber(cmbProvinces.Text)
        lngSectionsCode = ReturnSectionCodeEstate(cmbProvinces.Text, cmbCities.Text, cmbSections.Text, False)
        cmmEstate.CommandText = "SELECT Code FROM E" & strEstateNumber & " WHERE Convert(varchar(100), DecryptByPassPhrase('fgh147hgf$ t9p080602azpq',PelakNo)) LIKE '" & txtPelakNo.Text & "' AND SectionsCode = " & lngSectionsCode
        If cmmEstate.ExecuteScalar <> 0 Then
            Call ShowError("این ملک قبلاً وجود دارد.")
            Exit Sub
        End If
        If lngSectionsCode = 0 Then lngSectionsCode = ReturnSectionCodeEstate(cmbProvinces.Text, cmbCities.Text, cmbSections.Text, True)
        Call AddNewEstate(lngSectionsCode, strEstateNumber)
        If cmbProvinces.Text = frmForbidenEstates.cmbProvinces.Text Then
            Dim drNew As DataRow, dtForbidenEstates As DataTable = frmForbidenEstates.Tag
            drNew = dtForbidenEstates.NewRow
            drNew.Item(1) = txtPelakNo.Text
            drNew.Item(2) = cmbCities.Text
            drNew.Item(3) = cmbSections.Text
            drNew.Item(4) = txtSplitedNo.Text
            drNew.Item(5) = txtInPlace.Text
            drNew.Item(6) = txtENotes.Text
            dtForbidenEstates.Rows.Add(drNew)
            Call FillRowNumberColumn(frmForbidenEstates.dgrEstates)
        End If
        Me.Close()
    End Sub

    Public Function EditEstate() As Long
        EditEstate = 0
        If CheckValues() = False Then Exit Function
        Dim lngSectionsCode As Long
        Dim strEstateNumber As String = ReturnEstateTableNumber(cmbProvinces.Text)
        lngSectionsCode = ReturnSectionCodeEstate(cmbProvinces.Text, cmbCities.Text, cmbSections.Text, False)
        If cmbProvinces.Text = frmForbidenEstates.cmbProvinces.Text Then
            cmmEstate.CommandText = "SELECT Code FROM E" & strEstateNumber & " WHERE Convert(varchar(100), DecryptByPassPhrase('fgh147hgf$ t9p080602azpq',PelakNo)) LIKE '" & txtPelakNo.Text & "' AND SectionsCode = " & lngSectionsCode & " AND Code<>" & Me.Tag
        Else
            cmmEstate.CommandText = "SELECT Code FROM E" & strEstateNumber & " WHERE Convert(varchar(100), DecryptByPassPhrase('fgh147hgf$ t9p080602azpq',PelakNo)) LIKE '" & txtPelakNo.Text & "' AND SectionsCode = " & lngSectionsCode
        End If
        If Not IsNothing(cmmEstate.ExecuteScalar) Then
            Call ShowError("این ملک قبلاً وجود دارد.")
            Exit Function
        End If
        If lngSectionsCode = 0 Then lngSectionsCode = ReturnSectionCodeEstate(cmbProvinces.Text, cmbCities.Text, cmbSections.Text, True)
        cmmEstate.CommandText = "DELETE FROM Forbidens" & strEstateNumber & " WHERE TableType = 2 AND InTableCode IN (SELECT Code FROM Registers" & strEstateNumber & " WHERE E" & strEstateNumber & "Code = " & Me.Tag & ")"
        cmmEstate.ExecuteNonQuery()        
        cmmEstate.CommandText = "DELETE FROM Forbidens" & strEstateNumber & " WHERE TableType = 1 AND InTableCode = " & Me.Tag
        cmmEstate.ExecuteNonQuery()        
        cmmEstate.CommandText = "DELETE FROM Transes" & strEstateNumber & " WHERE Registers" & strEstateNumber & "Code IN (SELECT Code FROM Registers" & strEstateNumber & " WHERE E" & strEstateNumber & "Code = " & Me.Tag & ")"
        cmmEstate.ExecuteNonQuery()        
        cmmEstate.CommandText = "DELETE FROM Registers" & strEstateNumber & " WHERE E" & strEstateNumber & "Code = " & Me.Tag
        cmmEstate.ExecuteNonQuery()        
        cmmEstate.CommandText = "DELETE FROM E" & strEstateNumber & " WHERE Code = " & Me.Tag
        cmmEstate.ExecuteNonQuery()        
        Call AddNewEstate(lngSectionsCode, strEstateNumber)

        If cmbProvinces.Text = frmForbidenEstates.cmbProvinces.Text Then
            With frmForbidenEstates.dgrEstates.SelectedRows(0)
                .Cells(1).Value = txtPelakNo.Text
                .Cells(2).Value = cmbCities.Text
                .Cells(3).Value = cmbSections.Text
                .Cells(4).Value = txtSplitedNo.Text
                .Cells(5).Value = txtInPlace.Text
                .Cells(6).Value = txtENotes.Text
            End With
            Dim Index As Long = frmForbidenEstates.dgrEstates.SelectedRows(0).Index
            frmForbidenEstates.dgrEstates.Rows(Index).Selected = False
            frmForbidenEstates.dgrEstates.Rows(Index).Selected = True
        Else
            frmForbidenEstates.dgrEstates.Rows.Remove(frmForbidenEstates.dgrEstates.SelectedRows(0))
            If frmForbidenEstates.dgrEstates.SelectedRows.Count = 0 Then If frmForbidenEstates.dgrEstates.Rows.Count > 0 Then frmForbidenEstates.dgrEstates.Rows(frmForbidenEstates.dgrEstates.Rows.Count - 1).Selected = True
        End If
        Me.Close()
    End Function

    Private Sub AddNewEstate(ByVal lngSectionsCode As Long, ByVal strEstateNumber As String)
        Dim lngEstateCode As Long, lngRegisterCode As Long, lngTransCode As Long, lngForbidenCode As Long, Index As Long, Index1 As Long
        Dim strEstates As String = "E" & strEstateNumber
        Dim strRegisters As String = "Registers" & strEstateNumber
        Dim strTranses As String = "Transes" & strEstateNumber
        Dim strForbidens As String = "Forbidens" & strEstateNumber

        '1-Estates
        lngEstateCode = CreateNewCode(strEstates, cmmEstate)
        If lngEstateCode < 1000000000 Then lngEstateCode = 1000000000
        cmmEstate.CommandText = "INSERT INTO " & strEstates & "(Code, PelakNo, SectionsCode, SplitedNo, InPlace, Notes) VALUES (" & lngEstateCode & ",EncryptByPassPhrase('fgh147hgf$ t9p080602azpq',convert(varchar(100),'" & txtPelakNo.Text & "'))," & lngSectionsCode & ",'" & txtSplitedNo.Text & "','" & txtInPlace.Text & "','" & txtENotes.Text & "')"
        cmmEstate.ExecuteNonQuery()        
        For Index = 0 To dgrEForbidens.Rows.Count - 1
            lngForbidenCode = CreateNewCode(strForbidens, cmmEstate)
            cmmEstate.CommandText = "INSERT INTO " & strForbidens & "(Code, TableType, InTableCode, ForbidenTypesCode, ForbidenNo, ForbidenDate, ForbidenOrganizationsCode, ArrestNo, ArrestDate, ArrestOrganizationsCode, ScanFile, Notes) VALUES (" & lngForbidenCode & ",1," & lngEstateCode & "," & ReturnTableCode("ForbidenTypes", dgrEForbidens.Rows(Index).Cells(0).Value, cmmEstate, False) & ",'" & dgrEForbidens.Rows(Index).Cells(1).Value & "','" & dgrEForbidens.Rows(Index).Cells(2).Value & "'," & ReturnTableCode("Organizations", dgrEForbidens.Rows(Index).Cells(3).Value, cmmEstate, True) & ",'" & dgrEForbidens.Rows(Index).Cells(4).Value & "','" & dgrEForbidens.Rows(Index).Cells(5).Value & "'," & ReturnTableCode("Organizations", dgrEForbidens.Rows(Index).Cells(6).Value, cmmEstate, True) & ",'" & dgrEForbidens.Rows(Index).Cells(7).Value & "','" & dgrEForbidens.Rows(Index).Cells(8).Value & "')"
            cmmEstate.ExecuteNonQuery()            
        Next

        '2-Registers
        For Index = 0 To dgrRegisters.Rows.Count - 1
            lngRegisterCode = CreateNewCode(strRegisters, cmmEstate)
            cmmEstate.CommandText = "INSERT INTO " & strRegisters & "(Code, E" & strEstateNumber & "Code, PageNo, BookNo, RegisterNo, RegisterDate, PrintNo, OwnerName, Area, Limitis, Notes) VALUES (" & lngRegisterCode & "," & lngEstateCode & ",'" & dgrRegisters.Rows(Index).Cells(2).Value & "','" & dgrRegisters.Rows(Index).Cells(3).Value & "','" & dgrRegisters.Rows(Index).Cells(4).Value & "','" & dgrRegisters.Rows(Index).Cells(5).Value & "',EncryptByPassPhrase('fgh147hgf$ t9p080602azpq',convert(varchar(100),'" & dgrRegisters.Rows(Index).Cells(6).Value & "')),'" & dgrRegisters.Rows(Index).Cells(7).Value & "','" & dgrRegisters.Rows(Index).Cells(8).Value & "','" & dgrRegisters.Rows(Index).Cells(9).Value & "','" & dgrRegisters.Rows(Index).Cells(10).Value & "')"
            cmmEstate.ExecuteNonQuery()            
            For Index1 = 0 To dgrTranses.Rows.Count - 1
                If CStr(dgrTranses.Rows(Index1).Cells(0).Value) = CStr(dgrRegisters.Rows(Index).Cells(0).Value) Then dgrTranses.Rows(Index1).Cells(0).Value = lngRegisterCode
            Next
            For Index1 = 0 To dgrRForbidens.Rows.Count - 1
                If CStr(dgrRForbidens.Rows(Index1).Cells(0).Value) = CStr(dgrRegisters.Rows(Index).Cells(0).Value) Then dgrRForbidens.Rows(Index1).Cells(0).Value = lngRegisterCode
            Next
            dgrRegisters.Rows(Index).Cells(0).Value = lngRegisterCode
        Next
        For Index = 0 To dgrRForbidens.Rows.Count - 1
            lngForbidenCode = CreateNewCode(strForbidens, cmmEstate)
            cmmEstate.CommandText = "INSERT INTO " & strForbidens & "(Code, TableType, InTableCode, ForbidenTypesCode, ForbidenNo, ForbidenDate, ForbidenOrganizationsCode, ArrestNo, ArrestDate, ArrestOrganizationsCode, ScanFile, Notes) VALUES (" & lngForbidenCode & ",2," & dgrRForbidens.Rows(Index).Cells(0).Value & "," & ReturnTableCode("ForbidenTypes", dgrRForbidens.Rows(Index).Cells(1).Value, cmmEstate, False) & ",'" & dgrRForbidens.Rows(Index).Cells(2).Value & "','" & dgrRForbidens.Rows(Index).Cells(3).Value & "'," & ReturnTableCode("Organizations", dgrRForbidens.Rows(Index).Cells(4).Value, cmmEstate, True) & ",'" & dgrRForbidens.Rows(Index).Cells(5).Value & "','" & dgrRForbidens.Rows(Index).Cells(6).Value & "'," & ReturnTableCode("Organizations", dgrRForbidens.Rows(Index).Cells(7).Value, cmmEstate, True) & ",'" & dgrRForbidens.Rows(Index).Cells(8).Value & "','" & dgrRForbidens.Rows(Index).Cells(9).Value & "')"
            cmmEstate.ExecuteNonQuery()            
        Next

        '3-Transes
        For Index = 0 To dgrTranses.Rows.Count - 1
            lngTransCode = CreateNewCode(strTranses, cmmEstate)
            cmmEstate.CommandText = "INSERT INTO " & strTranses & "(Code, Registers" & strEstateNumber & "Code, DocNo, DocDate, NotariesCode, DocSerial, DocSery, DocTypesCode, OwnerName, Area, Notes) VALUES (" & lngTransCode & "," & dgrTranses.Rows(Index).Cells(0).Value & ",'" & dgrTranses.Rows(Index).Cells(2).Value & "','" & dgrTranses.Rows(Index).Cells(3).Value & "'," & ReturnNotaryCode(dgrTranses.Rows(Index).Cells(4).Value, dgrTranses.Rows(Index).Cells(5).Value, cmmEstate, True) & ",'" & dgrTranses.Rows(Index).Cells(6).Value & "','" & dgrTranses.Rows(Index).Cells(7).Value & "'," & ReturnTableCode("DocTypes", dgrTranses.Rows(Index).Cells(8).Value, cmmEstate, False) & ",'" & dgrTranses.Rows(Index).Cells(9).Value & "','" & dgrTranses.Rows(Index).Cells(10).Value & "','" & dgrTranses.Rows(Index).Cells(11).Value & "')"
            cmmEstate.ExecuteNonQuery()            
            dgrTranses.Rows(Index).Cells(0).Value = lngTransCode
        Next
    End Sub

    Private Sub frmForbidenEstatesAES_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load        
        Call FillComboBox(cmbProvinces, "SELECT Name FROM Provinces ORDER BY Name ASC", cnnEstate)
        SendMessage(cmbProvinces.Handle, CB_SETITEMHEIGHT, -1, 17)
        cmbProvinces.Text = frmForbidenEstates.cmbProvinces.Text
        txtPelakNo.Select()
        If Not IsNumeric(Me.Tag) Then Exit Sub
        If frmForbidenEstates.dgrEstates.SelectedRows.Count = 0 Then Exit Sub
        If CStr(frmForbidenEstates.dgrEstates.Rows(0).Cells(1).Value) = "" Then Exit Sub
        With frmForbidenEstates.dgrEstates.SelectedRows(0)
            txtPelakNo.Text = .Cells(1).Value
            cmbProvinces.Text = frmForbidenEstates.cmbProvinces.Text
            cmbCities.Text = .Cells(2).Value
            cmbSections.Text = .Cells(3).Value
            txtSplitedNo.Text = .Cells(4).Value
            txtInPlace.Text = .Cells(5).Value
            txtENotes.Text = .Cells(6).Value
        End With
        Dim strEstateNumber As String = ReturnEstateTableNumber(cmbProvinces.Text), Index As Long
        Call InsertRowsGrid(dgrEForbidens, "SELECT T.Name AS ""نوع"", F.ForbidenNo AS ""شماره دادنامه"", F.ForbidenDate AS ""تاریخ دادنامه"", FO.Name AS ""مرجع صدور"", F.ArrestNo AS ""شماره بازداشتي"", F.ArrestDate AS ""تاریخ بازداشتي"", AO.Name AS ""مرجع صدور"", F.ScanFile AS ""تصوير بخشنامه"", F.Notes AS ""توضیحات"" FROM Forbidens" & strEstateNumber & " F, Organizations FO, Organizations AO, ForbidenTypes T WHERE F.ForbidenTypesCode = T.Code AND F.ForbidenOrganizationsCode = FO.Code AND F.ArrestOrganizationsCode = AO.Code AND F.TableType = 1 AND F.InTableCode = " & Me.Tag & " ORDER BY F.ForbidenDate, F.ForbidenNo DESC", cnnEstate, 2, System.ComponentModel.ListSortDirection.Descending)
        Call InsertRowsGrid(dgrRegisters, "SELECT Code, 1 AS ""رديف"", PageNo AS ""ش صفحه"", BookNo AS ""ش جلد"", RegisterNo AS ""ش ثبتی"", RegisterDate AS ""تاريخ ثبتي"", Convert(varchar(100), DecryptByPassPhrase('fgh147hgf$ t9p080602azpq',PrintNo)) AS ""ش چاپي"", OwnerName AS ""نام مالک"", Area AS ""مساحت"", Limitis AS ""حدود"", Notes AS ""توضيحات"" FROM Registers" & strEstateNumber & " WHERE E" & strEstateNumber & "Code = " & Me.Tag & " ORDER BY Code ASC", cnnEstate, 2, System.ComponentModel.ListSortDirection.Ascending)
        Call FillVisibleRowNumberColumn(dgrRegisters, 1)
        Call InsertRowsGrid(dgrRForbidens, "SELECT F.InTableCode AS R, T.Name AS ""نوع"", F.ForbidenNo AS ""شماره دادنامه"", F.ForbidenDate AS ""تاریخ دادنامه"", FO.Name AS ""مرجع صدور"", F.ArrestNo AS ""شماره بازداشتي"", F.ArrestDate AS ""تاریخ بازداشتي"", AO.Name AS ""مرجع صدور"", F.ScanFile AS ""تصوير بخشنامه"", F.Notes AS ""توضیحات"" FROM Forbidens" & strEstateNumber & " F, Organizations FO, Organizations AO, ForbidenTypes T WHERE F.ForbidenTypesCode = T.Code AND F.ForbidenOrganizationsCode = FO.Code AND F.ArrestOrganizationsCode = AO.Code AND F.TableType = 2 AND F.InTableCode IN (SELECT Code FROM Registers" & strEstateNumber & " WHERE E" & strEstateNumber & "Code = " & Me.Tag & ") ORDER BY F.InTableCode ASC, F.Code ASC", cnnEstate, 3, System.ComponentModel.ListSortDirection.Descending)
        Call InsertRowsGrid(dgrTranses, "SELECT T.Registers" & strEstateNumber & "Code, 1 AS ""رديف"", T.DocNo AS ""ش سند"", T.DocDate AS ""تاريخ سند"", C.Name AS ""شهر دفترخانه"", N.NotaryNo AS ""ش دفترخانه"", T.DocSerial AS ""ش سريال"", T.DocSery AS ""ش سري"", DT.Name AS ""نوع سند"", T.OwnerName AS ""نام دارنده"", T.Area AS ""مقدار"", T.Notes AS ""توضيحات"" FROM Transes" & strEstateNumber & " T, DocTypes DT, Notaries N, Cities C WHERE T.NotariesCode = N.Code AND N.CitiesCode = C.Code AND T.DocTypesCode = DT.Code AND T.Registers" & strEstateNumber & "Code IN (SELECT Code FROM Registers" & strEstateNumber & " WHERE E" & strEstateNumber & "Code = " & Me.Tag & ") ORDER BY Registers" & strEstateNumber & "Code ASC, T.DocDate DESC, T.DocNo DESC", cnnEstate, 3, System.ComponentModel.ListSortDirection.Descending)
        Call FillVisibleRowNumberColumn(dgrTranses, 1)
        dgrRegisters.ClearSelection()
        If dgrRegisters.Rows.Count <> 0 Then dgrRegisters.Rows(0).Selected = True
        For Index = 0 To dgrTranses.Rows.Count - 1
            If CStr(dgrTranses.Rows(Index).Cells(0).Value) = CStr(dgrRegisters.SelectedRows(0).Cells(0).Value) Then dgrTranses.Rows(Index).Visible = True Else dgrTranses.Rows(Index).Visible = False
        Next
        For Index = 0 To dgrRForbidens.Rows.Count - 1
            If CStr(dgrRForbidens.Rows(Index).Cells(0).Value) = CStr(dgrRegisters.SelectedRows(0).Cells(0).Value) Then dgrRForbidens.Rows(Index).Visible = True Else dgrRForbidens.Rows(Index).Visible = False
        Next
    End Sub

    '-------------------------------------
    '-------------------------------------

    Private Function CheckValues() As Boolean
        CheckValues = False
        Call DeleteExBlanks(txtPelakNo.Text)
        If BeEmpty(txtPelakNo) Then Exit Function
        Call DeleteExBlanks(cmbSections.Text)
        If BeEmpty(cmbSections) Then Exit Function
        If dgrEForbidens.Rows.Count = 0 And dgrRForbidens.Rows.Count = 0 Then
            Call ShowError("هیچ بازداشت یا رفع بازداشتی اضافه نشده است.")
            Exit Function
        End If
        CheckValues = True
    End Function

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub


    Private Sub cmbProvinces_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbProvinces.SelectedIndexChanged
        Call FillComboBox(cmbCities, "SELECT C.Name FROM Cities C, Provinces P WHERE C.ProvincesCode = P.Code AND P.Name LIKE '" & cmbProvinces.Text & "' ORDER BY C.Name ASC", cnnEstate)
    End Sub

    Private Sub cmbCities_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCities.SelectedIndexChanged
        Call FillComboBox(cmbSections, "SELECT S.Name FROM Sections S, Cities C, Provinces P WHERE S.CitiesCode = C.Code AND C.ProvincesCode = P.Code AND P.Name LIKE '" & cmbProvinces.Text & "' AND C.Name LIKE '" & cmbCities.Text & "' ORDER BY S.Name ASC", cnnEstate)
    End Sub

    Private Sub cmdEFAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEFAdd.Click
        Call AddForbidenEstate(dgrEForbidens, 0, "اضافه نمودن ممنوعيت پلاك")
    End Sub

    Private Sub cmdEFEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEFEdit.Click
        Call EditForbidenEstate(dgrEForbidens, 0, "ويرايش ممنوعيت پلاك")
    End Sub

    Private Sub cmdEFDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEFDelete.Click
        Call DeleteForbidenEstate(dgrEForbidens)
    End Sub

    Private Sub cmdRFAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRFAdd.Click
        If dgrRegisters.SelectedRows.Count = 0 Then
            Call ShowError("ابتدا سند مالكيت را انتخاب كنيد.")
            Exit Sub
        End If
        Dim blnAdd As Boolean = AddForbidenEstate(dgrRForbidens, 1, "اضافه نمودن ممنوعيت سند مالكيت")
        If blnAdd Then dgrRForbidens.Rows(dgrRForbidens.Rows.Count - 1).Cells(0).Value = dgrRegisters.SelectedRows(0).Cells(0).Value
    End Sub

    Private Sub cmdRFEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRFEdit.Click
        If dgrRegisters.SelectedRows.Count = 0 Then
            Call ShowError("ابتدا سند مالكيت را انتخاب كنيد.")
            Exit Sub
        End If
        Call EditForbidenEstate(dgrRForbidens, 1, "ويرايش ممنوعيت سند مالكيت")
    End Sub

    Private Sub cmdRFDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRFDelete.Click
        Call DeleteForbidenEstate(dgrRForbidens)
    End Sub

    '--------------------------------------------
    'Registers

    Private Sub cmdRAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRAdd.Click
        With frmForbidenEstatesAESRegisters
            .Text = "اضافه نمودن سند مالكيت"
            .Tag = "Add"
            .cmdOK.Tag = dgrRegisters
            .cmdCancel.Tag = 2
            .ShowDialog()
            If .Tag <> "OK" Then
                .Dispose()
                Exit Sub
            End If
            dgrRegisters.Rows.Add()
            dgrRegisters.Rows(dgrRegisters.Rows.Count - 1).Cells(1).Value = dgrRegisters.Rows.Count
            dgrRegisters.Rows(dgrRegisters.Rows.Count - 1).Cells(2).Value = .txtPageNo.Text
            dgrRegisters.Rows(dgrRegisters.Rows.Count - 1).Cells(3).Value = .txtBookNo.Text
            dgrRegisters.Rows(dgrRegisters.Rows.Count - 1).Cells(4).Value = .txtRegisterNo.Text
            dgrRegisters.Rows(dgrRegisters.Rows.Count - 1).Cells(5).Value = .txtRegisterDate.Text
            dgrRegisters.Rows(dgrRegisters.Rows.Count - 1).Cells(6).Value = .txtPrintNo.Text
            dgrRegisters.Rows(dgrRegisters.Rows.Count - 1).Cells(7).Value = .txtOwnerName.Text
            dgrRegisters.Rows(dgrRegisters.Rows.Count - 1).Cells(8).Value = .txtArea.Text
            dgrRegisters.Rows(dgrRegisters.Rows.Count - 1).Cells(9).Value = .txtLimits.Text
            dgrRegisters.Rows(dgrRegisters.Rows.Count - 1).Cells(10).Value = .txtNotes.Text
            .Dispose()
            dgrRegisters.ClearSelection()
            dgrRegisters.Rows(dgrRegisters.Rows.Count - 1).Selected = True
        End With
        Call InsertRowCode(dgrRegisters, 0)
    End Sub

    Private Sub cmdREdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdREdit.Click
        If dgrRegisters.SelectedRows.Count = 0 Then Exit Sub
        With frmForbidenEstatesAESRegisters
            .Text = "ويرايش سند مالكيت"
            .Tag = "Edit"
            .cmdOK.Tag = dgrRegisters
            .cmdCancel.Tag = 2
            .ShowDialog()
            If .Tag <> "OK" Then
                .Dispose()
                Exit Sub
            End If
            dgrRegisters.SelectedRows(0).Cells(2).Value = .txtPageNo.Text
            dgrRegisters.SelectedRows(0).Cells(3).Value = .txtBookNo.Text
            dgrRegisters.SelectedRows(0).Cells(4).Value = .txtRegisterNo.Text
            dgrRegisters.SelectedRows(0).Cells(5).Value = .txtRegisterDate.Text
            dgrRegisters.SelectedRows(0).Cells(6).Value = .txtPrintNo.Text
            dgrRegisters.SelectedRows(0).Cells(7).Value = .txtOwnerName.Text
            dgrRegisters.SelectedRows(0).Cells(8).Value = .txtArea.Text
            dgrRegisters.SelectedRows(0).Cells(9).Value = .txtLimits.Text
            dgrRegisters.SelectedRows(0).Cells(10).Value = .txtNotes.Text
            .Dispose()
            Dim Index As Integer = dgrRegisters.SelectedRows(0).Index
            dgrRegisters.ClearSelection()
            dgrRegisters.Rows(Index).Selected = True
        End With
    End Sub

    Private Sub cmdRDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRDelete.Click
        If dgrRegisters.SelectedRows.Count = 0 Then Exit Sub
        If ShowWarning("آيا مطمئن هستيد؟") <> "Yes" Then Exit Sub
        For Index As Integer = 0 To dgrTranses.Rows.Count - 1
            If CStr(dgrTranses.Rows(Index).Cells(0).Value) = CStr(dgrRegisters.SelectedRows(0).Cells(0).Value) Then dgrTranses.Rows.RemoveAt(Index)
        Next
        For Index As Integer = 0 To dgrRForbidens.Rows.Count - 1
            If CStr(dgrRForbidens.Rows(Index).Cells(0).Value) = CStr(dgrRegisters.SelectedRows(0).Cells(0).Value) Then dgrRForbidens.Rows.RemoveAt(Index)
        Next
        dgrRegisters.Rows.RemoveAt(dgrRegisters.SelectedRows(0).Index)
        If dgrRegisters.Rows.Count <> 0 Then dgrRegisters.Rows(0).Selected = True
    End Sub

    Private Sub dgrRegisters_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgrRegisters.SelectionChanged
        If dgrRegisters.SelectedRows.Count = 0 Then Exit Sub
        Dim blnVisible As Boolean = False, Index As Integer
        dgrRForbidens.ClearSelection()
        dgrTranses.ClearSelection()
        For Index = 0 To dgrTranses.Rows.Count - 1
            If CStr(dgrTranses.Rows(Index).Cells(0).Value) = CStr(dgrRegisters.SelectedRows(0).Cells(0).Value) Then dgrTranses.Rows(Index).Visible = True Else dgrTranses.Rows(Index).Visible = False
        Next
        For Index = 0 To dgrRForbidens.Rows.Count - 1
            If CStr(dgrRForbidens.Rows(Index).Cells(0).Value) = CStr(dgrRegisters.SelectedRows(0).Cells(0).Value) Then dgrRForbidens.Rows(Index).Visible = True Else dgrRForbidens.Rows(Index).Visible = False
        Next
        For Index = 0 To dgrRForbidens.Rows.Count - 1
            If dgrRForbidens.Rows(Index).Visible = True Then
                blnVisible = True
                Exit For
            End If
        Next
        If blnVisible Then dgrRForbidens.Rows(Index).Selected = True
        blnVisible = False
        For Index = 0 To dgrTranses.Rows.Count - 1
            If dgrTranses.Rows(Index).Visible = True Then
                blnVisible = True
                Exit For
            End If
        Next
        If blnVisible Then dgrTranses.Rows(Index).Selected = True
        Call FillVisibleRowNumberColumn(dgrTranses, 1)
    End Sub

    '--------------------------------------------
    'Transes

    Private Sub cmdTAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTAdd.Click
        If dgrRegisters.SelectedRows.Count = 0 Then
            Call ShowError("ابتدا سند مالكيت را انتخاب كنيد.")
            Exit Sub
        End If
        Call AddTrans(dgrTranses, 0, "اضافه نمودن نقل و انتقال", dgrRegisters.SelectedRows(0).Cells(0).Value)
    End Sub

    Private Sub cmdTEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTEdit.Click
        If dgrRegisters.SelectedRows.Count = 0 Then
            Call ShowError("ابتدا سند مالكيت را انتخاب كنيد.")
            Exit Sub
        End If
        Call EditTrans(dgrTranses, 0, "ويرايش نقل و انتقال")
    End Sub

    Private Sub cmdTDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTDelete.Click
        Call DeleteTrans(dgrTranses)
    End Sub

    Private Sub InsertRowCode(ByVal dgrMyGridView As DataGridView, ByVal intColumnNumber As Integer)
        Dim Index As Long, intNumber As Integer, intMax As Integer = 0
        For Index = 0 To dgrMyGridView.Rows.Count - 2
            If Mid(dgrMyGridView.Rows(Index).Cells(intColumnNumber).Value, 1, 1) = "U" Then
                intNumber = Mid(dgrMyGridView.Rows(Index).Cells(intColumnNumber).Value, 2)
                If intMax < intNumber Then intMax = intNumber
            End If
        Next
        intMax += 1
        dgrMyGridView.Rows(Index).Cells(intColumnNumber).Value = "U" & intMax
    End Sub

    Private Sub dgrRegisters_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgrRegisters.Sorted
        Call FillVisibleRowNumberColumn(dgrRegisters, 1)
    End Sub

    Private Sub dgrTranses_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgrTranses.Sorted
        Call FillVisibleRowNumberColumn(dgrTranses, 1)
    End Sub

    Private Sub FillVisibleRowNumberColumn(ByVal dgrMyGridView As DataGridView, Optional ByVal intColumnNumber As Integer = 0)
        Dim Index As Integer, intRowNumber As Integer = 0
        For Index = 0 To dgrMyGridView.Rows.Count - 1
            If dgrMyGridView.Rows(Index).Visible = True Then intRowNumber += 1
            dgrMyGridView.Rows(Index).Cells(intColumnNumber).Value = intRowNumber
        Next
    End Sub

End Class