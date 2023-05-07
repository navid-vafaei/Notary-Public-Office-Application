Module modBuildDocs
    'All of subs dependts to frmDocuments.Don't move them!

    Public Function BuildDocumentNote(ByVal lngDocumentCode As Long, ByVal blnShowDocument As Boolean) As Boolean
        Dim dtDocument As New DataTable, dtPersons1 As New DataTable, dtPersons2 As New DataTable, dtPersons3 As New DataTable, dtCompanies1 As New DataTable, dtCompanies2 As New DataTable, dtCompanies3 As New DataTable, dtObjects As New DataTable, dtCars As New DataTable, dtAgriCars As New DataTable, dtWayCars As New DataTable, dtMotors As New DataTable, dtHouses As New DataTable, dtPronouns As New DataTable, dadNotary As SqlClient.SqlDataAdapter
        Dim Index As Integer, lngMyCode As String, blnIsPerson As Boolean, strType As String, strFieldNames() As String, rtbDocument As New RichTextBox, strText As String, prgProg As New ProgressBar

        BuildDocumentNote = False
        'If IsWordWindowOpen(lngDocumentCode) Then Exit Function
        frmDocuments.cmdBuild.Enabled = False
        frmDocuments.cmdView.Enabled = False
        Application.DoEvents()

        prgProg.Height = 10
        prgProg.Left = frmDocuments.dgrDocuments.Left
        prgProg.Top = frmDocuments.dgrDocuments.Top + frmDocuments.dgrDocuments.Height - prgProg.Height
        prgProg.Width = frmDocuments.dgrDocuments.Width
        prgProg.Parent = frmDocuments
        prgProg.BringToFront()
        prgProg.Visible = True
        Application.DoEvents()

        '----------------------------------------
        '1-Fill Tables
        dadNotary = New SqlClient.SqlDataAdapter("SELECT D.Code, B.Code, K.Name, T.Name, B.Name, D.TempNo, D.PageNo, D.SectionNo, D.DocumentNo, D.DocumentDate, D.DocumentTime, D.CompleteDate, CASE D.IsGoverment WHEN 0 THEN 'غيردولتي' ELSE 'دولتي' END, D.DocumentPrice, D.RegisterIncome, D.EditIncome, D.OtherIncome, D.TaxPrice, D.PagesIncome, D.PagesCount, D.RefrenceNo, D.RefrenceDate, CASE D.IsPaid WHEN 0 THEN 'تسويه شده' ELSE 'تسويه نشده' END, D.Reduction, D.Notes, T.TypeCode FROM Documents D, DocumentKinds K, DocumentTypes T, DocumentBranches B WHERE D.DocumentBranchesCode = B.Code AND B.DocumentTypesCode = T.Code AND T.DocumentKindsCode = K.Code AND D.Code = " & lngDocumentCode, cnnNotary)
        dadNotary.SelectCommand.CommandTimeout = 0
        dadNotary.Fill(dtDocument)
        dadNotary = New SqlClient.SqlDataAdapter("SELECT P.Code, CASE P.Sex WHEN 0 THEN 'خانم' ELSE 'آقاي' END, P.FName, P.LName, P.FaName, P.IDNo, PL.Name AS MyPlace, P.IDSerial, P.BDate, C.Name AS MyCity, P.NationalNo, M.Name, M.DocumentText AS MilitaryName, P.MilitaryNo, P.Tel, P.Address, P.PostalCode, P.EMail, P.Notes, R.Name, R.RelationCaption, A.RelationNotes, A.RowNumber, A.ActorType FROM DocumentActs A, Persons P, Militaries M, Places PL, Cities C, Relations R WHERE A.IsPerson = 1 AND A.EntityCode = P.Code AND A.RelationsCode = R.Code AND A.DocumentsCode = " & lngDocumentCode & " AND P.IDPlacesCode = PL.Code AND P.BCitiesCode = C.Code AND P.MilitariesCode = M.Code AND A.ActorType = 1 ORDER BY A.RowNumber ASC", cnnNotary)
        dadNotary.Fill(dtPersons1)
        dadNotary = New SqlClient.SqlDataAdapter("SELECT P.Code, CASE P.Sex WHEN 0 THEN 'خانم' ELSE 'آقاي' END, P.FName, P.LName, P.FaName, P.IDNo, PL.Name AS MyPlace, P.IDSerial, P.BDate, C.Name AS MyCity, P.NationalNo, M.Name, M.DocumentText AS MilitaryName, P.MilitaryNo, P.Tel, P.Address, P.PostalCode, P.EMail, P.Notes, R.Name, R.RelationCaption, A.RelationNotes, A.RowNumber, A.ActorType FROM DocumentActs A, Persons P, Militaries M, Places PL, Cities C, Relations R WHERE A.IsPerson = 1 AND A.EntityCode = P.Code AND A.RelationsCode = R.Code AND A.DocumentsCode = " & lngDocumentCode & " AND P.IDPlacesCode = PL.Code AND P.BCitiesCode = C.Code AND P.MilitariesCode = M.Code AND A.ActorType = 2 ORDER BY A.RowNumber ASC", cnnNotary)
        dadNotary.Fill(dtPersons2)
        dadNotary = New SqlClient.SqlDataAdapter("SELECT P.Code, CASE P.Sex WHEN 0 THEN 'خانم' ELSE 'آقاي' END, P.FName, P.LName, P.FaName, P.IDNo, PL.Name AS MyPlace, P.IDSerial, P.BDate, C.Name AS MyCity, P.NationalNo, M.Name, M.DocumentText AS MilitaryName, P.MilitaryNo, P.Tel, P.Address, P.PostalCode, P.EMail, P.Notes, R.Name, R.RelationCaption, A.RelationNotes, A.RowNumber, A.ActorType FROM DocumentActs A, Persons P, Militaries M, Places PL, Cities C, Relations R WHERE A.IsPerson = 1 AND A.EntityCode = P.Code AND A.RelationsCode = R.Code AND A.DocumentsCode = " & lngDocumentCode & " AND P.IDPlacesCode = PL.Code AND P.BCitiesCode = C.Code AND P.MilitariesCode = M.Code AND A.ActorType = 3 ORDER BY A.RowNumber ASC", cnnNotary)
        dadNotary.Fill(dtPersons3)
        dadNotary = New SqlClient.SqlDataAdapter("SELECT C.Code, C.Name, C.RegisterNo, C.ChiefName, C.ChiefPostName, CI.Name AS CityName, CASE C.IsPrivate WHEN 0 THEN 'دولتي' ELSE 'خصوصي' END, C.AdvertisementNo, C.SignNames, C.Tel1, C.Tel2, C.Fax, C.CIN, C.Address, C.PostalCode, C.WebSites, C.Notes, R.Name, R.RelationCaption, A.RelationNotes, A.RowNumber, A.ActorType FROM Companies C, Cities CI, DocumentActs A, Relations R WHERE C.RegisterCitiesCode = CI.Code AND A.IsPerson = 0 AND A.EntityCode = C.Code AND A.RelationsCode = R.Code AND A.DocumentsCode = " & lngDocumentCode & " AND A.ActorType = 1 ORDER BY A.RowNumber ASC", cnnNotary)
        dadNotary.Fill(dtCompanies1)
        dadNotary = New SqlClient.SqlDataAdapter("SELECT C.Code, C.Name, C.RegisterNo, C.ChiefName, C.ChiefPostName, CI.Name AS CityName, CASE C.IsPrivate WHEN 0 THEN 'دولتي' ELSE 'خصوصي' END, C.AdvertisementNo, C.SignNames, C.Tel1, C.Tel2, C.Fax, C.CIN, C.Address, C.PostalCode, C.WebSites, C.Notes, R.Name, R.RelationCaption, A.RelationNotes, A.RowNumber, A.ActorType FROM Companies C, Cities CI, DocumentActs A, Relations R WHERE C.RegisterCitiesCode = CI.Code AND A.IsPerson = 0 AND A.EntityCode = C.Code AND A.RelationsCode = R.Code AND A.DocumentsCode = " & lngDocumentCode & " AND A.ActorType = 2 ORDER BY A.RowNumber ASC", cnnNotary)
        dadNotary.Fill(dtCompanies2)
        dadNotary = New SqlClient.SqlDataAdapter("SELECT C.Code, C.Name, C.RegisterNo, C.ChiefName, C.ChiefPostName, CI.Name AS CityName, CASE C.IsPrivate WHEN 0 THEN 'دولتي' ELSE 'خصوصي' END, C.AdvertisementNo, C.SignNames, C.Tel1, C.Tel2, C.Fax, C.CIN, C.Address, C.PostalCode, C.WebSites, C.Notes, R.Name, R.RelationCaption, A.RelationNotes, A.RowNumber, A.ActorType FROM Companies C, Cities CI, DocumentActs A, Relations R WHERE C.RegisterCitiesCode = CI.Code AND A.IsPerson = 0 AND A.EntityCode = C.Code AND A.RelationsCode = R.Code AND A.DocumentsCode = " & lngDocumentCode & " AND A.ActorType = 3 ORDER BY A.RowNumber ASC", cnnNotary)
        dadNotary.Fill(dtCompanies3)
        dadNotary = New SqlClient.SqlDataAdapter("SELECT O.Code, E.FieldValue, O.FieldTypesCode FROM DocumentObjects O, DocumentEntities E WHERE O.DocumentBranchesCode = " & dtDocument.Rows(0).Item(1) & " AND O.Code = E.DocumentObjectsCode AND E.DocumentsCode = " & lngDocumentCode & " ORDER BY O.Code ASC", cnnNotary)
        dadNotary.Fill(dtObjects)
        If Mid(dtDocument.Rows(0).Item(25), 1, 1) = "1" Then
            dadNotary = New SqlClient.SqlDataAdapter("SELECT C.Code, C.Notes, C.Amount, K.Name, S.Name, CASE T.IsImport WHEN 0 THEN 'خارجي' ELSE 'داخلي' END, T.Name, C.Model, C.BuyerPelakNo, C.BuyerPelakSeri, C.BuyerPelakColor, C.SellerPelakNo, C.SellerPelakSeri, C.SellerPelakColor, C.Color, C.VIN, C.MotorNo, C.ShasyNo, C.DocNo, C.DocDate, R1.Name, C.PunishmentNo, C.PunishmentDate, C.FuelNo, C.TollNo, C.TollDate, C.TollPrice, R2.Name, C.GreenNo, C.GreenDate, R3.Name, R4.Name, C.InsuranceNo, C.InsuranceDate, C.InsuranceExit, R5.Name, C.MoveNo, C.MoveDate, R6.Name, C.TireNo, C.Capacity, C.SylandrNo FROM CarKinds K, CarSystems S, CarTypes T, Cars C, Refrences R1, Refrences R2, Refrences R3, Refrences R4, Refrences R5, Refrences R6 WHERE C.CarTypesCode = T.Code AND T.CarSystemsCode = S.Code AND S.CarKindsCode = K.Code AND C.DocRefrencesCode = R1.Code AND C.TollRefrencesCode = R2.Code AND C.PelakRefrencesCode = R3.Code AND C.CompanyRefrencesCode = R4.Code AND C.InsuranceRefrencesCode = R5.Code AND C.MoveRefrencesCode = R6.Code AND C.DocumentsCode = " & lngDocumentCode & " ORDER BY C.RowNumber ASC", cnnNotary)
            dadNotary.Fill(dtCars)
        End If
        If Mid(dtDocument.Rows(0).Item(25), 2, 1) = "1" Then
            dadNotary = New SqlClient.SqlDataAdapter("SELECT C.Code, C.Notes, C.Amount, K.Name, S.Name, CASE T.IsImport WHEN 0 THEN 'خارجي' ELSE 'داخلي' END, T.Name, C.Model, C.MotorNo, C.ShasyNo, C.DocNo, C.DocDate, R1.Name, C.FuelNo, C.TollNo, C.TollDate, C.TollPrice, R2.Name, C.GreenNo, C.GreenDate, R3.Name, C.MoveNo, C.MoveDate, R4.Name, C.TireNo FROM AgriCarKinds K, AgriCarSystems S, AgriCarTypes T, AgriCars C, Refrences R1, Refrences R2, Refrences R3, Refrences R4 WHERE C.AgriCarTypesCode = T.Code AND T.AgriCarSystemsCode = S.Code AND S.AgriCarKindsCode = K.Code AND C.DocRefrencesCode = R1.Code AND C.TollRefrencesCode = R2.Code AND C.CompanyRefrencesCode = R3.Code AND C.MoveRefrencesCode = R4.Code AND C.DocumentsCode = " & lngDocumentCode & " ORDER BY C.RowNumber ASC", cnnNotary)
            dadNotary.Fill(dtAgriCars)
        End If
        If Mid(dtDocument.Rows(0).Item(25), 3, 1) = "1" Then
            dadNotary = New SqlClient.SqlDataAdapter("SELECT C.Code, C.Notes, C.Amount, K.Name, S.Name, CASE T.IsImport WHEN 0 THEN 'خارجي' ELSE 'داخلي' END, T.Name, C.Model, C.MotorNo, C.ShasyNo, C.DocNo, C.DocDate, R1.Name, C.FuelNo, C.TollNo, C.TollDate, C.TollPrice, R2.Name, C.GreenNo, C.GreenDate, R3.Name, C.MoveNo, C.MoveDate, R4.Name, C.TireNo FROM WayCarKinds K, WayCarSystems S, WayCarTypes T, WayCars C, Refrences R1, Refrences R2, Refrences R3, Refrences R4 WHERE C.WayCarTypesCode = T.Code AND T.WayCarSystemsCode = S.Code AND S.WayCarKindsCode = K.Code AND C.DocRefrencesCode = R1.Code AND C.TollRefrencesCode = R2.Code AND C.CompanyRefrencesCode = R3.Code AND C.MoveRefrencesCode = R4.Code AND C.DocumentsCode = " & lngDocumentCode & " ORDER BY C.RowNumber ASC", cnnNotary)
            dadNotary.Fill(dtWayCars)
        End If
        If Mid(dtDocument.Rows(0).Item(25), 4, 1) = "1" Then
            dadNotary = New SqlClient.SqlDataAdapter("SELECT C.Code, C.Notes, C.Amount, K.Name, S.Name, CASE S.IsImport WHEN 0 THEN 'خارجي' ELSE 'داخلي' END, C.Model, C.BuyerPelakNo, C.BuyerPelakSeri, C.BuyerPelakColor, C.SellerPelakNo, C.SellerPelakSeri, C.SellerPelakColor, C.Color, C.VIN, C.MotorNo, C.ShasyNo, C.UnderUser, C.DocNo, C.DocDate, R1.Name, C.PunishmentNo, C.PunishmentDate, C.FuelNo, C.GreenNo, C.GreenDate, R2.Name, R3.Name, C.InsuranceNo, C.InsuranceDate, C.InsuranceExit, R4.Name, C.MoveNo, C.MoveDate, R5.Name, C.Capacity, C.SylandrNo FROM MotorKinds K, MotorSystems S, Motors C, Refrences R1, Refrences R2, Refrences R3, Refrences R4, Refrences R5 WHERE C.MotorSystemsCode = S.Code AND S.MotorKindsCode = K.Code AND C.DocRefrencesCode = R1.Code AND C.PelakRefrencesCode = R2.Code AND C.CompanyRefrencesCode = R3.Code AND C.InsuranceRefrencesCode = R4.Code AND MoveRefrencesCode = R5.Code AND C.DocumentsCode = " & lngDocumentCode & " ORDER BY C.RowNumber ASC", cnnNotary)
            dadNotary.Fill(dtMotors)
        End If
        If Mid(dtDocument.Rows(0).Item(25), 5, 1) = "1" Then
            dadNotary = New SqlClient.SqlDataAdapter("SELECT H.Code, MainPelak, LocalPelak, P.Name, C.Name, S.Name, HT.Name, H.Amount, H.Bargain, H.SplitedNo, H.Place, H.Address, H.Tel, H.Area, H.FloorNo, H.Height, H.Notes1, H.NorthNotes, H.EastNotes, H.SouthNotes, H.WestNotes, H.HouseRights, H.Notes2, Case WHEN H.IsShare = 0 THEN '' ELSE 'با قدرالسهم هر يک از عرصه مشاعي و مشاعات' END, H.ShareNotes, H.Notes3, H.MoveNo, H.MoveDate, R1.Name, H.MoveNotes, H.Notes4, H.DocCount, H.QuesNo, H.QuesDate, R2.Name, H.Notes5, H.RahnyDate, H.RahnyDuration, H.RahnyPrice, H.Notes6, CASE H.BenefitType WHEN 1 THEN 'منافع مورد معامله قبلاً به کسي واگذار نشده است و در تصرف خريدار است' WHEN 2 THEN 'منافع مورد معامله در حال حاضر در تصرف فروشنده است' WHEN 3 THEN 'منافع مورد معامله مادام الحيات متعلق به فروشنده و پس از فوت ايشان تابع عين است' WHEN 4 THEN '' END, H.EmptyDate, H.ChequeNo, H.EmptyDays, H.HouseCode, H.Notes7, H.MuniNo, H.MuniDate, R3.Name, CASE H.IsMuni100 WHEN 0 THEN '' ELSE 'در اجراي تبصره 8 ماده صد قانون شهرداري صادر شده است' END, H.Notes8, CASE H.IsNoChange WHEN 0 THEN '' ELSE 'هرگونه تغيير نسبت به مندرجات سند و مشاعات مورد تائيد شهرداري نمي باشد' END, H.TafkikNo, H.TafkikDate, R4.Name, H.TafkikNotes, H.Notes13, H.AccountNo, H.AccountDate, R5.Name, H.Notes9, CASE H.IsUnderstandAccount WHEN 0 THEN '' ELSE 'مفاد نامه به خريدار تفهيم شد و نامبرده با علم آگاهي كامل از آن و از كميت و كيفيت مورد معامله اقدام به خريد آن نموده است' END, CASE H.IsNotForbiden WHEN 0 THEN '' ELSE 'بنا به اظهار طرفين متعاملين هيچ يک از اشخاص ممنوع المعامله نمي باشند و مسئوليت آن را متضامناً به عهده گرفتند' END, CASE H.IsNoError WHEN 0 THEN '' ELSE 'فروشنده به موجب اين سند خريدار را نماينده خود قرار داد که در صورت بروز هر گونه اشتباه قلمي در تنظيم اين سند با حضور در دفترخانه اسناد رسمي نسبت به تنظيم و امضاي سند اصلاحي بدون تغيير در ارکان و ماهيت سند اقدام نمايند' END, H.TaxGovahyNo, H.TaxGovahyDate, R6.Name, CASE H.IsTaxGovahy187 WHEN 0 THEN '' ELSE 'در اجراي ماده 187 قانون مالياتهاي مستقيم صادر گرديده است' END, H.Notes10, CASE H.IsNotStore WHEN 0 THEN '' ELSE 'مورد معامله فاقد ارزش تجاري و سرقفلي مي باشد' END, CASE H.IsUnderstandGovahyTax WHEN 0 THEN '' ELSE 'مفاد گواهي مزبور کاملاً به خريدار تفهيم شد' END, H.NotArrestNo, H.NotArrestDate, R7.Name, H.Notes12 FROM Houses H, HouseTypes HT, Provinces P, Cities C, Sections S, Refrences R1, Refrences R2, Refrences R3, Refrences R4, Refrences R5, Refrences R6, Refrences R7 WHERE H.HouseTypesCode = HT.Code AND H.ProvincesCode = P.Code AND H.SectionsCode = S.Code AND S.CitiesCode = C.Code AND H.MoveRefrencesCode = R1.Code AND H.QuesRefrencesCode = R2.Code AND H.MuniRefrencesCode = R3.Code AND H.TafkikRefrencesCode = R4.Code AND H.AccountRefrencesCode = R5.Code AND H.TaxGovahyRefrencesCode = R6.Code AND H.NotArrestRefrencesCode = R7.Code AND H.DocumentsCode = " & lngDocumentCode & " ORDER BY H.RowNumber ASC", cnnNotary)
            dadNotary.Fill(dtHouses)
        End If
        dadNotary = New SqlClient.SqlDataAdapter("SELECT Code, M1Name, M2Name, M3Name, F1Name, F2Name, F3Name FROM Pronouns ORDER BY Code ASC", cnnNotary)
        dadNotary.Fill(dtPronouns)
        'If IsWordWindowOpen(dtDocument.Rows(0).Item(1)) Then Exit Function


        cmmNotary.CommandText = "SELECT NoteFields FROM DocumentBranches WHERE Code = " & dtDocument.Rows(0).Item(1)
        strFieldNames = Split(cmmNotary.ExecuteScalar, ",")
        If strFieldNames.Length = 0 Then GoTo lblEnd
        prgProg.Value = 20
        Application.DoEvents()

        Dim Keys(0) As DataColumn
        Keys(0) = dtObjects.Columns(0)
        dtObjects.PrimaryKey = Keys

        Dim dtIfs As New DataTable
        dadNotary = New SqlClient.SqlDataAdapter("SELECT FromChar, MyLength, FromObject, ObjCount, FieldName FROM DocumentNotesIfs WHERE DocumentBranchesCode = " & dtDocument.Rows(0).Item(1), cnnNotary)
        dadNotary.Fill(dtIfs)
        Dim blnIfs(dtIfs.Rows.Count - 1) As Boolean, intCount As Integer, blnDelObjects(strFieldNames.Length - 2) As Boolean
        For Index = 0 To dtIfs.Rows.Count - 1
            cmmNotary.CommandText = IIf(Mid(dtIfs.Rows(Index).Item(4), 1, 1) = "E", "SELECT Code FROM DocumentEntities WHERE " & dtIfs.Rows(Index).Item(4), "SELECT D.Code FROM Documents D WHERE Code = " & dtDocument.Rows(0).Item(0) & " AND (" & dtIfs.Rows(Index).Item(4) & ")")
            blnIfs(Index) = IIf(IsNothing(cmmNotary.ExecuteScalar), False, True)
        Next
        For Index = 0 To blnIfs.Length - 1
            If blnIfs(Index) = False Then
                For intCount = dtIfs.Rows(Index).Item(2) To dtIfs.Rows(Index).Item(2) + dtIfs.Rows(Index).Item(3) - 1
                    blnDelObjects(intCount - 1) = True
                Next
            End If
        Next

        'Call CheckWordApp(False)
        ' SetParent(ptrWordHandle, 0)
        Application.DoEvents()
        'While WordApp.Documents.Count >= 1
        'WordApp.Documents(1).Close(False)
        'End While
        Dim strPath As String = strNotaryPath & "\Docs\" & dtDocument.Rows(0).Item(1) & ".rtf"
        If Not IO.File.Exists(strPath) Then rtbDocument.SaveFile(strPath)
        Try
            rtbDocument.LoadFile(strPath)
        Catch ex As Exception
            Call CheckWordApp(False)
            While WordApp.Documents.Count >= 1
                WordApp.Documents(1).Close(False)
            End While
            Try
                rtbDocument.LoadFile(strPath)
            Catch ex1 As Exception
                Call ShowError("سند نتوانست ساخته شود")
            End Try
        End Try
        intCount = 0
        For Index = 0 To blnIfs.Length - 1
            rtbDocument.SelectionStart = dtIfs.Rows(Index).Item(0) - intCount
            rtbDocument.SelectionLength = dtIfs.Rows(Index).Item(1)
            If blnIfs(Index) Then
                rtbDocument.SelectionBackColor = rtbDocument.BackColor
            Else
                rtbDocument.SelectedText = ""
                intCount += dtIfs.Rows(Index).Item(1)
            End If
        Next
        dtIfs.Clear()
        prgProg.Value = 30
        Application.DoEvents()

        '----------------------------------------
        '2-Begin                
        If rtbDocument.Text <> "" Then
            rtbDocument.SelectionStart = 0
            rtbDocument.SelectionLength = 0
            For Index = 0 To strFieldNames.Length - 2
                If blnDelObjects(Index) = False Then
                    rtbDocument.Find("!!$##")
                    rtbDocument.SelectedText = ""
                    strType = Mid(strFieldNames(Index), 1, 1)
                    Select Case strType
                        Case "A"
                            If Not strFieldNames(Index).Contains("-") Then strFieldNames(Index) = Mid(strFieldNames(Index), 1, 1) & "0-" & Mid(strFieldNames(Index), 2)
                            lngMyCode = Mid(strFieldNames(Index), InStr(strFieldNames(Index), "-") + 1)
                            intCount = Mid(strFieldNames(Index), 2, InStr(strFieldNames(Index), "-") - 1)
                            Call BuildActText(rtbDocument, dtDocument.Rows(0), dtPersons1, dtCompanies1, lngMyCode, intCount)
                        Case "B"
                            If Not strFieldNames(Index).Contains("-") Then strFieldNames(Index) = Mid(strFieldNames(Index), 1, 1) & "0-" & Mid(strFieldNames(Index), 2)
                            lngMyCode = Mid(strFieldNames(Index), InStr(strFieldNames(Index), "-") + 1)
                            intCount = Mid(strFieldNames(Index), 2, InStr(strFieldNames(Index), "-") - 1)
                            Call BuildActText(rtbDocument, dtDocument.Rows(0), dtPersons2, dtCompanies2, lngMyCode, intCount)
                        Case "G"
                            If Not strFieldNames(Index).Contains("-") Then strFieldNames(Index) = Mid(strFieldNames(Index), 1, 1) & "0-" & Mid(strFieldNames(Index), 2)
                            lngMyCode = Mid(strFieldNames(Index), InStr(strFieldNames(Index), "-") + 1)
                            intCount = Mid(strFieldNames(Index), 2, InStr(strFieldNames(Index), "-") - 1)
                            Call BuildActText(rtbDocument, dtDocument.Rows(0), dtPersons3, dtCompanies3, lngMyCode, intCount)
                        Case "C"
                            If dtObjects.Rows.Count <> 0 Then
                                lngMyCode = Mid(strFieldNames(Index), 2)
                                Dim drMyObject As DataRow = dtObjects.Rows.Find(lngMyCode)
                                If Not IsNothing(drMyObject) Then 'ممکن است شيء در متن سند تازه ايجاد شده باشد و در سند مقداري براي آن ايجاد نشده باشد
                                    If drMyObject.Item(1) <> "" Then
                                        CopyTextToClipbaord(IIf(drMyObject.Item(2) = 4, ReverseMyDate(drMyObject.Item(1)), drMyObject.Item(1)))
                                        rtbDocument.Paste()
                                    End If
                                End If
                            End If
                        Case "D"
                            lngMyCode = Mid(strFieldNames(Index), 2)
                            strText = ReturnDocumentEntryNameFromTable(dtDocument.Rows(0), lngMyCode)
                            If strText <> "" Then
                                Call CopyTextToClipbaord(strText)
                                rtbDocument.Paste()
                            End If
                        Case "E", "F", "H"
                            Dim dtTempPersons As DataTable, dtTempCompanies As DataTable
                            If dtPronouns.Rows.Count <> 0 Then
                                lngMyCode = Mid(strFieldNames(Index), 2)
                                If strType = "E" Then
                                    dtTempPersons = dtPersons1
                                    dtTempCompanies = dtCompanies1
                                ElseIf strType = "F" Then
                                    dtTempPersons = dtPersons2
                                    dtTempCompanies = dtCompanies2
                                Else
                                    dtTempPersons = dtPersons3
                                    dtTempCompanies = dtCompanies3
                                End If
                                Dim intLeft As Integer = 0, intRight As Integer = dtPronouns.Rows.Count, Index1 As Integer
                                Index1 = (intLeft + intRight) / 2
                                While dtPronouns.Rows(Index1).Item(0) <> lngMyCode
                                    If dtPronouns.Rows(Index1).Item(0) < lngMyCode Then intLeft = Index1 Else intRight = Index1
                                    Index1 = (intLeft + intRight) / 2
                                End While
                                Select Case dtTempPersons.Rows.Count + dtTempCompanies.Rows.Count
                                    Case Is <= 1 'تعداد افراد يا شرکتها يک نفر يا هيچ يعني بدون متعمال
                                        If dtTempPersons.Rows.Count = 1 Then If dtTempPersons.Rows(0).Item(1) = "آقاي" Then CopyTextToClipbaord(dtPronouns.Rows(Index1).Item(1)) Else CopyTextToClipbaord(dtPronouns.Rows(Index1).Item(4)) Else CopyTextToClipbaord(dtPronouns.Rows(Index1).Item(1))
                                    Case Is = 2 'تعداد افراد يا شرکتها دو نفر
                                        If dtTempCompanies.Rows.Count = 0 Then
                                            If dtTempPersons.Rows(0).Item(1) = "آقاي" Or dtTempPersons.Rows(1).Item(1) = "آقاي" Then CopyTextToClipbaord(dtPronouns.Rows(Index1).Item(2)) Else CopyTextToClipbaord(dtPronouns.Rows(Index1).Item(5))
                                        Else
                                            CopyTextToClipbaord(dtPronouns.Rows(Index1).Item(1))
                                        End If
                                    Case Is >= 3 'تعداد افراد يا شرکتها سه نفر يا بيشتر
                                        If dtTempCompanies.Rows.Count = 0 Then
                                            blnIsPerson = False
                                            For intCount = 0 To dtTempPersons.Rows.Count - 1
                                                If dtTempPersons.Rows(intCount).Item(1) = "آقاي" Then
                                                    blnIsPerson = True
                                                    Exit For
                                                End If
                                            Next
                                            If blnIsPerson Then CopyTextToClipbaord(dtPronouns.Rows(Index1).Item(3)) Else CopyTextToClipbaord(dtPronouns.Rows(Index1).Item(6))
                                        Else
                                            CopyTextToClipbaord(dtPronouns.Rows(Index1).Item(3))
                                        End If
                                End Select
                                rtbDocument.Paste()
                            End If
                        Case "I" : Call PasteMyItem(rtbDocument, dtCars, strFieldNames(Index))
                        Case "K" : Call PasteMyItem(rtbDocument, dtAgriCars, strFieldNames(Index))
                        Case "L" : Call PasteMyItem(rtbDocument, dtWayCars, strFieldNames(Index))
                        Case "M" : Call PasteMyItem(rtbDocument, dtMotors, strFieldNames(Index))
                        Case "J" : Call PasteMyItem(rtbDocument, dtHouses, strFieldNames(Index))
                        Case "S" : Call PasteMyPattern(rtbDocument, dtCars, dtDocument.Rows(0), strFieldNames(Index))
                        Case "T" : Call PasteMyPattern(rtbDocument, dtAgriCars, dtDocument.Rows(0), strFieldNames(Index))
                        Case "U" : Call PasteMyPattern(rtbDocument, dtWayCars, dtDocument.Rows(0), strFieldNames(Index))
                        Case "V" : Call PasteMyPattern(rtbDocument, dtMotors, dtDocument.Rows(0), strFieldNames(Index))
                        Case "W" : Call PasteMyPattern(rtbDocument, dtHouses, dtDocument.Rows(0), strFieldNames(Index))
                        Case "+" : Call PasteMyPattern(rtbDocument, dtDocument, dtDocument.Rows(0), strFieldNames(Index))
                        Case "/"
                            Dim rtbTemp As New RichTextBox
                            rtbTemp.LoadFile(strNotaryPath & "\Pats\" & Mid(strFieldNames(Index), 2) & ".rtf")
                            rtbTemp.SelectAll()
                            rtbTemp.Copy()
                            rtbDocument.Paste()
                            rtbTemp.Dispose()
                    End Select
                End If
                prgProg.Value = 40 + Index / strFieldNames.Length * 40
                Application.DoEvents()
            Next
        End If
        strPath = GetDocumentPath(lngDocumentCode, True)
        Try
            rtbDocument.SaveFile(strPath)
        Catch ex As Exception
            Call CheckWordApp(False)
            While WordApp.Documents.Count >= 1
                WordApp.Documents(1).Close(False)
            End While
            Try
                rtbDocument.SaveFile(strPath)
            Catch ex1 As Exception
                Call ShowError("سند نتوانست ساخته شود")
            End Try
        End Try
        Erase strFieldNames
        Erase blnIfs
        Erase blnDelObjects
        prgProg.Value = 100
        Application.DoEvents()
        dtPersons1.Rows.Clear()
        dtPersons2.Rows.Clear()
        dtPersons3.Rows.Clear()
        dtCompanies1.Rows.Clear()
        dtCompanies3.Rows.Clear()
        dtCompanies2.Rows.Clear()
        dtCars.Rows.Clear()
        dtAgriCars.Rows.Clear()
        dtWayCars.Rows.Clear()
        dtMotors.Rows.Clear()
        dtHouses.Rows.Clear()
        dtObjects.Rows.Clear()
        dtDocument.Rows.Clear()
        dtPronouns.Rows.Clear()
        prgProg.Value = 0
        Application.DoEvents()
        prgProg.Dispose()
        BuildDocumentNote = True

        frmDocuments.cmdBuild.Enabled = True
        frmDocuments.cmdView.Enabled = True
        Exit Function

        'Open File and set changes and view it
        If blnShowDocument = False Then
        End If
        Call CheckWordApp(False)
        Dim MyDoc As Object = WordApp.Documents.Open(strPath)
        WordApp.ShowWindowsInTaskbar = True
        MyDoc.Application.Selection.WholeStory()
        MyDoc.Application.Selection.ParagraphFormat.Alignment = 3
        Dim frpDocument As New FastReport.Report
        frpDocument.Password = "qaz0qaz"
        frpDocument.Load(strNotaryPath & "\Reports\Paper.frx")
        Dim MyReportPage As FastReport.ReportPage = frpDocument.AllObjects(0), MyReportHeader As FastReport.PageHeaderBand = frpDocument.FindObject("PageHeader1"), MyReportFooter As FastReport.PageFooterBand = frpDocument.FindObject("PageFooter1")
        With MyDoc.PageSetup
            .TopMargin = 28.34646 * (MyReportPage.TopMargin / 10 + (Math.Round(MyReportHeader.Height / 37.795275591, 2)))
            .BottomMargin = 28.34646 * (MyReportPage.BottomMargin / 10 + (Math.Round(MyReportFooter.Height / 37.795275591, 2)))
            .LeftMargin = 2.834646 * MyReportPage.LeftMargin
            .RightMargin = 2.834646 * MyReportPage.RightMargin
            .Gutter = 28.34646 * 0
            .HeaderDistance = 28.34646 * 1.27
            .FooterDistance = 28.34646 * 1.27
            .PageWidth = 2.834646 * MyReportPage.PaperWidth
            .PageHeight = 2.834646 * MyReportPage.PaperHeight
        End With
        WordApp.Selection.ParagraphFormat().RightIndent = 0
        WordApp.Selection.ParagraphFormat().LeftIndent = 0
        Clipboard.Clear()
        MyDoc.UndoClear()
        MyDoc.Application.Selection.HomeKey(Unit:=6)
        If WordApp.ActiveWindow.View.SplitSpecial <> 0 Then WordApp.ActiveWindow.Panes(2).Close()
        WordApp.ActiveWindow.ActivePane.View.Type = 3
        WordApp.ActiveWindow.ActivePane.View.SeekView = 0
        WordApp.ActiveWindow.ActivePane.View.Zoom.Percentage = 80
        MyDoc.Saved = True 'for disabling ask save changes?
        MyDoc.Save()
        frmWord.grpView.BringToFront()
        Call SetEnabledWordSave(True)
        frmWord.lblCode.Text = lngDocumentCode
        frmWord.lblPattern.Tag = False
        Call ShowWordForm(False)
        WordApp.Activate()
        frmWord.Refresh() 'Force word to redraw itself correctly
        frmDocuments.cmdBuild.Enabled = True
        frmDocuments.cmdView.Enabled = True
lblEnd:
    End Function

    Public Sub BuildActText(ByRef rtbDocument As RichTextBox, ByRef drDocument As DataRow, ByRef dtPersons As DataTable, ByRef dtCompanies As DataTable, ByVal lngMyCode As Long, ByVal intCount As Integer)
        Dim intActNumber As Integer, intPerson As Integer, intCompany As Integer, intTotal As Integer, intBetweenActs As Integer, Index As Integer, intFileNo As Integer = FreeFile(), blnFound As Boolean = False
        FileOpen(intFileNo, strNotaryPath & "\Data\Slkws.bqz", OpenMode.Random, OpenAccess.ReadWrite, OpenShare.Shared)
        FileGet(intFileNo, intBetweenActs, 4) 'بين متعاملين
        FileClose(intFileNo)
        If intCount = 0 Then
            intActNumber = 1
            intPerson = 0
            intCompany = 0
            intTotal = dtPersons.Rows.Count + dtCompanies.Rows.Count
            For Index = 0 To intTotal - 1
                If intPerson < dtPersons.Rows.Count Then
                    If dtPersons.Rows(intPerson).Item(22) = intActNumber Then
                        If dtPersons.Rows(intPerson).Item(23) = 1 Then 'فقط براي اصيل شماره بگذريم
                            If intTotal > 1 Then If intBetweenActs = 2 Then CopyTextToClipbaord(intActNumber & "- ") Else If intBetweenActs = 3 Then CopyTextToClipbaord(intActNumber & ")")
                        End If
                        Call BuildText(rtbDocument, drDocument, dtPersons.Rows(intPerson), lngMyCode, "N", "SELECT N.Code FROM Persons N, Militaries NM, Places NP, Cities NC, DocumentActs NA, Relations NR WHERE N.MilitariesCode = NM.Code AND N.IDPlacesCode = NP.Code AND N.BCitiesCode = NC.Code AND NA.DocumentsCode = " & drDocument.Item(0) & " AND NA.EntityCode = N.Code AND NA.RelationsCode = NR.Code AND N.Code = " & dtPersons.Rows(intPerson).Item(0) & " AND (")
                        If intTotal > 1 Then
                            If intBetweenActs = 1 Then
                                If intActNumber <> intTotal Then
                                    CopyTextToClipbaord(" و ")
                                    If HasClipboardData("Text") Then rtbDocument.Paste()
                                End If
                            End If
                        End If
                        intPerson += 1
                    End If
                End If
                If intCompany < dtCompanies.Rows.Count Then
                    If dtCompanies.Rows(intCompany).Item(20) = intActNumber Then
                        If dtCompanies.Rows(intCompany).Item(21) = 1 Then 'فقط براي اصيل شماره بگذريم
                            If intTotal > 1 Then If intBetweenActs = 2 Then CopyTextToClipbaord(intActNumber & "- ") Else If intBetweenActs = 3 Then CopyTextToClipbaord(intActNumber & ")")
                        End If
                        Call BuildText(rtbDocument, drDocument, dtCompanies.Rows(intCompany), lngMyCode, "O", "SELECT O.Code FROM Companies O, Cities OC, DocumentActs OA, Relations R WHERE O.RegisterCitiesCode = OC.Code AND OA.DocumentsCode = " & drDocument.Item(0) & " AND OA.EntityCode = O.Code AND OA.RelationsCode = R.Code AND O.Code = " & dtCompanies.Rows(intCompany).Item(0) & " AND (")
                        If intTotal > 1 Then
                            If intBetweenActs = 1 Then
                                If intActNumber <> intTotal Then
                                    CopyTextToClipbaord(" و ")
                                    If HasClipboardText() Then rtbDocument.Paste()
                                End If
                            End If
                        End If
                        intCompany += 1
                    End If
                End If
                intActNumber += 1
            Next
        Else
            For Index = 0 To dtPersons.Rows.Count - 1
                If dtPersons.Rows(1).Item(22) = intCount Then
                    Call BuildText(rtbDocument, drDocument, dtPersons.Rows(Index), lngMyCode, "N", "SELECT N.Code FROM Persons N, Militaries NM, Places NP, Cities NC, DocumentActs NA, Relations NR WHERE N.MilitariesCode = NM.Code AND N.IDPlacesCode = NP.Code AND N.BCitiesCode = NC.Code AND NA.DocumentsCode = " & drDocument.Item(0) & " AND NA.EntityCode = N.Code AND NA.RelationsCode = NR.Code AND N.Code = " & dtPersons.Rows(Index).Item(0) & " AND (")
                    blnFound = True
                    Exit For
                End If
            Next
            If blnFound = False Then
                For Index = 0 To dtCompanies.Rows.Count - 1
                    If dtCompanies.Rows(1).Item(20) = intCount Then
                        Call BuildText(rtbDocument, drDocument, dtCompanies.Rows(Index), lngMyCode, "O", "SELECT O.Code FROM Companies O, Cities OC, DocumentActs OA, Relations OR WHERE O.CitiesCode = OC.Code AND OA.DocumentsCode = " & drDocument.Item(0) & " AND OA.EntityCode = O.Code AND OA.RelationsCode = OR.Code AND O.Code = " & dtCompanies.Rows(Index).Item(0) & " AND (")
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub BuildText(ByRef rtbDocument As RichTextBox, ByRef drDocument As DataRow, ByRef drMyRow As DataRow, ByVal lngPatCode As Long, ByVal strObjectCode As String, ByVal strIfSql As String)
        Dim rtbTemp As New RichTextBox, strPath As String
        If strObjectCode = "O" Then 'If we want companies patterns we should change code
            cmmNotary.CommandText = "SELECT Code FROM Patterns WHERE Name LIKE 'C' + (SELECT SUBSTRING(Name,2,100) FROM Patterns WHERE Code = " & lngPatCode & ")"
            lngPatCode = cmmNotary.ExecuteScalar
        End If
        cmmNotary.CommandText = "SELECT NoteFields FROM Patterns WHERE Code = " & lngPatCode
        Dim strObjects() As String = Split(cmmNotary.ExecuteScalar, ",")
        Dim dadNotary As New SqlClient.SqlDataAdapter("SELECT FromChar, MyLength, FromObject, ObjCount, FieldName FROM PatternsIfs WHERE PatternsCode = " & lngPatCode, cnnNotary), dtTemp As New DataTable
        dadNotary.Fill(dtTemp)
        Dim blnIfs(dtTemp.Rows.Count - 1) As Boolean, Index As Integer, intCount As Integer, blnDelObjects(strObjects.Length - 2) As Boolean
        For Index = 0 To dtTemp.Rows.Count - 1
            cmmNotary.CommandText = IIf(Mid(dtTemp.Rows(Index).Item(4), 1, 1) = strObjectCode, strIfSql, "SELECT D.Code FROM Documents D WHERE Code = " & drDocument.Item(0) & " AND (") & dtTemp.Rows(Index).Item(4) & ")"
            blnIfs(Index) = IIf(IsNothing(cmmNotary.ExecuteScalar), False, True)
        Next
        For Index = 0 To blnIfs.Length - 1
            If blnIfs(Index) = False Then
                For intCount = dtTemp.Rows(Index).Item(2) To dtTemp.Rows(Index).Item(2) + dtTemp.Rows(Index).Item(3) - 1
                    blnDelObjects(intCount - 1) = True
                Next
            End If
        Next
        Try
            While WordApp.Documents.Count >= 1
                WordApp.Documents(1).Close(True)
            End While
        Catch ex As Exception
        End Try
        strPath = strNotaryPath & "\Pats\" & lngPatCode & ".rtf"
        If Not IO.File.Exists(strPath) Then rtbTemp.SaveFile(strPath)
        rtbTemp.LoadFile(strPath)

        intCount = 0
        For Index = 0 To blnIfs.Length - 1
            rtbTemp.SelectionStart = dtTemp.Rows(Index).Item(0) - intCount
            rtbTemp.SelectionLength = dtTemp.Rows(Index).Item(1)
            If blnIfs(Index) Then
                rtbTemp.SelectionBackColor = rtbTemp.BackColor
            Else
                rtbTemp.SelectedText = ""
                intCount += dtTemp.Rows(Index).Item(1)
            End If
        Next
        If rtbTemp.Text <> "" Then
            rtbTemp.SelectionStart = 0
            rtbTemp.SelectionLength = 0
            For Index = 0 To strObjects.Length - 2
                If blnDelObjects(Index) = False Then
                    rtbTemp.Find("!!$##")
                    Select Case Mid(strObjects(Index), 1, 1)
                        Case "D"
                            rtbTemp.SelectedText = ReturnDocumentEntryNameFromTable(drDocument, Mid(strObjects(Index), 2))
                        Case "P"
                            Dim dtHouseAdds As New DataTable
                            dadNotary = New SqlClient.SqlDataAdapter("SELECT A.Code, HT.Name, A.Amount, A.LocalPelak, A.Place, A.Area, A.Height, A.NorthNotes, A.EastNotes, A.SouthNotes, A.WestNotes, A.Notes FROM HouseAdds A, HouseTypes HT WHERE A.HouseTypesCode = HT.Code AND " & IIf(drMyRow.Table.Columns.Count >= 70, "A.HousesCode = " & drMyRow.Item(0), "A.Code = " & drMyRow.Item(0)), cnnNotary)
                            dadNotary.Fill(dtHouseAdds)
                            Call PasteMyItem(rtbTemp, dtHouseAdds, strObjects(Index))
                            dtHouseAdds.Clear()
                        Case "Q"
                            Dim dtHouseRegisters As New DataTable
                            dadNotary = New SqlClient.SqlDataAdapter("SELECT R.Code, R.PageNo, R.BookNo, R.RegisterNo, R.RegisterDate, R.PrintNo, R.OwnerName, R.Area, R.Limitis, R.Notes FROM HouseRegisters R WHERE " & IIf(drMyRow.Table.Columns.Count >= 70, "R.HousesCode = " & drMyRow.Item(0), "R.Code = " & drMyRow.Item(0)), cnnNotary)
                            dadNotary.Fill(dtHouseRegisters)
                            Call PasteMyItem(rtbTemp, dtHouseRegisters, strObjects(Index))
                            dtHouseRegisters.Clear()
                        Case "R"
                            Dim dtHouseBranches As New DataTable
                            dadNotary = New SqlClient.SqlDataAdapter("SELECT B.Code, HBT.Name, B.BranchNo, B.LastFishNo, B.LastFishDate, B.Notes FROM HouseBranches B, HouseBranchTypes HBT WHERE B.HouseBranchTypesCode = HBT.Code AND " & IIf(drMyRow.Table.Columns.Count >= 70, "B.HousesCode = " & drMyRow.Item(0), "B.Code = " & drMyRow.Item(0)), cnnNotary)
                            dadNotary.Fill(dtHouseBranches)
                            Call PasteMyItem(rtbTemp, dtHouseBranches, strObjects(Index))
                            dtHouseBranches.Clear()
                        Case "X"
                            Dim dtHouseAdds As New DataTable
                            dadNotary = New SqlClient.SqlDataAdapter("SELECT A.Code, HT.Name, A.Amount, A.LocalPelak, A.Place, A.Area, A.Height, A.NorthNotes, A.EastNotes, A.SouthNotes, A.WestNotes, A.Notes FROM HouseAdds A, HouseTypes HT WHERE A.HouseTypesCode = HT.Code AND A.HousesCode = " & drMyRow.Item(0), cnnNotary)
                            dadNotary.Fill(dtHouseAdds)
                            If dtHouseAdds.Rows.Count = 0 Then rtbTemp.SelectedText = "" Else Call PasteMyPattern(rtbTemp, dtHouseAdds, drDocument, strObjects(Index))
                            dtHouseAdds.Clear()
                        Case "Y"
                            Dim dtHouseRegisters As New DataTable
                            dadNotary = New SqlClient.SqlDataAdapter("SELECT R.Code, R.PageNo, R.BookNo, R.RegisterNo, R.RegisterDate, R.PrintNo, R.OwnerName, R.Area, R.Limitis, R.Notes FROM HouseRegisters R WHERE R.HousesCode = " & drMyRow.Item(0), cnnNotary)
                            dadNotary.Fill(dtHouseRegisters)
                            If dtHouseRegisters.Rows.Count = 0 Then rtbTemp.SelectedText = "" Else Call PasteMyPattern(rtbTemp, dtHouseRegisters, drDocument, strObjects(Index))
                            dtHouseRegisters.Clear()
                        Case "Z"
                            Dim dtHouseBranches As New DataTable
                            dadNotary = New SqlClient.SqlDataAdapter("SELECT B.Code, HBT.Name, B.BranchNo, B.LastFishNo, B.LastFishDate, B.Notes FROM HouseBranches B, HouseBranchTypes HBT WHERE B.HouseBranchTypesCode = HBT.Code AND B.HousesCode = " & drMyRow.Item(0), cnnNotary)
                            dadNotary.Fill(dtHouseBranches)
                            If dtHouseBranches.Rows.Count = 0 Then rtbTemp.SelectedText = "" Else Call PasteMyPattern(rtbTemp, dtHouseBranches, drDocument, strObjects(Index))
                            dtHouseBranches.Clear()
                        Case Else : If strObjects(Index).Contains("-") Then rtbTemp.SelectedText = ReverseMyDate(drMyRow.Item(CInt(Mid(strObjects(Index), strObjects(Index).IndexOf("-") + 2)) + 1)) Else rtbTemp.SelectedText = ReverseMyDate(drMyRow.Item(CInt(Mid(strObjects(Index), 2)) + 1))
                    End Select
                End If
            Next
            Clipboard.Clear()
            For Index = 1 To 500
                Application.DoEvents()
            Next
            rtbTemp.SelectAll()
            rtbTemp.Copy()
            For Index = 1 To 500
                Application.DoEvents()
            Next
            rtbDocument.Select()
            rtbDocument.Paste()
        End If
        Erase strObjects
        Erase blnIfs
        Erase blnDelObjects
        dtTemp.Clear()
        rtbTemp.Dispose()
    End Sub

    Private Function ReturnDocumentEntryNameFromTable(ByRef drDocument As DataRow, ByRef lngEntryCode As Long) As String
        ReturnDocumentEntryNameFromTable = ""
        If lngEntryCode <= 30 Then
            If lngEntryCode <= 14 Then
                Select Case lngEntryCode
                    Case 0 : ReturnDocumentEntryNameFromTable = drDocument.Item(5)
                    Case 1 : ReturnDocumentEntryNameFromTable = ReturnAlphabetNumeric(drDocument.Item(5))
                    Case 2 : ReturnDocumentEntryNameFromTable = drDocument.Item(6)
                    Case 3 : ReturnDocumentEntryNameFromTable = ReturnAlphabetNumeric(drDocument.Item(6))
                    Case 4 : ReturnDocumentEntryNameFromTable = drDocument.Item(7)
                    Case 5 : ReturnDocumentEntryNameFromTable = ReturnAlphabetNumeric(drDocument.Item(7))
                    Case 6 : ReturnDocumentEntryNameFromTable = drDocument.Item(8)
                    Case 7 : ReturnDocumentEntryNameFromTable = ReturnAlphabetNumeric(drDocument.Item(8))
                    Case 8 : ReturnDocumentEntryNameFromTable = ReverseMyDate(drDocument.Item(9))
                    Case 9 : ReturnDocumentEntryNameFromTable = ReturnAlphabetDate(drDocument.Item(9))
                    Case 10 : ReturnDocumentEntryNameFromTable = PersianWeekDay(drDocument.Item(9))
                    Case 11 : ReturnDocumentEntryNameFromTable = FormatDateTime(drDocument.Item(10), DateFormat.ShortTime)
                    Case 12 : ReturnDocumentEntryNameFromTable = ReverseMyDate(drDocument.Item(11))
                    Case 13 : ReturnDocumentEntryNameFromTable = ReturnAlphabetDate(drDocument.Item(11))
                    Case 14 : ReturnDocumentEntryNameFromTable = PersianWeekDay(drDocument.Item(11))
                End Select
            Else
                Select Case lngEntryCode
                    Case 15 : ReturnDocumentEntryNameFromTable = drDocument.Item(12)
                    Case 16 : ReturnDocumentEntryNameFromTable = IIf(blnMoney, IIf(blnReverseNumber, ReverseMyString(Replace(FormatNumber(drDocument.Item(13), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers), strBetweenNumbers), Replace(FormatNumber(drDocument.Item(13), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers)), FormatNumber(drDocument.Item(13), 0, TriState.True, TriState.False, TriState.False))
                    Case 17 : ReturnDocumentEntryNameFromTable = ReturnAlphabetNumeric(drDocument.Item(13))
                    Case 18 : ReturnDocumentEntryNameFromTable = IIf(blnMoney, IIf(blnReverseNumber, ReverseMyString(Replace(FormatNumber(drDocument.Item(14), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers), strBetweenNumbers), Replace(FormatNumber(drDocument.Item(14), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers)), FormatNumber(drDocument.Item(14), 0, TriState.True, TriState.False, TriState.False))
                    Case 19 : ReturnDocumentEntryNameFromTable = ReturnAlphabetNumeric(drDocument.Item(14))
                    Case 20 : ReturnDocumentEntryNameFromTable = IIf(blnMoney, IIf(blnReverseNumber, ReverseMyString(Replace(FormatNumber(drDocument.Item(15), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers), strBetweenNumbers), Replace(FormatNumber(drDocument.Item(15), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers)), FormatNumber(drDocument.Item(15), 0, TriState.True, TriState.False, TriState.False))
                    Case 21 : ReturnDocumentEntryNameFromTable = ReturnAlphabetNumeric(drDocument.Item(15))
                    Case 22 : ReturnDocumentEntryNameFromTable = IIf(blnMoney, IIf(blnReverseNumber, ReverseMyString(Replace(FormatNumber(drDocument.Item(16), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers), strBetweenNumbers), Replace(FormatNumber(drDocument.Item(16), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers)), FormatNumber(drDocument.Item(16), 0, TriState.True, TriState.False, TriState.False))
                    Case 23 : ReturnDocumentEntryNameFromTable = ReturnAlphabetNumeric(drDocument.Item(16))
                    Case 24 : ReturnDocumentEntryNameFromTable = IIf(blnMoney, IIf(blnReverseNumber, ReverseMyString(Replace(FormatNumber(drDocument.Item(17), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers), strBetweenNumbers), Replace(FormatNumber(drDocument.Item(17), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers)), FormatNumber(drDocument.Item(17), 0, TriState.True, TriState.False, TriState.False))
                    Case 25 : ReturnDocumentEntryNameFromTable = ReturnAlphabetNumeric(drDocument.Item(17))
                    Case 26 : ReturnDocumentEntryNameFromTable = IIf(blnMoney, IIf(blnReverseNumber, ReverseMyString(Replace(FormatNumber(dblSaleTax * drDocument.Item(15) / 100, 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers), strBetweenNumbers), Replace(FormatNumber(dblSaleTax * drDocument.Item(15) / 100, 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers)), FormatNumber(dblSaleTax * drDocument.Item(15) / 100, 0, TriState.True, TriState.False, TriState.False))
                    Case 27 : ReturnDocumentEntryNameFromTable = ReturnAlphabetNumeric(dblSaleTax * drDocument.Item(15) / 100)
                    Case 28 : ReturnDocumentEntryNameFromTable = IIf(blnMoney, IIf(blnReverseNumber, ReverseMyString(Replace(FormatNumber(drDocument.Item(18), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers), strBetweenNumbers), Replace(FormatNumber(drDocument.Item(18), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers)), FormatNumber(drDocument.Item(18), 0, TriState.True, TriState.False, TriState.False))
                    Case 29 : ReturnDocumentEntryNameFromTable = ReturnAlphabetNumeric(drDocument.Item(18))
                    Case 30 : ReturnDocumentEntryNameFromTable = IIf(blnMoney, IIf(blnReverseNumber, ReverseMyString(Replace(FormatNumber(drDocument.Item(15) + drDocument.Item(18), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers), strBetweenNumbers), Replace(FormatNumber(drDocument.Item(15) + drDocument.Item(18), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers)), FormatNumber(drDocument.Item(15) + drDocument.Item(18), 0, TriState.True, TriState.False, TriState.False))
                End Select
            End If
        Else
            If lngEntryCode <= 45 Then
                Select Case lngEntryCode
                    Case 31 : ReturnDocumentEntryNameFromTable = ReturnAlphabetNumeric(drDocument.Item(15) + drDocument.Item(18))
                    Case 32 : ReturnDocumentEntryNameFromTable = IIf(blnMoney, IIf(blnReverseNumber, ReverseMyString(Replace(FormatNumber(drDocument.Item(15) + drDocument.Item(18) + (dblSaleTax * drDocument.Item(15) / 100), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers), strBetweenNumbers), Replace(FormatNumber(drDocument.Item(15) + drDocument.Item(18) + (dblSaleTax * drDocument.Item(15) / 100), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers)), FormatNumber(drDocument.Item(15) + drDocument.Item(18) + (dblSaleTax * drDocument.Item(15) / 100), 0, TriState.True, TriState.False, TriState.False))
                    Case 33 : ReturnDocumentEntryNameFromTable = ReturnAlphabetNumeric(drDocument.Item(15) + drDocument.Item(18) + (dblSaleTax * drDocument.Item(15) / 100))
                    Case 34 : ReturnDocumentEntryNameFromTable = IIf(blnMoney, IIf(blnReverseNumber, ReverseMyString(Replace(FormatNumber(drDocument.Item(15) + drDocument.Item(18) + (dblSaleTax * drDocument.Item(15) / 100) + drDocument.Item(14), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers), strBetweenNumbers), Replace(FormatNumber(drDocument.Item(15) + drDocument.Item(18) + (dblSaleTax * drDocument.Item(15) / 100) + drDocument.Item(14), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers)), FormatNumber(drDocument.Item(15) + drDocument.Item(18) + (dblSaleTax * drDocument.Item(15) / 100) + drDocument.Item(14), 0, TriState.True, TriState.False, TriState.False))
                    Case 35 : ReturnDocumentEntryNameFromTable = ReturnAlphabetNumeric(drDocument.Item(15) + drDocument.Item(18) + (dblSaleTax * drDocument.Item(15) / 100) + drDocument.Item(14))
                    Case 36 : ReturnDocumentEntryNameFromTable = IIf(blnMoney, IIf(blnReverseNumber, ReverseMyString(Replace(FormatNumber(drDocument.Item(15) + drDocument.Item(18) + (dblSaleTax * drDocument.Item(15) / 100) + drDocument.Item(14) + drDocument.Item(17), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers), strBetweenNumbers), Replace(FormatNumber(drDocument.Item(15) + drDocument.Item(18) + drDocument.Item(16) + drDocument.Item(14) + drDocument.Item(17), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers)), FormatNumber(drDocument.Item(15) + drDocument.Item(18) + drDocument.Item(16) + drDocument.Item(14) + drDocument.Item(17), 0, TriState.True, TriState.False, TriState.False))
                    Case 37 : ReturnDocumentEntryNameFromTable = ReturnAlphabetNumeric(drDocument.Item(15) + drDocument.Item(18) + (dblSaleTax * drDocument.Item(15) / 100) + drDocument.Item(14) + drDocument.Item(17))
                    Case 38 : ReturnDocumentEntryNameFromTable = drDocument.Item(20)
                    Case 39 : ReturnDocumentEntryNameFromTable = ReverseMyDate(drDocument.Item(21))
                    Case 40 : ReturnDocumentEntryNameFromTable = ReturnAlphabetDate(drDocument.Item(21))
                    Case 41
                        Dim dadNotary As New SqlClient.SqlDataAdapter("SELECT P.Serial FROM Papers P, DocumentPapers DP WHERE DP.DocumentsCode = " & drDocument.Item(0) & " AND DP.PapersCode = P.Code ORDER BY P.Serial", cnnNotary), dtPapers As New DataTable
                        dadNotary.Fill(dtPapers)
                        If dtPapers.Rows.Count = 1 Then
                            ReturnDocumentEntryNameFromTable &= dtPapers.Rows(0).Item(0)
                        ElseIf dtPapers.Rows.Count = 2 Then
                            ReturnDocumentEntryNameFromTable &= dtPapers.Rows(0).Item(0) & " و " & dtPapers.Rows(1).Item(0)
                        ElseIf dtPapers.Rows.Count >= 3 Then
                            Dim Index As Long = 0, intCount As Long = 0
                            While (Index + 1) < dtPapers.Rows.Count
                                Index += 1
                                intCount = 0
                                While CLng(dtPapers.Rows(Index).Item(0)) = CLng(dtPapers.Rows(Index - 1).Item(0)) + 1
                                    Index += 1
                                    intCount += 1
                                    If Index = dtPapers.Rows.Count Then Exit While
                                End While
                                If intCount > 1 Then
                                    ReturnDocumentEntryNameFromTable &= dtPapers.Rows(Index - intCount - 1).Item(0) & " الي " & dtPapers.Rows(Index - 1).Item(0) & " و "
                                ElseIf intCount = 1 Then
                                    ReturnDocumentEntryNameFromTable &= dtPapers.Rows(Index - 2).Item(0) & " و " & dtPapers.Rows(Index - 1).Item(0) & " و "
                                Else
                                    ReturnDocumentEntryNameFromTable &= dtPapers.Rows(Index - 1).Item(0) & " و "
                                End If
                            End While
                            ReturnDocumentEntryNameFromTable = Mid(ReturnDocumentEntryNameFromTable, 1, Len(ReturnDocumentEntryNameFromTable) - 3)
                        End If
                    Case 42
                        Dim dadNotary As New SqlClient.SqlDataAdapter("SELECT P.PaperPass FROM Papers P, DocumentPapers DP WHERE DP.DocumentsCode = " & drDocument.Item(0) & " AND DP.PapersCode = P.Code ORDER BY P.PaperPass", cnnNotary), dtPapers As New DataTable
                        dadNotary.Fill(dtPapers)
                        For Index As Integer = 0 To dtPapers.Rows.Count - 2
                            ReturnDocumentEntryNameFromTable &= dtPapers.Rows(Index).Item(0) & " و "
                        Next
                        If dtPapers.Rows.Count <> 0 Then ReturnDocumentEntryNameFromTable &= dtPapers.Rows(dtPapers.Rows.Count - 1).Item(0)
                    Case 43 : ReturnDocumentEntryNameFromTable = drDocument.Item(19)
                    Case 44 : ReturnDocumentEntryNameFromTable = ReturnAlphabetNumeric(drDocument.Item(19))
                    Case 45 : ReturnDocumentEntryNameFromTable = drDocument.Item(22)
                End Select
            Else
                Select Case lngEntryCode
                    Case 46 : ReturnDocumentEntryNameFromTable = IIf(blnMoney, IIf(blnReverseNumber, ReverseMyString(Replace(FormatNumber(drDocument.Item(23), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers), strBetweenNumbers), Replace(FormatNumber(drDocument.Item(23), 0, TriState.True, TriState.False, TriState.True), ",", strBetweenNumbers)), FormatNumber(drDocument.Item(23), 0, TriState.True, TriState.False, TriState.False))
                    Case 47 : ReturnDocumentEntryNameFromTable = ReturnAlphabetNumeric(drDocument.Item(23))
                    Case 48 : ReturnDocumentEntryNameFromTable = drDocument.Item(24)
                    Case 49 : ReturnDocumentEntryNameFromTable = strNotaryNo
                    Case 50 : ReturnDocumentEntryNameFromTable = strNotaryCity
                    Case 53 : ReturnDocumentEntryNameFromTable = strNotaryFName
                    Case 54 : ReturnDocumentEntryNameFromTable = strNotaryLName
                    Case 51 To 52, 55 To 62
                        Dim intFileNo As Integer = FreeFile(), strName As String = ""
                        FileOpen(intFileNo, strNotaryPath & "\Data\Slkws.bqz", OpenMode.Random, OpenAccess.ReadWrite, OpenShare.LockReadWrite)
                        FileGet(intFileNo, strName, IIf(lngEntryCode = 51 Or lngEntryCode = 52, lngEntryCode - 19, lngEntryCode - 21))
                        FileClose(intFileNo)
                        ReturnDocumentEntryNameFromTable = strName
                End Select
            End If
        End If
    End Function

    Private Sub PasteMyItem(ByRef rtbDocument As RichTextBox, ByRef dtTemp As DataTable, ByRef strField As String)
        Dim blnFound As Boolean = False, intColumn As Integer, intRowNumber As Integer, strName As String = ""
        If strField.Contains("-") Then
            intColumn = Mid(strField, InStr(strField, "-") + 1)
            If Mid(strField, 2, 1) = "0" Then
                blnFound = True
            Else
                intRowNumber = Mid(strField, 2, InStr(strField, "-") - 2)
            End If
        Else
            intColumn = Mid(strField, 2)
            blnFound = True
        End If
        intColumn += 1 'because all of tables has extra field code at first column
        If blnFound Then
            For Index As Integer = 0 To dtTemp.Rows.Count - 1
                strName &= dtTemp.Rows(Index).Item(intColumn) & " و "
            Next
            If strName <> "" Then strName = Mid(strName, 1, Len(strName) - 3)
        Else
            strName = dtTemp.Rows(intRowNumber).Item(intColumn)
        End If
        If strName = "" Then
            rtbDocument.SelectedText = ""
        Else
            CopyTextToClipbaord(ReverseMyDate(strName))
            rtbDocument.Paste()
            Application.DoEvents()
        End If
    End Sub

    Private Sub PasteMyPattern(ByRef rtbDocument As RichTextBox, ByRef dtTemp As DataTable, ByVal drDocument As DataRow, ByRef strField As String)
        Dim intRowNumber As Integer, lngMyCode As Long
        lngMyCode = Mid(strField, InStr(strField, "-") + 1)
        If Mid(strField, 2, 1) = "0" Then
            Select Case Mid(strField, 1, 1)
                Case "S"
                    For Index As Integer = 0 To dtTemp.Rows.Count - 1
                        Call BuildText(rtbDocument, drDocument, dtTemp.Rows(Index), lngMyCode, "I", "SELECT I.Code FROM Cars I, CarKinds IK, CarSystems IS, CarTypes IT WHERE I.CarTypesCode = IT.Code AND IT.CarSystemsCode = IS.Code AND IS.CarKindsCode = IK.Code AND I.Code = " & dtTemp.Rows(Index).Item(0) & " AND (")
                        If Index <> dtTemp.Rows.Count - 1 Then
                            CopyTextToClipbaord(" و ")
                            rtbDocument.Paste()
                            Application.DoEvents()
                        End If
                    Next
                Case "T"
                    For Index As Integer = 0 To dtTemp.Rows.Count - 1
                        Call BuildText(rtbDocument, drDocument, dtTemp.Rows(Index), lngMyCode, "K", "SELECT K.Code FROM AgriCars K, AgriCarKinds KK, AgriCarSystems KS, AgriCarTypes KT WHERE K.AgriCarTypesCode = KT.Code AND KT.AgriCarSystemsCode = KS.Code AND KS.AgriCarKindsCode = KK.Code AND K.Code = " & dtTemp.Rows(Index).Item(0) & " AND (")
                        If Index <> dtTemp.Rows.Count - 1 Then
                            CopyTextToClipbaord(" و ")
                            rtbDocument.Paste()
                            Application.DoEvents()
                        End If
                    Next
                Case "U"
                    For Index As Integer = 0 To dtTemp.Rows.Count - 1
                        Call BuildText(rtbDocument, drDocument, dtTemp.Rows(Index), lngMyCode, "L", "SELECT L.Code FROM WayCars L, WayCarKinds LK, WayCarSystems LS, WayCarTypes LT WHERE L.WayCarTypesCode = LT.Code AND LT.WayCarSystemsCode = LS.Code AND LS.WayCarKindsCode = LK.Code AND L.Code = " & dtTemp.Rows(Index).Item(0) & " AND (")
                        If Index <> dtTemp.Rows.Count - 1 Then
                            CopyTextToClipbaord(" و ")
                            rtbDocument.Paste()
                            Application.DoEvents()
                        End If
                    Next
                Case "V"
                    For Index As Integer = 0 To dtTemp.Rows.Count - 1
                        Call BuildText(rtbDocument, drDocument, dtTemp.Rows(Index), lngMyCode, "M", "SELECT M.Code FROM Motors M, MotorKinds MK, MotorSystems MS WHERE M.MotorSystemsCode = MS.Code AND MS.MotorKindsCode = MK.Code AND M.Code = " & dtTemp.Rows(Index).Item(0) & " AND (")
                        If Index <> dtTemp.Rows.Count - 1 Then
                            CopyTextToClipbaord(" و ")
                            rtbDocument.Paste()
                            Application.DoEvents()
                        End If
                    Next
                Case "W"
                    For Index As Integer = 0 To dtTemp.Rows.Count - 1
                        Call BuildText(rtbDocument, drDocument, dtTemp.Rows(Index), lngMyCode, "J", "SELECT J.Code FROM Houses J, Sections JS, Cities JC, Provinces JP, HouseTypes JT WHERE J.HouseTypesCode = JT.Code AND J.ProvincesCode = JP.Code AND J.SectionsCode = JS.Code AND JS.CitiesCode = JC.Code AND J.Code = " & dtTemp.Rows(Index).Item(0) & " AND (")
                        If Index <> dtTemp.Rows.Count - 1 Then
                            CopyTextToClipbaord(" و ")
                            rtbDocument.Paste()
                        End If
                    Next
                Case "X"
                    For Index As Integer = 0 To dtTemp.Rows.Count - 1
                        Call BuildText(rtbDocument, drDocument, dtTemp.Rows(Index), lngMyCode, "P", "SELECT P.Code FROM HouseAdds P, HouseTypes PT WHERE P.HouseTypesCode = PT.Code AND P.Code = " & dtTemp.Rows(Index).Item(0) & " AND (")
                        If Index <> dtTemp.Rows.Count - 1 Then
                            CopyTextToClipbaord(" و ")
                            rtbDocument.Paste()
                            Application.DoEvents()
                        End If
                    Next
                Case "Y"
                    For Index As Integer = 0 To dtTemp.Rows.Count - 1
                        Call BuildText(rtbDocument, drDocument, dtTemp.Rows(Index), lngMyCode, "Q", "SELECT Q.Code FROM HouseRegisters Q WHERE Q.Code = " & dtTemp.Rows(Index).Item(0) & " AND (")
                        If Index <> dtTemp.Rows.Count - 1 Then
                            CopyTextToClipbaord(" و ")
                            rtbDocument.Paste()
                            Application.DoEvents()
                        End If
                    Next
                Case "Z"
                    For Index As Integer = 0 To dtTemp.Rows.Count - 1
                        Call BuildText(rtbDocument, drDocument, dtTemp.Rows(Index), lngMyCode, "R", "SELECT R.Code FROM HouseBranches R, HouseBranchTypes RT WHERE R.HouseBranchTypesCode = RT.Code AND R.Code = " & dtTemp.Rows(Index).Item(0) & " AND (")
                        If Index <> dtTemp.Rows.Count - 1 Then
                            CopyTextToClipbaord(" و ")
                            rtbDocument.Paste()
                            Application.DoEvents()
                        End If
                    Next
            End Select
        Else
            If strField.Contains("-") Then intRowNumber = Mid(strField, 2, InStr(strField, "-") - 2) Else intRowNumber = Mid(strField, 2)
            Select Case Mid(strField, 1, 1)
                Case "S" : Call BuildText(rtbDocument, drDocument, dtTemp.Rows(intRowNumber), lngMyCode, "I", "SELECT I.Code FROM Cars I, CarKinds IK, CarSystems IS, CarTypes IT WHERE I.CarTypesCode = IT.Code AND IT.CarSystemsCode = IS.Code AND IS.CarKindsCode = IK.Code AND I.Code = " & dtTemp.Rows(intRowNumber).Item(0) & " AND (")
                Case "T" : Call BuildText(rtbDocument, drDocument, dtTemp.Rows(intRowNumber), lngMyCode, "K", "SELECT K.Code FROM AgriCars K, AgriCarKinds KK, AgriCarSystems KS, AgriCarTypes KT WHERE K.AgriCarTypesCode = KT.Code AND KT.AgriCarSystemsCode = KS.Code AND KS.AgriCarKindsCode = KK.Code AND K.Code = " & dtTemp.Rows(intRowNumber).Item(0) & " AND (")
                Case "U" : Call BuildText(rtbDocument, drDocument, dtTemp.Rows(intRowNumber), lngMyCode, "L", "SELECT L.Code FROM WayCars L, WayCarKinds LK, WayCarSystems LS, WayCarTypes LT WHERE L.WayCarTypesCode = LT.Code AND LT.WayCarSystemsCode = LS.Code AND LS.WayCarKindsCode = LK.Code AND L.Code = " & dtTemp.Rows(intRowNumber).Item(0) & " AND (")
                Case "V" : Call BuildText(rtbDocument, drDocument, dtTemp.Rows(intRowNumber), lngMyCode, "M", "SELECT M.Code FROM Motors M, MotorKinds MK, MotorSystems MS, MotorCapacities MC WHERE M.MotorSystemsCode = MS.Code AND MS.MotorKindsCode = MK.Code AND M.MotorCapacitiesCode = MC.Code AND M.Code = " & dtTemp.Rows(intRowNumber).Item(0) & " AND (")
                Case "W" : Call BuildText(rtbDocument, drDocument, dtTemp.Rows(intRowNumber), lngMyCode, "J", "SELECT J.Code FROM Houses J, Sections JS, Cities JC, Provinces JP, HouseTypes JP WHERE J.HouseTypesCode = JT.Code AND J.ProvincesCode = JP.Code AND J.SectionsCode = JS.Code AND JS.CitiesCode = JC.Code AND J.Code = " & dtTemp.Rows(intRowNumber).Item(0) & " AND (")
                Case "X" : Call BuildText(rtbDocument, drDocument, dtTemp.Rows(intRowNumber), lngMyCode, "P", "SELECT P.Code FROM HouseAdss P, HouseTypes PT WHERE P.HouseTypesCode = PT.Code AND P.Code = " & dtTemp.Rows(intRowNumber).Item(0) & " AND (")
                Case "Y" : Call BuildText(rtbDocument, drDocument, dtTemp.Rows(intRowNumber), lngMyCode, "Q", "SELECT Q.Code FROM HouseRegisters Q WHERE Q.Code = " & dtTemp.Rows(intRowNumber).Item(0) & " AND (")
                Case "Z" : Call BuildText(rtbDocument, drDocument, dtTemp.Rows(intRowNumber), lngMyCode, "R", "SELECT R.Code FROM HouseBranches R, HouseBranchesTypes RT WHERE R.HouseBranchesTypesCode = RT.Code AND R.Code = " & dtTemp.Rows(intRowNumber).Item(0) & " AND (")
                Case "+" : Call BuildText(rtbDocument, drDocument, drDocument, lngMyCode, "D", "SELECT D.Code FROM Documents D WHERE Code = " & drDocument.Item(0) & " AND (")
            End Select
        End If
    End Sub

    Public Function GetDocumentPath(ByVal lngDocumentCode As Long, ByVal blnIsWrite As Boolean) As String
        Dim IntFileNo As Integer = FreeFile(), strDrive As String
        GetDocumentPath = ""
        FileOpen(IntFileNo, strNotaryPath & "\Data\Slkws.bqz", OpenMode.Random, OpenAccess.ReadWrite, OpenShare.LockReadWrite)
        FileGet(IntFileNo, GetDocumentPath, 1)
        FileClose(IntFileNo)
        strDrive = Mid(GetDocumentPath, 1, 3)
        If strDrive.Length = 3 Then
            Dim MyDriveInfo As New IO.DriveInfo(strDrive)
            If Not MyDriveInfo.IsReady Then Shell("explorer.exe " & strDrive, AppWinStyle.Hide)
        End If
        Dim blnExist As Boolean = False
        If GetDocumentPath Like "?:\" Then
            For Index = 0 To My.Computer.FileSystem.Drives.Count - 1
                If My.Computer.FileSystem.Drives(Index).Name = GetDocumentPath Then
                    blnExist = True
                    Exit For
                End If
            Next
        Else
            If IO.Directory.Exists(Mid(GetDocumentPath, 1, GetDocumentPath.Length - 1)) Then blnExist = True
        End If
        If blnExist = False Then
            If blnIsWrite Then Call ShowError("ذخيره سند در مسير داده شده امکان پذير نبود. سند در درايو C ذخيره مي شود.")
            GetDocumentPath = "C:\"
        End If
        GetDocumentPath &= lngDocumentCode & ".rtf"
    End Function

    Public Sub ViewDocument(ByVal lngDocumentCode As Long, ByVal strFormText As String)
        Dim strPath As String = GetDocumentPath(lngDocumentCode, False)
        If Not IO.File.Exists(strPath) Then If BuildDocumentNote(lngDocumentCode, True) = False Then Exit Sub
        Try
            frmDocuments.cmdBuild.Enabled = False
            frmDocuments.cmdView.Enabled = False
            Application.DoEvents()

            Call CheckWordApp(False)
            SetParent(ptrWordHandle, 0)
            WordApp.ScreenUpdating = False
            Application.DoEvents()
            While WordApp.Documents.Count >= 1
                WordApp.Documents(1).Close(False)
            End While
            Dim MyDoc As Object = WordApp.Documents.Open(strPath)
            WordApp.ShowWindowsInTaskbar = True
            frmWord.grpView.Visible = True
            If WordApp.ActiveWindow.View.SplitSpecial <> 0 Then WordApp.ActiveWindow.Panes(2).Close()
            WordApp.ActiveWindow.ActivePane.View.Type = 6
            WordApp.ActiveWindow.ActivePane.View.SeekView = 0
            WordApp.ActiveWindow.ActivePane.View.Zoom.Percentage = 80
            Call SetEnabledWordSave(True)
            frmWord.lblCode.Text = lngDocumentCode
            frmWord.lblPattern.Tag = False
            frmWord.Text = strFormText
            Call ShowWordForm(False)
            WordApp.ScreenUpdating = True
            WordApp.Keyboard(1065)
            WordApp.Activate()
            frmDocuments.cmdBuild.Enabled = True
            frmDocuments.cmdView.Enabled = True

        Catch ex As Exception
            Call ShowError("سند قابل نمايش نيست. احتمالاً سندي با اين نام باز مي باشد.")
        End Try
    End Sub

    Public Sub MoveDocToPos(ByVal lngDocumentCode As Long)
        Dim intFileNo As Integer = FreeFile(), strWindowName As String = "", strName As String = ""
        FileOpen(intFileNo, strNotaryPath & "\Data\Slkws.bqz", OpenMode.Random, OpenAccess.ReadWrite, OpenShare.Shared)
        FileGet(intFileNo, strName, 44)
        FileGet(intFileNo, strWindowName, 60)
        FileClose(intFileNo)
        Dim lpszParentWindow As String = strWindowName, ParenthWnd As New IntPtr(0)
        ParenthWnd = FindWindowNullClassName(0, lpszParentWindow)
        If ParenthWnd.Equals(IntPtr.Zero) Then
            Call ShowError("فرم پرداخت پنجره پيدا نشد.")
            Exit Sub
        End If
        SetForegroundWindow(ParenthWnd)
        For Index As Integer = 1 To 20000
            Application.DoEvents()
        Next

        Dim dadNotary = New SqlClient.SqlDataAdapter("SELECT CASE (SELECT IsPerson FROM DocumentActs WHERE DocumentsCode = D.Code AND ActorType = 1 AND RowNumber = 1) WHEN 0 THEN (SELECT C.Name FROM Companies C, DocumentActs A WHERE C.Code = A.EntityCode AND A.DocumentsCode = D.Code AND A.ActorType = 1 AND A.RowNumber = 1 AND A.IsPerson = 0) ELSE (SELECT P.FName + ' ' + P.LName FROM Persons P, DocumentActs A WHERE P.Code = A.EntityCode AND A.DocumentsCode = D.Code AND A.ActorType = 1 AND A.RowNumber = 1 AND A.IsPerson = 1) END, CASE (SELECT IsPerson FROM DocumentActs WHERE DocumentsCode = D.Code AND ActorType = 1 AND RowNumber = 1) WHEN 0 THEN '' ELSE (SELECT P.NationalNo FROM Persons P, DocumentActs A WHERE P.Code = A.EntityCode AND A.DocumentsCode = D.Code AND A.ActorType = 1 AND A.RowNumber = 1 AND A.IsPerson = 1) END, CASE WHEN (SELECT Code FROM DocumentActs WHERE DocumentsCode = D.Code AND ActorType = 2 AND RowNumber = 1) IS NULL THEN '' ELSE CASE (SELECT IsPerson FROM DocumentActs WHERE DocumentsCode = D.Code AND ActorType = 2 AND RowNumber = 1) WHEN 0 THEN (SELECT C.Name FROM Companies C, DocumentActs A WHERE C.Code = A.EntityCode AND A.DocumentsCode = D.Code AND A.ActorType = 2 AND A.RowNumber = 1 AND A.IsPerson = 0) ELSE (SELECT P.FName + ' ' + P.LName FROM Persons P, DocumentActs A WHERE P.Code = A.EntityCode AND A.DocumentsCode = D.Code AND A.ActorType = 2 AND A.RowNumber = 1 AND A.IsPerson = 1) END END, CASE WHEN (SELECT Code FROM DocumentActs WHERE DocumentsCode = D.Code AND ActorType = 2 AND RowNumber = 1) IS NULL THEN '' ELSE CASE (SELECT IsPerson FROM DocumentActs WHERE DocumentsCode = D.Code AND ActorType = 2 AND RowNumber = 1) WHEN 0 THEN '' ELSE (SELECT P.NationalNo FROM Persons P, DocumentActs A WHERE P.Code = A.EntityCode AND A.DocumentsCode = D.Code AND A.ActorType = 2 AND A.RowNumber = 1 AND A.IsPerson = 1) END END, P.RowNumber, PS.RowNumber, D.DocumentNo, D.DocumentPrice, D.RegisterIncome, D.EditIncome + D.PagesIncome + D.OtherIncome, D.TaxPrice FROM DocumentKinds K, DocumentTypes T, DocumentBranches B, Documents D, Poses P, PoseSubs PS WHERE B.DocumentTypesCode = T.Code AND T.DocumentKindsCode = K.Code AND T.PoseSubsCode = PS.Code AND PS.PosesCode = P.Code AND D.DocumentBranchesCode = B.Code AND D.Code = " & lngDocumentCode, cnnNotary), dtDoc As New DataTable
        dadNotary.SelectCommand.CommandTimeout = 0
        dadNotary.Fill(dtDoc)
        If dtDoc.Rows.Count = 0 Then 'If Pos not set, it doesn't any row
            dtDoc = New DataTable
            dadNotary = New SqlClient.SqlDataAdapter("SELECT CASE (SELECT IsPerson FROM DocumentActs WHERE DocumentsCode = D.Code AND ActorType = 1 AND RowNumber = 1) WHEN 0 THEN (SELECT C.Name FROM Companies C, DocumentActs A WHERE C.Code = A.EntityCode AND A.DocumentsCode = D.Code AND A.ActorType = 1 AND A.RowNumber = 1 AND A.IsPerson = 0) ELSE (SELECT P.FName + ' ' + P.LName FROM Persons P, DocumentActs A WHERE P.Code = A.EntityCode AND A.DocumentsCode = D.Code AND A.ActorType = 1 AND A.RowNumber = 1 AND A.IsPerson = 1) END, CASE (SELECT IsPerson FROM DocumentActs WHERE DocumentsCode = D.Code AND ActorType = 1 AND RowNumber = 1) WHEN 0 THEN '' ELSE (SELECT P.NationalNo FROM Persons P, DocumentActs A WHERE P.Code = A.EntityCode AND A.DocumentsCode = D.Code AND A.ActorType = 1 AND A.RowNumber = 1 AND A.IsPerson = 1) END, CASE WHEN (SELECT Code FROM DocumentActs WHERE DocumentsCode = D.Code AND ActorType = 2 AND RowNumber = 1) IS NULL THEN '' ELSE CASE (SELECT IsPerson FROM DocumentActs WHERE DocumentsCode = D.Code AND ActorType = 2 AND RowNumber = 1) WHEN 0 THEN (SELECT C.Name FROM Companies C, DocumentActs A WHERE C.Code = A.EntityCode AND A.DocumentsCode = D.Code AND A.ActorType = 2 AND A.RowNumber = 1 AND A.IsPerson = 0) ELSE (SELECT P.FName + ' ' + P.LName FROM Persons P, DocumentActs A WHERE P.Code = A.EntityCode AND A.DocumentsCode = D.Code AND A.ActorType = 2 AND A.RowNumber = 1 AND A.IsPerson = 1) END END, CASE WHEN (SELECT Code FROM DocumentActs WHERE DocumentsCode = D.Code AND ActorType = 2 AND RowNumber = 1) IS NULL THEN '' ELSE CASE (SELECT IsPerson FROM DocumentActs WHERE DocumentsCode = D.Code AND ActorType = 2 AND RowNumber = 1) WHEN 0 THEN '' ELSE (SELECT P.NationalNo FROM Persons P, DocumentActs A WHERE P.Code = A.EntityCode AND A.DocumentsCode = D.Code AND A.ActorType = 2 AND A.RowNumber = 1 AND A.IsPerson = 1) END END, 0, '', D.DocumentNo, D.DocumentPrice, D.RegisterIncome, D.EditIncome + D.PagesIncome + D.OtherIncome, D.TaxPrice FROM DocumentKinds K, DocumentTypes T, DocumentBranches B, Documents D WHERE B.DocumentTypesCode = T.Code AND T.DocumentKindsCode = K.Code AND D.DocumentBranchesCode = B.Code AND D.Code = " & lngDocumentCode, cnnNotary)
            dadNotary.SelectCommand.CommandTimeout = 0
            dadNotary.Fill(dtDoc)
        End If
        If Mid(strName, 1, 1) <> "0" Then
            If IsDBNull(dtDoc.Rows(0).Item(0)) Then dtDoc.Rows(0).Item(0) = ""
            If Mid(strName, 1, 1) = "2" Then
                dtDoc.Rows(0).Item(0) = "-"
            Else
                cmmNotary.CommandText = "SELECT COUNT(Code) FROM DocumentActs WHERE ActorType = 1 AND DocumentsCode = " & lngDocumentCode
                If cmmNotary.ExecuteScalar = 2 Then dtDoc.Rows(0).Item(0) &= " و شريک" Else If cmmNotary.ExecuteScalar >= 3 Then dtDoc.Rows(0).Item(0) &= " و شرکا"
            End If
            SendKeys.SendWait(dtDoc.Rows(0).Item(0))
        End If
        SendKeys.SendWait("{TAB}")
        If Mid(strName, 2, 1) <> "0" Then
            If IsDBNull(dtDoc.Rows(0).Item(1)) Then dtDoc.Rows(0).Item(1) = ""
            If Mid(strName, 2, 1) = "2" Then
                dtDoc.Rows(0).Item(1) = "-"
            Else
                cmmNotary.CommandText = "SELECT COUNT(Code) FROM DocumentActs WHERE ActorType = 2 AND DocumentsCode = " & lngDocumentCode
                If cmmNotary.ExecuteScalar = 2 Then dtDoc.Rows(0).Item(2) &= " و شريک" Else If cmmNotary.ExecuteScalar >= 3 Then dtDoc.Rows(0).Item(2) &= " و شرکا"
            End If
            SendKeys.SendWait(dtDoc.Rows(0).Item(1))
        End If
        SendKeys.SendWait("{TAB}")
        If Mid(strName, 3, 1) = "2" Then dtDoc.Rows(0).Item(2) = "-"
        If Mid(strName, 3, 1) <> "0" Then SendKeys.SendWait(dtDoc.Rows(0).Item(2))
        SendKeys.SendWait("{TAB}")
        If Mid(strName, 4, 1) = "2" Then dtDoc.Rows(0).Item(3) = "-"
        If Mid(strName, 4, 1) <> "0" Then SendKeys.SendWait(dtDoc.Rows(0).Item(3))
        SendKeys.SendWait("{TAB}")
        If Mid(strName, 5, 1) <> "0" Then
            For Index As Integer = 1 To dtDoc.Rows(0).Item(4)
                SendKeys.SendWait("{DOWN}")
            Next
        End If
        SendKeys.SendWait("{TAB}")
        If Mid(strName, 6, 1) <> "0" Then
            If IsNumeric(dtDoc.Rows(0).Item(5)) Then
                For Index As Integer = 1 To dtDoc.Rows(0).Item(5) - 1
                    SendKeys.SendWait("{DOWN}")
                Next
            End If
        End If
        SendKeys.SendWait("{TAB}")
        If Mid(strName, 7, 1) = "2" Then dtDoc.Rows(0).Item(6) = 0
        If Mid(strName, 7, 1) <> "0" Then SendKeys.SendWait(FormatNumber(dtDoc.Rows(0).Item(6), 0, TriState.True, TriState.False, TriState.False))
        SendKeys.SendWait("{TAB}")
        If Mid(strName, 8, 1) = "2" Then dtDoc.Rows(0).Item(7) = 0
        If Mid(strName, 8, 1) <> "0" Then SendKeys.SendWait(FormatNumber(dtDoc.Rows(0).Item(7), 0, TriState.True, TriState.False, TriState.False))
        SendKeys.SendWait("{TAB}")
        If Mid(strName, 9, 1) = "2" Then dtDoc.Rows(0).Item(8) = 0
        If Mid(strName, 9, 1) <> "0" Then SendKeys.SendWait(FormatNumber(dtDoc.Rows(0).Item(8), 0, TriState.True, TriState.False, TriState.False))
        SendKeys.SendWait("{TAB}")
        If Mid(strName, 10, 1) = "2" Then dtDoc.Rows(0).Item(9) = 0
        If Mid(strName, 10, 1) <> "0" Then SendKeys.SendWait(FormatNumber(dtDoc.Rows(0).Item(9), 0, TriState.True, TriState.False, TriState.False))
        SendKeys.SendWait("{TAB}")
        If Mid(strName, 11, 1) = "2" Then dtDoc.Rows(0).Item(10) = 0
        If Mid(strName, 11, 1) <> "0" Then SendKeys.SendWait(FormatNumber(dtDoc.Rows(0).Item(10), 0, TriState.True, TriState.False, TriState.False))
        dtDoc.Clear()
    End Sub

    Private Function ReturnAlphabetNumeric(ByVal Number As Long) As String
        Dim Yekan As Object = New Object() {"صفر", "يک", "دو", "سه", "چهار", "پنج", "شش", "هفت", "هشت", "نه"}
        Dim YD As Object = New Object() {"ده", "يازده", "دوازده", "سيزده", "چهارده", "پانزده", "شانزده", "هفده", "هجده", "نوزده"}
        Dim Dahgan As Object = New Object() {"بيست", "سي", "چهل", "پنجاه", "شصت", "هفتاد", "هشتاد", "نود"}
        Dim Sadgan As Object = New Object() {"صد", "دويست", "سيصد", "چهارصد", "پانصد", "ششصد", "هفتصد", "هشتصد", "نهصد"}
        Dim Str As String = ""

        If Number = 0 Then
            ReturnAlphabetNumeric = Yekan(0)
            Exit Function
        End If
        While Number > 0
            Select Case Number
                Case 1000000000 To 2147483647 : Str &= TAlphabet(Number \ 1000000000, Yekan, Dahgan, YD, Sadgan) & " مليارد" & Str
                    Number = Number Mod 1000000000
                Case 1000000 To 999999999 : Str &= TAlphabet(Number \ 1000000, Yekan, Dahgan, YD, Sadgan) & " ميليون"
                    Number = Number Mod 1000000
                Case 1000 To 999999 : Str &= TAlphabet(Number \ 1000, Yekan, Dahgan, YD, Sadgan) & " هزار"
                    Number = Number Mod 1000
                Case 0 To 999 : Str &= TAlphabet(Number, Yekan, Dahgan, YD, Sadgan)
                    Number = Number Mod 1
            End Select
            If Number > 0 Then
                Str &= " و "
            End If
        End While
        If Right(Str, 1) = "و" And Right(Str, 2) <> "دو" Then
            Str = Left(Str, Len(Str) - 1)
        End If
        ReturnAlphabetNumeric = Str
    End Function

    Private Function TAlphabet(ByVal Number As Long, ByRef Yekan As Object, ByRef Dahgan As Object, ByRef YD As Object, ByRef Sadgan As Object) As String
        Dim Str As String
        Str = ""
        If Number >= 100 Then
            Str = Sadgan(Number \ 100 - 1) & " و "
        End If
        Number = Number Mod 100
        Select Case Number
            Case 1 To 9 : TAlphabet = Str & Yekan(Number)
            Case 10 To 19 : TAlphabet = Str & YD(Number - 10)
            Case 20 To 99
                If Number Mod 10 = 0 Then
                    TAlphabet = Str & Dahgan(Number \ 10 - 2)
                Else
                    TAlphabet = Str & Dahgan(Number \ 10 - 2) & " و " & Yekan(Number Mod 10)
                End If
            Case Else : TAlphabet = Left(Str, Len(Str) - 3)
        End Select
    End Function

    Private Function ReturnAlphabetDate(ByVal strMyDate As String) As String
        Dim MyYear As Integer, MyMonth As Integer, MyDay As Integer, strDayName As String
        ReturnAlphabetDate = ""
        If strMyDate = "" Then Exit Function
        MyYear = Mid(strMyDate, 1, 4)
        MyMonth = Mid(strMyDate, 6, 2)
        MyDay = Mid(strMyDate, 9, 2)
        strDayName = ReturnAlphabetNumeric(MyDay)
        If Right(strDayName, 2) = "سه" Then strDayName = IIf(Len(strDayName) > 2, Mid(strDayName, 1, Len(strDayName) - 2), "") & "سو"
        Select Case MyMonth
            Case 1 : ReturnAlphabetDate = strDayName & "م فروردين ماه " & MyYear
            Case 2 : ReturnAlphabetDate = strDayName & "م ارديبهشت ماه " & MyYear
            Case 3 : ReturnAlphabetDate = strDayName & "م خرداد ماه " & MyYear
            Case 4 : ReturnAlphabetDate = strDayName & "م تير ماه " & MyYear
            Case 5 : ReturnAlphabetDate = strDayName & "م مرداد ماه " & MyYear
            Case 6 : ReturnAlphabetDate = strDayName & "م شهريور ماه " & MyYear
            Case 7 : ReturnAlphabetDate = strDayName & "م مهر ماه " & MyYear
            Case 8 : ReturnAlphabetDate = strDayName & "م آبان ماه " & MyYear
            Case 9 : ReturnAlphabetDate = strDayName & "م آذر ماه " & MyYear
            Case 10 : ReturnAlphabetDate = strDayName & "م دي ماه " & MyYear
            Case 11 : ReturnAlphabetDate = strDayName & "م بهمن ماه " & MyYear
            Case 12 : ReturnAlphabetDate = strDayName & "م اسفند ماه " & MyYear
        End Select
    End Function

    Private Function ReverseMyString(ByVal strMyString As String, ByVal strBetween As String) As String
        Dim strArray() As String, intFrom As Integer = 1, intTo As Integer
        ReDim strArray(0)
        ReverseMyString = ""
        intTo = InStr(intFrom, strMyString, strBetween)
        While intTo <> 0
            ReDim Preserve strArray(strArray.Length)
            strArray(strArray.Length - 1) = Mid(strMyString, intFrom, intTo - intFrom)
            intFrom = intTo + 1
            intTo = InStr(intFrom, strMyString, strBetween)
        End While
        ReDim Preserve strArray(strArray.Length)
        strArray(strArray.Length - 1) = Mid(strMyString, intFrom)
        For intFrom = strArray.Length - 1 To 1 Step -1
            ReverseMyString &= strArray(intFrom) & strBetween
        Next
        ReverseMyString = Mid(ReverseMyString, 1, Len(ReverseMyString) - 1)
    End Function

    Public Sub CheckDocumentNote(ByRef rtbDocument As RichTextBox, ByVal lngDocumentCode As Long)
        Dim intFileNo As Integer = FreeFile(), strName() As String, Index As Integer, intFrom As Integer, intTo As Integer, strExp As String, blnValue As Boolean, objValue As Object
        FileOpen(intFileNo, strNotaryPath & "\Data\Acts.bqz", OpenMode.Random, OpenAccess.ReadWrite, OpenShare.Shared)
        FileGet(intFileNo, Index, 1)
        ReDim strName(Index)
        For intFrom = 1 To Index
            FileGet(intFileNo, strName(intFrom), intFrom + 1)
        Next
        FileClose(intFileNo)
        For Index = 1 To strName.Length - 1
            intFrom = InStr(strName(Index), "!!$##")
            strExp = Mid(strName(Index), 1, intFrom - 1)
            If rtbDocument.Find(strExp, RichTextBoxFinds.NoHighlight) Then
                intFrom += 5
                blnValue = Mid(strName(Index), intFrom + 2, 1) = "1"
                intTo = InStr(intFrom, strName(Index), "!!$##")
                intFrom += 3
                objValue = Mid(strName(Index), intFrom, intTo - intFrom)
                If blnValue Then ShowError(objValue)

                intFrom = intTo + 5
                blnValue = Mid(strName(Index), intFrom + 2, 1)
                intTo = InStr(intFrom, strName(Index), "!!$##")
                intFrom += 3
                objValue = Mid(strName(Index), intFrom, intTo - intFrom)
                If blnValue Then
                    cmmNotary.CommandText = "UPDATE Documents SET DocumentPrice = " & CDec(objValue) & " WHERE Code = " & lngDocumentCode
                    cmmNotary.ExecuteNonQuery()
                End If

                intFrom = intTo + 5
                blnValue = Mid(strName(Index), intFrom + 2, 1)
                intTo = InStr(intFrom, strName(Index), "!!$##")
                intFrom += 3
                objValue = Mid(strName(Index), intFrom, intTo - intFrom)
                If blnValue Then
                    cmmNotary.CommandText = "UPDATE Documents SET RegisterIncome = " & CDec(objValue) & " WHERE Code = " & lngDocumentCode
                    cmmNotary.ExecuteNonQuery()
                End If

                intFrom = intTo + 5
                blnValue = Mid(strName(Index), intFrom + 2, 1)
                intTo = InStr(intFrom, strName(Index), "!!$##")
                intFrom += 3
                objValue = Mid(strName(Index), intFrom, intTo - intFrom)
                If blnValue Then
                    cmmNotary.CommandText = "UPDATE Documents SET EditIncome = " & CDec(objValue) & " WHERE Code = " & lngDocumentCode
                    cmmNotary.ExecuteNonQuery()
                End If

                intFrom = intTo + 5
                blnValue = Mid(strName(Index), intFrom + 2, 1)
                intTo = InStr(intFrom, strName(Index), "!!$##")
                intFrom += 3
                objValue = Mid(strName(Index), intFrom, intTo - intFrom)
                If blnValue Then
                    cmmNotary.CommandText = "UPDATE Documents SET TaxPrice = " & CDec(objValue) & " WHERE Code = " & lngDocumentCode
                    cmmNotary.ExecuteNonQuery()
                End If

                intFrom = intTo + 5
                blnValue = Mid(strName(Index), intFrom + 2, 1)
                intTo = InStr(intFrom, strName(Index), "!!$##")
                intFrom += 3
                objValue = Mid(strName(Index), intFrom, intTo - intFrom)
                If blnValue Then
                    cmmNotary.CommandText = "UPDATE Documents SET OtherIncome = " & CDec(objValue) & " WHERE Code = " & lngDocumentCode
                    cmmNotary.ExecuteNonQuery()
                End If

                intFrom = intTo + 5
                blnValue = Mid(strName(Index), intFrom + 2, 1)
                intTo = InStr(intFrom, strName(Index), "!!$##")
                intFrom += 3
                objValue = Mid(strName(Index), intFrom, intTo - intFrom)
                If blnValue Then
                    cmmNotary.CommandText = "UPDATE Documents SET DocumentDate = '" & objValue & "' WHERE Code = " & lngDocumentCode
                    cmmNotary.ExecuteNonQuery()
                    frmDocuments.dgrDocuments.SelectedRows(0).Cells(6).Value = objValue
                End If

                intFrom = intTo + 5
                blnValue = Mid(strName(Index), intFrom + 2, 1)
                intTo = InStr(intFrom, strName(Index), "!!$##")
                intFrom += 3
                objValue = Mid(strName(Index), intFrom, intTo - intFrom)
                If blnValue Then
                    cmmNotary.CommandText = "UPDATE Documents SET CompleteDate = '" & objValue & "' WHERE Code = " & lngDocumentCode
                    cmmNotary.ExecuteNonQuery()
                    frmDocuments.dgrDocuments.SelectedRows(0).Cells(12).Value = objValue
                End If

                intFrom = intTo + 5
                blnValue = Mid(strName(Index), intFrom + 2, 1)
                intTo = InStr(intFrom, strName(Index), "!!$##")
                intFrom += 3
                objValue = Mid(strName(Index), intFrom, intTo - intFrom)
                If blnValue Then
                    'cmmNotary.CommandText = "UPDATE Documents SET EditDate = '" & objValue & "' WHERE Code = " & lngDocumentCode
                    'cmmNotary.ExecuteNonQuery()
                End If

                intFrom = intTo + 5
                blnValue = Mid(strName(Index), intFrom + 2, 1)
                intTo = InStr(intFrom, strName(Index), "!!$##")
                intFrom += 3
                objValue = Mid(strName(Index), intFrom, intTo - intFrom)
                If blnValue Then
                    'cmmNotary.CommandText = "UPDATE Documents SET IsPos = " & objValue & " WHERE Code = " & lngDocumentCode
                    'cmmNotary.ExecuteNonQuery()
                End If

                intFrom = intTo + 5
                blnValue = Mid(strName(Index), intFrom + 2, 1)
                intTo = InStr(intFrom, strName(Index), "!!$##")
                intFrom += 3
                objValue = Mid(strName(Index), intFrom, intTo - intFrom)
                If blnValue Then
                    cmmNotary.CommandText = "UPDATE Documents SET IsGoverment = " & objValue & " WHERE Code = " & lngDocumentCode
                    cmmNotary.ExecuteNonQuery()
                End If

                intFrom = intTo + 5
                blnValue = Mid(strName(Index), intFrom + 2, 1)
                intTo = InStr(intFrom, strName(Index), "!!$##")
                intFrom += 3
                objValue = Mid(strName(Index), intFrom, intTo - intFrom)
                If blnValue Then
                    cmmNotary.CommandText = "UPDATE Documents SET PagesIncome = " & CDec(objValue) & " WHERE Code = " & lngDocumentCode
                    cmmNotary.ExecuteNonQuery()
                End If
            End If
        Next
    End Sub

    Public Sub LoadAndFillReportPaper(ByVal lngDocumentCode As Long, ByVal blnIsPrint As Boolean, ByVal blnFromOpenedWord As Boolean, ByVal blnSpaces As Boolean, Optional ByVal intTotalPapers As Integer = 1)
        'If dgrDocuments.SelectedRows.Count = 0 Then Exit Sub
        'If IsWordWindowOpen(dgrDocuments.SelectedRows(0).Cells(1).Value) Then Exit Sub
        Dim frpDocument As New FastReport.Report
        Dim txt As FastReport.TextObject, rtb As FastReport.RichObject, strRtf As String, Index As Integer, intFrom As Integer, intTo As Integer, strNumbers() As String, strSource As String, strDest As String
        Dim intFileNo As Integer = FreeFile(), blnIsTempVisible As Boolean, blnIsBackGround As Boolean
        '1-Load Report and fill header
        Dim dadNotary As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter("SELECT D.PageNo, D.SectionNo, D.DocumentNo, D.DocumentDate, D.RegisterIncome, D.EditIncome, D.OtherIncome, (D.RegisterIncome + D.EditIncome + D.OtherIncome), K.Name, 'شماره موقت: ' + convert(varchar(100), D.TempNo) FROM Documents D, DocumentKinds K, DocumentTypes T, DocumentBranches B WHERE D.DocumentBranchesCode = B.Code AND B.DocumentTypesCode = T.Code AND T.DocumentKindsCode = K.Code AND D.Code = " & lngDocumentCode, cnnNotary), dtDocument As New DataTable
        dadNotary.SelectCommand.CommandTimeout = 0
        dadNotary.Fill(dtDocument)
        frpDocument.Password = "qaz0qaz"
        frpDocument.Load(strNotaryPath & "\Reports\Paper.frx")
        For Index = 0 To 9
            txt = frpDocument.FindObject(strTitleNames(Index))
            If Index >= 4 And Index <= 7 Then txt.Text = FormatNumber(dtDocument.Rows(0).Item(Index), 0, TriState.True, TriState.False, TriState.True) Else txt.Text = dtDocument.Rows(0).Item(Index)
        Next
        Dim MyReportPage As FastReport.ReportPage = frpDocument.AllObjects(0), MyReportHeader As FastReport.PageHeaderBand = frpDocument.AllObjects(2), MyReportFooter As FastReport.PageFooterBand = frpDocument.FindObject("PageFooter1")
        rtb = frpDocument.FindObject("rtbDocument")
        rtb.Text = "" 'We should empty richtext box first otherwise in some documents we get error in text
        FileOpen(intFileNo, strNotaryPath & "\Data\Slkws.bqz", OpenMode.Random, OpenAccess.ReadWrite, OpenShare.LockReadWrite)
        FileGet(intFileNo, blnIsTempVisible, 71)
        FileGet(intFileNo, blnIsBackGround, 72)
        FileClose(intFileNo)
        txt = frpDocument.FindObject(strTitleNames(9))
        txt.Visible = blnIsTempVisible
        If blnIsBackGround Then
            MyReportPage.Watermark.Enabled = True
            MyReportPage.Watermark.Text = "دفتر شماره " & strNotaryNo & " " & strNotaryCity
            MyReportPage.Watermark.Font = New Font("Nazanin", 72)
        End If

        '2-Delete Extra Spaces
        Dim rtbDocument As New RichTextBox, strPath As String = GetDocumentPath(lngDocumentCode, False)
        'Call CheckWordApp(True)
        'If WordApp.Documents.Count = 0 Then Exit Sub
        If blnFromOpenedWord Then
            WordApp.Selection.WholeStory()
            WordApp.Selection.Copy()
            rtbDocument.Paste()
            WordApp.Selection.HomeKey(Unit:=6)
            While rtbDocument.Text.EndsWith(" ")
                rtbDocument.SelectionLength = 1
                rtbDocument.SelectionStart = rtbDocument.TextLength - 1
                rtbDocument.SelectedText = ""
            End While
        Else
            If Not IO.File.Exists(strPath) Then If BuildDocumentNote(lngDocumentCode, False) = False Then Exit Sub
            Try
                rtbDocument.LoadFile(strPath)
            Catch ex As Exception
                Call CheckWordApp(False)
                While WordApp.Documents.Count >= 1
                    WordApp.Documents(1).Close(False)
                End While
                Try
                    rtbDocument.LoadFile(strPath)
                Catch ex1 As Exception
                    Call ShowError("سند قابل نمايش نيست. اگر باز است آنرا ببنديد.")
                    Exit Sub
                End Try
            End Try
        End If
        strRtf = rtbDocument.Rtf

        '3-Justify alignment
        strRtf = Replace(strRtf, "\qr", "\qj")
        strRtf = Replace(strRtf, "\ql", "\qj")
        If Not strRtf.Contains("\qj") Then strRtf = Replace(strRtf, "\slmult0", "\slmult0\qj")

        '4-We should change formats like 123/45/67 to 67/45/123 in rtf
        Index = strRtf.IndexOf("/", 0)
        While Index <> -1
            intFrom = Index - 1
            While IsNumeric(strRtf.Substring(intFrom, 1)) And intFrom > 1
                intFrom -= 1
            End While
            intTo = Index + 1
            While (IsNumeric(strRtf.Substring(intTo, 1)) Or strRtf.Substring(intTo, 1) = "/") And intFrom < strRtf.Length
                intTo += 1
            End While
            strSource = Mid(strRtf, intFrom + 2, intTo - intFrom - 1)
            If strSource.Length > 2 Then
                strNumbers = Split(strSource, "/")
                strDest = ""
                For Index = strNumbers.Length - 1 To 0 Step -1
                    strDest &= "/" & strNumbers(Index)
                Next
                strDest = strDest.Remove(0, 1)
                strRtf = strRtf.Substring(0, intFrom + 1) & strDest & strRtf.Substring(intTo)
                'strRtf = Replace(strRtf, strSource, strDest, , 1)
            End If
            Index = strRtf.IndexOf("/", intTo)
        End While
        rtb.Width = (MyReportPage.PaperWidth - MyReportPage.LeftMargin - MyReportPage.RightMargin) * 3.7795275591

        '5-Set Spaces if need
        If blnSpaces Then
            strRtf = Replace(strRtf, "\slmult1", "\slmult0")
            While strRtf.Contains("\sl")
                Index = strRtf.IndexOf("\sl")
                Index = strRtf.Substring(Index + 3, strRtf.IndexOf("\", Index + 2) - Index - 3)
                strRtf = Replace(strRtf, "\sl" & Index & "\slmult0", "")
            End While
            strRtf = Replace(strRtf, "\pard", "")
            strRtf = Replace(strRtf, "\viewkind4\uc1", "\viewkind4\uc1\pard\sl-180\slmult0")
            Dim sngPaperHeight As Single = (MyReportPage.PaperHeight - MyReportPage.TopMargin - MyReportPage.BottomMargin) * 3.7795275591 - MyReportHeader.Height - MyReportFooter.Height
            rtb.Text = strRtf
            Dim sngRichHeight As Single = rtb.CalcHeight
            ' IIf(sngPaperHeight > sngRichHeight, 1, sngRichHeight / sngPaperHeight + 1)
            Dim intNewSpace As Integer = (sngPaperHeight * intTotalPapers) / sngRichHeight * 180
            strRtf = Replace(strRtf, "\sl-180", "\sl-" & intNewSpace)
            rtb.Text = strRtf
            sngRichHeight = rtb.CalcHeight
            While sngRichHeight > ((sngPaperHeight - 10) * intTotalPapers) 'If it is less no problem but if is bigger, small it
                intNewSpace -= 5
                strRtf = Replace(strRtf, "\sl-" & intNewSpace + 5, "\sl-" & intNewSpace)
                rtb.Text = strRtf
                sngRichHeight = rtb.CalcHeight
            End While
        Else
            rtb.Text = strRtf
        End If
        If blnIsPrint Then frpDocument.Print() Else frpDocument.Show()
    End Sub

End Module
