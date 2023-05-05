Imports Microsoft.Office.Tools.Ribbon

Public Class SapCoAbcRibbon
    Private aSapCon
    Private aSapGeneral
    Private aTlPar As SAPCommon.TStr
    Private aIntPar As SAPCommon.TStr
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private Sub SapCoAbcRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        aSapGeneral = New SapGeneral
    End Sub

    Private Sub ButtonLogoff_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogoff.Click
        log.Debug("ButtonLogoff_Click - " & "starting logoff")
        If Not aSapCon Is Nothing Then
            log.Debug("ButtonLogoff_Click - " & "calling aSapCon.SAPlogoff()")
            aSapCon.SAPlogoff()
            aSapCon = Nothing
        End If
        log.Debug("ButtonLogoff_Click - " & "exit")
    End Sub

    Private Sub ButtonLogon_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogon.Click
        Dim aConRet As Integer

        log.Debug("ButtonLogon_Click - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            log.Debug("ButtonLogon_Click - " & "Version check failed")
            Exit Sub
        End If
        log.Debug("ButtonLogon_Click - " & "creating SapCon")
        If aSapCon Is Nothing Then
            aSapCon = New SapCon()
        End If
        log.Debug("ButtonLogon_Click - " & "calling SapCon.checkCon()")
        aConRet = aSapCon.checkCon()
        If aConRet = 0 Then
            log.Debug("ButtonLogon_Click - " & "connection successfull")
            MsgBox("SAP-Logon successful! ", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Sap CO-ABC")
        Else
            log.Debug("ButtonLogon_Click - " & "connection failed")
            aSapCon = Nothing
        End If
    End Sub

    Private Function getTlParameters() As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim akey As String
        Dim aName As String
        Dim i As Integer

        log.Debug("getTlParameters - " & "reading Parameter")
        aWB = Globals.SapCoAbcAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SAP CO-ABC Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-ABC")
            getTlParameters = False
            Exit Function
        End Try
        aName = "SAPTemplate"
        akey = CStr(aPws.Cells(1, 1).Value)
        If akey <> aName Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key " & aName & ". Check if the current workbook is a valid SAP CO-ABC Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-ABC")
            getTlParameters = False
            Exit Function
        End If
        i = 2
        aTlPar = New SAPCommon.TStr
        Do
            aTlPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 4).value), pFORMAT:=CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
        ' no obligatory parameters for TlDoc - otherwise check here
        getTlParameters = True
    End Function

    Private Function getTlImpParameters(pPar As SAPCommon.TStr) As SAPCommon.TStr
        Dim aKvb As KeyValuePair(Of String, SAPCommon.TStrRec)
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aNewHdrRec As New TDataRec(aIntPar)
        Dim aPar As SAPCommon.TStr = New SAPCommon.TStr
        For Each aKvb In pPar.getData()
            aTStrRec = aKvb.Value
            If aTStrRec.Fieldname = "CONTROLLINGAREA" Or aTStrRec.Fieldname = "ENVIRONMENT" Then
                aPar.add(aTStrRec.Strucname, aTStrRec.Fieldname & "_IMP", aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
            Else
                aPar.add(aTStrRec.Strucname, aTStrRec.Fieldname, aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
            End If
        Next
        getTlImpParameters = aPar
    End Function

    Private Function getIntParameters() As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim i As Integer

        log.Debug("getIntParameters - " & "reading Parameter")
        aWB = Globals.SapCoAbcAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter_Int")
        Catch Exc As System.Exception
            MsgBox("No Parameter_Int Sheet in current workbook. Check if the current workbook is a valid SAP CO-ABC Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-ABC")
            getIntParameters = False
            Exit Function
        End Try
        i = 2
        aIntPar = New SAPCommon.TStr
        Do
            aIntPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
        ' no obligatory parameters check - we should know what we are doing
        getIntParameters = True
    End Function

    Private Function checkCon() As Integer
        Dim aSapConRet As Integer
        Dim aSapVersionRet As Integer
        checkCon = False
        log.Debug("checkCon - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            Exit Function
        End If
        log.Debug("checkCon - " & "checking Connection")
        aSapConRet = 0
        If aSapCon Is Nothing Then
            Try
                aSapCon = New SapCon()
            Catch ex As SystemException
                log.Warn("checkCon-New SapCon - )" & ex.ToString)
            End Try
        End If
        Try
            aSapConRet = aSapCon.checkCon()
        Catch ex As SystemException
            log.Warn("checkCon-aSapCon.checkCon - )" & ex.ToString)
        End Try
        If aSapConRet = 0 Then
            log.Debug("checkCon - " & "checking version in SAP")
            Try
                aSapVersionRet = aSapGeneral.checkVersionInSAP(aSapCon)
            Catch ex As SystemException
                log.Warn("checkCon - )" & ex.ToString)
            End Try
            log.Debug("checkCon - " & "aSapVersionRet=" & CStr(aSapVersionRet))
            If aSapVersionRet = True Then
                log.Debug("checkCon - " & "checkCon = True")
                checkCon = True
            Else
                log.Debug("checkCon - " & "connection check failed")
            End If
        End If
    End Function

    Private Sub ButtonSapTLRead_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSapTLRead.Click
        If checkCon() = True Then
            ButtonSapTLRead_Exec()
        End If
    End Sub

    Private Sub ButtonSapTLRead_Exec()
        Dim aSAPTemplateCO As New SAPTemplateCO(aSapCon)
        ' get posting parameters
        If Not getTlParameters() Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters() Then
            Exit Sub
        End If

        Dim jMax As UInt64 = 0
        Dim aTllLOff As Integer = If(aIntPar.value("LOFF", "TLLIST") <> "", CInt(aIntPar.value("LOFF", "TLLIST")), 4)
        Dim aPosLOff As Integer = If(aIntPar.value("LOFF", "TLPOS") <> "", CInt(aIntPar.value("LOFF", "TLPOS")), 4)
        Dim aCfvLOff As Integer = If(aIntPar.value("LOFF", "TLCFV") <> "", CInt(aIntPar.value("LOFF", "TLCFV")), 4)
        Dim aCsoLOff As Integer = If(aIntPar.value("LOFF", "TLCSO") <> "", CInt(aIntPar.value("LOFF", "TLCSO")), 4)
        Dim aFfsLOff As Integer = If(aIntPar.value("LOFF", "TLFFS") <> "", CInt(aIntPar.value("LOFF", "TLFFS")), 4)
        Dim aTllWsName As String = If(aIntPar.value("WS", "TLLIST") <> "", aIntPar.value("WS", "TLLIST"), "TL_List")
        Dim aPosWsName As String = If(aIntPar.value("WS", "TLPOS") <> "", aIntPar.value("WS", "TLPOS"), "TL_Positions")
        Dim aCfvWsName As String = If(aIntPar.value("WS", "TLCFV") <> "", aIntPar.value("WS", "TLCFV"), "TL_CellFixValues")
        Dim aCSoWsName As String = If(aIntPar.value("WS", "TLCSO") <> "", aIntPar.value("WS", "TLCSO"), "TL_CellSources")
        Dim aFfsWsName As String = If(aIntPar.value("WS", "TLFFS") <> "", aIntPar.value("WS", "TLFFS"), "TL_FlexFuncSources")
        Dim aTllWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value("COL", "DATAMSG") <> "", aIntPar.value("COL", "DATAMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aTltClmn As String = If(aIntPar.value("COL", "TLTXT") <> "", aIntPar.value("COL", "TLTXT"), "TEMPLATETEXT")
        Dim aTltClmnNr As Integer = 0
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value("TL_RET", "OKMSG") <> "", aIntPar.value("TL_RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.SapCoAbcAddIn.Application.ActiveWorkbook
        Try
            aTllWs = aWB.Worksheets(aTllWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aTllWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP CO-ABC Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-ABC")
            Exit Sub
        End Try

        Do
            jMax += 1
            If CStr(aTllWs.Cells(1, jMax).value) = aMsgClmn Then
                aMsgClmnNr = jMax
            ElseIf CStr(aTllWs.Cells(1, jMax).value) = aTltClmn Then
                aTltClmnNr = jMax
            End If
        Loop While CStr(aTllWs.Cells(aTllLOff - 3, jMax + 1).value) <> ""

        aTllWs.Activate()
        Try
            log.Debug("ButtonSapTLRead_Exec - " & "processing data - disabling events, screen update, cursor")
            Globals.SapCoAbcAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapCoAbcAddIn.Application.EnableEvents = False
            Globals.SapCoAbcAddIn.Application.ScreenUpdating = False
            Dim i As UInt64 = aTllLOff + 1
            Dim aKey As String
            Do
                If Left(CStr(aTllWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    aKey = CStr(i)
                    Dim aTlItems As New TData(aIntPar)
                    aTlItems.addValue(aKey, CStr(aTllWs.Cells(aTllLOff - 3, 1).value), CStr(aTllWs.Cells(i, 1).value), "", "")
                    Dim aTSAP_TLData As New TSAP_TLData(aTlPar, aIntPar, aSAPTemplateCO, "GetDetail")
                    If aTSAP_TLData.fillHeader(aTlItems) Then
                        log.Debug("ButtonSapTLRead_Exec - " & "calling aSAPTemplateCO.GetDetail")
                        aRetStr = aSAPTemplateCO.GetDetail(aTSAP_TLData, pOKMsg:=aOKMsg)
                        log.Debug("ButtonSapTLRead_Exec - " & "aSAPTemplateCO.GetDetail returned, aRetStr=" & aRetStr)
                        aTllWs.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                        ' output the data now
                        Dim aClear As Boolean = False
                        If i = aTllLOff + 1 Then
                            aClear = True
                        End If
                        Dim aTData As TData
                        Dim aTemplate As String = aTSAP_TLData.aHdrRec.getTemplate()
                        aTllWs.Cells(i, aTltClmnNr) = aTSAP_TLData.aHdrRec.getText()
                        If aTSAP_TLData.aDataDic.aTDataDic.ContainsKey("POSITIONS") Then
                            aTData = aTSAP_TLData.aDataDic.aTDataDic("POSITIONS")
                            aTData.ws_output(aPosWsName, aPosLOff, 1, pClear:=aClear, pKey:=aTemplate)
                        End If
                        If aTSAP_TLData.aDataDic.aTDataDic.ContainsKey("CELLFIXVALUES") Then
                            aTData = aTSAP_TLData.aDataDic.aTDataDic("CELLFIXVALUES")
                            aTData.ws_output(aCfvWsName, aCfvLOff, 1, pClear:=aClear, pKey:=aTemplate)
                        End If
                        If aTSAP_TLData.aDataDic.aTDataDic.ContainsKey("CELLSOURCES") Then
                            aTData = aTSAP_TLData.aDataDic.aTDataDic("CELLSOURCES")
                            aTData.ws_output(aCSoWsName, aCsoLOff, 1, pClear:=aClear, pKey:=aTemplate)
                        End If
                        If aTSAP_TLData.aDataDic.aTDataDic.ContainsKey("FLEXFUNCSOURCES") Then
                            aTData = aTSAP_TLData.aDataDic.aTDataDic("FLEXFUNCSOURCES")
                            aTData.ws_output(aFfsWsName, aFfsLOff, 1, pClear:=aClear, pKey:=aTemplate)
                        End If
                    End If
                End If
                i += 1
            Loop While CStr(aTllWs.Cells(i, 1).value) <> ""

            log.Debug("ButtonSapTLRead_Exec - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapCoAbcAddIn.Application.EnableEvents = True
            Globals.SapCoAbcAddIn.Application.ScreenUpdating = True
            Globals.SapCoAbcAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapCoAbcAddIn.Application.EnableEvents = True
            Globals.SapCoAbcAddIn.Application.ScreenUpdating = True
            Globals.SapCoAbcAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("ButtonSapTLRead_Exec failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-ABC")
            log.Error("ButtonSapTLRead_Exec - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Private Sub ButtonSapTLDeleteCheck_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSapTLDeleteCheck.Click
        If checkCon() = True Then
            ButtonSapTLDelete_Exec(pTest:=True)
        End If
    End Sub

    Private Sub ButtonSapTLDeletePost_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSapTLDeletePost.Click
        If checkCon() = True Then
            ButtonSapTLDelete_Exec(pTest:=False)
        End If
    End Sub

    Private Sub ButtonSapTLDelete_Exec(pTest As String)
        Dim aSAPTemplateCO As New SAPTemplateCO(aSapCon)
        ' get posting parameters
        If Not getTlParameters() Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters() Then
            Exit Sub
        End If

        Dim jMax As UInt64 = 0
        Dim aTllLOff As Integer = If(aIntPar.value("LOFF", "TLLIST") <> "", CInt(aIntPar.value("LOFF", "TLLIST")), 4)
        Dim aTllWsName As String = If(aIntPar.value("WS", "TLLIST") <> "", aIntPar.value("WS", "TLLIST"), "TL_List")
        Dim aTllWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value("COL", "DATAMSG") <> "", aIntPar.value("COL", "DATAMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aTltClmn As String = If(aIntPar.value("COL", "TLTXT") <> "", aIntPar.value("COL", "TLTXT"), "TEMPLATETEXT")
        Dim aTltClmnNr As Integer = 0
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value("TL_RET", "OKMSG") <> "", aIntPar.value("TL_RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.SapCoAbcAddIn.Application.ActiveWorkbook
        Try
            aTllWs = aWB.Worksheets(aTllWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aTllWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP CO-ABC Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-ABC")
            Exit Sub
        End Try

        Do
            jMax += 1
            If CStr(aTllWs.Cells(1, jMax).value) = aMsgClmn Then
                aMsgClmnNr = jMax
            ElseIf CStr(aTllWs.Cells(1, jMax).value) = aTltClmn Then
                aTltClmnNr = jMax
            End If
        Loop While CStr(aTllWs.Cells(aTllLOff - 3, jMax + 1).value) <> ""

        aTllWs.Activate()
        Try
            log.Debug("ButtonSapTLDelete_Exec - " & "processing data - disabling events, screen update, cursor")
            Globals.SapCoAbcAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            ' Globals.SapCoAbcAddIn.Application.EnableEvents = False
            ' Globals.SapCoAbcAddIn.Application.ScreenUpdating = False
            Dim i As UInt64 = aTllLOff + 1
            Dim aKey As String
            Do
                If Left(CStr(aTllWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    aKey = CStr(i)
                    Dim aTlItems As New TData(aIntPar)
                    aTlItems.addValue(aKey, CStr(aTllWs.Cells(aTllLOff - 3, 1).value), CStr(aTllWs.Cells(i, 1).value), "", "")
                    Dim aTSAP_TLData As New TSAP_TLData(aTlPar, aIntPar, aSAPTemplateCO, "Delete")
                    If aTSAP_TLData.fillHeader(aTlItems) Then
                        log.Debug("ButtonSapTLDelete_Exec - " & "calling aSAPTemplateCO.Delete")
                        Globals.SapCoAbcAddIn.Application.StatusBar = "Posting at line " & i
                        aRetStr = aSAPTemplateCO.Delete(aTSAP_TLData, pCheck:=pTest, pOKMsg:=aOKMsg)
                        log.Debug("ButtonSapTLDelete_Exec - " & "aSAPTemplateCO.Delete returned, aRetStr=" & aRetStr)
                        aTllWs.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                        ' output the data now
                        Dim aClear As Boolean = False
                        If i = aTllLOff + 1 Then
                            aClear = True
                        End If
                    End If
                End If
                i += 1
            Loop While CStr(aTllWs.Cells(i, 1).value) <> ""

            log.Debug("ButtonSapTLDelete_Exec - " & "all data processed - enabling events, screen update, cursor")
            ' Globals.SapCoAbcAddIn.Application.EnableEvents = True
            ' Globals.SapCoAbcAddIn.Application.ScreenUpdating = True
            Globals.SapCoAbcAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            ' Globals.SapCoAbcAddIn.Application.EnableEvents = True
            ' Globals.SapCoAbcAddIn.Application.ScreenUpdating = True
            Globals.SapCoAbcAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("ButtonSapTLDelete_Exec failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-ABC")
            log.Error("ButtonSapTLDelete_Exec - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try

    End Sub

    Private Sub ButtonSapTLCreateCheck_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSapTLCreateCheck.Click
        If checkCon() = True Then
            ButtonSapTLCreate_Exec(pTest:=True)
        End If
    End Sub

    Private Sub ButtonSapTLCreatePost_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSapTLCreatePost.Click
        If checkCon() = True Then
            ButtonSapTLCreate_Exec(pTest:=False)
        End If
    End Sub

    Private Sub ButtonSapTLCreate_Exec(pTest As Boolean)
        Dim aSAPTemplateCO As New SAPTemplateCO(aSapCon)
        ' get posting parameters
        If Not getTlParameters() Then
            Exit Sub
        End If
        aTlPar = getTlImpParameters(aTlPar)
        ' get internal parameters
        If Not getIntParameters() Then
            Exit Sub
        End If

        Dim jMax As UInt64 = 0
        Dim aTllLOff As Integer = If(aIntPar.value("LOFF", "TLLIST") <> "", CInt(aIntPar.value("LOFF", "TLLIST")), 4)
        Dim aPosLOff As Integer = If(aIntPar.value("LOFF", "TLPOS") <> "", CInt(aIntPar.value("LOFF", "TLPOS")), 4)
        Dim aCfvLOff As Integer = If(aIntPar.value("LOFF", "TLCFV") <> "", CInt(aIntPar.value("LOFF", "TLCFV")), 4)
        Dim aCsoLOff As Integer = If(aIntPar.value("LOFF", "TLCSO") <> "", CInt(aIntPar.value("LOFF", "TLCSO")), 4)
        Dim aFfsLOff As Integer = If(aIntPar.value("LOFF", "TLFFS") <> "", CInt(aIntPar.value("LOFF", "TLFFS")), 4)
        Dim aTllWsName As String = If(aIntPar.value("WS", "TLLIST") <> "", aIntPar.value("WS", "TLLIST"), "TL_List")
        Dim aPosWsName As String = If(aIntPar.value("WS", "TLPOS") <> "", aIntPar.value("WS", "TLPOS"), "TL_Positions")
        Dim aCfvWsName As String = If(aIntPar.value("WS", "TLCFV") <> "", aIntPar.value("WS", "TLCFV"), "TL_CellFixValues")
        Dim aCSoWsName As String = If(aIntPar.value("WS", "TLCSO") <> "", aIntPar.value("WS", "TLCSO"), "TL_CellSources")
        Dim aFfsWsName As String = If(aIntPar.value("WS", "TLFFS") <> "", aIntPar.value("WS", "TLFFS"), "TL_FlexFuncSources")
        Dim aTllWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value("COL", "DATAMSG") <> "", aIntPar.value("COL", "DATAMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value("TL_RET", "OKMSG") <> "", aIntPar.value("TL_RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.SapCoAbcAddIn.Application.ActiveWorkbook
        Try
            aTllWs = aWB.Worksheets(aTllWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aTllWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP CO-ABC Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-ABC")
            Exit Sub
        End Try

        Do
            jMax += 1
            If CStr(aTllWs.Cells(1, jMax).value) = aMsgClmn Then
                aMsgClmnNr = jMax
            End If
        Loop While CStr(aTllWs.Cells(aTllLOff - 3, jMax + 1).value) <> ""

        aTllWs.Activate()
        Try
            log.Debug("ButtonSapTLCreate_Exec - " & "processing data - disabling events, screen update, cursor")
            Globals.SapCoAbcAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            '            Globals.SapCoAbcAddIn.Application.EnableEvents = False
            '            Globals.SapCoAbcAddIn.Application.ScreenUpdating = False
            Dim i As UInt64 = aTllLOff + 1
            Dim aKey As String
            Do
                If Left(CStr(aTllWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    aKey = CStr(i)
                    Dim aTemplate As String = CStr(aTllWs.Cells(i, 1).value)
                    Dim aTlItems As New TData(aIntPar)
                    aTlItems.addValue(aKey, CStr(aTllWs.Cells(aTllLOff - 3, 1).value) & "_IMP", aTemplate, "", "")
                    aTlItems.addValue(aKey, CStr(aTllWs.Cells(aTllLOff - 3, 2).value), CStr(aTllWs.Cells(i, 2).value), "", "")
                    Dim aTSAP_TLData As New TSAP_TLData(aTlPar, aIntPar, aSAPTemplateCO, "CreateFromData")
                    ' read POSITIONS
                    aTlItems.ws_parse(aTemplate, aPosWsName, aPosLOff, 1)
                    ' read CELLFIXVALUES
                    aTlItems.ws_parse(aTemplate, aCfvWsName, aCfvLOff, 1)
                    ' read CELLSOURCES
                    aTlItems.ws_parse(aTemplate, aCSoWsName, aCsoLOff, 1)
                    ' read FLEXFUNCSOURCES
                    aTlItems.ws_parse(aTemplate, aFfsWsName, aFfsLOff, 1)
                    If aTSAP_TLData.fillHeader(aTlItems) And aTSAP_TLData.fillData(aTlItems) Then
                        log.Debug("ButtonSapTLCreate_Exec - " & "calling aSAPTemplateCO.CreateFromData")
                        Globals.SapCoAbcAddIn.Application.StatusBar = "Posting at line " & i
                        aRetStr = aSAPTemplateCO.CreateFromData(aTSAP_TLData, pCheck:=pTest, pOKMsg:=aOKMsg)
                        log.Debug("ButtonSapTLCreate_Exec - " & "aSAPTemplateCO.CreateFromData returned, aRetStr=" & aRetStr)
                        aTllWs.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                    End If
                End If
                i += 1
            Loop While CStr(aTllWs.Cells(i, 1).value) <> ""

            log.Debug("ButtonSapTLCreate_Exec - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapCoAbcAddIn.Application.EnableEvents = True
            Globals.SapCoAbcAddIn.Application.ScreenUpdating = True
            Globals.SapCoAbcAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapCoAbcAddIn.Application.EnableEvents = True
            Globals.SapCoAbcAddIn.Application.ScreenUpdating = True
            Globals.SapCoAbcAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("ButtonSapTLCreate_Exec failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-ABC")
            log.Error("ButtonSapTLCreate_Exec - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Private Sub ButtonSapTLGenerate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSapTLGenerate.Click
        If checkCon() = True Then
            ButtonSapTLGenerate_Exec()
        End If
    End Sub


    Private Sub ButtonSapTLGenerate_Exec()
        Dim aSAPTemplateCO As New SAPTemplateCO(aSapCon)
        ' get posting parameters
        If Not getTlParameters() Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters() Then
            Exit Sub
        End If

        Dim jMax As UInt64 = 0
        Dim aTllLOff As Integer = If(aIntPar.value("LOFF", "TLLIST") <> "", CInt(aIntPar.value("LOFF", "TLLIST")), 4)
        Dim aTllWsName As String = If(aIntPar.value("WS", "TLLIST") <> "", aIntPar.value("WS", "TLLIST"), "TL_List")
        Dim aTllWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value("COL", "DATAMSG") <> "", aIntPar.value("COL", "DATAMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aTltClmn As String = If(aIntPar.value("COL", "TLTXT") <> "", aIntPar.value("COL", "TLTXT"), "TEMPLATETEXT")
        Dim aTltClmnNr As Integer = 0
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value("TL_RET", "OKMSG") <> "", aIntPar.value("TL_RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.SapCoAbcAddIn.Application.ActiveWorkbook
        Try
            aTllWs = aWB.Worksheets(aTllWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aTllWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP CO-ABC Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-ABC")
            Exit Sub
        End Try

        Do
            jMax += 1
            If CStr(aTllWs.Cells(1, jMax).value) = aMsgClmn Then
                aMsgClmnNr = jMax
            ElseIf CStr(aTllWs.Cells(1, jMax).value) = aTltClmn Then
                aTltClmnNr = jMax
            End If
        Loop While CStr(aTllWs.Cells(aTllLOff - 3, jMax + 1).value) <> ""

        aTllWs.Activate()
        Try
            log.Debug("ButtonSapTLGenerate_Exec - " & "processing data - disabling events, screen update, cursor")
            Globals.SapCoAbcAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            ' Globals.SapCoAbcAddIn.Application.EnableEvents = False
            ' Globals.SapCoAbcAddIn.Application.ScreenUpdating = False
            Dim i As UInt64 = aTllLOff + 1
            Dim aKey As String
            Do
                If Left(CStr(aTllWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    aKey = CStr(i)
                    Dim aTlItems As New TData(aIntPar)
                    aTlItems.addValue(aKey, CStr(aTllWs.Cells(aTllLOff - 3, 1).value), CStr(aTllWs.Cells(i, 1).value), "", "")
                    Dim aTSAP_TLData As New TSAP_TLData(aTlPar, aIntPar, aSAPTemplateCO, "Generate")
                    If aTSAP_TLData.fillHeader(aTlItems) Then
                        log.Debug("ButtonSapTLGenerate_Exec - " & "calling aSAPTemplateCO.Generate")
                        Globals.SapCoAbcAddIn.Application.StatusBar = "Posting at line " & i
                        aRetStr = aSAPTemplateCO.Generate(aTSAP_TLData, pOKMsg:=aOKMsg)
                        log.Debug("ButtonSapTLGenerate_Exec - " & "aSAPTemplateCO.Generate returned, aRetStr=" & aRetStr)
                        aTllWs.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                        ' output the data now
                        Dim aClear As Boolean = False
                        If i = aTllLOff + 1 Then
                            aClear = True
                        End If
                    End If
                End If
                i += 1
            Loop While CStr(aTllWs.Cells(i, 1).value) <> ""

            log.Debug("ButtonSapTLGenerate_Exec - " & "all data processed - enabling events, screen update, cursor")
            ' Globals.SapCoAbcAddIn.Application.EnableEvents = True
            ' Globals.SapCoAbcAddIn.Application.ScreenUpdating = True
            Globals.SapCoAbcAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            ' Globals.SapCoAbcAddIn.Application.EnableEvents = True
            ' Globals.SapCoAbcAddIn.Application.ScreenUpdating = True
            Globals.SapCoAbcAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("ButtonSapTLGenerate_Exec failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-ABC")
            log.Error("ButtonSapTLGenerate_Exec - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try

    End Sub
End Class
