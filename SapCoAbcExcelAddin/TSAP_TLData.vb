Public Class TSAP_TLData

    Public aHdrRec As TDataRec
    Public aDataDic As TDataDic

    Private Template_Fields_GetDetail() As String = {"CONTROLLINGAREA", "ENVIRONMENT", "TEMPLATE", "LANGUAGE"}
    Private Template_Fields_CreateFromData() As String = {"CONTROLLINGAREA_IMP", "ENVIRONMENT_IMP", "TEMPLATE_IMP", "TEMPLATETEXT", "LANGUAGE", "TESTRUN"}
    Private Template_Fields_Delete() As String = {"CONTROLLINGAREA", "ENVIRONMENT", "TEMPLATE", "TESTRUN"}
    Private Template_Fields_Generate() As String = {"CONTROLLINGAREA", "ENVIRONMENT", "TEMPLATE"}

    Private Pos_Fields() As String = {}
    Private Cfv_Fields() As String = {}
    Private Cso_Fields() As String = {}
    Private Ffs_Fields() As String = {}

    Private Const sPos As String = "POSITIONS"
    Private Const sCfv As String = "CELLFIXVALUES"
    Private Const sCso As String = "CELLSOURCES"
    Private Const sFfs As String = "FLEXFUNCSOURCES"

    Private aTlPar As SAPCommon.TStr
    Private aIntPar As SAPCommon.TStr
    Private aFunction As String
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Sub New(ByRef pTlPar As SAPCommon.TStr, ByRef pIntPar As SAPCommon.TStr, ByRef pSAPTemplateCO As SAPTemplateCO, pFunction As String)
        aTlPar = pTlPar
        aIntPar = pIntPar
        aFunction = pFunction
        aDataDic = New TDataDic(aIntPar)
        aHdrRec = New TDataRec(aIntPar)
        ' get Metadata
        If pFunction = "CreateFromData" Then
            pSAPTemplateCO.getMeta_CreateFromData(Pos_Fields, Cfv_Fields, Cso_Fields, Ffs_Fields)
        ElseIf pFunction = "GetDetail" Then
            pSAPTemplateCO.getMeta_GetDetail(Pos_Fields, Cfv_Fields, Cso_Fields, Ffs_Fields)
        End If
    End Sub

    Public Function fillHeader(pData As TData) As Boolean
        Dim aKvb As KeyValuePair(Of String, SAPCommon.TStrRec)
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aNewHdrRec As New TDataRec(aIntPar)
        For Each aKvb In aTlPar.getData()
            aTStrRec = aKvb.Value
            If valid_Template_Field(aTStrRec) Then
                aNewHdrRec.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmptyChar:="")
            End If
        Next
        ' First fill the value from the paramters and tehn overwrite them from the posting record
        Dim aPostRec As New TDataRec(aIntPar)
        aPostRec = pData.getFirstRecord()
        For Each aTStrRec In aPostRec.aTDataRecCol
            If valid_Template_Field(aTStrRec) Then
                aNewHdrRec.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
            End If
        Next
        aHdrRec = aNewHdrRec
        fillHeader = True
    End Function

    Public Function fillData(pData As TData) As Boolean
        Dim aKvB As KeyValuePair(Of String, TDataRec)
        Dim aTDataRec As TDataRec
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aCnt As UInt64
        aDataDic = New TDataDic(aIntPar)
        fillData = True
        aCnt = 1
        For Each aKvB In pData.aTDataDic
            aTDataRec = aKvB.Value
            For Each aTStrRec In aTDataRec.aTDataRecCol
                If valid_Pos_Field(aTStrRec) Then
                    aDataDic.addValue(CStr(aCnt), aTStrRec, pNewStrucname:=sPos)
                ElseIf valid_cfv_Field(aTStrRec) Then
                    aDataDic.addValue(CStr(aCnt), aTStrRec, pNewStrucname:=sCfv)
                ElseIf valid_Cso_Field(aTStrRec) Then
                    aDataDic.addValue(CStr(aCnt), aTStrRec, pNewStrucname:=sCso)
                ElseIf valid_Ffs_Field(aTStrRec) Then
                    aDataDic.addValue(CStr(aCnt), aTStrRec, pNewStrucname:=sFfs)
                End If
            Next
            aCnt += 1
        Next
    End Function

    Public Function valid_Template_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        valid_Template_Field = False
        If String.IsNullOrEmpty(pTStrRec.Strucname) Then
            Select Case aFunction
                Case "GetDetail"
                    valid_Template_Field = isInArray(pTStrRec.Fieldname, Template_Fields_GetDetail)
                Case "CreateFromData"
                    valid_Template_Field = isInArray(pTStrRec.Fieldname, Template_Fields_CreateFromData)
                Case "Delete"
                    valid_Template_Field = isInArray(pTStrRec.Fieldname, Template_Fields_Delete)
                Case "Generate"
                    valid_Template_Field = isInArray(pTStrRec.Fieldname, Template_Fields_Generate)
            End Select
        End If
    End Function


    Public Function valid_Pos_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Pos_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("POSITIONS", aStrucName) Then
            valid_Pos_Field = isInArray(pTStrRec.Fieldname, Pos_Fields)
        End If
    End Function

    Public Function valid_Cfv_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Cfv_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("CELLFIXVALUES", aStrucName) Then
            valid_Cfv_Field = isInArray(pTStrRec.Fieldname, Cfv_Fields)
        End If
    End Function

    Public Function valid_Cso_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Cso_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("CELLSOURCES", aStrucName) Then
            valid_Cso_Field = isInArray(pTStrRec.Fieldname, Cso_Fields)
        End If
    End Function

    Public Function valid_Ffs_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Ffs_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("FLEXFUNCSOURCES", aStrucName) Then
            valid_Ffs_Field = isInArray(pTStrRec.Fieldname, Ffs_Fields)
        End If
    End Function

    Private Function isInArray(pString As String, pArray As Object) As Boolean
        Dim st As String, M As String
        M = "$"
        st = M & Join(pArray, M) & M
        isInArray = InStr(st, M & pString & M) > 0
        ' isInArray = (UBound(Filter(pArray, pString)) > -1)
    End Function

    Public Sub dumpHeader()
        Dim dumpHd As String = If(aIntPar.value("DBG", "DUMPHEADER") <> "", aIntPar.value("DBG", "DUMPHEADER"), "")
        If dumpHd <> "" Then
            Dim aDWS As Excel.Worksheet
            Dim aWB As Excel.Workbook
            Dim aRange As Excel.Range
            aWB = Globals.SapCoAbcAddIn.Application.ActiveWorkbook
            Try
                aDWS = aWB.Worksheets(dumpHd)
                aDWS.Activate()
            Catch Exc As System.Exception
                log.Warn("dumpHeader - " & "No " & dumpHd & " Sheet in current workbook.")
                MsgBox("No " & dumpHd & " Sheet in current workbook. Check the DBG-DUMPHEADR Parameter",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
                Exit Sub
            End Try
            log.Debug("dumpHeader - " & "dumping to " & dumpHd)
            ' clear the Header
            If CStr(aDWS.Cells(1, 1).Value) <> "" Then
                aRange = aDWS.Range(aDWS.Cells(1, 1), aDWS.Cells(1000, 1))
                aRange.EntireRow.Delete()
            End If
            ' dump the Header
            Dim aTStrRec As New SAPCommon.TStrRec
            Dim aFieldArray() As String = {}
            Dim aValueArray() As String = {}
            For Each aTStrRec In aHdrRec.aTDataRecCol
                Array.Resize(aFieldArray, aFieldArray.Length + 1)
                aFieldArray(aFieldArray.Length - 1) = aTStrRec.getKey()
                Array.Resize(aValueArray, aValueArray.Length + 1)
                aValueArray(aValueArray.Length - 1) = aTStrRec.formated()
            Next
            aRange = aDWS.Range(aDWS.Cells(1, 1), aDWS.Cells(1, aFieldArray.Length))
            aRange.Value = aFieldArray
            aRange = aDWS.Range(aDWS.Cells(2, 1), aDWS.Cells(2, aValueArray.Length))
            aRange.Value = aValueArray
        End If
    End Sub

    Public Sub dumpData()
        Dim dumpDt As String = If(aIntPar.value("DBG", "DUMPDATA") <> "", aIntPar.value("DBG", "DUMPDATA"), "")
        If dumpDt <> "" Then
            Dim aDWS As Excel.Worksheet
            Dim aWB As Excel.Workbook
            Dim aRange As Excel.Range
            aWB = Globals.SapCoAbcAddIn.Application.ActiveWorkbook
            Try
                aDWS = aWB.Worksheets(dumpDt)
                aDWS.Activate()
            Catch Exc As System.Exception
                log.Warn("dumpData - " & "No " & dumpDt & " Sheet in current workbook.")
                MsgBox("No " & dumpDt & " Sheet in current workbook. Check the DBG-DUMPDATA Parameter",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
                Exit Sub
            End Try
            log.Debug("dumpData - " & "dumping to " & dumpDt)
            ' clear the Data
            If CStr(aDWS.Cells(5, 1).Value) <> "" Then
                aRange = aDWS.Range(aDWS.Cells(5, 1), aDWS.Cells(1000, 1))
                aRange.EntireRow.Delete()
            End If

            Dim aKvB_Dic As KeyValuePair(Of String, TData)
            Dim aKvB_Rec As KeyValuePair(Of String, TDataRec)
            Dim aData As TData
            Dim aData_Am As New TData(aIntPar)
            Dim aDataRec As New TDataRec(aIntPar)
            Dim aDataRec_Am As New TDataRec(aIntPar)
            Dim i As Int64
            Dim aTStrRec As New SAPCommon.TStrRec
            i = 6
            For Each aKvB_Dic In aDataDic.aTDataDic
                aData = aKvB_Dic.Value
                aDWS.Cells(i, 1).Value = aKvB_Dic.Key
                For Each aKvB_Rec In aData.aTDataDic
                    aDataRec = aKvB_Rec.Value
                    Dim aFieldArray() As String = {}
                    Dim aValueArray() As String = {}
                    For Each aTStrRec In aDataRec.aTDataRecCol
                        Array.Resize(aFieldArray, aFieldArray.Length + 1)
                        aFieldArray(aFieldArray.Length - 1) = aTStrRec.getKey()
                        Array.Resize(aValueArray, aValueArray.Length + 1)
                        aValueArray(aValueArray.Length - 1) = aTStrRec.formated()
                    Next
                    aRange = aDWS.Range(aDWS.Cells(i, 1), aDWS.Cells(i, aFieldArray.Length))
                    aRange.Value = aFieldArray
                    aRange = aDWS.Range(aDWS.Cells(i + 1, 1), aDWS.Cells(i + 1, aValueArray.Length))
                    aRange.Value = aValueArray
                    i += 2
                Next
                i += 2
            Next
        End If
    End Sub

End Class
