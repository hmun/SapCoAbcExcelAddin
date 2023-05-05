Imports SAP.Middleware.Connector

Public Class TData

    Public aTDataDic As Dictionary(Of String, TDataRec)
    Private aPar As SAPCommon.TStr
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Sub New(ByRef pPar As SAPCommon.TStr)
        aTDataDic = New Dictionary(Of String, TDataRec)
        aPar = pPar
    End Sub

    Public Sub addValue(pKey As String, pNAME As String, pVALUE As String, pCURRENCY As String, pFORMAT As String,
                        Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set")
        Dim aTDataRec As TDataRec
        If aTDataDic.ContainsKey(pKey) Then
            aTDataRec = aTDataDic(pKey)
            aTDataRec.setValues(pNAME, pVALUE, pCURRENCY, pFORMAT, pEmty, pEmptyChar, pOperation)
        Else
            aTDataRec = New TDataRec(aPar)
            aTDataRec.setValues(pNAME, pVALUE, pCURRENCY, pFORMAT, pEmty, pEmptyChar, pOperation)
            aTDataDic.Add(pKey, aTDataRec)
        End If
    End Sub

    Public Sub addValue(pKey As String, ByRef oStruc As IRfcStructure, Optional pStrucName As String = "")
        If Not oStruc Is Nothing Then
            Dim aStrucName As String = If(pStrucName = "", oStruc.Metadata.Name, pStrucName)
            For j As Integer = 0 To oStruc.Count - 1
                addValue(pKey, aStrucName & "-" & oStruc(j).Metadata.Name, CStr(oStruc(j).GetValue), "", "")
            Next
        End If
    End Sub

    Public Sub addValues(ByRef oTable As IRfcTable, Optional pStrucName As String = "")
        Dim oStruc As IRfcStructure = Nothing
        Dim aStrucName As String
        If Not oTable Is Nothing Then
            aStrucName = If(pStrucName = "", oTable(0).Metadata.Name, pStrucName)
            For i As Integer = 0 To oTable.Count - 1
                addValue(CStr(i), oTable(i), aStrucName)
            Next
            End If
    End Sub

    Public Sub addValue(pKey As String, pTStrRec As SAPCommon.TStrRec,
                        Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set",
                        Optional pNewStrucname As String = "")
        Dim aTDataRec As TDataRec
        Dim aName As String
        If pNewStrucname <> "" Then
            aName = pNewStrucname & "-" & pTStrRec.Fieldname
        Else
            aName = pTStrRec.Strucname & "-" & pTStrRec.Fieldname
        End If
        If aTDataDic.ContainsKey(pKey) Then
            aTDataRec = aTDataDic(pKey)
            aTDataRec.setValues(aName, pTStrRec.Value, pTStrRec.Currency, pTStrRec.Format, pEmty, pEmptyChar, pOperation)
        Else
            aTDataRec = New TDataRec(aPar)
            aTDataRec.setValues(aName, pTStrRec.Value, pTStrRec.Currency, pTStrRec.Format, pEmty, pEmptyChar, pOperation)
            aTDataDic.Add(pKey, aTDataRec)
        End If
    End Sub

    Public Sub delData(pKey As String)
        aTDataDic.Remove(pKey)
    End Sub

    Public Function getPostingRecord() As TDataRec
        Dim aTDataRec As TDataRec = Nothing
        Dim aKvb As KeyValuePair(Of String, TDataRec)
        For Each aKvb In aTDataDic
            aTDataRec = aKvb.Value
            If aTDataRec.getPost(aPar) <> "" Then
                getPostingRecord = aTDataRec
                Exit Function
            End If
        Next
        getPostingRecord = Nothing
    End Function

    Public Function getFirstRecord() As TDataRec
        Dim aTDataRec As TDataRec = Nothing
        Dim aKvb As KeyValuePair(Of String, TDataRec)
        aKvb = aTDataDic.ElementAt(0)
        getFirstRecord = Nothing
        If Not IsNothing(aKvb) Then
            getFirstRecord = aKvb.Value
        End If
    End Function

    Public Sub ws_parse(pKey As String, pWsName As String, ByRef pLoff As Integer, pCoff As Integer)
        Dim aDWS As Excel.Worksheet
        Dim aWB As Excel.Workbook
        aWB = Globals.SapCoAbcAddIn.Application.ActiveWorkbook
        Try
            aDWS = aWB.Worksheets(pWsName)
        Catch Exc As System.Exception
            log.Warn("ws_parse - " & "No " & pWsName & " Sheet in current workbook.")
            MsgBox("No " & pWsName & " Sheet in current workbook. Check the WS Parameters",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-ABC")
            Exit Sub
        End Try
        Dim jMax As UInt64 = pCoff + 1
        Do
            jMax += 1
        Loop While Not String.IsNullOrEmpty(aDWS.Cells(1, jMax).value)
        Dim i As UInt64 = pLoff + 1
        Dim j As Integer
        Dim aKey As String
        Do
            If CStr(aDWS.Cells(i, 1).value) = pKey Then
                aKey = CStr(i)
                For j = pCoff + 1 To jMax
                    If CStr(aDWS.Cells(pLoff - 3, j).value) <> "N/A" And CStr(aDWS.Cells(pLoff - 3, j).value) <> "" Then
                        addValue(aKey, CStr(aDWS.Cells(pLoff - 3, j).value), CStr(aDWS.Cells(i, j).value),
                                               CStr(aDWS.Cells(pLoff - 2, j).value), CStr(aDWS.Cells(pLoff - 1, j).value),
                                               pEmptyChar:="")
                    End If
                Next
            End If
            i += 1
        Loop While Not String.IsNullOrEmpty(aDWS.Cells(i, 1).value)
    End Sub

    Public Sub ws_output(pWsName As String, ByRef pLoff As Integer, pCoff As Integer, Optional pClear As Boolean = True, Optional pKey As String = "")
        Dim aDWS As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aRange As Excel.Range
        aWB = Globals.SapCoAbcAddIn.Application.ActiveWorkbook
        Try
            aDWS = aWB.Worksheets(pWsName)
        Catch Exc As System.Exception
            log.Warn("ws_output - " & "No " & pWsName & " Sheet in current workbook.")
            MsgBox("No " & pWsName & " Sheet in current workbook. Check the WS Parameters",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-ABC")
            Exit Sub
        End Try
        log.Debug("ws_output - " & "output to " & pWsName)
        Dim i As UInt64 = pLoff + 1
        Dim iMax As UInt64 = i - 1
        If pClear Then
            Do
                iMax += 1
            Loop While Not String.IsNullOrEmpty(aDWS.Cells(iMax, 1).value)
            If iMax > i Then
                aRange = aDWS.Range(aDWS.Cells(i, 1), aDWS.Cells(iMax, 1))
                aRange.EntireRow.Delete()
            End If
        End If
        ' read the header fields
        Dim j As UInt64 = pCoff + 1
        Dim aFieldArray() As String = {}
        Dim aOutArray() As String = {}
        Do
            Array.Resize(aFieldArray, aFieldArray.Length + 1)
            aFieldArray(aFieldArray.Length - 1) = CStr(aDWS.Cells(1, j).value)
            j += 1
        Loop While Not String.IsNullOrEmpty(aDWS.Cells(1, j).value)
        ' output
        Dim aKvB_Rec As KeyValuePair(Of String, TDataRec)
        Dim aDataRec As New TDataRec(aPar)
        For Each aKvB_Rec In aTDataDic
            aDataRec = aKvB_Rec.Value
            aOutArray = aDataRec.toArray(aFieldArray)
            aRange = aDWS.Range(aDWS.Cells(i, 1 + pCoff), aDWS.Cells(i, aFieldArray.Length + pCoff))
            aRange.Value = aOutArray
            If Not String.IsNullOrEmpty(pKey) Then
                aDWS.Cells(i, 1).value = pKey
            End If
            i += 1
        Next
        pLoff = i - 1
    End Sub

End Class
