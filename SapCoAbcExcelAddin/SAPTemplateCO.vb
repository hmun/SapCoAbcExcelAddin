' Copyright 2022 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPTemplateCO

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        Try
            log.Debug("New - " & "checking connection")
            sapcon = aSapCon
            aSapCon.getDestination(destination)
            sapcon.checkCon()
        Catch ex As System.Exception
            log.Error("New - Exception=" & ex.ToString)
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPTemplateCO")
        End Try
    End Sub

    Public Sub getMeta_GetDetail(ByRef pPos_Fields() As String, ByRef pCfv_Fields() As String, ByRef pCso_Fields() As String, ByRef pFfs_Fields() As String)
        Try
            log.Debug("getMeta_GetDetail - " & "creating Function BAPI_TEMPLATECO_GET_DETAIL")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_TEMPLATECO_GET_DETAIL")
            Dim oPOSITIONS As IRfcTable = oRfcFunction.GetTable("POSITIONS")
            Dim oCELLFIXVALUES As IRfcTable = oRfcFunction.GetTable("CELLFIXVALUES")
            Dim oCELLSOURCES As IRfcTable = oRfcFunction.GetTable("CELLSOURCES")
            Dim oFLEXFUNCSOURCES As IRfcTable = oRfcFunction.GetTable("FLEXFUNCSOURCES")
            ' Pos_Fields
            pPos_Fields = {}
            For i As Integer = 0 To oPOSITIONS.Metadata.LineType.FieldCount - 1
                Array.Resize(pPos_Fields, pPos_Fields.Length + 1)
                pPos_Fields(pPos_Fields.Length - 1) = oPOSITIONS.Metadata.LineType(i).Name
            Next
            ' Cfv_Fields
            pCfv_Fields = {}
            For i As Integer = 0 To oCELLFIXVALUES.Metadata.LineType.FieldCount - 1
                Array.Resize(pCfv_Fields, pCfv_Fields.Length + 1)
                pCfv_Fields(pCfv_Fields.Length - 1) = oCELLFIXVALUES.Metadata.LineType(i).Name
            Next
            ' Cso_Fields
            pCso_Fields = {}
            For i As Integer = 0 To oCELLSOURCES.Metadata.LineType.FieldCount - 1
                Array.Resize(pCso_Fields, pCso_Fields.Length + 1)
                pCso_Fields(pCso_Fields.Length - 1) = oCELLSOURCES.Metadata.LineType(i).Name
            Next
            ' Ffs_Fields
            pFfs_Fields = {}
            For i As Integer = 0 To oFLEXFUNCSOURCES.Metadata.LineType.FieldCount - 1
                Array.Resize(pFfs_Fields, pFfs_Fields.Length + 1)
                pFfs_Fields(pFfs_Fields.Length - 1) = oFLEXFUNCSOURCES.Metadata.LineType(i).Name
            Next
        Catch Ex As System.Exception
            log.Error("getMeta_GetDetail - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPTemplateCO")
        Finally
            log.Debug("getMeta_GetDetail - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Sub getMeta_CreateFromData(ByRef pPos_Fields() As String, ByRef pCfv_Fields() As String, ByRef pCso_Fields() As String, ByRef pFfs_Fields() As String)
        Try
            log.Debug("getMeta_CreateFromData - " & "creating Function BAPI_TEMPLATECO_CREATE")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_TEMPLATECO_CREATE")
            Dim oPOSITIONS As IRfcTable = oRfcFunction.GetTable("POSITIONS")
            Dim oCELLFIXVALUES As IRfcTable = oRfcFunction.GetTable("CELLFIXVALUES")
            Dim oCELLSOURCES As IRfcTable = oRfcFunction.GetTable("CELLSOURCES")
            Dim oFLEXFUNCSOURCES As IRfcTable = oRfcFunction.GetTable("FLEXFUNCSOURCES")
            ' Pos_Fields
            pPos_Fields = {}
            For i As Integer = 0 To oPOSITIONS.Metadata.LineType.FieldCount - 1
                Array.Resize(pPos_Fields, pPos_Fields.Length + 1)
                pPos_Fields(pPos_Fields.Length - 1) = oPOSITIONS.Metadata.LineType(i).Name
            Next
            ' Cfv_Fields
            pCfv_Fields = {}
            For i As Integer = 0 To oCELLFIXVALUES.Metadata.LineType.FieldCount - 1
                Array.Resize(pCfv_Fields, pCfv_Fields.Length + 1)
                pCfv_Fields(pCfv_Fields.Length - 1) = oCELLFIXVALUES.Metadata.LineType(i).Name
            Next
            ' Cso_Fields
            pCso_Fields = {}
            For i As Integer = 0 To oCELLSOURCES.Metadata.LineType.FieldCount - 1
                Array.Resize(pCso_Fields, pCso_Fields.Length + 1)
                pCso_Fields(pCso_Fields.Length - 1) = oCELLSOURCES.Metadata.LineType(i).Name
            Next
            ' Ffs_Fields
            pFfs_Fields = {}
            For i As Integer = 0 To oFLEXFUNCSOURCES.Metadata.LineType.FieldCount - 1
                Array.Resize(pFfs_Fields, pFfs_Fields.Length + 1)
                pFfs_Fields(pFfs_Fields.Length - 1) = oFLEXFUNCSOURCES.Metadata.LineType(i).Name
            Next
        Catch Ex As System.Exception
            log.Error("getMeta_GetDetail - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPTemplateCO")
        Finally
            log.Debug("getMeta_GetDetail - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Function Delete(pData As TSAP_TLData, Optional pOKMsg As String = "OK", Optional pCheck As Boolean = False) As String
        Delete = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_TEMPLATECO_DELETE")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oRETURN.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            If pCheck Then
                oRfcFunction.SetValue("TESTRUN", "X")
            Else
                oRfcFunction.SetValue("TESTRUN", "")
            End If
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                If oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    Delete = Delete & ";" & oRETURN(i).GetValue("MESSAGE")
                    If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "W" Then
                        aErr = True
                    End If
                End If
            Next i
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            Delete = If(Delete = "", pOKMsg, If(aErr = False, pOKMsg & Delete, "Error" & Delete))

        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPTemplateCO")
            Delete = "Error: Exception in Delete"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function GetDetail(ByRef pData As TSAP_TLData, Optional pOKMsg As String = "OK") As String
        GetDetail = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_TEMPLATECO_GET_DETAIL")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            Dim oPOSITIONS As IRfcTable = oRfcFunction.GetTable("POSITIONS")
            Dim oCELLFIXVALUES As IRfcTable = oRfcFunction.GetTable("CELLFIXVALUES")
            Dim oCELLSOURCES As IRfcTable = oRfcFunction.GetTable("CELLSOURCES")
            Dim oFLEXFUNCSOURCES As IRfcTable = oRfcFunction.GetTable("FLEXFUNCSOURCES")
            oRETURN.Clear()
            oPOSITIONS.Clear()
            oCELLFIXVALUES.Clear()
            oCELLSOURCES.Clear()
            oFLEXFUNCSOURCES.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next

            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                If oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    GetDetail = GetDetail & ";" & oRETURN(i).GetValue("MESSAGE")
                    If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "W" Then
                        aErr = True
                    End If
                End If
            Next i
            GetDetail = If(GetDetail = "", pOKMsg, If(aErr = False, pOKMsg & GetDetail, "Error" & GetDetail))

            If aErr = False Then
                ' return the header data
                pData.aHdrRec.setValues("-TEMPLATETEXT", oRfcFunction.GetValue("TEMPLATETEXT"), "", "", pEmptyChar:="")
                ' return the positions
                pData.aDataDic.addValues(oTable:=oPOSITIONS, pStrucName:="POSITIONS")
                ' return the cell fix values
                pData.aDataDic.addValues(oTable:=oCELLFIXVALUES, pStrucName:="CELLFIXVALUES")
                ' return the cell fix values
                pData.aDataDic.addValues(oTable:=oCELLSOURCES, pStrucName:="CELLSOURCES")
                ' return the cell fix values
                pData.aDataDic.addValues(oTable:=oFLEXFUNCSOURCES, pStrucName:="FLEXFUNCSOURCES")
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPTemplateCO")
            GetDetail = "Error: Exception in GetDetail"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function CreateFromData(pData As TSAP_TLData, Optional pOKMsg As String = "OK", Optional pCheck As Boolean = False) As String
        CreateFromData = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_TEMPLATECO_CREATE")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            Dim oPOSITIONS As IRfcTable = oRfcFunction.GetTable("POSITIONS")
            Dim oCELLFIXVALUES As IRfcTable = oRfcFunction.GetTable("CELLFIXVALUES")
            Dim oCELLSOURCES As IRfcTable = oRfcFunction.GetTable("CELLSOURCES")
            Dim oFLEXFUNCSOURCES As IRfcTable = oRfcFunction.GetTable("FLEXFUNCSOURCES")
            oRETURN.Clear()
            oPOSITIONS.Clear()
            oCELLFIXVALUES.Clear()
            oCELLSOURCES.Clear()
            oFLEXFUNCSOURCES.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            If pCheck Then
                oRfcFunction.SetValue("TESTRUN", "X")
            Else
                oRfcFunction.SetValue("TESTRUN", "")
            End If
            ' set the data values
            pData.aDataDic.to_IRfcTable(pKey:="POSITIONS", pIRfcTable:=oPOSITIONS)
            pData.aDataDic.to_IRfcTable(pKey:="CELLFIXVALUES", pIRfcTable:=oCELLFIXVALUES)
            pData.aDataDic.to_IRfcTable(pKey:="CELLSOURCES", pIRfcTable:=oCELLSOURCES)
            pData.aDataDic.to_IRfcTable(pKey:="FLEXFUNCSOURCES", pIRfcTable:=oFLEXFUNCSOURCES)
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                If oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    CreateFromData = CreateFromData & ";" & oRETURN(i).GetValue("MESSAGE")
                    If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "W" Then
                        aErr = True
                    End If
                End If
            Next i
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            CreateFromData = If(CreateFromData = "", pOKMsg, If(aErr = False, pOKMsg & CreateFromData, "Error" & CreateFromData))

        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPTemplateCO")
            CreateFromData = "Error: Exception in CreateFromData"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function Generate(pData As TSAP_TLData, Optional pOKMsg As String = "OK") As String
        Generate = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("ZCO_ABC_COTPL_READ_GENERATE")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oRETURN.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the table fields
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
                Generate = Generate & ";" & oRETURN(i).GetValue("MESSAGE")
            Next i
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            Generate = If(Generate = "", pOKMsg, If(aErr = False, pOKMsg & Generate, "Error" & Generate))

        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPTemplateCO")
            Generate = "Error: Exception in Generate"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class
