' Copyright 2022 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class SapHrMdRibbon_HrM
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Public Function getGenParameters(ByRef pPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        Dim aName As String
        Dim i As Integer
        log.Debug("SapHrMdRibbon_HrM getGenParametrs - " & "reading Parameter")
        aWB = Globals.SapHrMdExcelAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SapHrMd Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapHrMd")
            getGenParameters = False
            Exit Function
        End Try
        aName = "SAPHrMdMaster"
        aKey = CStr(aPws.Cells(1, 1).Value)
        If aKey <> aName Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key " & aName & ". Check if the current workbook is a valid SapHrMd Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapHrMd")
            getGenParameters = False
            Exit Function
        End If
        i = 2
        pPar = New SAPCommon.TStr
        Do While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
            pPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 4).value), pFORMAT:=CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop
        getGenParameters = True
    End Function

    Private Function getIntParameters(ByRef pIntPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim i As Integer

        log.Debug("getIntParameters - " & "reading Parameter")
        aWB = Globals.SapHrMdExcelAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter_Int")
        Catch Exc As System.Exception
            MsgBox("No Parameter_Int Sheet in current workbook. Check if the current workbook is a valid SapHrMd Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapHrMd")
            getIntParameters = False
            Exit Function
        End Try
        i = 2
        pIntPar = New SAPCommon.TStr
        Do
            pIntPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
        ' no obligatory parameters check - we should know what we are doing
        getIntParameters = True
    End Function

    Public Sub Save(ByRef pSapCon As SapCon)
        Dim aSAPHrMaster As New SAPHrMaster(pSapCon)

        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr

        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If

        Dim jMax As UInt64 = 0
        Dim aHrMLOff As Integer = If(aIntPar.value("LOFF", "HRM_DATA") <> "", CInt(aIntPar.value("LOFF", "HRM_DATA")), 4)
        Dim aHrMWsName As String = If(aIntPar.value("WS", "HRM_DATA") <> "", aIntPar.value("WS", "HRM_DATA"), "Data")
        Dim aHrMWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value("COL", "DATAMSG") <> "", aIntPar.value("COL", "DATAMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value("RET", "OKMSG") <> "", aIntPar.value("RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.SapHrMdExcelAddIn.Application.ActiveWorkbook
        Try
            aHrMWs = aWB.Worksheets(aHrMWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aHrMWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP HrM Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP HrM")
            Exit Sub
        End Try
        parseHeaderLine(aHrMWs, jMax, aMsgClmn, aMsgClmnNr, pHdrLine:=aHrMLOff - 3)
        Try
            log.Debug("SapHrMdRibbon_HrM.Save - " & "processing data - disabling events, screen update, cursor")
            Globals.SapHrMdExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapHrMdExcelAddIn.Application.EnableEvents = False
            '            Globals.SapHrMdExcelAddIn.Application.ScreenUpdating = False
            Dim i As UInt64 = aHrMLOff + 1
            Dim aKey As String
            Dim aHrMItems As New TData(aIntPar)
            Dim aTSAP_Data_HrM As New TSAP_Data_HrM(aPar, aIntPar, aSAPHrMaster, "SaveReplMulti")
            Do
                If Left(CStr(aHrMWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    aKey = CStr(i)
                    ' read DATA
                    aHrMItems.ws_parse_line_simple(aHrMWs, aHrMLOff, i, jMax)
                    If aTSAP_Data_HrM.fillHeader(aHrMItems) And aTSAP_Data_HrM.fillData(aHrMItems) Then
                        log.Debug("SapHrMdRibbon_HrM.Save - " & "calling aSAPHrMaster.SaveReplMulti")
                        Globals.SapHrMdExcelAddIn.Application.StatusBar = "Calling SAP-BAPI at line " & i
                        aRetStr = aSAPHrMaster.SaveReplMulti(aTSAP_Data_HrM, pOKMsg:=aOKMsg)
                        log.Debug("SapHrMdRibbon_HrM.Save - " & "aSAPHrMaster.SaveReplMulti returned, aRetStr=" & aRetStr)
                        For Each aKey In aHrMItems.aTDataDic.Keys
                            aHrMWs.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                        Next
                    End If
                    aHrMItems = New TData(aIntPar)
                    aTSAP_Data_HrM = New TSAP_Data_HrM(aPar, aIntPar, aSAPHrMaster, "SaveReplMulti")
                End If
                i += 1
            Loop While Not String.IsNullOrEmpty(CStr(aHrMWs.Cells(i, 1).value))
            log.Debug("SapHrMdRibbon_HrM.Save - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapHrMdExcelAddIn.Application.EnableEvents = True
            Globals.SapHrMdExcelAddIn.Application.ScreenUpdating = True
            Globals.SapHrMdExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapHrMdExcelAddIn.Application.EnableEvents = True
            Globals.SapHrMdExcelAddIn.Application.ScreenUpdating = True
            Globals.SapHrMdExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapHrMdRibbon_HrM.Save failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP HrM AddIn")
            log.Error("SapHrMdRibbon_HrM.Save - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try

    End Sub

    Private Sub parseHeaderLine(ByRef pWs As Excel.Worksheet, ByRef pMaxJ As Integer, Optional pMsgClmn As String = "", Optional ByRef pMsgClmnNr As Integer = 0, Optional pHdrLine As Integer = 1)
        pMaxJ = 0
        Do
            pMaxJ += 1
            If Not String.IsNullOrEmpty(pMsgClmn) And CStr(pWs.Cells(pHdrLine, pMaxJ).value) = pMsgClmn Then
                pMsgClmnNr = pMaxJ
            End If
        Loop While CStr(pWs.Cells(pHdrLine, pMaxJ + 1).value) <> ""
    End Sub

End Class
