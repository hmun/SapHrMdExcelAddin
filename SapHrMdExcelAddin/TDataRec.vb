' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class TDataRec

    Public aTDataRecCol As Collection
    Private aIntPar As SAPCommon.TStr

    Public Sub New(ByRef pIntPar As SAPCommon.TStr)
        aTDataRecCol = New Collection
        aIntPar = pIntPar
    End Sub

    Public Sub setValues(pNAME As String, pVALUE As String, pCURRENCY As String, pFORMAT As String,
                         Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set", Optional pUseAsEmpty As String = "#")
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aNameArray() As String
        Dim aKey As String
        Dim aSTRUCNAME As String = ""
        Dim aFIELDNAME As String = ""
        Dim aValue As String
        If pVALUE = pUseAsEmpty Then
            aValue = " "
        Else
            aValue = pVALUE
            If Not pEmty And aValue = pEmptyChar Then
                Exit Sub
            End If
        End If
        ' do not add empty values

        If InStr(pNAME, "-") <> 0 Then
            aNameArray = Split(pNAME, "-")
            aSTRUCNAME = aNameArray(0)
            aFIELDNAME = aNameArray(1)
        Else
            aSTRUCNAME = ""
            aFIELDNAME = pNAME
        End If
        aKey = pNAME
        If aTDataRecCol.Contains(aKey) Then
            aTStrRec = aTDataRecCol(aKey)
            Select Case pOperation
                Case "add"
                    aTStrRec.addValues(aSTRUCNAME, aFIELDNAME, aValue, pCURRENCY, pFORMAT)
                Case "sub"
                    aTStrRec.subValues(aSTRUCNAME, aFIELDNAME, aValue, pCURRENCY, pFORMAT)
                Case "mul"
                    aTStrRec.mulValues(aSTRUCNAME, aFIELDNAME, aValue, pCURRENCY, pFORMAT)
                Case "div"
                    aTStrRec.divValues(aSTRUCNAME, aFIELDNAME, aValue, pCURRENCY, pFORMAT)
                Case Else
                    aTStrRec.setValues(aSTRUCNAME, aFIELDNAME, aValue, pCURRENCY, pFORMAT)
            End Select
        Else
            aTStrRec = New SAPCommon.TStrRec
            aTStrRec.setValues(aSTRUCNAME, aFIELDNAME, aValue, pCURRENCY, pFORMAT)
            aTDataRecCol.Add(aTStrRec, aKey)
        End If
    End Sub

    Public Sub setValues(pTDataRec As TDataRec, Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set")
        Dim aTStrRec As SAPCommon.TStrRec
        For Each aTStrRec In pTDataRec.aTDataRecCol
            setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmty, pEmptyChar, pOperation)
        Next
    End Sub

    Public Sub addValues(pTDataRec As TDataRec, Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#")
        Dim aTStrRec As SAPCommon.TStrRec
        For Each aTStrRec In pTDataRec.aTDataRecCol
            If aTStrRec.Currency <> "" Then
                setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmty, pEmptyChar, pOperation:="add")
            Else
                setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmty, pEmptyChar, pOperation:="set")
            End If
        Next
    End Sub

    Public Function getColumn(pClmn As String) As SAPCommon.TStrRec
        Dim aTStrRec As SAPCommon.TStrRec
        If aTDataRecCol.Contains(pClmn) Then
            aTStrRec = aTDataRecCol(pClmn)
            getColumn = aTStrRec
        End If
    End Function

    Public Function getMaterial() As String
        Dim aTlClmn As String = If(aIntPar.value("COL", "MATERIAL") <> "", aIntPar.value("COL", "MATERIAL"), "MATERIAL")
        Dim aTStrRec As SAPCommon.TStrRec
        getMaterial = ""
        If aTDataRecCol.Contains(aTlClmn) Then
            aTStrRec = aTDataRecCol(aTlClmn)
            getMaterial = aTStrRec.Value
        End If
    End Function

    Public Function getPost(ByRef pPar As SAPCommon.TStr) As String
        Dim aClmn As String = If(pPar.value("COL", "DATAPOST") <> "", pPar.value("COL", "DATAPOST"), "INT-POST")
        Dim aTStrRec As SAPCommon.TStrRec
        getPost = ""
        If aTDataRecCol.Contains(aClmn) Then
            aTStrRec = aTDataRecCol(aClmn)
            getPost = aTStrRec.Value
        End If
    End Function

    Public Sub toRange(pFields() As String, pIsValue() As String, ByRef aRange As Excel.Range)
        Dim aTStrRec As SAPCommon.TStrRec
        For i = 0 To pFields.Count - 1
            If aTDataRecCol.Contains(pFields(i)) Then
                aTStrRec = aTDataRecCol(pFields(i))
                If pIsValue(i) = "X" Then
                    aRange(1, i + 1).Value = CDbl(aTStrRec.formated())
                Else
                    aRange(1, i + 1).Value = CStr(aTStrRec.formated())
                End If
            End If
        Next
    End Sub

End Class
