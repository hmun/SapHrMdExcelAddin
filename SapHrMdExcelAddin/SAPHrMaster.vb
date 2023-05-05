' Copyright 2022 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPHrMaster

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
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPHrMaster")
        End Try
    End Sub

    Private Sub addToStrucDic(pArrayName As String, pRfcStructureMetadata As RfcStructureMetadata, ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        If pStrucDic.ContainsKey(pArrayName) Then
            pStrucDic.Remove(pArrayName)
            pStrucDic.Add(pArrayName, pRfcStructureMetadata)
        Else
            pStrucDic.Add(pArrayName, pRfcStructureMetadata)
        End If
    End Sub

    Private Sub addToFieldDic(pArrayName As String, pRfcStructureMetadata As RfcParameterMetadata, ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata))
        If pFieldDic.ContainsKey(pArrayName) Then
            pFieldDic.Remove(pArrayName)
            pFieldDic.Add(pArrayName, pRfcStructureMetadata)
        Else
            pFieldDic.Add(pArrayName, pRfcStructureMetadata)
        End If
    End Sub

    Public Sub getMeta_SaveReplMulti(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {}
        Dim aImports As String() = {}
        Dim aTables As String() = {"HR_OBJECT_HEADER", "HR_INFOTYPE_HEADER", "EMPLOYEE_PERSONNEL_ACTION", "EMPLOYEE_ORG_ASSIGNMENT", "EMPLOYEE_PERSONAL_DATA", "EMPLOYEE_PAYROLL_STATUS", "EMPLOYEE_PRIVATE_ADDRESS", "EMPLOYEE_BANK_DETAILS", "EMPLOYEE_TRIP_PRIVILEGES", "EMPLOYEE_TIME_RECORDING_INFO", "EMPLOYEE_COMMUNICATIONS", "EMPLOYEE_COMPANY_CARS", "EMPLOYEE_ARCHIVED_OBJECTS", "EMPLOYEE_CATS_DEFAULT_VALUES", "EMPLOYEE_SALES_ORG_ASSIGNMENT", "PD_OBJECT", "PD_OBJECT_RELATIONSHIPS", "PD_REL_IS_EQUIPPED", "PD_REL_REQUIRES", "PD_REL_RESERVES", "PD_REL_TAKES_PART_IN", "PD_REL_IS_PREBOOKED_FOR", "PD_REL_QUALIFICATIONS_REQMNTS", "PD_REL_CAREER_DEVELOP_ACTION", "PD_REL_APPRAISAL_MODEL", "PD_REL_APPLICATION_FOR", "PD_REL_APPRAISES", "PD_REL_SESSION", "PD_REL_STILL_REQUIRES", "PD_REL_JOB_REQMNTS", "PD_REL_SPECIFIED_SUBSTITUTE", "PD_REL_MESSAGE_TYPE", "PD_OBJECT_DESCRIPTION", "PD_OBJECT_DEPARTMENT_STAFF", "PD_OBJECT_TASK_CHARACTER", "PD_OBJECT_PLANNED_COMPENSATION", "PD_OBJECT_RESTRICTIONS", "PD_OBJECT_VACANCY", "PD_OBJECT_POSTING_SPECS", "PD_OBJECT_HEALTH_EXAMINATIONS", "PD_OBJECT_AUTHORITY_RESOURCES", "PD_OBJECT_WORK_SCHEDULE", "PD_OBJECT_EMPL_GROUP_SUBGROUP", "PD_OBJECT_OBSOLETE", "PD_OBJECT_COST_PLANNING", "PD_OBJECT_STANDARD_PROFILES", "PD_OBJECT_PD_PROFILES", "PD_OBJECT_COST_DISTRIBUTION", "PD_OBJECT_HEAD_COUNT_BUDGET", "PD_OBJECT_B_EVENT_PRICE", "PD_OBJECT_B_EVENT_AVAILABILITY", "PD_OBJECT_B_EVENT_CAPACITY", "PD_OBJECT_VALIDITY", "PD_OBJECT_B_EVENT_INFO", "PD_OBJECT_SITE_DEPENDENT_INFO", "PD_OBJECT_ADDRESS", "PD_OBJECT_B_EVENT_TYPE", "PD_OBJECT_B_EVENT_PROCEDURE", "PD_OBJECT_ROOM_RESERV_INFO", "PD_OBJECT_MAIL_ADDRESS", "PD_OBJECT_WORK_REQMNT_SCALE", "PD_OBJECT_B_EVENT_NAME_FORMAT", "PD_OBJECT_B_EVENT_SCHEDULE", "PD_OBJECT_B_EVENT_COSTS", "PD_OBJECT_COST_ALLOC_INVOICE", "PD_OBJECT_EXTERNAL_KEY", "PD_OBJECT_SHIFT_GROUP", "PD_OBJECT_OVERRIDE_REQMNT", "PD_OBJECT_B_EVENT_BLOCKS", "PD_OBJECT_B_EVENT_SCHED_MODEL", "PD_OBJECT_APPRAISAL_SAMPLE", "PD_OBJECT_APPRAISAL_SCALE", "PD_OBJECT_VALUATION", "PD_OBJECT_PROCESSING_FM", "PD_OBJECT_PROFICIENCY_DESC", "PD_OBJECT_WORK_EVALUAT_RESULT", "PD_OBJECT_JOB_MARKET_INFO", "PD_OBJECT_B_EVENT_DEMAND", "PD_OBJECT_B_EVENTKNOWLEDGELINK", "RETURN"}
        Try
            log.Debug("getMeta_SaveReplMulti - " & "creating Function BAPI_HRMASTER_SAVE_REPL_MULT")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_HRMASTER_SAVE_REPL_MULT")
            Dim oStructure As IRfcStructure
            Dim oTable As IRfcTable
            ' Imports
            For s As Integer = 0 To aImports.Length - 1
                addToFieldDic("I|" & aImports(s), oRfcFunction.Metadata.Item(aImports(s)), pFieldDic)
            Next
            ' Import Strcutures
            For s As Integer = 0 To aStructures.Length - 1
                oStructure = oRfcFunction.GetStructure(aStructures(s))
                addToStrucDic("S|" & aStructures(s), oStructure.Metadata, pStrucDic)
            Next
            For s As Integer = 0 To aTables.Length - 1
                oTable = oRfcFunction.GetTable(aTables(s))
                addToStrucDic("T|" & aTables(s), oTable.Metadata.LineType, pStrucDic)
            Next
        Catch Ex As System.Exception
            log.Error("getMeta_SaveReplMulti - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPHrMaster")
        Finally
            log.Debug("getMeta_SaveReplMulti - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Function SaveReplMulti(pData As TSAP_Data_HrM, Optional pOKMsg As String = "OK", Optional pCheck As Boolean = False) As String
        SaveReplMulti = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_HRMASTER_SAVE_REPL_MULT")
            RfcSessionManager.BeginContext(destination)
            Dim oHR_OBJECT_HEADER As IRfcTable = oRfcFunction.GetTable("HR_OBJECT_HEADER")
            Dim oHR_INFOTYPE_HEADER As IRfcTable = oRfcFunction.GetTable("HR_INFOTYPE_HEADER")
            Dim oEMPLOYEE_PERSONNEL_ACTION As IRfcTable = oRfcFunction.GetTable("EMPLOYEE_PERSONNEL_ACTION")
            Dim oEMPLOYEE_ORG_ASSIGNMENT As IRfcTable = oRfcFunction.GetTable("EMPLOYEE_ORG_ASSIGNMENT")
            Dim oEMPLOYEE_PERSONAL_DATA As IRfcTable = oRfcFunction.GetTable("EMPLOYEE_PERSONAL_DATA")
            Dim oEMPLOYEE_PAYROLL_STATUS As IRfcTable = oRfcFunction.GetTable("EMPLOYEE_PAYROLL_STATUS")
            Dim oEMPLOYEE_PRIVATE_ADDRESS As IRfcTable = oRfcFunction.GetTable("EMPLOYEE_PRIVATE_ADDRESS")
            Dim oEMPLOYEE_BANK_DETAILS As IRfcTable = oRfcFunction.GetTable("EMPLOYEE_BANK_DETAILS")
            Dim oEMPLOYEE_TRIP_PRIVILEGES As IRfcTable = oRfcFunction.GetTable("EMPLOYEE_TRIP_PRIVILEGES")
            Dim oEMPLOYEE_TIME_RECORDING_INFO As IRfcTable = oRfcFunction.GetTable("EMPLOYEE_TIME_RECORDING_INFO")
            Dim oEMPLOYEE_COMMUNICATIONS As IRfcTable = oRfcFunction.GetTable("EMPLOYEE_COMMUNICATIONS")
            Dim oEMPLOYEE_COMPANY_CARS As IRfcTable = oRfcFunction.GetTable("EMPLOYEE_COMPANY_CARS")
            Dim oEMPLOYEE_ARCHIVED_OBJECTS As IRfcTable = oRfcFunction.GetTable("EMPLOYEE_ARCHIVED_OBJECTS")
            Dim oEMPLOYEE_CATS_DEFAULT_VALUES As IRfcTable = oRfcFunction.GetTable("EMPLOYEE_CATS_DEFAULT_VALUES")
            Dim oEMPLOYEE_SALES_ORG_ASSIGNMENT As IRfcTable = oRfcFunction.GetTable("EMPLOYEE_SALES_ORG_ASSIGNMENT")
            Dim oPD_OBJECT As IRfcTable = oRfcFunction.GetTable("PD_OBJECT")
            Dim oPD_OBJECT_RELATIONSHIPS As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_RELATIONSHIPS")
            Dim oPD_REL_IS_EQUIPPED As IRfcTable = oRfcFunction.GetTable("PD_REL_IS_EQUIPPED")
            Dim oPD_REL_REQUIRES As IRfcTable = oRfcFunction.GetTable("PD_REL_REQUIRES")
            Dim oPD_REL_RESERVES As IRfcTable = oRfcFunction.GetTable("PD_REL_RESERVES")
            Dim oPD_REL_TAKES_PART_IN As IRfcTable = oRfcFunction.GetTable("PD_REL_TAKES_PART_IN")
            Dim oPD_REL_IS_PREBOOKED_FOR As IRfcTable = oRfcFunction.GetTable("PD_REL_IS_PREBOOKED_FOR")
            Dim oPD_REL_QUALIFICATIONS_REQMNTS As IRfcTable = oRfcFunction.GetTable("PD_REL_QUALIFICATIONS_REQMNTS")
            Dim oPD_REL_CAREER_DEVELOP_ACTION As IRfcTable = oRfcFunction.GetTable("PD_REL_CAREER_DEVELOP_ACTION")
            Dim oPD_REL_APPRAISAL_MODEL As IRfcTable = oRfcFunction.GetTable("PD_REL_APPRAISAL_MODEL")
            Dim oPD_REL_APPLICATION_FOR As IRfcTable = oRfcFunction.GetTable("PD_REL_APPLICATION_FOR")
            Dim oPD_REL_APPRAISES As IRfcTable = oRfcFunction.GetTable("PD_REL_APPRAISES")
            Dim oPD_REL_SESSION As IRfcTable = oRfcFunction.GetTable("PD_REL_SESSION")
            Dim oPD_REL_STILL_REQUIRES As IRfcTable = oRfcFunction.GetTable("PD_REL_STILL_REQUIRES")
            Dim oPD_REL_JOB_REQMNTS As IRfcTable = oRfcFunction.GetTable("PD_REL_JOB_REQMNTS")
            Dim oPD_REL_SPECIFIED_SUBSTITUTE As IRfcTable = oRfcFunction.GetTable("PD_REL_SPECIFIED_SUBSTITUTE")
            Dim oPD_REL_MESSAGE_TYPE As IRfcTable = oRfcFunction.GetTable("PD_REL_MESSAGE_TYPE")
            Dim oPD_OBJECT_DESCRIPTION As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_DESCRIPTION")
            Dim oPD_OBJECT_DEPARTMENT_STAFF As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_DEPARTMENT_STAFF")
            Dim oPD_OBJECT_TASK_CHARACTER As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_TASK_CHARACTER")
            Dim oPD_OBJECT_PLANNED_COMPENSATION As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_PLANNED_COMPENSATION")
            Dim oPD_OBJECT_RESTRICTIONS As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_RESTRICTIONS")
            Dim oPD_OBJECT_VACANCY As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_VACANCY")
            Dim oPD_OBJECT_POSTING_SPECS As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_POSTING_SPECS")
            Dim oPD_OBJECT_HEALTH_EXAMINATIONS As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_HEALTH_EXAMINATIONS")
            Dim oPD_OBJECT_AUTHORITY_RESOURCES As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_AUTHORITY_RESOURCES")
            Dim oPD_OBJECT_WORK_SCHEDULE As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_WORK_SCHEDULE")
            Dim oPD_OBJECT_EMPL_GROUP_SUBGROUP As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_EMPL_GROUP_SUBGROUP")
            Dim oPD_OBJECT_OBSOLETE As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_OBSOLETE")
            Dim oPD_OBJECT_COST_PLANNING As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_COST_PLANNING")
            Dim oPD_OBJECT_STANDARD_PROFILES As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_STANDARD_PROFILES")
            Dim oPD_OBJECT_PD_PROFILES As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_PD_PROFILES")
            Dim oPD_OBJECT_COST_DISTRIBUTION As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_COST_DISTRIBUTION")
            Dim oPD_OBJECT_HEAD_COUNT_BUDGET As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_HEAD_COUNT_BUDGET")
            Dim oPD_OBJECT_B_EVENT_PRICE As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_B_EVENT_PRICE")
            Dim oPD_OBJECT_B_EVENT_AVAILABILITY As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_B_EVENT_AVAILABILITY")
            Dim oPD_OBJECT_B_EVENT_CAPACITY As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_B_EVENT_CAPACITY")
            Dim oPD_OBJECT_VALIDITY As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_VALIDITY")
            Dim oPD_OBJECT_B_EVENT_INFO As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_B_EVENT_INFO")
            Dim oPD_OBJECT_SITE_DEPENDENT_INFO As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_SITE_DEPENDENT_INFO")
            Dim oPD_OBJECT_ADDRESS As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_ADDRESS")
            Dim oPD_OBJECT_B_EVENT_TYPE As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_B_EVENT_TYPE")
            Dim oPD_OBJECT_B_EVENT_PROCEDURE As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_B_EVENT_PROCEDURE")
            Dim oPD_OBJECT_ROOM_RESERV_INFO As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_ROOM_RESERV_INFO")
            Dim oPD_OBJECT_MAIL_ADDRESS As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_MAIL_ADDRESS")
            Dim oPD_OBJECT_WORK_REQMNT_SCALE As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_WORK_REQMNT_SCALE")
            Dim oPD_OBJECT_B_EVENT_NAME_FORMAT As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_B_EVENT_NAME_FORMAT")
            Dim oPD_OBJECT_B_EVENT_SCHEDULE As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_B_EVENT_SCHEDULE")
            Dim oPD_OBJECT_B_EVENT_COSTS As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_B_EVENT_COSTS")
            Dim oPD_OBJECT_COST_ALLOC_INVOICE As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_COST_ALLOC_INVOICE")
            Dim oPD_OBJECT_EXTERNAL_KEY As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_EXTERNAL_KEY")
            Dim oPD_OBJECT_SHIFT_GROUP As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_SHIFT_GROUP")
            Dim oPD_OBJECT_OVERRIDE_REQMNT As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_OVERRIDE_REQMNT")
            Dim oPD_OBJECT_B_EVENT_BLOCKS As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_B_EVENT_BLOCKS")
            Dim oPD_OBJECT_B_EVENT_SCHED_MODEL As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_B_EVENT_SCHED_MODEL")
            Dim oPD_OBJECT_APPRAISAL_SAMPLE As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_APPRAISAL_SAMPLE")
            Dim oPD_OBJECT_APPRAISAL_SCALE As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_APPRAISAL_SCALE")
            Dim oPD_OBJECT_VALUATION As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_VALUATION")
            Dim oPD_OBJECT_PROCESSING_FM As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_PROCESSING_FM")
            Dim oPD_OBJECT_PROFICIENCY_DESC As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_PROFICIENCY_DESC")
            Dim oPD_OBJECT_WORK_EVALUAT_RESULT As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_WORK_EVALUAT_RESULT")
            Dim oPD_OBJECT_JOB_MARKET_INFO As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_JOB_MARKET_INFO")
            Dim oPD_OBJECT_B_EVENT_DEMAND As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_B_EVENT_DEMAND")
            Dim oPD_OBJECT_B_EVENTKNOWLEDGELINK As IRfcTable = oRfcFunction.GetTable("PD_OBJECT_B_EVENTKNOWLEDGELINK")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oHR_OBJECT_HEADER.Clear()
            oHR_INFOTYPE_HEADER.Clear()
            oEMPLOYEE_PERSONNEL_ACTION.Clear()
            oEMPLOYEE_ORG_ASSIGNMENT.Clear()
            oEMPLOYEE_PERSONAL_DATA.Clear()
            oEMPLOYEE_PAYROLL_STATUS.Clear()
            oEMPLOYEE_PRIVATE_ADDRESS.Clear()
            oEMPLOYEE_BANK_DETAILS.Clear()
            oEMPLOYEE_TRIP_PRIVILEGES.Clear()
            oEMPLOYEE_TIME_RECORDING_INFO.Clear()
            oEMPLOYEE_COMMUNICATIONS.Clear()
            oEMPLOYEE_COMPANY_CARS.Clear()
            oEMPLOYEE_ARCHIVED_OBJECTS.Clear()
            oEMPLOYEE_CATS_DEFAULT_VALUES.Clear()
            oEMPLOYEE_SALES_ORG_ASSIGNMENT.Clear()
            oPD_OBJECT.Clear()
            oPD_OBJECT_RELATIONSHIPS.Clear()
            oPD_REL_IS_EQUIPPED.Clear()
            oPD_REL_REQUIRES.Clear()
            oPD_REL_RESERVES.Clear()
            oPD_REL_TAKES_PART_IN.Clear()
            oPD_REL_IS_PREBOOKED_FOR.Clear()
            oPD_REL_QUALIFICATIONS_REQMNTS.Clear()
            oPD_REL_CAREER_DEVELOP_ACTION.Clear()
            oPD_REL_APPRAISAL_MODEL.Clear()
            oPD_REL_APPLICATION_FOR.Clear()
            oPD_REL_APPRAISES.Clear()
            oPD_REL_SESSION.Clear()
            oPD_REL_STILL_REQUIRES.Clear()
            oPD_REL_JOB_REQMNTS.Clear()
            oPD_REL_SPECIFIED_SUBSTITUTE.Clear()
            oPD_REL_MESSAGE_TYPE.Clear()
            oPD_OBJECT_DESCRIPTION.Clear()
            oPD_OBJECT_DEPARTMENT_STAFF.Clear()
            oPD_OBJECT_TASK_CHARACTER.Clear()
            oPD_OBJECT_PLANNED_COMPENSATION.Clear()
            oPD_OBJECT_RESTRICTIONS.Clear()
            oPD_OBJECT_VACANCY.Clear()
            oPD_OBJECT_POSTING_SPECS.Clear()
            oPD_OBJECT_HEALTH_EXAMINATIONS.Clear()
            oPD_OBJECT_AUTHORITY_RESOURCES.Clear()
            oPD_OBJECT_WORK_SCHEDULE.Clear()
            oPD_OBJECT_EMPL_GROUP_SUBGROUP.Clear()
            oPD_OBJECT_OBSOLETE.Clear()
            oPD_OBJECT_COST_PLANNING.Clear()
            oPD_OBJECT_STANDARD_PROFILES.Clear()
            oPD_OBJECT_PD_PROFILES.Clear()
            oPD_OBJECT_COST_DISTRIBUTION.Clear()
            oPD_OBJECT_HEAD_COUNT_BUDGET.Clear()
            oPD_OBJECT_B_EVENT_PRICE.Clear()
            oPD_OBJECT_B_EVENT_AVAILABILITY.Clear()
            oPD_OBJECT_B_EVENT_CAPACITY.Clear()
            oPD_OBJECT_VALIDITY.Clear()
            oPD_OBJECT_B_EVENT_INFO.Clear()
            oPD_OBJECT_SITE_DEPENDENT_INFO.Clear()
            oPD_OBJECT_ADDRESS.Clear()
            oPD_OBJECT_B_EVENT_TYPE.Clear()
            oPD_OBJECT_B_EVENT_PROCEDURE.Clear()
            oPD_OBJECT_ROOM_RESERV_INFO.Clear()
            oPD_OBJECT_MAIL_ADDRESS.Clear()
            oPD_OBJECT_WORK_REQMNT_SCALE.Clear()
            oPD_OBJECT_B_EVENT_NAME_FORMAT.Clear()
            oPD_OBJECT_B_EVENT_SCHEDULE.Clear()
            oPD_OBJECT_B_EVENT_COSTS.Clear()
            oPD_OBJECT_COST_ALLOC_INVOICE.Clear()
            oPD_OBJECT_EXTERNAL_KEY.Clear()
            oPD_OBJECT_SHIFT_GROUP.Clear()
            oPD_OBJECT_OVERRIDE_REQMNT.Clear()
            oPD_OBJECT_B_EVENT_BLOCKS.Clear()
            oPD_OBJECT_B_EVENT_SCHED_MODEL.Clear()
            oPD_OBJECT_APPRAISAL_SAMPLE.Clear()
            oPD_OBJECT_APPRAISAL_SCALE.Clear()
            oPD_OBJECT_VALUATION.Clear()
            oPD_OBJECT_PROCESSING_FM.Clear()
            oPD_OBJECT_PROFICIENCY_DESC.Clear()
            oPD_OBJECT_WORK_EVALUAT_RESULT.Clear()
            oPD_OBJECT_JOB_MARKET_INFO.Clear()
            oPD_OBJECT_B_EVENT_DEMAND.Clear()
            oPD_OBJECT_B_EVENTKNOWLEDGELINK.Clear()
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
            pData.aDataDic.to_IRfcTable(pKey:="HR_OBJECT_HEADER", pIRfcTable:=oHR_OBJECT_HEADER)
            pData.aDataDic.to_IRfcTable(pKey:="HR_INFOTYPE_HEADER", pIRfcTable:=oHR_INFOTYPE_HEADER)
            pData.aDataDic.to_IRfcTable(pKey:="EMPLOYEE_PERSONNEL_ACTION", pIRfcTable:=oEMPLOYEE_PERSONNEL_ACTION)
            pData.aDataDic.to_IRfcTable(pKey:="EMPLOYEE_ORG_ASSIGNMENT", pIRfcTable:=oEMPLOYEE_ORG_ASSIGNMENT)
            pData.aDataDic.to_IRfcTable(pKey:="EMPLOYEE_PERSONAL_DATA", pIRfcTable:=oEMPLOYEE_PERSONAL_DATA)
            pData.aDataDic.to_IRfcTable(pKey:="EMPLOYEE_PAYROLL_STATUS", pIRfcTable:=oEMPLOYEE_PAYROLL_STATUS)
            pData.aDataDic.to_IRfcTable(pKey:="EMPLOYEE_PRIVATE_ADDRESS", pIRfcTable:=oEMPLOYEE_PRIVATE_ADDRESS)
            pData.aDataDic.to_IRfcTable(pKey:="EMPLOYEE_BANK_DETAILS", pIRfcTable:=oEMPLOYEE_BANK_DETAILS)
            pData.aDataDic.to_IRfcTable(pKey:="EMPLOYEE_TRIP_PRIVILEGES", pIRfcTable:=oEMPLOYEE_TRIP_PRIVILEGES)
            pData.aDataDic.to_IRfcTable(pKey:="EMPLOYEE_TIME_RECORDING_INFO", pIRfcTable:=oEMPLOYEE_TIME_RECORDING_INFO)
            pData.aDataDic.to_IRfcTable(pKey:="EMPLOYEE_COMMUNICATIONS", pIRfcTable:=oEMPLOYEE_COMMUNICATIONS)
            pData.aDataDic.to_IRfcTable(pKey:="EMPLOYEE_COMPANY_CARS", pIRfcTable:=oEMPLOYEE_COMPANY_CARS)
            pData.aDataDic.to_IRfcTable(pKey:="EMPLOYEE_ARCHIVED_OBJECTS", pIRfcTable:=oEMPLOYEE_ARCHIVED_OBJECTS)
            pData.aDataDic.to_IRfcTable(pKey:="EMPLOYEE_CATS_DEFAULT_VALUES", pIRfcTable:=oEMPLOYEE_CATS_DEFAULT_VALUES)
            pData.aDataDic.to_IRfcTable(pKey:="EMPLOYEE_SALES_ORG_ASSIGNMENT", pIRfcTable:=oEMPLOYEE_SALES_ORG_ASSIGNMENT)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT", pIRfcTable:=oPD_OBJECT)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_RELATIONSHIPS", pIRfcTable:=oPD_OBJECT_RELATIONSHIPS)
            pData.aDataDic.to_IRfcTable(pKey:="PD_REL_IS_EQUIPPED", pIRfcTable:=oPD_REL_IS_EQUIPPED)
            pData.aDataDic.to_IRfcTable(pKey:="PD_REL_REQUIRES", pIRfcTable:=oPD_REL_REQUIRES)
            pData.aDataDic.to_IRfcTable(pKey:="PD_REL_RESERVES", pIRfcTable:=oPD_REL_RESERVES)
            pData.aDataDic.to_IRfcTable(pKey:="PD_REL_TAKES_PART_IN", pIRfcTable:=oPD_REL_TAKES_PART_IN)
            pData.aDataDic.to_IRfcTable(pKey:="PD_REL_IS_PREBOOKED_FOR", pIRfcTable:=oPD_REL_IS_PREBOOKED_FOR)
            pData.aDataDic.to_IRfcTable(pKey:="PD_REL_QUALIFICATIONS_REQMNTS", pIRfcTable:=oPD_REL_QUALIFICATIONS_REQMNTS)
            pData.aDataDic.to_IRfcTable(pKey:="PD_REL_CAREER_DEVELOP_ACTION", pIRfcTable:=oPD_REL_CAREER_DEVELOP_ACTION)
            pData.aDataDic.to_IRfcTable(pKey:="PD_REL_APPRAISAL_MODEL", pIRfcTable:=oPD_REL_APPRAISAL_MODEL)
            pData.aDataDic.to_IRfcTable(pKey:="PD_REL_APPLICATION_FOR", pIRfcTable:=oPD_REL_APPLICATION_FOR)
            pData.aDataDic.to_IRfcTable(pKey:="PD_REL_APPRAISES", pIRfcTable:=oPD_REL_APPRAISES)
            pData.aDataDic.to_IRfcTable(pKey:="PD_REL_SESSION", pIRfcTable:=oPD_REL_SESSION)
            pData.aDataDic.to_IRfcTable(pKey:="PD_REL_STILL_REQUIRES", pIRfcTable:=oPD_REL_STILL_REQUIRES)
            pData.aDataDic.to_IRfcTable(pKey:="PD_REL_JOB_REQMNTS", pIRfcTable:=oPD_REL_JOB_REQMNTS)
            pData.aDataDic.to_IRfcTable(pKey:="PD_REL_SPECIFIED_SUBSTITUTE", pIRfcTable:=oPD_REL_SPECIFIED_SUBSTITUTE)
            pData.aDataDic.to_IRfcTable(pKey:="PD_REL_MESSAGE_TYPE", pIRfcTable:=oPD_REL_MESSAGE_TYPE)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_DESCRIPTION", pIRfcTable:=oPD_OBJECT_DESCRIPTION)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_DEPARTMENT_STAFF", pIRfcTable:=oPD_OBJECT_DEPARTMENT_STAFF)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_TASK_CHARACTER", pIRfcTable:=oPD_OBJECT_TASK_CHARACTER)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_PLANNED_COMPENSATION", pIRfcTable:=oPD_OBJECT_PLANNED_COMPENSATION)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_RESTRICTIONS", pIRfcTable:=oPD_OBJECT_RESTRICTIONS)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_VACANCY", pIRfcTable:=oPD_OBJECT_VACANCY)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_POSTING_SPECS", pIRfcTable:=oPD_OBJECT_POSTING_SPECS)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_HEALTH_EXAMINATIONS", pIRfcTable:=oPD_OBJECT_HEALTH_EXAMINATIONS)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_AUTHORITY_RESOURCES", pIRfcTable:=oPD_OBJECT_AUTHORITY_RESOURCES)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_WORK_SCHEDULE", pIRfcTable:=oPD_OBJECT_WORK_SCHEDULE)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_EMPL_GROUP_SUBGROUP", pIRfcTable:=oPD_OBJECT_EMPL_GROUP_SUBGROUP)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_OBSOLETE", pIRfcTable:=oPD_OBJECT_OBSOLETE)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_COST_PLANNING", pIRfcTable:=oPD_OBJECT_COST_PLANNING)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_STANDARD_PROFILES", pIRfcTable:=oPD_OBJECT_STANDARD_PROFILES)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_PD_PROFILES", pIRfcTable:=oPD_OBJECT_PD_PROFILES)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_COST_DISTRIBUTION", pIRfcTable:=oPD_OBJECT_COST_DISTRIBUTION)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_HEAD_COUNT_BUDGET", pIRfcTable:=oPD_OBJECT_HEAD_COUNT_BUDGET)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_B_EVENT_PRICE", pIRfcTable:=oPD_OBJECT_B_EVENT_PRICE)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_B_EVENT_AVAILABILITY", pIRfcTable:=oPD_OBJECT_B_EVENT_AVAILABILITY)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_B_EVENT_CAPACITY", pIRfcTable:=oPD_OBJECT_B_EVENT_CAPACITY)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_VALIDITY", pIRfcTable:=oPD_OBJECT_VALIDITY)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_B_EVENT_INFO", pIRfcTable:=oPD_OBJECT_B_EVENT_INFO)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_SITE_DEPENDENT_INFO", pIRfcTable:=oPD_OBJECT_SITE_DEPENDENT_INFO)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_ADDRESS", pIRfcTable:=oPD_OBJECT_ADDRESS)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_B_EVENT_TYPE", pIRfcTable:=oPD_OBJECT_B_EVENT_TYPE)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_B_EVENT_PROCEDURE", pIRfcTable:=oPD_OBJECT_B_EVENT_PROCEDURE)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_ROOM_RESERV_INFO", pIRfcTable:=oPD_OBJECT_ROOM_RESERV_INFO)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_MAIL_ADDRESS", pIRfcTable:=oPD_OBJECT_MAIL_ADDRESS)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_WORK_REQMNT_SCALE", pIRfcTable:=oPD_OBJECT_WORK_REQMNT_SCALE)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_B_EVENT_NAME_FORMAT", pIRfcTable:=oPD_OBJECT_B_EVENT_NAME_FORMAT)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_B_EVENT_SCHEDULE", pIRfcTable:=oPD_OBJECT_B_EVENT_SCHEDULE)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_B_EVENT_COSTS", pIRfcTable:=oPD_OBJECT_B_EVENT_COSTS)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_COST_ALLOC_INVOICE", pIRfcTable:=oPD_OBJECT_COST_ALLOC_INVOICE)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_EXTERNAL_KEY", pIRfcTable:=oPD_OBJECT_EXTERNAL_KEY)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_SHIFT_GROUP", pIRfcTable:=oPD_OBJECT_SHIFT_GROUP)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_OVERRIDE_REQMNT", pIRfcTable:=oPD_OBJECT_OVERRIDE_REQMNT)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_B_EVENT_BLOCKS", pIRfcTable:=oPD_OBJECT_B_EVENT_BLOCKS)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_B_EVENT_SCHED_MODEL", pIRfcTable:=oPD_OBJECT_B_EVENT_SCHED_MODEL)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_APPRAISAL_SAMPLE", pIRfcTable:=oPD_OBJECT_APPRAISAL_SAMPLE)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_APPRAISAL_SCALE", pIRfcTable:=oPD_OBJECT_APPRAISAL_SCALE)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_VALUATION", pIRfcTable:=oPD_OBJECT_VALUATION)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_PROCESSING_FM", pIRfcTable:=oPD_OBJECT_PROCESSING_FM)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_PROFICIENCY_DESC", pIRfcTable:=oPD_OBJECT_PROFICIENCY_DESC)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_WORK_EVALUAT_RESULT", pIRfcTable:=oPD_OBJECT_WORK_EVALUAT_RESULT)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_JOB_MARKET_INFO", pIRfcTable:=oPD_OBJECT_JOB_MARKET_INFO)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_B_EVENT_DEMAND", pIRfcTable:=oPD_OBJECT_B_EVENT_DEMAND)
            pData.aDataDic.to_IRfcTable(pKey:="PD_OBJECT_B_EVENTKNOWLEDGELINK", pIRfcTable:=oPD_OBJECT_B_EVENTKNOWLEDGELINK)
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                If oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    SaveReplMulti = SaveReplMulti & ";" & oRETURN(i).GetValue("MESSAGE")
                    If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "W" Then
                        aErr = True
                    End If
                End If
            Next i
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            SaveReplMulti = If(SaveReplMulti = "", pOKMsg, If(aErr = False, pOKMsg & SaveReplMulti, "Error" & SaveReplMulti))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPMaterial")
            SaveReplMulti = "Error: Exception in SaveReplMulti"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class
