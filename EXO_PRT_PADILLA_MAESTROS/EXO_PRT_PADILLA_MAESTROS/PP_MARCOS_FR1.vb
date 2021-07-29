Imports System.Xml
Imports SAPbouiCOM
Public Class PP_MARCOS_FR1
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        cargamenu()
        If actualizar Then
            cargaCampos()
            'ParametrizacionGeneral()
        End If
    End Sub
#Region "Inicialización"
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults
    End Sub
    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_FILTROS.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
        Return filtro
    End Function
    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function
    Private Sub cargaCampos()
        Dim sXML As String = ""
        Dim res As String = ""

        'If objGlobal.refDi.comunes.esAdministrador Then
        '    'Campos de usuario en Factura de clientes
        '    sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_PP_MARCOS_FR1.xml")
        '    objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_PP_MARCOS_FR1", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        '    objGlobal.refDi.comunes.LoadBDFromXML(sXML)
        '    res = objGlobal.SBOApp.GetLastBatchResults
        '    'objGlobal.SBOApp.StatusBar.SetText(res, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        'End If
    End Sub
    'Private Sub ParametrizacionGeneral()
    '    If Not objGlobal.refDi.OGEN.existeVariable("EXO_PATH_EDI_FACTURAS") Then
    '        objGlobal.refDi.OGEN.fijarValorVariable("EXO_PATH_EDI_FACTURAS", "\\" & objGlobal.compañia.Server.Split(CChar(":"))(0) & "\B1_SHF\EDIFACT\" & objGlobal.compañia.CompanyDB)
    '    End If
    'End Sub
#End Region
#Region "Eventos"
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then

            Else

                Select Case infoEvento.MenuUID
                    Case "PP-MnRCMAR"
                        'Cargamos UDO
                        objGlobal.SBOApp.OpenForm(BoFormObjectEnum.fo_UserDefinedObject, "PP_MARCOS_FR1", "")
                End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)

            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_PP_MARCOS_FR1"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_PP_MARCOS_FR1"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                        End Select
                End Select
            End If

            Return MyBase.SBOApp_FormDataEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_PP_MARCOS_FR1"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_PP_MARCOS_FR1"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_PP_MARCOS_FR1"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_Form_Visible(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_PP_MARCOS_FR1"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                            End Select
                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        End Try
    End Function
    Private Function EventHandler_Choose_FromList_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

        EventHandler_Choose_FromList_After = False
        Dim sCod As String = "" : Dim sDes As String = ""
        Try
            Dim oDataTable As SAPbouiCOM.IChooseFromListEvent = CType(pVal, SAPbouiCOM.IChooseFromListEvent)
            If pVal.ItemUID = "0_U_E" Or pVal.ItemUID = "0_U_G" Then
                If oDataTable IsNot Nothing Then
                    Try
                        Select Case oForm.ChooseFromLists.Item(oDataTable.ChooseFromListUID).ObjectType
                            Case "PP_MARCOS"
                                Try
                                    sCod = oDataTable.SelectedObjects.GetValue("DocEntry", 0).ToString
                                    sDes = oDataTable.SelectedObjects.GetValue("U_PP_NAME", 0).ToString

                                    oForm.DataSources.DBDataSources.Item("@PP_MARCOS_FR1").SetValue("Name", 0, sDes)
                                Catch ex As Exception
                                    oForm.DataSources.DBDataSources.Item("@PP_MARCOS_FR1").SetValue("Code", 0, sCod)
                                    oForm.DataSources.DBDataSources.Item("@PP_MARCOS_FR1").SetValue("Name", 0, sDes)
                                End Try
                            Case "PP_LCONDICION"
                                Try
                                    sCod = oDataTable.SelectedObjects.GetValue("Code", 0).ToString
                                    sDes = oDataTable.SelectedObjects.GetValue("Name", 0).ToString

                                    oForm.DataSources.DBDataSources.Item("@PP_MARCOS_FR2").SetValue("U_PP_CONDNAME", oForm.DataSources.DBDataSources.Item("@PP_MARCOS_FR2").Offset, sDes)
                                Catch ex As Exception
                                    oForm.DataSources.DBDataSources.Item("@PP_MARCOS_FR2").SetValue("U_PP_CONDICION", oForm.DataSources.DBDataSources.Item("@PP_MARCOS_FR2").Offset, sCod)
                                    oForm.DataSources.DBDataSources.Item("@PP_MARCOS_FR2").SetValue("U_PP_CONDNAME", oForm.DataSources.DBDataSources.Item("@PP_MARCOS_FR2").Offset, sDes)
                                End Try
                            Case "PP_FRAGMENTOS"
                                Try
                                    sCod = oDataTable.SelectedObjects.GetValue("Code", 0).ToString
                                    sDes = oDataTable.SelectedObjects.GetValue("Name", 0).ToString

                                    oForm.DataSources.DBDataSources.Item("@PP_MARCOS_FR2").SetValue("U_PP_FRAGNAME", oForm.DataSources.DBDataSources.Item("@PP_MARCOS_FR2").Offset, sDes)
                                Catch ex As Exception
                                    oForm.DataSources.DBDataSources.Item("@PP_MARCOS_FR2").SetValue("U_PP_FRAGMENTO", oForm.DataSources.DBDataSources.Item("@PP_MARCOS_FR2").Offset, sCod)
                                    oForm.DataSources.DBDataSources.Item("@PP_MARCOS_FR2").SetValue("U_PP_FRAGNAME", oForm.DataSources.DBDataSources.Item("@PP_MARCOS_FR2").Offset, sDes)
                                End Try
                        End Select
                        If oForm.Mode = BoFormMode.fm_OK_MODE Then oForm.Mode = BoFormMode.fm_UPDATE_MODE
                    Catch ex As Exception
                        Throw ex
                    End Try
                End If
            End If

            EventHandler_Choose_FromList_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_Form_Visible(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oItem As SAPbouiCOM.Item = Nothing
        EventHandler_Form_Visible = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                Dim sObjType As String = oForm.TypeEx.Replace("UDO_FT_", "").Replace("_FR1", "")
                sSQL = "SELECT * FROM ""@PP_FRAGMENTOS"" WHERE ""U_PP_MAESTRO""='" & sObjType & "'"
                oRs.DoQuery(sSQL)
                oConds = New SAPbouiCOM.Conditions
                oCond = oConds.Add
                oCond.Alias = "Code"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "0"
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR

                For i = 0 To oRs.RecordCount - 1
                    oCond = oConds.Add
                    oCond.Alias = "Code"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = oRs.Fields.Item("Code").Value.ToString
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oRs.MoveNext()
                Next
                If oConds.Count > 0 Then oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE
                oForm.ChooseFromLists.Item("CFLFRG").SetConditions(oConds)

            End If

            EventHandler_Form_Visible = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oItem, Object))
        End Try
    End Function
#End Region
End Class
