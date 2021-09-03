Imports System.Reflection
Imports System.Xml
Imports SAPbouiCOM

Public Class PP_FRAGMENTOS
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

        If objGlobal.refDi.comunes.esAdministrador Then
            'Campos de usuario en Factura de clientes
            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_PP_FRAGMENTOS.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_PP_FRAGMENTOS", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults
            'objGlobal.SBOApp.StatusBar.SetText(res, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
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

                'Select Case infoEvento.MenuUID
                '    Case "PP-MnMFRG"
                '        'Cargamos UDO
                '        objGlobal.SBOApp.OpenForm(BoFormObjectEnum.fo_UserDefinedObject, "PP_FRAGMENTOS", "")
                'End Select
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
            oForm.Freeze(True)
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_PP_FRAGMENTOS"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_PP_FRAGMENTOS"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If oForm.Visible = True Then
                                    'If CType(oForm.Items.Item("13_U_C").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                                    '    Dim sSQL As String = ""
                                    '    Dim sTipo As String = CType(oForm.Items.Item("13_U_C").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                                    '    sSQL = "SELECT DISTINCT ""AbsEntry"", ""Code""  FROM ""ORST"" WHERE ""U_PP_TLIN""='O' and ""U_PP_TIPO""='" & sTipo & "' Order BY ""AbsEntry"" "
                                    '    objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").ValidValues, sSQL)
                                    'End If
                                    CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_5").Visible = False
                                    CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_6").Visible = False
                                    'For i = 1 To CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).RowCount
                                    '    If CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_5").Cells.Item(i).Specific, SAPbouiCOM.CheckBox).Checked = True Then
                                    '        CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).CommonSetting.SetCellEditable(i, 6, True)
                                    '    Else
                                    '        CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).CommonSetting.SetCellEditable(i, 6, False)
                                    '    End If
                                    'Next
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                        End Select
                End Select
            End If

            Return MyBase.SBOApp_FormDataEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            oForm.Freeze(False)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            oForm.Freeze(False)
            Return False
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_PP_FRAGMENTOS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    'If EventHandler_COMBO_SELECT_After(infoEvento) = False Then
                                    '    GC.Collect()
                                    '    Return False
                                    'End If
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        Return False
                                    End If
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
                        Case "UDO_FT_PP_FRAGMENTOS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_Before(infoEvento) = False Then
                                        Return False
                                    End If
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
                        Case "UDO_FT_PP_FRAGMENTOS"
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
                        Case "UDO_FT_PP_FRAGMENTOS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_Before(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
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
    'Private Function EventHandler_COMBO_SELECT_After(ByRef pVal As ItemEvent) As Boolean
    '    Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
    '    Dim sSQL As String = ""
    '    EventHandler_COMBO_SELECT_After = False
    '    Try

    '        If pVal.ItemUID = "0_U_G" And pVal.ColUID = "C_0_1" Then
    '            Dim iRegistros As Integer = oForm.DataSources.DBDataSources.Item("@PP_FRAGMENTOSL").Size
    '            Dim iRegActivo As Integer = oForm.DataSources.DBDataSources.Item("@PP_FRAGMENTOSL").Offset + 1
    '            If iRegistros = iRegActivo Then
    '                CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).AddRow()
    '                CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).ClearRowData(iRegActivo + 1)
    '                CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).FlushToDataSource()
    '            End If

    '        ElseIf pVal.ItemUID = "13_U_C" Then
    '            If CType(oForm.Items.Item("13_U_C").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
    '                Dim sTipo As String = CType(oForm.Items.Item("13_U_C").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
    '                sSQL = "SELECT DISTINCT ""AbsEntry"", ""Code""  FROM ""ORST"" WHERE ""U_PP_TLIN""='O' and ""U_PP_TIPO""='" & sTipo & "' Order BY ""AbsEntry"" "
    '                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").ValidValues, sSQL)
    '            End If
    '        End If

    '        EventHandler_COMBO_SELECT_After = True

    '    Catch exCOM As System.Runtime.InteropServices.COMException
    '        objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
    '    Catch ex As Exception
    '        objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
    '    Finally
    '        EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
    '    End Try
    'End Function
    Private Function EventHandler_Choose_FromList_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

        EventHandler_Choose_FromList_After = False
        Dim sCod As String = "" : Dim sDes As String = ""
        Try
            Dim oDataTable As SAPbouiCOM.IChooseFromListEvent = CType(pVal, SAPbouiCOM.IChooseFromListEvent)
            If pVal.ItemUID = "0_U_G" Then
                If oDataTable IsNot Nothing Then
                    Try
                        Select Case oForm.ChooseFromLists.Item(oDataTable.ChooseFromListUID).ObjectType
                            Case "PP_OPERACION"
                                With oForm.DataSources.DBDataSources.Item("@PP_FRAGMENTOSL")
                                    .SetValue("U_PP_CODOP", pVal.Row - 1, oDataTable.SelectedObjects.GetValue("DocEntry", 0).ToString)
                                    .SetValue("U_PP_DOPE", pVal.Row - 1, oDataTable.SelectedObjects.GetValue("U_PP_CODE", 0).ToString)
                                End With
                                CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).LoadFromDataSourceEx(False)
                                AddNewLine(oForm, True)
                                If oForm.Mode = BoFormMode.fm_OK_MODE Then oForm.Mode = BoFormMode.fm_UPDATE_MODE
                            Case "PP_CENTROS"
                                With oForm.DataSources.DBDataSources.Item("@PP_FRAGMENTOSL")
                                    .SetValue("U_PP_CENT", pVal.Row - 1, oDataTable.SelectedObjects.GetValue("DocEntry", 0).ToString)
                                    .SetValue("U_PP_DENT", pVal.Row - 1, oDataTable.SelectedObjects.GetValue("U_PP_NAME", 0).ToString)
                                End With
                                CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).LoadFromDataSourceEx(False)
                                If oForm.Mode = BoFormMode.fm_OK_MODE Then oForm.Mode = BoFormMode.fm_UPDATE_MODE
                            Case "4"
                                With oForm.DataSources.DBDataSources.Item("@PP_FRAGMENTOSL")
                                    .SetValue("U_PP_ITEM", pVal.Row - 1, oDataTable.SelectedObjects.GetValue("ItemCode", 0).ToString)
                                    .SetValue("U_PP_NITE", pVal.Row - 1, oDataTable.SelectedObjects.GetValue("ItemName", 0).ToString)
                                End With
                                CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).LoadFromDataSourceEx(False)
                                If oForm.Mode = BoFormMode.fm_OK_MODE Then oForm.Mode = BoFormMode.fm_UPDATE_MODE

                                'Case "PP_FRAGMENTOS"
                                '    Try
                                '        sCod = oDataTable.SelectedObjects.GetValue("ItemCode", 0).ToString
                                '        sDes = oDataTable.SelectedObjects.GetValue("ItemName", 0).ToString

                                '        'oForm.DataSources.DBDataSources.Item("@PP_FRAGMENTOSL").SetValue("U_PP_CODART", oForm.DataSources.DBDataSources.Item("@PP_FRAGMENTOSL").Offset, sCod)
                                '        'Try
                                '        '    oForm.DataSources.DBDataSources.Item("@PP_FRAGMENTOSL").SetValue("U_PP_DESART", oForm.DataSources.DBDataSources.Item("@PP_FRAGMENTOSL").Offset, sDes)
                                '        'Catch ex As Exception

                                '        'End Try

                                '        Try
                                '            oForm.DataSources.DBDataSources.Item("@PP_FRAGMENTOSL").SetValue("U_PP_CODART", pVal.Row - 1, sCod)
                                '            oForm.DataSources.DBDataSources.Item("@PP_FRAGMENTOSL").SetValue("U_PP_DESART", pVal.Row - 1, sDes)
                                '            CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).LoadFromDataSourceEx(False)
                                '            'CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_3").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = sDes
                                '        Catch ex As Exception

                                '        End Try

                                '    Catch ex As Exception
                                '        oForm.DataSources.DBDataSources.Item("@PP_FRAGMENTOSL").SetValue("U_PP_CODART", pVal.Row - 1, sCod)
                                '        oForm.DataSources.DBDataSources.Item("@PP_FRAGMENTOSL").SetValue("U_PP_DESART", pVal.Row - 1, sDes)
                                '        CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).LoadFromDataSourceEx(False)
                                '    End Try
                                'Case "4"
                                '    Try
                                '        sCod = oDataTable.SelectedObjects.GetValue("ItemCode", 0).ToString
                                '        sDes = oDataTable.SelectedObjects.GetValue("ItemName", 0).ToString

                                '        oForm.DataSources.DBDataSources.Item("@PP_FRAGMENTOSL").SetValue("U_PP_CODART", pVal.Row - 1, sCod)
                                '        oForm.DataSources.DBDataSources.Item("@PP_FRAGMENTOSL").SetValue("U_PP_DESART", pVal.Row - 1, sDes)
                                '        CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).LoadFromDataSourceEx(False)

                                '    Catch ex As Exception
                                '        oForm.DataSources.DBDataSources.Item("@PP_FRAGMENTOSL").SetValue("U_PP_CODART", pVal.Row - 1, sCod)
                                '        oForm.DataSources.DBDataSources.Item("@PP_FRAGMENTOSL").SetValue("U_PP_DESART", pVal.Row - 1, sDes)
                                '        CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).LoadFromDataSourceEx(False)
                                '    End Try
                        End Select
                        'If oForm.Mode = BoFormMode.fm_OK_MODE Then oForm.Mode = BoFormMode.fm_UPDATE_MODE
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
    Private Function EventHandler_Choose_FromList_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing
        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = Nothing
        EventHandler_Choose_FromList_Before = False
        Dim sCod As String = "" : Dim sDes As String = ""
        Try

            oCFLEvento = CType(pVal, SAPbouiCOM.IChooseFromListEvent)

            Select Case oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).ObjectType
                Case "PP_OPERACION"
                    oConds = New SAPbouiCOM.Conditions
                    oCond = oConds.Add
                    oCond.Alias = "U_PP_ACTIVO"
                    oCond.Operation = BoConditionOperation.co_EQUAL
                    oCond.CondVal = "Y"

                Case "PP_CENTROS"
                    Dim oRs As SAPbobsCOM.Recordset = Nothing
                    oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                    With oForm.DataSources.DBDataSources.Item("@PP_FRAGMENTOSL")
                        If .GetValue("U_PP_DOPE", pVal.Row - 1).ToString = "" Then
                            oConds = New SAPbouiCOM.Conditions
                            oCond = oConds.Add
                            oCond.Alias = "DocEntry"
                            oCond.Operation = BoConditionOperation.co_EQUAL
                            oCond.CondVal = "-1"
                        Else
                            oRs.DoQuery("SELECT T1.""U_PP_ID"" FROM ""@PP_OPERACION"" T0 INNER JOIN ""@PP_OPERACION_C"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                        WHERE T0.""DocEntry"" =" & .GetValue("U_PP_CODOP", pVal.Row - 1).ToString)
                            If oRs.RecordCount > 0 Then
                                oConds = New SAPbouiCOM.Conditions
                                While oRs.EoF = False
                                    If oConds.Count > 0 Then
                                        oCond.Relationship = BoConditionRelationship.cr_OR
                                    End If
                                    oCond = oConds.Add
                                    oCond.Alias = "DocEntry"
                                    oCond.Operation = BoConditionOperation.co_EQUAL
                                    oCond.CondVal = oRs.Fields.Item(0).Value.ToString
                                    oRs.MoveNext()
                                End While
                            Else
                                oConds = New SAPbouiCOM.Conditions
                                oCond = oConds.Add
                                oCond.Alias = "DocEntry"
                                oCond.Operation = BoConditionOperation.co_EQUAL
                                oCond.CondVal = "-1"
                            End If
                        End If
                    End With

                Case "4"
                    oConds = New SAPbouiCOM.Conditions
                    oCond = oConds.Add
                    oCond.Alias = "frozenFor"
                    oCond.Operation = BoConditionOperation.co_EQUAL
                    oCond.CondVal = "N"
                    'Resto
                    'Case Else
                    '    oConds = New SAPbouiCOM.Conditions
                    '    oCond = oConds.Add
                    '    oCond.Alias = "U_PP_ACTIVO"
                    '    oCond.Operation = BoConditionOperation.co_EQUAL
                    '    oCond.CondVal = "Y"
            End Select
            'CType(oForm.Items.Item(Replace(NOMBRE_MATRIX, "X", (oForm.PaneLevel - 1).ToString)).Specific, SAPbouiCOM.Matrix).FlushToDataSource()
            CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).FlushToDataSource()
            If Not oConds Is Nothing Then
                oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).SetConditions(oConds)
            Else
                oConds = New SAPbouiCOM.Conditions
                oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).SetConditions(oConds)
            End If
            EventHandler_Choose_FromList_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.FormCondition(oCond)
            EXO_CleanCOM.CLiberaCOM.FormConditions(oConds)
        End Try
    End Function
    Private Function EventHandler_Form_Visible(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim oItem As SAPbouiCOM.Item = Nothing
        EventHandler_Form_Visible = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                If oForm.Mode <> BoFormMode.fm_ADD_MODE Then
                    oForm.Items.Item("0_U_E").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
                Else
                    oForm.Items.Item("0_U_E").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_True)
                End If

                'If oForm.Mode <> BoFormMode.fm_ADD_MODE Then
                '    oForm.Items.Item("0_U_E").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
                'End If
                sSQL = "Select ""UF"".""FldValue"", ""UF"".""Descr"" From ""CUFD"" ""UT"" "
                sSQL &= " INNER JOIN ""UFD1"" ""UF"" ON ""UT"".""FieldID""=""UF"".""FieldID"" And ""UT"".""TableID""=""UF"".""TableID"" "
                sSQL &= " WHERE ""UT"".""TableID""='ORST' and ""UT"".""AliasID""='PP_TIPO'  Order by ""UF"".""FldValue"" "
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("13_U_C").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)

                CType(oForm.Items.Item("13_U_C").Specific, SAPbouiCOM.ComboBox).ExpandType = BoExpandType.et_DescriptionOnly
                oForm.Items.Item("13_U_C").DisplayDesc = True

                'If CType(oForm.Items.Item("13_U_C").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                '    Dim sTipo As String = CType(oForm.Items.Item("13_U_C").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                '    sSQL = "SELECT DISTINCT ""AbsEntry"", ""Code""  FROM ""ORST"" WHERE ""U_PP_TLIN""='O' and ""U_PP_TIPO""='" & sTipo & "' Order BY ""AbsEntry"" "
                '    objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").ValidValues, sSQL)
                'Else
                '    sSQL = "SELECT DISTINCT ""AbsEntry"", ""Code""  FROM ""ORST"" WHERE ""U_PP_TLIN""='O' and ""U_PP_TIPO""='Z' Order BY ""AbsEntry"" "
                '    objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").ValidValues, sSQL)
                'End If
                'For i = 1 To CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).RowCount
                '    If CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_5").Cells.Item(i).Specific, SAPbouiCOM.CheckBox).Checked = True Then
                '        CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).CommonSetting.SetCellEditable(i, 6, True)
                '    Else
                '        CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).CommonSetting.SetCellEditable(i, 6, False)
                '    End If
                'Next
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
    Private Function EventHandler_ItemPressed_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = "" : Dim oRs As SAPbobsCOM.Recordset = Nothing
        EventHandler_ItemPressed_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "B_ADD"
                    Return AddNewLine(oForm)
                Case "B_DEL"
                    Return RemoveLine(oForm)
                    'Case "C_0_5"
                    '    If CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_5").Cells.Item(pVal.Row).Specific, SAPbouiCOM.CheckBox).Checked = True Then
                    '        CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).CommonSetting.SetCellEditable(pVal.Row, 6, True)
                    '    Else
                    '        CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).CommonSetting.SetCellEditable(pVal.Row, 6, False)
                    '    End If
            End Select

            EventHandler_ItemPressed_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "1" And pVal.ActionSuccess = True Then
                RefreshList()
                oForm.Items.Item("2").Click(BoCellClickType.ct_Regular)
            End If

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function

    Private Function RefreshList() As Boolean

        Dim Assembly As Assembly = Nothing
        Dim calledType As Type = Nothing
        Dim obj As Object = Nothing
        Try

            For i = 0 To objGlobal.SBOApp.Forms.Count - 1
                If objGlobal.SBOApp.Forms.Item(i).TypeEx = "LIST_PP_FRAGMENTOS" Then
                    LIST_PP_FRAGMENTOS.Load_Grid(objGlobal, objGlobal.SBOApp.Forms.Item(i))
                End If
            Next

            Return True
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            If objGlobal.tipoCliente = EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente.Clasico Then
            End If
        End Try
    End Function
    Private Function AddNewLine(ByRef oform As SAPbouiCOM.Form, Optional ByVal Forzado As Boolean = False) As Boolean

        Dim _oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim _SQL As String = ""
        Dim NombreColumnaControl As String = ""
        Dim CodigoColumnaControl As String = ""
        Dim Tablas As String = ""
        Dim Columna As Integer = 0

        Try
            _oMatrix = CType(oform.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix)
            Tablas = "@PP_FRAGMENTOSL"

            'Pasar a la fuente de datos lo que tenga la matrix
            _oMatrix.FlushToDataSource()
            'Mirar si tengo que cargar los combos
            If _oMatrix.RowCount > 0 Then
                'Mirar si tengo línea en blanco
                With oform.DataSources.DBDataSources.Item(Tablas)
                    If .Size <> 0 Then
                        If .GetValue("U_PP_DOPE", .Size - 1) = "" Then
                            If Forzado = False Then
                                objGlobal.SBOApp.MessageBox("Ya tiene una línea disponible")
                            End If
                            _oMatrix.AutoResizeColumns()
                            Return True
                        Else
                            .InsertRecord(.Size)
                            _oMatrix.LoadFromDataSourceEx(True)
                            _oMatrix.AutoResizeColumns()
                        End If
                    End If
                End With
            Else
                _oMatrix.AddRow()
                _oMatrix.FlushToDataSource()
                _oMatrix.AutoResizeColumns()
            End If

            Return True
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            If objGlobal.tipoCliente = EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente.Clasico Then
                EXO_CleanCOM.CLiberaCOM.FormMatrix(_oMatrix)
            End If
        End Try
    End Function
    Private Function RemoveLine(ByRef oform As SAPbouiCOM.Form) As Boolean

        Dim _oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim Tablas As String = ""
        Dim _Row As Integer = 0
        Try
            _oMatrix = CType(oform.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix)
            Tablas = "@PP_FRAGMENTOSL"
            Try
                _Row = _oMatrix.GetNextSelectedRow(0, BoOrderType.ot_RowOrder)
            Catch ex As Exception
                _Row = -1
            End Try
            If _Row <= 0 Then
                objGlobal.SBOApp.MessageBox("Debe seleccionar un detalle para eliminar!")
                Return False
            End If
            If objGlobal.SBOApp.MessageBox("¿Está seguro que quiere eliminar el registro?", 2, "OK", "Cancelar") = 2 Then
                Return False
            End If
            With oform.DataSources.DBDataSources.Item(Tablas)
                .RemoveRecord(_Row - 1)
            End With
            _oMatrix.LoadFromDataSourceEx(True)
            _oMatrix.AutoResizeColumns()
            If oform.Mode = BoFormMode.fm_OK_MODE Then
                oform.Mode = BoFormMode.fm_UPDATE_MODE
            End If
            Return AddNewLine(oform, True)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            If objGlobal.tipoCliente = EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente.Clasico Then
                EXO_CleanCOM.CLiberaCOM.FormMatrix(_oMatrix)
            End If
        End Try
    End Function
#End Region
End Class
