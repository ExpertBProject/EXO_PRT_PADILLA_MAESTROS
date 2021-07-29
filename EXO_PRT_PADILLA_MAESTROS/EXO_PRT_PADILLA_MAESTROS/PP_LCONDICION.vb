﻿Imports System.Xml
Imports SAPbouiCOM
Public Class PP_LCONDICION
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        cargamenu()
        If actualizar Then
            cargaCampos()
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
            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_PP_LCONDICION.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_PP_LCONDICION", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults
            'objGlobal.SBOApp.StatusBar.SetText(res, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
    End Sub
#End Region
#Region "Eventos"
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then

            Else

                Select Case infoEvento.MenuUID
                    Case "PP-MnLCND"
                        'Cargamos UDO
                        'objGlobal.funcionesUI.cargaFormUdoBD("PP_LCONDICION")
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
                    Case "UDO_FT_PP_LCONDICION"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                Borra_lin_vacia(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                Borra_lin_vacia(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_PP_LCONDICION"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If oForm.Visible = True Then
                                    CargaCombo_Valor(oForm)

                                    If CType(oForm.Items.Item("cbCOND").Specific, SAPbouiCOM.ComboBox).Selected.Value.Trim = "LIST" Or CType(oForm.Items.Item("cbCOND").Specific, SAPbouiCOM.ComboBox).Selected.Value.Trim = "NOTLIST" Then
                                        CType(oForm.Items.Item("btnL").Specific, SAPbouiCOM.Button).Item.Visible = True
                                        oForm.Items.Item("btnL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                        CType(oForm.Items.Item("btnL").Specific, SAPbouiCOM.Button).Item.Visible = True
                                    Else
                                        oForm.Items.Item("btnL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                        CType(oForm.Items.Item("btnL").Specific, SAPbouiCOM.Button).Item.Visible = False
                                        CType(oForm.Items.Item("btnL").Specific, SAPbouiCOM.Button).Item.Visible = True
                                    End If
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

            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Sub Borra_lin_vacia(ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Try
            sSQL = "DELETE  FROM ""@PP_LCONDICIONL""  T0 WHERE ""Code""='" & CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Value.ToString & "' and IFNULL(T0.""U_PP_MCONDICION"",'')=''"
            objGlobal.refDi.SQL.sqlUpdB1(sSQL)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub CargaCombo_Valor(ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Dim sCodigo As String = ""
        Dim sTabla As String = "" : Dim sCampo As String = ""
        Try
            sCodigo = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(1).Specific, SAPbouiCOM.ComboBox).Value.ToString
            sSQL = "SELECT ""U_PP_TABLA"" FROM ""@PP_MCONDICION"" Where ""Code""='" & sCodigo & "' "
            sTabla = objGlobal.refDi.SQL.sqlStringB1(sSQL)
            sSQL = "SELECT ""U_PP_CAMPO"" FROM ""@PP_MCONDICION"" Where ""Code""='" & sCodigo & "' "
            sCampo = objGlobal.refDi.SQL.sqlStringB1(sSQL)
            sSQL = "SELECT """ & sCampo & """ FROM ""@" & sTabla & """ "
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_3").ValidValues, sSQL)
            CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_3").ExpandType = BoExpandType.et_ValueOnly

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_PP_LCONDICION"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    If EventHandler_COMBO_SELECT_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
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
                        Case "UDO_FT_PP_LCONDICION"
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
                        Case "UDO_FT_PP_LCONDICION"
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
                        Case "UDO_FT_PP_LCONDICION"
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
    Private Function EventHandler_COMBO_SELECT_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim sSQL As String = ""
        EventHandler_COMBO_SELECT_After = False
        Try

            If pVal.ItemChanged = True Then
                Select Case pVal.ItemUID
                    Case "cbValor"
                        If oForm.DataSources.UserDataSources.Item("UDVMAT").ValueEx = "" Then
                            oForm.DataSources.UserDataSources.Item("UDVMAT").ValueEx = CType(oForm.Items.Item("cbValor").Specific, SAPbouiCOM.ComboBox).Selected.Value.Trim
                            oForm.DataSources.UserDataSources.Item("UDVALORD").ValueEx = CType(oForm.Items.Item("cbValor").Specific, SAPbouiCOM.ComboBox).Selected.Description.Trim
                        Else
                            oForm.DataSources.UserDataSources.Item("UDVMAT").ValueEx &= "," & CType(oForm.Items.Item("cbValor").Specific, SAPbouiCOM.ComboBox).Selected.Value.Trim
                            oForm.DataSources.UserDataSources.Item("UDVALORD").ValueEx &= "," & CType(oForm.Items.Item("cbValor").Specific, SAPbouiCOM.ComboBox).Selected.Description.Trim
                        End If
                    Case "cbCOND"
                        If CType(oForm.Items.Item("cbCOND").Specific, SAPbouiCOM.ComboBox).Selected.Value.Trim = "LIST" Or CType(oForm.Items.Item("cbCOND").Specific, SAPbouiCOM.ComboBox).Selected.Value.Trim = "NOTLIST" Then
                            CType(oForm.Items.Item("btnL").Specific, SAPbouiCOM.Button).Item.Visible = True
                            oForm.Items.Item("btnL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            CType(oForm.Items.Item("txtVMAT").Specific, SAPbouiCOM.EditText).Item.Visible = True
                        Else
                            oForm.Items.Item("btnL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            CType(oForm.Items.Item("btnL").Specific, SAPbouiCOM.Button).Item.Visible = False
                            CType(oForm.Items.Item("txtVMAT").Specific, SAPbouiCOM.EditText).Item.Visible = False
                        End If
                End Select

            End If

            EventHandler_COMBO_SELECT_After = True

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
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim oItem As SAPbouiCOM.Item = Nothing
        EventHandler_Form_Visible = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                sSQL = "SELECT ""Code"", ""Name"" FROM ""@PP_MCONDICION"" Order BY ""Name"" "
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").ValidValues, sSQL)

                sSQL = "SELECT ""Code"", ""Name"" FROM ""@PP_CONDICION"" Order BY ""Name"" "
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").ValidValues, sSQL)
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbCOND").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)

                Mostrar_Ocultar(False, oForm)
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
    Private Sub Mostrar_Ocultar(ByVal bValor As Boolean, ByRef oform As SAPbouiCOM.Form)
        Try
            CType(oform.Items.Item("lblCCOND").Specific, SAPbouiCOM.StaticText).Item.Visible = bValor
            CType(oform.Items.Item("txtCCOND").Specific, SAPbouiCOM.EditText).Item.Visible = bValor
            CType(oform.Items.Item("txtCNAME").Specific, SAPbouiCOM.EditText).Item.Visible = bValor
            CType(oform.Items.Item("lblC").Specific, SAPbouiCOM.StaticText).Item.Visible = bValor
            CType(oform.Items.Item("cbCOND").Specific, SAPbouiCOM.ComboBox).Item.Visible = bValor
            CType(oform.Items.Item("lblValor").Specific, SAPbouiCOM.StaticText).Item.Visible = bValor
            CType(oform.Items.Item("cbValor").Specific, SAPbouiCOM.ComboBox).Item.Visible = bValor
            CType(oform.Items.Item("txtValor").Specific, SAPbouiCOM.EditText).Item.Visible = bValor
            CType(oform.Items.Item("txtVMAT").Specific, SAPbouiCOM.EditText).Item.Visible = bValor

            CType(oform.Items.Item("btnAceptar").Specific, SAPbouiCOM.Button).Item.Visible = bValor
            CType(oform.Items.Item("btnC").Specific, SAPbouiCOM.Button).Item.Visible = bValor
            CType(oform.Items.Item("btnL").Specific, SAPbouiCOM.Button).Item.Visible = False

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub Mostrar_Ocultar_Botones(ByVal bValor As Boolean, ByRef oform As SAPbouiCOM.Form)
        Try
            Select Case bValor
                Case True
                    oform.Items.Item("btnA").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                    oform.Items.Item("btnE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                    oform.Items.Item("btnB").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                    oform.Items.Item("btnL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                Case Else
                    oform.Items.Item("btnA").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    oform.Items.Item("btnE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    oform.Items.Item("btnB").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                    oform.Items.Item("btnL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End Select

            'CType(oform.Items.Item("btnA").Specific, SAPbouiCOM.Button).Item.Enabled = bValor
            'CType(oform.Items.Item("btnE").Specific, SAPbouiCOM.Button).Item.Enabled = bValor
            'CType(oform.Items.Item("btnB").Specific, SAPbouiCOM.Button).Item.Enabled = bValor
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = "" : Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim Row As Integer = 0
        Dim sCod As String = ""
        Dim ID As String = "" : Dim sValorMatrix As String = ""

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "0_U_G"
                    If pVal.Row > 0 Then
                        If CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).IsRowSelected(pVal.Row) = True Then
                            oForm.DataSources.UserDataSources.Item("UDROW").ValueEx = pVal.Row.ToString
                        Else
                            oForm.DataSources.UserDataSources.Item("UDROW").ValueEx = ""
                            Exit Function
                        End If
                    Else
                        oForm.DataSources.UserDataSources.Item("UDROW").ValueEx = ""
                        Exit Function
                    End If

                Case "btnC" 'Cancelar
#Region "Cancelar"
                    CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.EditText).Active = True
                    ' Se limpian los campos y se ocultan
                    oForm.DataSources.UserDataSources.Item("UDCCOND").ValueEx = ""
                    oForm.DataSources.UserDataSources.Item("UDNAME").ValueEx = ""
                    oForm.DataSources.UserDataSources.Item("UDCOND").ValueEx = ""
                    oForm.DataSources.UserDataSources.Item("UDVALORC").ValueEx = ""
                    oForm.DataSources.UserDataSources.Item("UDVALORT").ValueEx = ""
                    oForm.DataSources.UserDataSources.Item("UDVALORD").ValueEx = ""
                    oForm.DataSources.UserDataSources.Item("UDVMAT").ValueEx = ""
                    Mostrar_Ocultar(False, oForm)
                    Mostrar_Ocultar_Botones(True, oForm)
#End Region
                Case "btnA" 'Añadir
#Region "Añadir"
                    Mostrar_Ocultar_Botones(False, oForm)
                    Mostrar_Ocultar(True, oForm)
                    oForm.DataSources.UserDataSources.Item("UDACC").ValueEx = "A"
#End Region
                Case "btnB" 'Borrar
#Region "Borrar"
                    'Comprobamos si ha seleccionado alguna fila
                    Try
                        If oForm.DataSources.UserDataSources.Item("UDROW").Value = "" Then
                            Row = -1
                        Else
                            Row = CInt(oForm.DataSources.UserDataSources.Item("UDROW").Value)
                        End If

                    Catch ex As Exception
                        Row = -1
                    End Try
                    If Row < 0 Then
                        objGlobal.SBOApp.MessageBox("Debe seleccionar una línea!")
                        Return False
                    End If
                    CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).DeleteRow(Row)
                    'oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").RemoveRecord(Row - 1)
                    ' CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).LoadFromDataSource()
                    If oForm.Mode = BoFormMode.fm_OK_MODE Then oForm.Mode = BoFormMode.fm_UPDATE_MODE
#End Region
                Case "btnE" 'Editar
#Region "Editar"
                    'Comprobamos si ha seleccionado alguna fila
                    Try
                        If oForm.DataSources.UserDataSources.Item("UDROW").Value = "" Then
                            Row = -1
                        Else
                            Row = CInt(oForm.DataSources.UserDataSources.Item("UDROW").Value)
                        End If

                    Catch ex As Exception
                        Row = -1
                    End Try
                    If Row < 0 Then
                        objGlobal.SBOApp.MessageBox("Debe seleccionar una línea!")
                        Return False
                    End If
                    ID = oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").GetValue("LineId", Row - 1)
                    oForm.DataSources.UserDataSources.Item("UDACC").ValueEx = "E"
                    Mostrar_Ocultar_Botones(False, oForm)
                    Mostrar_Ocultar(True, oForm)
                    'Rellenamos los datos
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Cargando datos de la línea...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    sCod = oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").GetValue("U_PP_MCONDICION", Row - 1)
                    sValorMatrix = oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").GetValue("U_PP_VALOR", Row - 1)
                    oForm.DataSources.UserDataSources.Item("UDVMAT").ValueEx = sValorMatrix
                    oForm.DataSources.UserDataSources.Item("UDCCOND").ValueEx = sCod
                    'Buscamos el nombre para introducirlo en el campo
                    sSQL = "SELECT ""Name"" FROM ""@PP_MCONDICION"" Where ""Code""='" & sCod & "' "
                    oForm.DataSources.UserDataSources.Item("UDNAME").ValueEx = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                    oForm.DataSources.UserDataSources.Item("UDCOND").ValueEx = oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").GetValue("U_PP_CONDICION", Row - 1)
                    'Ahora tenemos que ver si es lista o texto para introducir el valor
                    If sCod <> "" Then
                        Dim sCondicion As String = sCod
                        Dim sTipo As String = "" : Dim sCampo As String = ""
                        Dim sEntidad As String = "" : Dim sCampoV As String = ""
                        sSQL = "SELECT *  FROM ""@PP_MCONDICION"" WHERE ""Code""='" & sCondicion & "' "
                        oRs.DoQuery(sSQL)
                        For i = 0 To oRs.RecordCount - 1
                            sTipo = oRs.Fields.Item("U_PP_TCAMPOCFG").Value.ToString
                            sCampo = oRs.Fields.Item("U_PP_CAMPOCFG").Value.ToString.Replace("U_", "")
                            sEntidad = oRs.Fields.Item("U_PP_ENTIDAD").Value.ToString.Replace("U_", "")
                            sCampoV = oRs.Fields.Item("U_PP_CAMPO").Value.ToString.Replace("U_", "")
                            Dim SearchForThis As String = "IsValidValues"

                            Dim FirstCharacter As Integer = sTipo.IndexOf(SearchForThis)
                            If FirstCharacter > 0 Then 'ValidValues
                                CType(oForm.Items.Item("cbValor").Specific, SAPbouiCOM.ComboBox).Item.Visible = True
                                CType(oForm.Items.Item("txtValor").Specific, SAPbouiCOM.EditText).Item.Visible = False
                                sSQL = "SELECT F1.""FldValue"", F1.""Descr"" From ""UFD1"" F1 "
                                sSQL &= " INNER JOIN ""CUFD"" CF ON CF.""TableID""=F1.""TableID"" and CF.""FieldID""=F1.""FieldID"" "
                                sSQL &= " Where F1.""TableID""='@PP_OITM' and CF.""AliasID""='" & sCampo & "'"
                                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbValor").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                                If CType(oForm.Items.Item("cbCOND").Specific, SAPbouiCOM.ComboBox).Selected.Value.Trim <> "LIST" And CType(oForm.Items.Item("cbCOND").Specific, SAPbouiCOM.ComboBox).Selected.Value.Trim <> "NOTLIST" Then
                                    CType(oForm.Items.Item("cbValor").Specific, SAPbouiCOM.ComboBox).Select(sValorMatrix, BoSearchKey.psk_ByValue)
                                End If
                            Else
                                CType(oForm.Items.Item("cbValor").Specific, SAPbouiCOM.ComboBox).Item.Visible = False
                                CType(oForm.Items.Item("txtValor").Specific, SAPbouiCOM.EditText).Item.Visible = True
                                Dim ocfls As SAPbouiCOM.ChooseFromListCollection

                                ocfls = oForm.ChooseFromLists

                                Dim ocfl As SAPbouiCOM.ChooseFromList

                                Dim cflcrepa As SAPbouiCOM.ChooseFromListCreationParams

                                cflcrepa = CType(objGlobal.SBOApp.CreateObject(BoCreatableObjectType.cot_ChooseFromListCreationParams), ChooseFromListCreationParams)

                                cflcrepa.MultiSelection = False
                                'If sEntidad = "PP_ACABADOS" Then
                                '    sEntidad = "PP_ACABADO"
                                'End If
                                cflcrepa.ObjectType = sEntidad
                                cflcrepa.UniqueID = “CFL" & sEntidad
                                Try
                                    ocfl = ocfls.Add(cflcrepa)
                                Catch ex As Exception

                                End Try

                                CType(oForm.Items.Item("txtValor").Specific, SAPbouiCOM.EditText).ChooseFromListUID = “CFL" & sEntidad
                                sSQL = " SELECT ""ObjectType"" FROM ""OUTB"" WHERE ""TableName""='" & sEntidad & "' "
                                Dim sTipoUDO As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                Select Case sTipoUDO
                                    Case "3" : CType(oForm.Items.Item("txtValor").Specific, SAPbouiCOM.EditText).ChooseFromListAlias = "DocEntry"
                                    Case Else : CType(oForm.Items.Item("txtValor").Specific, SAPbouiCOM.EditText).ChooseFromListAlias = "Code"
                                End Select
                                If CType(oForm.Items.Item("cbCOND").Specific, SAPbouiCOM.ComboBox).Selected.Value.Trim <> "LIST" And CType(oForm.Items.Item("cbCOND").Specific, SAPbouiCOM.ComboBox).Selected.Value.Trim <> "NOTLIST" Then
                                    CType(oForm.Items.Item("txtValor").Specific, SAPbouiCOM.EditText).Value = sValorMatrix
                                End If

                            End If
                            If CType(oForm.Items.Item("cbCOND").Specific, SAPbouiCOM.ComboBox).Selected.Value.Trim = "LIST" Or CType(oForm.Items.Item("cbCOND").Specific, SAPbouiCOM.ComboBox).Selected.Value.Trim = "NOTLIST" Then
                                CType(oForm.Items.Item("btnL").Specific, SAPbouiCOM.Button).Item.Visible = True
                                oForm.Items.Item("btnL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                CType(oForm.Items.Item("txtVMAT").Specific, SAPbouiCOM.EditText).Item.Visible = True
                            Else
                                oForm.Items.Item("btnL").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                CType(oForm.Items.Item("btnL").Specific, SAPbouiCOM.Button).Item.Visible = False
                                CType(oForm.Items.Item("txtVMAT").Specific, SAPbouiCOM.EditText).Item.Visible = False
                            End If
                            oRs.MoveNext()
                        Next
                    End If
#End Region
                Case "btnAceptar" 'Aceptar
                    Select Case oForm.DataSources.UserDataSources.Item("UDACC").Value.Trim
                        Case "A" 'Añadir
#Region "Añadir"
                            If oForm.DataSources.UserDataSources.Item("UDVMAT").ValueEx.Trim <> "" Then
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Cargando línea...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Dim i As Integer = 0
                                i = oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").Size - 1
                                'Cragamos los datos en la matrix
                                CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).FlushToDataSource()
                                oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").InsertRecord(i)
                                oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").Offset = i
                                oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").SetValue("U_PP_MCONDICION", oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").Offset, oForm.DataSources.UserDataSources.Item("UDCCOND").ValueEx)
                                oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").SetValue("U_PP_CONDICION", oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").Offset, oForm.DataSources.UserDataSources.Item("UDCOND").ValueEx)
                                'Dim sValor As String = ""
                                'If oForm.DataSources.UserDataSources.Item("UDVALORT").ValueEx <> "" Then
                                '    sValor = oForm.DataSources.UserDataSources.Item("UDVALORT").ValueEx
                                'Else
                                '    sValor = oForm.DataSources.UserDataSources.Item("UDVALORC").ValueEx
                                'End If
                                oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").SetValue("U_PP_VALOR", oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").Offset, oForm.DataSources.UserDataSources.Item("UDVMAT").ValueEx.Trim)
                                'Tenemos que poner la descripción
                                oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").SetValue("U_PP_VALORD", oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").Offset, oForm.DataSources.UserDataSources.Item("UDVALORD").Value.Trim)
                                CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).LoadFromDataSource()
                            Else
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - No se ha indicado valor. Por favor, indique un valor disponible.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If

#End Region
                        Case "E" 'Editar
#Region "Editar"
                            Row = CInt(oForm.DataSources.UserDataSources.Item("UDROW").Value)
                            If Row < 0 Then
                                objGlobal.SBOApp.MessageBox("Se ha perdido la linea a actualizar!!")
                                Return False
                            End If
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Actualizando línea...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").SetValue("U_PP_MCONDICION", Row - 1, oForm.DataSources.UserDataSources.Item("UDCCOND").ValueEx)
                            oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").SetValue("U_PP_CONDICION", Row - 1, oForm.DataSources.UserDataSources.Item("UDCOND").ValueEx)
                            'Dim sValor As String = ""
                            'If oForm.DataSources.UserDataSources.Item("UDVALORT").ValueEx <> "" Then
                            '    sValor = oForm.DataSources.UserDataSources.Item("UDVALORT").ValueEx
                            'Else
                            '    sValor = oForm.DataSources.UserDataSources.Item("UDVALORC").ValueEx
                            'End If
                            oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").SetValue("U_PP_VALOR", Row - 1, oForm.DataSources.UserDataSources.Item("UDVMAT").Value.Trim)
                            'Tenemos que poner la descripción
                            oForm.DataSources.DBDataSources.Item("@PP_LCONDICIONL").SetValue("U_PP_VALORD", Row - 1, oForm.DataSources.UserDataSources.Item("UDVALORD").Value.Trim)
                            CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).LoadFromDataSource()
#End Region
                    End Select

                    ' Se limpian los campos
                    oForm.DataSources.UserDataSources.Item("UDCCOND").ValueEx = ""
                    oForm.DataSources.UserDataSources.Item("UDNAME").ValueEx = ""
                    oForm.DataSources.UserDataSources.Item("UDCOND").ValueEx = ""
                    oForm.DataSources.UserDataSources.Item("UDVALORC").ValueEx = ""
                    oForm.DataSources.UserDataSources.Item("UDVALORT").ValueEx = ""
                    oForm.DataSources.UserDataSources.Item("UDVALORD").ValueEx = ""
                    oForm.DataSources.UserDataSources.Item("UDVMAT").ValueEx = ""

                    Mostrar_Ocultar_Botones(False, oForm)
                    If oForm.Mode = BoFormMode.fm_OK_MODE Then oForm.Mode = BoFormMode.fm_UPDATE_MODE
                Case "btnL" 'Limpiar la lista
                    oForm.DataSources.UserDataSources.Item("UDVALORD").ValueEx = ""
                    oForm.DataSources.UserDataSources.Item("UDVMAT").ValueEx = ""
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_Choose_FromList_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        EventHandler_Choose_FromList_After = False
        Dim sCod As String = "" : Dim sDes As String = ""
        Try
            Dim oDataTable As SAPbouiCOM.IChooseFromListEvent = CType(pVal, SAPbouiCOM.IChooseFromListEvent)
            If pVal.ItemUID = "txtCCOND" Then
                If oDataTable IsNot Nothing Then
                    Try
                        Select Case oForm.ChooseFromLists.Item(oDataTable.ChooseFromListUID).ObjectType
                            Case "PP_MCONDICION"
                                Try
                                    sCod = oDataTable.SelectedObjects.GetValue("Code", 0).ToString
                                    sDes = oDataTable.SelectedObjects.GetValue("Name", 0).ToString
                                    oForm.DataSources.UserDataSources.Item("UDCCOND").ValueEx = sCod
                                    oForm.DataSources.UserDataSources.Item("UDNAME").ValueEx = sDes
                                Catch ex As Exception
                                    oForm.DataSources.UserDataSources.Item("UDCCOND").ValueEx = sCod
                                    oForm.DataSources.UserDataSources.Item("UDNAME").ValueEx = sDes
                                End Try
                        End Select
                        If sCod <> "" Then
                            Dim sCondicion As String = sCod
                            Dim sTipo As String = "" : Dim sCampo As String = ""
                            Dim sEntidad As String = "" : Dim sCampoV As String = ""
                            sSQL = "SELECT *  FROM ""@PP_MCONDICION"" WHERE ""Code""='" & sCondicion & "' "
                            oRs.DoQuery(sSQL)
                            For i = 0 To oRs.RecordCount - 1
                                sTipo = oRs.Fields.Item("U_PP_TCAMPOCFG").Value.ToString
                                sCampo = oRs.Fields.Item("U_PP_CAMPOCFG").Value.ToString.Replace("U_", "")
                                sEntidad = oRs.Fields.Item("U_PP_ENTIDAD").Value.ToString.Replace("U_", "")
                                sCampoV = oRs.Fields.Item("U_PP_CAMPO").Value.ToString.Replace("U_", "")
                                Dim SearchForThis As String = "IsValidValues"

                                Dim FirstCharacter As Integer = sTipo.IndexOf(SearchForThis)
                                If FirstCharacter > 0 Then 'ValidValues
                                    CType(oForm.Items.Item("cbValor").Specific, SAPbouiCOM.ComboBox).Item.Visible = True
                                    CType(oForm.Items.Item("txtValor").Specific, SAPbouiCOM.EditText).Item.Visible = False
                                    sSQL = "SELECT F1.""FldValue"", F1.""Descr"" From ""UFD1"" F1 "
                                    sSQL &= " INNER JOIN ""CUFD"" CF ON CF.""TableID""=F1.""TableID"" and CF.""FieldID""=F1.""FieldID"" "
                                    sSQL &= " Where F1.""TableID""='@PP_OITM' and CF.""AliasID""='" & sCampo & "'"
                                    objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbValor").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                                Else
                                    CType(oForm.Items.Item("cbValor").Specific, SAPbouiCOM.ComboBox).Item.Visible = False
                                    CType(oForm.Items.Item("txtValor").Specific, SAPbouiCOM.EditText).Item.Visible = True
                                    Dim ocfls As SAPbouiCOM.ChooseFromListCollection

                                    ocfls = oForm.ChooseFromLists

                                    Dim ocfl As SAPbouiCOM.ChooseFromList

                                    Dim cflcrepa As SAPbouiCOM.ChooseFromListCreationParams

                                    cflcrepa = CType(objGlobal.SBOApp.CreateObject(BoCreatableObjectType.cot_ChooseFromListCreationParams), ChooseFromListCreationParams)

                                    cflcrepa.MultiSelection = False
                                    'If sEntidad = "PP_ACABADOS" Then
                                    '    sEntidad = "PP_ACABADO"
                                    'End If
                                    cflcrepa.ObjectType = sEntidad
                                    cflcrepa.UniqueID = “CFL" & sEntidad
                                    Try
                                        ocfl = ocfls.Add(cflcrepa)
                                    Catch ex As Exception

                                    End Try

                                    CType(oForm.Items.Item("txtValor").Specific, SAPbouiCOM.EditText).ChooseFromListUID = “CFL" & sEntidad
                                    sSQL = " SELECT ""ObjectType"" FROM ""OUTB"" WHERE ""TableName""='" & sEntidad & "' "
                                    Dim sTipoUDO As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                    Select Case sTipoUDO
                                        Case "3" : CType(oForm.Items.Item("txtValor").Specific, SAPbouiCOM.EditText).ChooseFromListAlias = "DocEntry"
                                        Case Else : CType(oForm.Items.Item("txtValor").Specific, SAPbouiCOM.EditText).ChooseFromListAlias = "Code"
                                    End Select
                                End If
                                oRs.MoveNext()
                            Next
                        End If

                    Catch ex As Exception
                        Throw ex
                    End Try
                End If
                If oForm.Mode = BoFormMode.fm_UPDATE_MODE Then oForm.Mode = BoFormMode.fm_OK_MODE
                Mostrar_Ocultar_Botones(False, oForm)
            ElseIf pVal.ItemUID = "txtValor" Then
                If oDataTable IsNot Nothing Then
                    Try
                        sCod = oDataTable.SelectedObjects.GetValue("Code", 0).ToString
                        sDes = oDataTable.SelectedObjects.GetValue("Name", 0).ToString
                    Catch ex As Exception
                        sCod = oDataTable.SelectedObjects.GetValue("DocEntry", 0).ToString
                        sDes = oDataTable.SelectedObjects.GetValue("U_PP_NAME", 0).ToString
                    End Try
                    oForm.DataSources.UserDataSources.Item("UDVALORT").ValueEx = sCod

                    If CType(oForm.Items.Item("cbCOND").Specific, SAPbouiCOM.ComboBox).Selected.Value.Trim = "LIST" Or CType(oForm.Items.Item("cbCOND").Specific, SAPbouiCOM.ComboBox).Selected.Value.Trim = "NOTLIST" Then
                        If oForm.DataSources.UserDataSources.Item("UDVMAT").ValueEx = "" Then
                            oForm.DataSources.UserDataSources.Item("UDVMAT").ValueEx = oForm.DataSources.UserDataSources.Item("UDVALORT").ValueEx.Trim
                            oForm.DataSources.UserDataSources.Item("UDVALORD").ValueEx = sDes
                        Else
                            oForm.DataSources.UserDataSources.Item("UDVMAT").ValueEx &= "," & oForm.DataSources.UserDataSources.Item("UDVALORT").ValueEx.Trim
                            oForm.DataSources.UserDataSources.Item("UDVALORD").ValueEx &= "," & sDes
                        End If
                    Else
                        oForm.DataSources.UserDataSources.Item("UDVMAT").ValueEx = oForm.DataSources.UserDataSources.Item("UDVALORT").ValueEx.Trim
                        oForm.DataSources.UserDataSources.Item("UDVALORD").ValueEx = sDes
                    End If
                End If
                If oForm.Mode = BoFormMode.fm_UPDATE_MODE Then oForm.Mode = BoFormMode.fm_OK_MODE
                Mostrar_Ocultar_Botones(False, oForm)
            End If

            EventHandler_Choose_FromList_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
#End Region
End Class
