Imports System.Xml
Imports SAPbouiCOM

Public Class PP_CONDFRAG
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
            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_PP_CONDFRAG.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_PP_CONDFRAG", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
            oForm.Freeze(True)
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_PP_CONDFRAG"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_PP_CONDFRAG"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If oForm.Visible = True Then
                                    If CType(oForm.Items.Item("14_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                                        Dim sCondicion As String = CType(oForm.Items.Item("14_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                                        Dim sTipo As String = "" : Dim sCampo As String = "" : Dim sSQL As String = ""
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
                                                CType(oForm.Items.Item("15_U_Cb").Specific, SAPbouiCOM.ComboBox).Item.Visible = True
                                                CType(oForm.Items.Item("16_U_E").Specific, SAPbouiCOM.EditText).Item.Visible = False
                                                sSQL = "SELECT F1.""FldValue"", F1.""Descr"" From ""UFD1"" F1 "
                                                sSQL &= " INNER JOIN ""CUFD"" CF ON CF.""TableID""=F1.""TableID"" and CF.""FieldID""=F1.""FieldID"" "
                                                sSQL &= " Where F1.""TableID""='@PP_OITM' and CF.""AliasID""='" & sCampo & "'"
                                                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("15_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                                            Else
                                                CType(oForm.Items.Item("15_U_Cb").Specific, SAPbouiCOM.ComboBox).Item.Visible = False
                                                CType(oForm.Items.Item("16_U_E").Specific, SAPbouiCOM.EditText).Item.Visible = True
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

                                                CType(oForm.Items.Item("16_U_E").Specific, SAPbouiCOM.EditText).ChooseFromListUID = “CFL" & sEntidad
                                                sSQL = " SELECT ""ObjectType"" FROM ""OUTB"" WHERE ""TableName""='" & sEntidad & "' "
                                                Dim sTipoUDO As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                                Select Case sTipoUDO
                                                    Case "3" : CType(oForm.Items.Item("16_U_E").Specific, SAPbouiCOM.EditText).ChooseFromListAlias = "DocEntry"
                                                    Case Else : CType(oForm.Items.Item("16_U_E").Specific, SAPbouiCOM.EditText).ChooseFromListAlias = "Code"
                                                End Select
                                            End If
                                            oRs.MoveNext()
                                        Next
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
            oForm.Freeze(False)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            oForm.Freeze(False)
            Return False
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_PP_CONDFRAG"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    If EventHandler_COMBO_SELECT_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
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
                        Case "UDO_FT_PP_CONDFRAG"
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
                        Case "UDO_FT_PP_CONDFRAG"
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
                        Case "UDO_FT_PP_CONDFRAG"
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
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        EventHandler_COMBO_SELECT_After = False
        Try

            If pVal.ItemUID = "14_U_Cb" Then
                If CType(oForm.Items.Item("14_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                    Dim sCondicion As String = CType(oForm.Items.Item("14_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
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
                            CType(oForm.Items.Item("15_U_Cb").Specific, SAPbouiCOM.ComboBox).Item.Visible = True
                            CType(oForm.Items.Item("16_U_E").Specific, SAPbouiCOM.EditText).Item.Visible = False
                            sSQL = "SELECT F1.""FldValue"", F1.""Descr"" From ""UFD1"" F1 "
                            sSQL &= " INNER JOIN ""CUFD"" CF ON CF.""TableID""=F1.""TableID"" and CF.""FieldID""=F1.""FieldID"" "
                            sSQL &= " Where F1.""TableID""='@PP_OITM' and CF.""AliasID""='" & sCampo & "'"
                            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("15_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                        Else
                            CType(oForm.Items.Item("15_U_Cb").Specific, SAPbouiCOM.ComboBox).Item.Visible = False
                            CType(oForm.Items.Item("16_U_E").Specific, SAPbouiCOM.EditText).Item.Visible = True
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

                            CType(oForm.Items.Item("16_U_E").Specific, SAPbouiCOM.EditText).ChooseFromListUID = “CFL" & sEntidad
                            sSQL = " SELECT ""ObjectType"" FROM ""OUTB"" WHERE ""TableName""='" & sEntidad & "' "
                            Dim sTipoUDO As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                            Select Case sTipoUDO
                                Case "3" : CType(oForm.Items.Item("16_U_E").Specific, SAPbouiCOM.EditText).ChooseFromListAlias = "DocEntry"
                                Case Else : CType(oForm.Items.Item("16_U_E").Specific, SAPbouiCOM.EditText).ChooseFromListAlias = "Code"
                            End Select


                        End If
                        oRs.MoveNext()
                    Next
                End If
            End If

            EventHandler_COMBO_SELECT_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function EventHandler_Choose_FromList_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

        EventHandler_Choose_FromList_After = False
        Dim sCod As String = "" : Dim sDes As String = ""
        Try
            Dim oDataTable As SAPbouiCOM.IChooseFromListEvent = CType(pVal, SAPbouiCOM.IChooseFromListEvent)
            If pVal.ItemUID = "0_U_G" Then
                If pVal.ItemUID = "0_U_G" And pVal.ColUID = "C_0_1" Then
                    Dim iRegistros As Integer = oForm.DataSources.DBDataSources.Item("@PP_CODFRAGL").Size
                    Dim iRegActivo As Integer = pVal.Row
                    If iRegistros = iRegActivo Then
                        CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).AddRow()
                        CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).ClearRowData(iRegActivo + 1)
                        CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).FlushToDataSource()
                    End If

                End If
                If oDataTable IsNot Nothing Then
                    Try
                        Select Case oForm.ChooseFromLists.Item(oDataTable.ChooseFromListUID).ObjectType
                            Case "PP_FRAGMENTOS"
                                Try
                                    sCod = oDataTable.SelectedObjects.GetValue("Code", 0).ToString
                                    sDes = oDataTable.SelectedObjects.GetValue("Name", 0).ToString

                                    'oForm.DataSources.DBDataSources.Item("@PP_CODFRAGL").SetValue("U_PP_FRAGMENTO", pVal.Row - 1, sCod)
                                    oForm.DataSources.DBDataSources.Item("@PP_CODFRAGL").SetValue("U_PP_FRAGNAME", pVal.Row - 1, sDes)
                                    'CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_3").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = sDes

                                Catch ex As Exception
                                    oForm.DataSources.DBDataSources.Item("@PP_CODFRAGL").SetValue("U_PP_FRAGMENTO", pVal.Row - 1, sCod)
                                    oForm.DataSources.DBDataSources.Item("@PP_CODFRAGL").SetValue("U_PP_FRAGNAME", pVal.Row - 1, sDes)
                                End Try
                            Case "PP_LCONDICION"
                                Try
                                    sCod = oDataTable.SelectedObjects.GetValue("Code", 0).ToString
                                    sDes = oDataTable.SelectedObjects.GetValue("Name", 0).ToString

                                    oForm.DataSources.DBDataSources.Item("@PP_CODFRAGL").SetValue("U_PP_CODNAME", pVal.Row - 1, sDes)
                                Catch ex As Exception
                                    oForm.DataSources.DBDataSources.Item("@PP_CODFRAGL").SetValue("U_PP_CONDICION", pVal.Row - 1, sCod)
                                    oForm.DataSources.DBDataSources.Item("@PP_CODFRAGL").SetValue("U_PP_CODNAME", pVal.Row - 1, sDes)
                                End Try
                        End Select
                        If oForm.Mode = BoFormMode.fm_OK_MODE Then oForm.Mode = BoFormMode.fm_UPDATE_MODE
                    Catch ex As Exception
                        Throw ex
                    End Try
                End If
            ElseIf pVal.ItemUID = "16_U_E" Then
                Try
                    sCod = oDataTable.SelectedObjects.GetValue("Code", 0).ToString

                Catch ex As Exception
                    sCod = oDataTable.SelectedObjects.GetValue("DocEntry", 0).ToString
                End Try
                oForm.DataSources.DBDataSources.Item("@PP_CONDFRAG").SetValue("U_PP_Valor", 0, sCod)
                If oForm.Mode = BoFormMode.fm_OK_MODE Then oForm.Mode = BoFormMode.fm_UPDATE_MODE
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
                sSQL = "SELECT ""Code"",""Name"" From ""@PP_MCONDICION""  "
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("14_U_Cb").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
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
