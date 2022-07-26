Imports System.Xml
Imports SAPbouiCOM
Public Class PP_MCONDICION
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
            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_PP_MCONDICION.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_PP_MCONDICION", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
                Select Case infoEvento.MenuUID
                    Case "1281", "1282"
                        Select Case objGlobal.SBOApp.Forms.ActiveForm.TypeEx
                            Case "UDO_FT_PP_MCONDICION"
                                objGlobal.SBOApp.Forms.ActiveForm.Items.Item("0_U_E").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_True)
                        End Select
                End Select


            Else

                Select Case infoEvento.MenuUID
                    Case "PP-MnMCND"
                        'Cargamos UDO
                        objGlobal.funcionesUI.cargaFormUdoBD("PP_MCONDICION")
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
                    Case "UDO_FT_PP_MCONDICION"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_PP_MCONDICION"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If oForm.Visible = True Then
                                    Dim sUDO As String = "" : Dim sTabla As String = "" : Dim sSQL As String = ""

                                    If oForm.Mode <> BoFormMode.fm_ADD_MODE Then
                                        oForm.Items.Item("0_U_E").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
                                    Else
                                        oForm.Items.Item("0_U_E").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_True)
                                    End If

                                    With oForm.DataSources.DBDataSources.Item("@PP_MCONDICION")
                                        sSQL = "SELECT ""COLUMN_NAME"", ""COLUMN_NAME"" ""Descripción"" FROM TABLE_COLUMNS WHERE ""TABLE_NAME""='@" & .GetValue("U_PP_TCONF", .Offset).Trim & "' and ""SCHEMA_NAME""='" & objGlobal.compañia.CompanyDB & "' "
                                        sSQL &= " And ""COLUMN_NAME"" in ('Code', 'DocEntry', 'DocNum')  "
                                        sSQL &= " UNION ALL "
                                        sSQL &= "SELECT Concat('U_',""AliasID"") ""AliasID"",Concat(Concat(""AliasID"",' - '),""Descr"") ""Descripción"" "
                                        sSQL &= " From ""CUFD"" Where ""TableID""='@" & .GetValue("U_PP_TCONF", .Offset).Trim & "' "
                                    End With
                                    PP_UTILITIES.EO.Formularios.MyFormFunctions.Init_Combo(objGlobal, oForm, "16_U_C", sSQL)

                                    'SELECT T0."TableName", T0."Descr" FROM OUTB T0 WHERE T0."TableName" LIKE 'PP_OITM%'
                                    'PP_UTILITIES.EO.Formularios.MyFormFunctions.Init_Combo(objGlobal, oForm, "", "")

                                    sUDO = CType(oForm.Items.Item("13_U_C").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                                    If sUDO = "-1" Then
                                        sSQL = "Select '-1','None' FROM DUMMY"
                                        objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("15_U_C").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                                        CType(oForm.Items.Item("15_U_C").Specific, SAPbouiCOM.ComboBox).ExpandType = BoExpandType.et_ValueDescription
                                        CType(oForm.Items.Item("14_U_E").Specific, SAPbouiCOM.EditText).Value = "Ninguna"
                                    Else
                                        sSQL = "SELECT ""TableName"" FROM ""OUDO"" WHERE ""Code""='" & sUDO & "' "
                                        sTabla = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                        CType(oForm.Items.Item("14_U_E").Specific, SAPbouiCOM.EditText).Value = sTabla

                                        sSQL = "SELECT ""COLUMN_NAME"", ""COLUMN_NAME"" ""Descripción"" FROM TABLE_COLUMNS WHERE ""TABLE_NAME""='@" & sTabla & "' and ""SCHEMA_NAME""='" & objGlobal.compañia.CompanyDB & "' "
                                        sSQL &= " And ""COLUMN_NAME"" in ('Code', 'DocEntry', 'DocNum')  "
                                        sSQL &= " UNION ALL "
                                        sSQL &= "SELECT Concat('U_',""AliasID"") ""AliasID"",Concat(Concat(""AliasID"",' - '),""Descr"") ""Descripción"" "
                                        sSQL &= " From ""CUFD"" Where ""TableID""='@" & sTabla & "' "
                                        objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("15_U_C").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                                        CType(oForm.Items.Item("15_U_C").Specific, SAPbouiCOM.ComboBox).ExpandType = BoExpandType.et_ValueDescription

                                    End If
                                    sSQL = "SELECT ""COLUMN_NAME"", ""COLUMN_NAME"" ""Descripción"" FROM TABLE_COLUMNS WHERE ""TABLE_NAME""='@PP_OITM' and ""SCHEMA_NAME""='" & objGlobal.compañia.CompanyDB & "' "
                                    sSQL &= " And ""COLUMN_NAME"" in ('Code', 'DocEntry', 'DocNum')  "
                                    sSQL &= " UNION ALL "
                                    sSQL &= "SELECT Concat('U_',""AliasID"") ""AliasID"",Concat(Concat(""AliasID"",' - '),""Descr"") ""Descripción"" "
                                    sSQL &= " From ""CUFD"" Where ""TableID""='@PP_OITM' "
                                    objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("16_U_C").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                                    CType(oForm.Items.Item("16_U_C").Specific, SAPbouiCOM.ComboBox).ExpandType = BoExpandType.et_ValueDescription

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

    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_PP_MCONDICION"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    If EventHandler_COMBO_SELECT(objGlobal, infoEvento) = False Then
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
                        Case "UDO_FT_PP_MCONDICION"
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
                        Case "UDO_FT_PP_MCONDICION"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_Form_Visible(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_PP_MCONDICION"
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
    Private Function EventHandler_COMBO_SELECT(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oBOB As SAPbobsCOM.SBObob = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        Dim sSQL As String = ""
        EventHandler_COMBO_SELECT = False
        Dim sUDO As String = "" : Dim sTabla As String = ""
        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                If pVal.ItemChanged And pVal.ItemUID = "19" Then
                    With oForm.DataSources.DBDataSources.Item("@PP_MCONDICION")
                        sSQL = "SELECT ""COLUMN_NAME"", ""COLUMN_NAME"" ""Descripción"" FROM TABLE_COLUMNS WHERE ""TABLE_NAME""='@" & .GetValue("U_PP_TCONF", .Offset).Trim & "' and ""SCHEMA_NAME""='" & objGlobal.compañia.CompanyDB & "' "
                        sSQL &= " And ""COLUMN_NAME"" in ('Code', 'DocEntry', 'DocNum')  "
                        sSQL &= " UNION ALL "
                        sSQL &= "SELECT Concat('U_',""AliasID"") ""AliasID"",Concat(Concat(""AliasID"",' - '),""Descr"") ""Descripción"" "
                        sSQL &= " From ""CUFD"" Where ""TableID""='@" & .GetValue("U_PP_TCONF", .Offset).Trim & "' "
                    End With
                    PP_UTILITIES.EO.Formularios.MyFormFunctions.Init_Combo(objGlobal, oForm, "16_U_C", sSQL)
                End If
                If pVal.ItemUID = "13_U_C" Then
                    sUDO = CType(oForm.Items.Item("13_U_C").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                    If sUDO = "-1" Then
                        sSQL = "Select '-1','None' FROM DUMMY"
                        objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("15_U_C").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                        CType(oForm.Items.Item("15_U_C").Specific, SAPbouiCOM.ComboBox).ExpandType = BoExpandType.et_ValueDescription
                        oForm.DataSources.DBDataSources.Item("@PP_MCONDICION").SetValue("U_PP_CAMPO", oForm.DataSources.DBDataSources.Item("@PP_MCONDICION").Offset, "-1")
                        CType(oForm.Items.Item("14_U_E").Specific, SAPbouiCOM.EditText).Value = "Ninguna"
                    Else
                        sSQL = "SELECT ""TableName"" FROM ""OUDO"" WHERE ""Code""='" & sUDO & "' "
                        sTabla = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                        CType(oForm.Items.Item("14_U_E").Specific, SAPbouiCOM.EditText).Value = sTabla

                        sSQL = "SELECT ""COLUMN_NAME"", ""COLUMN_NAME"" ""Descripción"" FROM TABLE_COLUMNS WHERE ""TABLE_NAME""='@" & sTabla & "' and ""SCHEMA_NAME""='" & objGlobal.compañia.CompanyDB & "' "
                        sSQL &= " And ""COLUMN_NAME"" in ('Code', 'DocEntry', 'DocNum')  "
                        sSQL &= " UNION ALL "
                        sSQL &= "SELECT Concat('U_',""AliasID"") ""AliasID"",Concat(Concat(""AliasID"",' - '),""Descr"") ""Descripción"" "
                        sSQL &= " From ""CUFD"" Where ""TableID""='@" & sTabla & "' "
                        objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("15_U_C").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                        CType(oForm.Items.Item("15_U_C").Specific, SAPbouiCOM.ComboBox).ExpandType = BoExpandType.et_ValueDescription
                    End If

                    'sSQL = "SELECT ""COLUMN_NAME"", ""COLUMN_NAME"" ""Descripción"" FROM TABLE_COLUMNS WHERE ""TABLE_NAME""='@PP_OITM' and ""SCHEMA_NAME""='" & objGlobal.compañia.CompanyDB & "' "
                    'sSQL &= " And ""COLUMN_NAME"" in ('Code', 'DocEntry', 'DocNum')  "
                    'sSQL &= " UNION ALL "
                    'sSQL &= "SELECT Concat('U_',""AliasID"") ""AliasID"",Concat(Concat(""AliasID"",' - '),""Descr"") ""Descripción"" "
                    'sSQL &= " From ""CUFD"" Where ""TableID""='@PP_OITM' "
                    'objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("16_U_C").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                    'CType(oForm.Items.Item("16_U_C").Specific, SAPbouiCOM.ComboBox).ExpandType = BoExpandType.et_ValueDescription
                ElseIf pVal.ItemUID = "16_U_C" Then
                    Dim sValor As String = ""
                    oBOB = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge), SAPbobsCOM.SBObob)
                    'Buscamos campos de la tabla
                    With oForm.DataSources.DBDataSources.Item("@PP_MCONDICION")
                        oRs = oBOB.GetTableFieldList("@" & .GetValue("U_PP_TCONF", .Offset).Trim)
                    End With

                    If CType(oForm.Items.Item("16_U_C").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                        sValor = CType(oForm.Items.Item("16_U_C").Specific, SAPbouiCOM.ComboBox).Selected.Value
                    End If
                    While oRs.EoF = False
                        If sValor = oRs.Fields.Item("FieldName").Value.ToString Then
                            Dim sTipo As String = "" : Dim sValidValue As String = "" : Dim sTipoCampo As String = ""
                            sTipo = oRs.Fields.Item("FieldType").Value.ToString
                            sValidValue = oRs.Fields.Item("IsValidValues").Value.ToString()
                            Select Case sTipo
                                Case "0" : sTipoCampo = "0 - Alfanumérico"
                                Case "1" : sTipoCampo = "1 - Fecha"
                                Case "2" : sTipoCampo = "2 - Numérico"
                                Case "3" : sTipoCampo = "3 - Precio"
                                Case "4" : sTipoCampo = "4 - Cantidad"
                                Case "5" : sTipoCampo = "5 - Importe"
                            End Select
                            Select Case sValidValue
                                Case "1" : sTipoCampo &= " - 1 - IsValidValues"
                            End Select
                            CType(oForm.Items.Item("17_U_E").Specific, SAPbouiCOM.EditText).Value = sTipoCampo
                            Exit While
                        Else
                            CType(oForm.Items.Item("17_U_E").Specific, SAPbouiCOM.EditText).Value = ""
                        End If
                        oRs.MoveNext()
                    End While
                End If

            End If

            EventHandler_COMBO_SELECT = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oBOB, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function


    Private Function EventHandler_Form_Visible(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Dim sSQL As String = ""

        EventHandler_Form_Visible = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then

                sSQL = "SELECT * FROM (SELECT '-1' ""Code"",' None' ""Name"" FROM DUMMY UNION ALL SELECT T0.""Code"" ""Code"", T0.""Name"" ""Name"" FROM ""OUDO"" T0) T0 ORDER BY T0.""Name"" "
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("13_U_C").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                CType(oForm.Items.Item("13_U_C").Specific, SAPbouiCOM.ComboBox).ExpandType = BoExpandType.et_ValueDescription
                If oForm.Mode <> BoFormMode.fm_ADD_MODE Then
                    oForm.Items.Item("0_U_E").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False)
                Else
                    oForm.Items.Item("0_U_E").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_True)
                End If

                PP_UTILITIES.EO.Formularios.MyFormFunctions.Init_Combo(objGlobal, oForm, "19", "SELECT T0.""TableName"", T0.""Descr"" FROM OUTB T0 WHERE T0.""TableName"" LIKE 'PP_OITM%'")

            End If

            EventHandler_Form_Visible = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
#End Region
End Class
