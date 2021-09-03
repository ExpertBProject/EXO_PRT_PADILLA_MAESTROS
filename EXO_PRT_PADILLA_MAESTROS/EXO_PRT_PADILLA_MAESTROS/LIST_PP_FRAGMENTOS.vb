Imports System.Xml
Imports SAPbouiCOM
Imports System.Reflection
Public Class LIST_PP_FRAGMENTOS
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        cargamenu()
        If actualizar Then
            'cargaCampos()
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
#End Region
#Region "Eventos"
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then

            Else

                Select Case infoEvento.MenuUID
                    Case "PP-MnMFRG"
                        If CargarForm() = False Then
                            Exit Function
                        End If
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
    Public Function CargarForm() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_UIAPI.EXO_XML(objGlobal)

        CargarForm = False

        Try
            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "LIST_PP_FRAGMENTOS.xml")

            Try
                oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                    Exit Function
                End If
            End Try
            Load_Grid(objGlobal, oForm)

            'Llamamos a la función de inicialización
            If Init_Load(objGlobal, oForm) = False Then
                oForm.Items.Item("2").Click(BoCellClickType.ct_Regular)
                Return False
            End If

            CargarForm = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function Init_Load(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oform As SAPbouiCOM.Form) As Boolean

        Dim _MyActions As String() = {"N:Nuevo", "E:Editar", "D:Eliminar", "R:Refrescar"}

        Try
            'Cargamos las acciones que dispone el botón de acción.
            CType(oform.Items.Item("B1").Specific, SAPbouiCOM.ButtonCombo).ExpandType = BoExpandType.et_DescriptionOnly
            CType(oform.Items.Item("B1").Specific, SAPbouiCOM.ButtonCombo).Caption = "Acciones"
            'Cargo las opciones
            For i = 0 To _MyActions.Length - 1
                CType(oform.Items.Item("B1").Specific, SAPbouiCOM.ButtonCombo).ValidValues.Add(_MyActions(i).Split(CChar(":"))(0), _MyActions(i).Split(CChar(":"))(1))
            Next
            oform.DataSources.UserDataSources.Item("EO_ACTION").Value = "Acciones"

            Return Load_Grid(objGlobal, oform)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            If objGlobal.tipoCliente = EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente.Clasico Then
            End If
        End Try
    End Function
    Public Shared Function Load_Grid(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oform As SAPbouiCOM.Form) As Boolean

        Dim SQL As String = ""
        Dim oGrid As SAPbouiCOM.Grid = Nothing
        Dim i As Integer = 0
        Try
            oform.Freeze(True)
            oGrid = CType(oform.Items.Item("G1").Specific, SAPbouiCOM.Grid)
            SQL = "SELECT ""FRG"".""Code"" ""Código"", ""FRG"".""Name"" ""Nombre"", ""FRG"".""U_PP_ENTIDAD"" ""Tipo"", ""UF"".""Descr"" ""Descripción"", ""FRG"".""U_PP_CANTIDAD"" ""Cantidad"", ""FRG"". ""U_PP_ACTIVO"" ""Activo"" "
            SQL &= " FROM ""@PP_FRAGMENTOS"" ""FRG"" "
            SQL &= " LEFT JOIN ""UFD1"" ""UF"" ON ""FRG"".""U_PP_ENTIDAD""=""UF"".""FldValue"" "
            SQL &= " LEFT JOIN ""CUFD"" ""UT"" ON ""UT"".""FieldID""=""UF"".""FieldID"" And ""UT"".""TableID""=""UF"".""TableID"" "
            SQL &= " WHERE ""UT"".""TableID""='ORST' and ""UT"".""AliasID""='PP_TIPO'"
            oGrid.DataTable.ExecuteQuery(SQL)
            For i = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item(i).TitleObject.Sortable = True
                oGrid.Columns.Item(i).Editable = False
            Next
            oGrid.Columns.Item("Activo").Type = BoGridColumnType.gct_CheckBox
            oGrid.Columns.Item("Cantidad").RightJustified = True
            If oGrid.DataTable.IsEmpty = True Then
                CType(oform.Items.Item("L2").Specific, SAPbouiCOM.StaticText).Caption = "Registros encontrado 0"
            Else
                CType(oform.Items.Item("L2").Specific, SAPbouiCOM.StaticText).Caption = "Registros encontrado " & oGrid.DataTable.Rows.Count.ToString
            End If
            oGrid.AutoResizeColumns()


            Return True
        Catch ex As Exception
            oform.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            oform.Freeze(False)
            If objGlobal.SBOApp.ClientType = BoClientType.ct_Desktop Then
                EXO_CleanCOM.CLiberaCOM.FormGrid(oGrid)
            End If

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
                    Case "LIST_PP_FRAGMENTOS"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "LIST_PP_FRAGMENTOS"
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
                        Case "LIST_PP_FRAGMENTOS"
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
                        Case "LIST_PP_FRAGMENTOS"
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
                        Case "LIST_PP_FRAGMENTOS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "LIST_PP_FRAGMENTOS"
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
    Private Function Lista_Tratar_Registro(ByRef ObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oform As SAPbouiCOM.Form) As Boolean

        Dim oGrid As SAPbouiCOM.Grid = Nothing
        Dim ocombo As SAPbouiCOM.ButtonCombo = Nothing
        Dim Accion As String = ""
        Dim Row As Integer = 0

        Dim ID As String = ""

        Try
            'Acciones:
            'N: Nuevo
            'E: Editar
            'D: Eliminar
            'R: Refrescar

            'EO.Formularios.MyFormFunctions.Freeze(ObjGlobal, oform)

            'Validar selección
            oGrid = CType(oform.Items.Item("G1").Specific, SAPbouiCOM.Grid)
            ocombo = CType(oform.Items.Item("B1").Specific, SAPbouiCOM.ButtonCombo)
            If ocombo.Selected Is Nothing Then
                Return False
            End If
            Accion = ocombo.Selected.Value.Trim

            Select Case Accion
                Case "N" 'Nuevo
                    ObjGlobal.SBOApp.OpenForm(BoFormObjectEnum.fo_UserDefinedObject, "PP_FRAGMENTOS", "")
                Case "E" 'Editar
                    Try
                        Row = oGrid.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)
                    Catch ex As Exception
                        Row = -1
                    End Try
                    If Row < 0 Then
                        ObjGlobal.SBOApp.MessageBox("Debe seleccionar una línea!")
                        Return False
                    End If
                    ID = oGrid.DataTable.GetValue("Código", oGrid.GetDataTableRowIndex(Row)).ToString.Trim
                    'Cargamos UDO
                    ObjGlobal.SBOApp.OpenForm(BoFormObjectEnum.fo_UserDefinedObject, "PP_FRAGMENTOS", ID)
                Case "D" 'Eliminar
                    Try
                        Row = oGrid.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)
                    Catch ex As Exception
                        Row = -1
                    End Try
                    If Row < 0 Then
                        ObjGlobal.SBOApp.MessageBox("Debe seleccionar una línea!")
                        Return False
                    End If
                    ID = oGrid.DataTable.GetValue("Código", oGrid.GetDataTableRowIndex(Row)).ToString.Trim
                    If Eliminar_Registro(ObjGlobal, ID) = False Then
                        Return False
                    End If
                    oform.Freeze(True)
                    Load_Grid(ObjGlobal, oform)
                Case "R" 'Refrescar
                    Return Load_Grid(ObjGlobal, oform)
            End Select
            Return True
        Catch ex As Exception
            oform.Freeze(False)
            ObjGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            oform.DataSources.UserDataSources.Item("EO_ACTION").Value = "Acciones"
            oform.Freeze(False)
            If ObjGlobal.tipoCliente = EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente.Clasico Then
                EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(ocombo, Object))
                EXO_CleanCOM.CLiberaCOM.FormGrid(oGrid)
            End If
        End Try
    End Function
    Public Function Eliminar_Registro(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByVal ID As String) As Boolean

        Dim oDI_COM As EXO_DIAPI.EXO_UDOEntity = Nothing 'Instancia del UDO para Insertar datos
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim SQL As String = ""
        Try
            If objGlobal.SBOApp.MessageBox("¿Está seguro que quiere eliminar el registro?", 2, "Si", "No") = 2 Then
                Return False
            End If
            oDI_COM = New EXO_DIAPI.EXO_UDOEntity(objGlobal.refDi.comunes, "PP_FRAGMENTOS")

            If oDI_COM.GetByKey(ID) = False Then
                objGlobal.SBOApp.MessageBox("El registro con clave " & ID & " no se encontró en Fragmentos de la BBDD!")
                Return False
            Else
                If oDI_COM.UDO_Delete(ID) = False Then
                    objGlobal.SBOApp.MessageBox("El registro con clave " & ID & " no se pudo eliminar de Fragmentos:" & oDI_COM.GetLastError)
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            If objGlobal.tipoCliente = EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente.Clasico Then
            End If
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function EventHandler_COMBO_SELECT_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim sSQL As String = ""
        EventHandler_COMBO_SELECT_After = False
        Try

            If pVal.ItemChanged = True Then
                Select Case pVal.ItemUID
                    Case "B1"
                        Return Lista_Tratar_Registro(objGlobal, oForm)
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

    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = "" : Dim oRs As SAPbobsCOM.Recordset = Nothing
        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)


            Select Case pVal.ItemUID
                Case "B1"
                    Return Lista_Tratar_Registro(objGlobal, oForm)
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
#End Region
End Class
