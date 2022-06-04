Imports System.Xml
Imports System.IO
Imports OfficeOpenXml
Imports SAPbouiCOM
Public Class EXO_CEXCEL
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        cargamenu()

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
                    Case "PP-MnCFRG"
                        If CargarForm() = False Then
                            Return False
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
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_CEXCEL.srf")

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
            CargaComboFormato(oForm)

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
    Private Function CargaComboFormato(ByRef oForm As SAPbouiCOM.Form) As Boolean

        CargaComboFormato = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oForm.Freeze(True)
            sSQL = "(Select 'EXCEL' as ""Code"",'EXCEL' as ""Name"" FROM ""DUMMY"") "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Select("EXCEL", BoSearchKey.psk_ByValue)
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Por favor, antes de continuar, revise la parametrización.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
            CargaComboFormato = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function

    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_CEXCEL"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        GC.Collect()
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
                        Case "EXO_CEXCEL"
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
                        Case "EXO_CEXCEL"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_CEXCEL"
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
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sTipoArchivo As String = ""
        Dim sArchivoOrigen As String = ""
        Dim sArchivo As String = ""
        Dim sNomFICH As String = ""
        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            'Comprobamos el directorio de trabajo
            sArchivo = objGlobal.refDi.SQL.sqlStringB1("SELECT T0.""U_EXO_PATH"" FROM ""@EXO_OGEN""  T0")
            sArchivo &= "\08.Historico\DOC_CARGADOS\FRAGMENTOS\"


            'Comprobamos que exista el directorio y sino, lo creamos
            If System.IO.Directory.Exists(sArchivo) = False Then
                System.IO.Directory.CreateDirectory(sArchivo)
            End If

            Select Case pVal.ItemUID
                Case "btn_Carga"
                    If objGlobal.SBOApp.MessageBox("¿Está seguro que quiere tratar el fichero seleccionado?", 1, "Sí", "No") = 1 Then
                        sArchivoOrigen = CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value
                        sNomFICH = IO.Path.GetFileName(sArchivoOrigen)
                        sArchivo = sArchivo & sNomFICH
                        'Hacemos copia de seguridad para tratarlo
                        Copia_Seguridad(sArchivoOrigen, sArchivo)
                        oForm.Items.Item("btn_Carga").Enabled = False
                        objGlobal.SBOApp.StatusBar.SetText("Cargando datos de Fragmentos ... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        'Ahora abrimos el fichero para tratarlo
                        TratarFichero(sArchivo, CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString, oForm)
                        oForm.Freeze(False)
                        objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log para ver las operaciones realizadas.")
                        oForm.Items.Item("btn_Carga").Enabled = True
                    Else
                        objGlobal.SBOApp.StatusBar.SetText("Se ha cancelado el proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
                Case "btn_Fich"
                    'Cargar Fichero para leer
                    If CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString <> "--" Then
                        sTipoArchivo = "Libro de Excel|*.xlsx|Excel 97-2003|*.xls"

                        'Tenemos que controlar que es cliente o web
                        If objGlobal.SBOApp.ClientType = SAPbouiCOM.BoClientType.ct_Browser Then
                            sArchivoOrigen = objGlobal.SBOApp.GetFileFromBrowser() 'Modificar
                        Else
                            'Controlar el tipo de fichero que vamos a abrir según campo de formato
                            sArchivoOrigen = objGlobal.funciones.OpenDialogFiles("Abrir archivo como", sTipoArchivo)
                        End If

                        If Len(sArchivoOrigen) = 0 Then
                            CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value = ""
                            objGlobal.SBOApp.MessageBox("Debe indicar un archivo a importar.")
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Debe indicar un archivo a importar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            oForm.Items.Item("btn_Carga").Enabled = False
                            Exit Function
                        Else
                            CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value = sArchivoOrigen
                            oForm.Items.Item("btn_Carga").Enabled = True
                        End If
                    Else
                        objGlobal.SBOApp.MessageBox("No ha seleccionado el formato a importar." & ChrW(10) & ChrW(13) & " Antes de continuar Seleccione un formato de los que se ha creado en la parametrización.")
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - No ha seleccionado el formato a importar." & ChrW(10) & ChrW(13) & " Antes de continuar Seleccione un formato de los que se ha creado en la parametrización.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Active = True
                    End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Sub Copia_Seguridad(ByVal sArchivoOrigen As String, ByVal sArchivo As String)
        'Comprobamos el directorio de copia que exista
        Dim sPath As String = ""
        sPath = IO.Path.GetDirectoryName(sArchivo)
        If IO.Directory.Exists(sPath) = False Then
            IO.Directory.CreateDirectory(sPath)
        End If
        If IO.File.Exists(sArchivo) = True Then
            IO.File.Delete(sArchivo)
        End If
        'Subimos el archivo
        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Comienza la Copia de seguridad del fichero - " & sArchivoOrigen & " -.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        If objGlobal.SBOApp.ClientType = BoClientType.ct_Browser Then
            Dim fs As FileStream = New FileStream(sArchivoOrigen, FileMode.Open, FileAccess.Read)
            Dim b(CInt(fs.Length() - 1)) As Byte
            fs.Read(b, 0, b.Length)
            fs.Close()
            Dim fs2 As New System.IO.FileStream(sArchivo, IO.FileMode.Create, IO.FileAccess.Write)
            fs2.Write(b, 0, b.Length)
            fs2.Close()
        Else
            My.Computer.FileSystem.CopyFile(sArchivoOrigen, sArchivo)
        End If
        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Copia de Seguridad realizada Correctamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    End Sub
    Private Sub TratarFichero(ByVal sArchivo As String, ByVal sTipoArchivo As String, ByRef oForm As SAPbouiCOM.Form)
        Dim myStream As StreamReader = Nothing
        Dim Reader As XmlTextReader = New XmlTextReader(myStream)
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sExiste As String = "" ' Para comprobar si existen los datos
        Dim sDelimitador As String = ""
        Try
            Select Case sTipoArchivo
                Case "EXCEL"
                    TratarFichero_Excel(sArchivo, oForm)
                Case Else
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) -El tipo de fichero a importar no está contemplado. Avise a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    objGlobal.SBOApp.MessageBox("El tipo de fichero a importar no está contemplado. Avise a su Administrador.")
                    Exit Sub
            End Select
            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Se ha leido correctamente el fichero.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            oForm.Freeze(True)

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            Throw exCOM
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        Finally
            oForm.Freeze(False)
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub
    Private Sub TratarFichero_Excel(ByVal sArchivo As String, ByRef oForm As SAPbouiCOM.Form)
#Region "Variables"
        Dim pck As ExcelPackage = Nothing
        Dim sAnno As String = "" : Dim sPeriodo As String = "" : Dim sCodeFrag As String = "" : Dim sNameFrag As String = "" : Dim sMaestro As String = "" : Dim dCantidadCab As Double = 1
        Dim sItemCode As String = "" : Dim sItemName As String = "" : Dim dCant As Double = 0 : Dim sFormula As String = "" : Dim sOperacion As String = ""
        Dim sExiste As String = ""
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oDI_COM As EXO_DIAPI.EXO_UDOEntity = Nothing 'Instancia del UDO para Insertar datos
#End Region
        Try
            ' miramos si existe el fichero y cargamos
            If File.Exists(sArchivo) Then
                Dim excel As New FileInfo(sArchivo)
                ExcelPackage.LicenseContext = LicenseContext.Commercial
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial
                pck = New ExcelPackage(excel)
                Dim workbook = pck.Workbook
                Dim worksheet = workbook.Worksheets.First()
                Dim startCell As ExcelCellAddress = worksheet.Dimension.Start
                Dim endCell As ExcelCellAddress = worksheet.Dimension.End
                For iRow As Integer = 2 To endCell.Row
                    sCodeFrag = worksheet.Cells(iRow, 1).Text
                    If sCodeFrag = "" Then
                        objGlobal.SBOApp.StatusBar.SetText("El registro Nº" & iRow & " tiene el cod de fragmento vacío. Se deja de leer el fichero. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Exit Sub
                    End If
                    sNameFrag = worksheet.Cells(iRow, 2).Text.Replace(",", "")
                    sMaestro = worksheet.Cells(iRow, 3).Text.Replace(",", "")
                    dCantidadCab = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, worksheet.Cells(iRow, 4).Text)
                    sItemCode = worksheet.Cells(iRow, 5).Text : sItemName = worksheet.Cells(iRow, 6).Text.Replace(",", "")
                    dCant = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, worksheet.Cells(iRow, 7).Text)
                    sFormula = worksheet.Cells(iRow, 8).Text
                    sOperacion = worksheet.Cells(iRow, 9).Text

                    'Buscamos si existe el Artículo
                    If sItemCode.Trim <> "" Then
                        sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OITM""", """ItemCode""", """ItemCode"" ='" & sItemCode.Trim & "' ")
                        If sItemName.Trim = "" And sExiste.Trim <> "" Then
                            sItemName = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OITM""", """ItemName""", """ItemCode"" ='" & sItemCode.Trim & "' ")
                        End If
                    ElseIf sItemName <> "" Then
                        sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OITM""", """ItemCode""", """ItemName"" ='" & sItemName.Trim & "' ")
                        If sExiste <> "" Then
                            sItemCode = sExiste
                        End If
                    Else
                        objGlobal.SBOApp.StatusBar.SetText("El registro Nº" & iRow & " No tiene marcado el artículo. Se deja de leer el fichero. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Exit Sub
                    End If
                    If sExiste = "" Then
                        objGlobal.SBOApp.StatusBar.SetText("El registro Nº" & iRow & " no encuentra el artículo " & sCodeFrag.Trim & " - " & sNameFrag.Trim & ". No puede continuar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Exit Sub
                    End If
                    If sOperacion.Trim <> "" Then
                        sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """@PP_OPERACION""", """U_PP_CODE""", """U_PP_CODE""='" & sOperacion.Trim & "' ")
                    End If
                    If sExiste = "" Then
                        objGlobal.SBOApp.StatusBar.SetText("El registro Nº" & iRow & " no encuentra la operación " & sOperacion.Trim & ". No puede continuar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Exit Sub
                    End If
                    'Ahora empezamos a generar el registro
                    oDI_COM = New EXO_DIAPI.EXO_UDOEntity(objGlobal.refDi.comunes, "PP_FRAGMENTOS") 'UDO 
                    Dim sCode As String = sCodeFrag
                    sSQL = "SELECT * FROM ""@PP_FRAGMENTOS"" WHERE ""Code""='" & sCode & "' "
                    oRs.DoQuery(sSQL)
                    If oRs.RecordCount > 0 Then
                        oDI_COM.GetByKey(sCode)
                        CrearCamposLíneas(oDI_COM, sCode, sItemCode, sItemName, dCant, sFormula, sOperacion)
                        If oDI_COM.UDO_Update = False Then
                            Throw New Exception("(EXO) - Error al añadir registro Nº" & iRow.ToString & ". " & oDI_COM.GetLastError)
                        End If
                    Else
                        oDI_COM.GetNew()
                        oDI_COM.SetValue("Code") = sCode
                        oDI_COM.SetValue("Name") = sNameFrag
                        oDI_COM.SetValue("U_PP_MAESTRO") = sMaestro
                        oDI_COM.SetValue("U_PP_CANT") = dCantidadCab
                        CrearCamposLíneas(oDI_COM, sCode, sItemCode, sItemName, dCant, sFormula, sOperacion)
                        If oDI_COM.UDO_Add = False Then
                            Throw New Exception("(EXO) - Error al añadir registro Nº" & iRow.ToString & ". " & oDI_COM.GetLastError)
                        End If
                    End If
                Next
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha encontrado el fichero excel a cargar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDI_COM, Object))
        End Try
    End Sub
    Private Sub CrearCamposLíneas(ByRef oDI_COM As EXO_DIAPI.EXO_UDOEntity, ByVal sCodigo As String, ByVal sItemCode As String, ByVal sItemName As String, ByVal dCant As Double,
                                       ByVal sFormula As String, ByVal sOperacion As String)
        Try
            oDI_COM.GetNewChild("PP_FRAGMENTOSL")
            oDI_COM.SetValueChild("U_PP_CODART") = sItemCode
            oDI_COM.SetValueChild("U_PP_DESART") = sItemName
            oDI_COM.SetValueChild("U_PP_CANT") = dCant
            If sFormula.Trim <> "" Then
                oDI_COM.SetValueChild("U_EXO_FORM") = "Y"
                oDI_COM.SetValueChild("U_EXO_FRMLA") = sFormula.Trim
            End If
            oDI_COM.SetValueChild("U_PP_CODOP") = sOperacion
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region
End Class
