Imports EXOCryptorLib


Module MyGlobals
    Public VersionSAP As Integer = 0
    Public CmpSrv As SAPbobsCOM.CompanyService = Nothing
    Public ClassName As String = "EXO_PRT_PADILLA_MAESTROS."
    Public NamespaceClass As String = "PuertasPadilla.Maestros."
    Public PathDLL As String = ""

    '******* CONSTANTE DE MENÚ ***********************
    Public Const _SAP_MENU_DELETE As String = "1283"
    Public Const _SAP_MENU_FIND As String = "1281"
    Public Const _SAP_MENU_NEW As String = "1282"
    Public Const _SAP_MENU_FIRST As String = "1290"
    Public Const _SAP_MENU_LAST As String = "1291"
    Public Const _SAP_MENU_NEXT As String = "1288"
    Public Const _SAP_MENU_PREVIOUS As String = "1289"
    Public Const _SAP_MENU_CLOSE As String = "1286"
    Public Const _SAP_MENU_DUPLICATE As String = "1287"
    Public Const _SAP_MENU_REFRESH As String = "1304"
    Public Const _SAP_MENU_CANCEL As String = "1284"
    Public Const _SAP_MENU_FILTER As String = "4870"
    Public Const _SAP_MENU_ORDER As String = "4869"
    Public Const _SAP_MENU_PARAM As String = "5890"
    Public Const _SAP_MENU_ADD_LINEA As String = "1292"
    Public Const _SAP_MENU_DEL_LINEA As String = "1293"
    Public Const _SAP_MENU_CLOSE_LINEA As String = "1299"
    Public Const _SAP_MENU_DUPLICATE_LINEA As String = "1294"
    Public Const _SAP_MENU_REOPEN_LINEA As String = "1312"

    Public GetError As String = ""
    'Public cryptor As Cryptor = Nothing
    'Public GetTranslator As EO.Genericos.MyMessages = Nothing

    Public Function GetParamsXsjs(ByVal NodeXsjs As Xml.XmlNode) As String
        Dim Idx As Integer = 0
        Dim Params As String = ""
        Try
            For Idx = 0 To NodeXsjs.ChildNodes.Count - 1
                If Params.Length > 0 Then
                    Params = Params & "&"
                Else
                    Params = "?"
                End If
                Params = Params & NodeXsjs.ChildNodes.Item(Idx).Attributes("Property").Value.Trim() & "="
                Params = Params & NodeXsjs.ChildNodes.Item(Idx).Attributes("Value").Value.Trim()
            Next
            Return Params
        Catch ex As Exception
            Return ""
        End Try
    End Function
End Module

