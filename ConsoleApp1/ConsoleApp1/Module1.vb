Imports System.IO
Imports Microsoft.VisualBasic
Imports System.Xml
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System.Text
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Net.Mime.MediaTypeNames
Imports System.Collections
Imports System.Data.OleDb


Module Module1

    Private archivoExcelCarga As String
    Private archivoExcelCheck As String
    Private archivoExcelPlanif As String
    Private archivoExcelContacto As String
    Private archivoExcelItemSeguim As String
    Private archivoExcelCheck_inicial As String
    Private archivoExcelCheckListInicial As String
    Private archivoExcelRP As String

    Private idCategoria1 As String
    Private idCategoria2 As String
    Private idGantCubo As String
    Private idJefeProyecto As String
    Private idLider As String
    Private idIF As String
    Private idIFBackup As String
    Private initData As String
    Private iniciativa As String
    Private cod_iniciativa As String
    Private estado As String
    Private desc_ejecutiva As String
    Private entregable_proyecto As String
    Private canal_impactado_suc As String
    Private canal_impactado_int As String
    Private canal_impactado_aut As String
    Private canal_impactado_otr As String
    Private responsable_eje As String
    Private id_iniCod As String
    Private query_id_iniCod As String
    Private idGestores As String
    Private id_imp_Gestor As String
    Private id_paso_prod_Coment As String
    Private iD_iniciativa As String
    Private id_piloto As String
    Private id_despliegue As String
    Private id_historico_despliegue As String
    Private id_ambito As String
    Private id_tema_relevante As String
    Private id_historico_piloto As String
    Private idEspif As String
    Private id_pack_norm As String
    Private idCheckList As String
    Private idCheckList_h As String
    Private idInsertCheckList As String
    Private idplanific As String
    Private idplanific_h As String
    Private idContactoIF As String
    Private idContactoIF_h As String
    Private idContactoItemSeg As String
    Private idContactoItemSeg_Hist As String
    Private idReplanificacion As String
    Private idEvidencias As String
    Private idCheckList_inicial As String
    Private idCheckListInicial As String
    Private id_tema_rel_historico As String
    Private idReunionP As String
    Private conEncabezado As String = "conEncabezado"
    Private sinEncabezado As String = "sinEncabezado"
    Dim Withheader As String = "Si"
    Dim Withoutheader As String = "No"
    Dim idHistoricoPiloto As String = ""
    Dim idHistoricoDespliegue As String = ""
    Dim filePlanificacion As New List(Of String)
    Dim fileCheckList As New List(Of String)
    Dim fileAdminCompromisos As New List(Of String)

    Private appPath As String
    Private filePath As String
    Private dbStrConexion As String

    '********************************************************CONEXIONES*************************************************************************
    Private Function GetConnectionString(ByRef tipoConexion As Integer) As String

        Select Case tipoConexion
            Case 0
                Return dbStrConexion '"Driver={SQL Server};Server=167.28.65.134;Database=IMPLANTACION;Uid=usrSAF;Pwd=pwdSAF;"
            Case 1
                Return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & archivoExcelCarga & ";Extended Properties=""Excel 12.0 Xml;HDR=NO;IMEX=1"""
            Case 2
                Return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & archivoExcelCheck & ";Extended Properties=""Excel 12.0 Xml;HDR=NO"""
            Case 3
                Return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & archivoExcelPlanif & ";Extended Properties=""Excel 12.0 Xml;HDR=NO;IMEX=1"""
            Case 4
                Return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & archivoExcelContacto & ";Extended Properties=""Excel 12.0 Xml;HDR=NO;IMEX=1"""

            Case 5
                Return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & archivoExcelItemSeguim & ";Extended Properties=""Excel 12.0 Xml;HDR=NO;IMEX=1"""

            Case 6
                Return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & archivoExcelCheck_inicial & ";Extended Properties=""Excel 12.0 Xml;HDR=NO;IMEX=1"""

            Case 7
                Return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & archivoExcelCheckListInicial & ";Extended Properties=""Excel 12.0 Xml;HDR=NO;IMEX=1"""

            Case 8
                Return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & archivoExcelRP & ";Extended Properties=""Excel 12.0 Xml;HDR=NO;IMEX=1"""

        End Select



    End Function

    '****************************************************************************ARCHIVO LOGS*********************************************************************
    Private Sub Log(ByVal logMessage As String, _
                    ByVal tipoMessage As String, _
                    Optional ByVal tipoArchivo As String = "ext")

        If appPath = "" Then
            appPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)
            appPath = Replace(appPath, "file:\", "")
        End If

        If filePath = "" Then
            'filePath = String.Format(appPath & "\logs\logs_" & tipoArchivo & "_{0}.txt", DateTime.Today.ToString("yyyy-MM-dd"))
            filePath = String.Format(appPath & "\logs\logs_" & tipoArchivo & "_{0}_{1}.txt", DateTime.Today.ToString("yyyyMMdd"), DateTime.Now.ToString("hhmmss"))
        End If

        Using writer As New StreamWriter(filePath, True)
            'If File.Exists(filePath) And (tipoMessage = "error") Then
            If LCase(tipoMessage = "error") Then
                writer.WriteLine("Error -- " & DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") & " : " & logMessage & ".")
            Else
                writer.WriteLine(DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") & " : " & logMessage & ".")
                'writer.WriteLine("log de errores hoy dia:" & logMessage & ".")
            End If
        End Using

    End Sub

    '***************************************************************************************************************************'

    Private Function checkCategoria(ByVal nom As String,
                                    ByVal rsp As String,
                                    ByVal cod_Iniciativa As String,
                                    ByVal tblNom As String,
                                          Optional ByVal opt As Integer = 1) As String

        Dim dbConexion As Data.Odbc.OdbcConnection
        Dim dbcommand As Data.Odbc.OdbcCommand
        Dim dbdata As Data.Odbc.OdbcDataReader
        Dim dbconsulta As String = ""
        Dim dbresultados As String = ""


        If (nom <> "") Then

            'If opt = 0 Then dbconsulta = "select cat_ide from [dbo].[" & tblNom & "] where cat_nom='" & nom & "' and cat_due='" & rsp & "'"
            'If opt = 1 Then dbconsulta = "select cat2_ide from [dbo].[" & tblNom & "] where cat2_nom='" & nom & "' and cat2_rsp='" & rsp & "'"

            If opt = 0 Then dbconsulta = "select cat_ide from  [dbo].[" & tblNom & "] where cat_cod_ini='" & cod_Iniciativa & "' "
            If opt = 1 Then dbconsulta = "select cat2_ide from [dbo].[" & tblNom & "] where cat2_cod_ini='" & cod_Iniciativa & "' "



            Try
                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(dbconsulta, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                dbdata = dbcommand.ExecuteReader

                If dbdata.HasRows = True Then
                    dbdata.Read()
                    dbresultados = dbdata.Item(0).ToString

                    dbdata.Close()
                    dbConexion.Close()

                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing

                    Return dbresultados
                Else
                    If opt = 0 Then dbconsulta = "insert [dbo].[" & tblNom & "] (cat_cod_ini,cat_nom,cat_des,cat_tri,cat_due) values('" & cod_Iniciativa & "','" & nom & "','','','" & rsp & "')"
                    If opt = 1 Then dbconsulta = "insert [dbo].[" & tblNom & "] (cat2_cod_ini,cat2_nom,cat2_des,cat2_tri,cat2_rsp) values('" & cod_Iniciativa & "','" & nom & "','','','" & rsp & "')"

                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbconsulta, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbcommand.ExecuteNonQuery()

                    If opt = 0 Then dbconsulta = "select cat_ide from [dbo].[" & tblNom & "] where cat_nom='" & nom & "' and cat_due='" & rsp & "'"
                    If opt = 1 Then dbconsulta = "select cat2_ide from [dbo].[" & tblNom & "] where cat2_nom='" & nom & "' and cat2_rsp='" & rsp & "'"

                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbconsulta, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbdata = dbcommand.ExecuteReader

                    If dbdata.HasRows = True Then
                        dbdata.Read()
                        dbresultados = dbdata.Item(0).ToString

                        dbdata.Close()
                        dbConexion.Close()

                        dbConexion = Nothing
                        dbcommand = Nothing
                        dbdata = Nothing

                        Return dbresultados
                    Else
                        dbConexion = Nothing
                        dbcommand = Nothing
                        dbdata = Nothing

                        Return "0"
                    End If
                End If

            Catch ex As Exception
                Log("Se ha producido un error en la funcion checkCategoria " & ex.Message, "error")
                Console.WriteLine("Se ha producido un error " & ex.Message)
                Console.ReadLine()
                Return "0"
            End Try
        End If

    End Function

    '******************************************GANTT CUBO*******************************************************************'
    Private Function checkGantCubo(ByVal nom As String,
                                   ByVal cod_Iniciativa As String,
                                   ByVal tblNom As String) As String

        Dim dbConexion As Data.Odbc.OdbcConnection
        Dim dbcommand As Data.Odbc.OdbcCommand
        Dim dbdata As Data.Odbc.OdbcDataReader
        Dim dbconsulta As String = ""
        Dim dbresultados As String = ""
        Dim dbinsert_ini_gc As String = ""
        Dim dbquery As String = ""
        Dim dbRowCount As String = ""
        Dim nomb_aux As String = ""
        Dim cadEnter_nomb_gant() As String

        Try

            'consulto en la tabla imp_gc si el nombre esta repetido
            If (nom <> "") Then
                dbRowCount = "select COUNT(*) AS contador from [dbo].[" & tblNom & "] where gc_cod_ini='" & nom & "'"
                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(dbRowCount, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                dbdata = dbcommand.ExecuteReader

                If dbdata.HasRows = True Then
                    dbdata.Read()
                    dbresultados = dbdata.Item(0).ToString

                    dbdata.Close()
                    dbConexion.Close()

                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing
                End If

                If (dbresultados = "0") Then
                    cadEnter_nomb_gant = nom.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)

                    For i As Integer = 0 To cadEnter_nomb_gant.Length - 1

                        If (nom = "") Then
                            nomb_aux = "Sin Asignar"
                        Else
                            nomb_aux = cadEnter_nomb_gant(i)

                        End If


                        dbconsulta = "insert [dbo].[" & tblNom & "] (gc_cod_ini,gc_nombre_gant,gc_alc,gc_fec_crn,gc_fec_ini,gc_fec_ter,gc_est) values('" & cod_Iniciativa & "','" & nomb_aux & "','','','','','')"
                        'Debug.Print(dbconsulta)

                        dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                        dbcommand = New Data.Odbc.OdbcCommand(dbconsulta, dbConexion)
                        dbcommand.CommandType = CommandType.Text
                        dbConexion.Open()
                        dbcommand.ExecuteNonQuery()

                        'Busco El id que se genero con la inserción
                        dbquery = "select gc_ide from [dbo].[" & tblNom & "] where gc_cod_ini='" & cod_Iniciativa & "'"
                        'Debug.Print(dbquery)

                        dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                        dbcommand = New Data.Odbc.OdbcCommand(dbquery, dbConexion)
                        dbcommand.CommandType = CommandType.Text
                        dbConexion.Open()
                        dbdata = dbcommand.ExecuteReader

                        If dbdata.HasRows = True Then
                            dbdata.Read()
                            dbresultados = dbdata.Item(0).ToString
                            dbdata.Close()
                            dbConexion.Close()

                            'Inserto en la tabla imp_ini_gc
                            dbinsert_ini_gc = "INSERT INTO [dbo].[imp_ini_gc] " _
                           & "(ingc_ini_ide,ingc_gc_ide) " _
                           & "values('" & id_iniCod & "', '" & dbresultados & "' ) "

                            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                            dbcommand = New Data.Odbc.OdbcCommand(dbinsert_ini_gc, dbConexion)
                            dbcommand.CommandType = CommandType.Text
                            dbConexion.Open()
                            dbcommand.ExecuteNonQuery()

                            dbConexion.Close()
                            dbdata.Close()

                            dbConexion = Nothing
                            dbcommand = Nothing
                            dbdata = Nothing

                        End If
                    Next

                    Log("Se Inserto con Exito en la tabla IMP INI GC de la funcion checkGantCubo", "exito")

                    Console.WriteLine("Se Inserto con Exito en la tabla IMP INI GC... ")
                End If

            End If



            'Else
            '        dbConexion = Nothing
            '        dbcommand = Nothing
            '        dbdata = Nothing
            '        Return "0"



        Catch ex As Exception
            Log("Se ha producido un error en la funcion checkGantCubo " & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try

salir:
        Return 0

    End Function

    '******************************************************************************************************************************************************
    Private Function check_JP_LDR_IF(ByVal nom As String,
                                     ByVal cod_iniciativa As String,
                                     ByVal nombre As String,
                                     ByVal apellido As String,
                                     ByVal tblNom As String,
                                           Optional ByVal opt As Integer = 1) As String

        Dim dbConexion As Data.Odbc.OdbcConnection
        Dim dbcommand As Data.Odbc.OdbcCommand
        Dim dbdata As Data.Odbc.OdbcDataReader
        Dim dbconsulta As String = ""
        Dim dbinsert As String = ""
        Dim dbresultados As String = ""
        Dim partes() As String
        Dim cadNoasignada As String = " "
        Dim dbRowCount As String = ""
        Dim index As Integer = 0
        Dim dbControw As Integer = 0

        'If nom = "" Or InStr(1, nom, ".", CompareMethod.Text) = 0 Then GoTo salir
        If nom = "" Or InStr(1, nom, ".", CompareMethod.Text) = 0 Then

            dbRowCount = "select COUNT(" & nombre & ") AS contador from [dbo].[" & tblNom & "]  WHERE " & nombre & " = '" & cadNoasignada & "'  "
            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(dbRowCount, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader

            If dbdata.HasRows = True Then
                dbdata.Read()
                dbControw = dbdata.Item(0).ToString

                'If (dbControw = 0) Then
                If opt = 0 Then dbinsert = "insert [dbo].[" & tblNom & "] (jp_cod_ini,jp_nom,jp_pat,jp_mat,jp_als) values('" & cod_iniciativa & "','" & cadNoasignada & "','" & cadNoasignada & "','" & cadNoasignada & "','')"
                If opt = 1 Then dbinsert = "insert [dbo].[" & tblNom & "] (ldr_cod_ini,ldr_nom,ldr_pat,ldr_mat,ldr_als) values('" & cod_iniciativa & "','" & cadNoasignada & "','" & cadNoasignada & "','" & cadNoasignada & "','')"
                If opt = 2 Then dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom,if_pat,if_mat,if_als) values('" & cod_iniciativa & "','" & cadNoasignada & "','" & cadNoasignada & "','" & cadNoasignada & "','')"
                If opt = 3 Then dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom,if_pat,if_mat,if_als) values('" & cod_iniciativa & "','" & cadNoasignada & "','" & cadNoasignada & "','" & cadNoasignada & "','')"

                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(dbinsert, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                dbcommand.ExecuteNonQuery()
                dbConexion.Close()
                Return dbControw
                GoTo salir
                'Else
                GoTo salir
                'End If
            End If

        End If
        partes = Split(nom, ".")

        If UBound(partes) = 2 Then
            If opt = 0 Then dbconsulta = "select jp_ide from [dbo].[" & tblNom & "] where jp_nom='" & partes(0) & "' and jp_pat='" & partes(1) & "' and jp_mat='" & partes(2) & "'"
            If opt = 1 Then dbconsulta = "select ldr_ide from [dbo].[" & tblNom & "] where ldr_nom='" & partes(0) & "' and ldr_pat='" & partes(1) & "' and ldr_mat='" & partes(2) & "'"
            If opt = 2 Then dbconsulta = "select if_ide from [dbo].[" & tblNom & "] where if_nom='" & partes(0) & "' and if_pat='" & partes(1) & "' and if_mat='" & partes(2) & "'"
            If opt = 3 Then dbconsulta = "select if_ide from [dbo].[" & tblNom & "] where if_nom='" & partes(0) & "' and if_pat='" & partes(1) & "' and if_mat='" & partes(2) & "'"
        End If

        If UBound(partes) = 1 Then
            If opt = 0 Then dbconsulta = "select jp_ide from [dbo].[" & tblNom & "] where jp_nom='" & partes(0) & "' and jp_pat='" & partes(1) & "' and jp_mat=''"
            If opt = 1 Then dbconsulta = "select ldr_ide from [dbo].[" & tblNom & "] where ldr_nom='" & partes(0) & "' and ldr_pat='" & partes(1) & "' and ldr_mat=''"
            If opt = 2 Then dbconsulta = "select if_ide from [dbo].[" & tblNom & "] where if_nom='" & partes(0) & "' and if_pat='" & partes(1) & "' and if_mat=''"
            If opt = 3 Then dbconsulta = "select if_ide from [dbo].[" & tblNom & "] where if_nom='" & partes(0) & "' and if_pat='" & partes(1) & "' and if_mat=''"
        End If

        If UBound(partes) = 0 Then
            If opt = 0 Then dbconsulta = "select jp_ide from [dbo].[" & tblNom & "] where jp_nom='" & partes(0) & "' and jp_pat='' and jp_mat=''"
            If opt = 1 Then dbconsulta = "select ldr_ide from [dbo].[" & tblNom & "] where ldr_nom='" & partes(0) & "' and ldr_pat='' and ldr_mat=''"
            If opt = 2 Then dbconsulta = "select if_ide from [dbo].[" & tblNom & "] where if_nom='" & partes(0) & "' and if_pat='' and if_mat=''"
            If opt = 3 Then dbconsulta = "select if_ide from [dbo].[" & tblNom & "] where if_nom='" & partes(0) & "' and if_pat='' and if_mat=''"
        End If

        'Debug.Print(dbconsulta)

        Try
            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(dbconsulta, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader

            'If dbdata.HasRows = True Then
            '    dbdata.Read()
            '    dbresultados = dbdata.Item(0).ToString

            '    dbdata.Close()
            '    dbConexion.Close()

            '    dbConexion = Nothing
            '    dbcommand = Nothing
            '    dbdata = Nothing

            '    Return dbresultados
            'Else

            dbRowCount = "select COUNT(*) AS contador from [dbo].[" & tblNom & "]  WHERE " & nombre & " = '" & partes(0) & "' AND " & apellido & " = '" & partes(1) & "' "
            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(dbRowCount, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader

            If dbdata.HasRows = True Then
                dbdata.Read()
                dbRowCount = dbdata.Item(0).ToString
                'If (dbRowCount = 0) Then

                If UBound(partes) = 2 Then

                    If opt = 0 Then dbinsert = "insert [dbo].[" & tblNom & "] (jp_cod_ini,jp_nom,jp_pat,jp_mat,jp_als) values('" & cod_iniciativa & "','" & partes(0) & "','" & partes(1) & "','" & partes(2) & "','')"
                    If opt = 1 Then dbinsert = "insert [dbo].[" & tblNom & "] (ldr_cod_ini,ldr_nom,ldr_pat,ldr_mat,ldr_als) values('" & cod_iniciativa & "','" & partes(0) & "','" & partes(1) & "','" & partes(2) & "','')"
                    If opt = 2 Then dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom,if_pat,if_mat,if_als) values('" & cod_iniciativa & "','" & partes(0) & "','" & partes(1) & "','" & partes(2) & "','')"
                    If opt = 3 Then dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom,if_pat,if_mat,if_als) values('" & cod_iniciativa & "','" & partes(0) & "','" & partes(1) & "','" & partes(2) & "','')"

                    Log("Se inserto con exito en la tabla: " & tblNom, "exito")
                End If



                If UBound(partes) = 1 Then
                    If opt = 0 Then dbinsert = "insert [dbo].[" & tblNom & "] (jp_cod_ini,jp_nom,jp_pat,jp_mat,jp_als) values('" & cod_iniciativa & "','" & partes(0) & "','" & partes(1) & "','','')"
                    If opt = 1 Then dbinsert = "insert [dbo].[" & tblNom & "] (ldr_cod_ini,ldr_nom,ldr_pat,ldr_mat,ldr_als) values('" & cod_iniciativa & "','" & partes(0) & "','" & partes(1) & "','','')"
                    If opt = 2 Then dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom,if_pat,if_mat,if_als) values('" & cod_iniciativa & "','" & partes(0) & "','" & partes(1) & "','','')"
                    If opt = 3 Then dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom,if_pat,if_mat,if_als) values('" & cod_iniciativa & "','" & partes(0) & "','" & partes(1) & "','','')"

                    Log("Se inserto con exito en la tabla: " & tblNom, "exito")
                End If

                If UBound(partes) = 0 Then
                    If opt = 0 Then dbinsert = "insert [dbo].[" & tblNom & "] (jp_cod_ini,jp_nom,jp_pat,jp_mat,jp_als) values('" & cod_iniciativa & "','" & partes(0) & "','','','')"
                    If opt = 1 Then dbinsert = "insert [dbo].[" & tblNom & "] (ldr_cod_ini,ldr_nom,ldr_pat,ldr_mat,ldr_als) values('" & cod_iniciativa & "','" & partes(0) & "','','','')"
                    If opt = 2 Then dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom,if_pat,if_mat,if_als) values('" & cod_iniciativa & "','" & partes(0) & "','','','')"
                    If opt = 3 Then dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom,if_pat,if_mat,if_als) values('" & cod_iniciativa & "','" & partes(0) & "','','','')"

                    Log("Se inserto con exito en la tabla: " & tblNom, "exito")
                End If

                'Debug.Print(dbinsert)

                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(dbinsert, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                dbcommand.ExecuteNonQuery()

                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(dbconsulta, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                dbdata = dbcommand.ExecuteReader

                If dbdata.HasRows = True Then
                    dbdata.Read()
                    dbresultados = dbdata.Item(0).ToString

                    dbdata.Close()
                    dbConexion.Close()

                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing

                    Return dbresultados
                Else
                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing

                    Return "0"
                End If

            End If
            'End If
            ' End If
        Catch ex As Exception
            Log("Se ha producido un error en la Función check_JP_LDR_IF  " & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try

salir:
        Return 0

    End Function
    '*******************************************************************************************************************************'
    Private Function check_Esp_IF_Back(ByVal nom_esp As String,
                                       ByVal nombre_esp_back As String,
                                       ByVal cod_iniciativa As String,
                                       ByVal tblNom As String,
                                       Optional ByVal opt As Integer = 1) As String

        Dim dbConexion As Data.Odbc.OdbcConnection
        Dim dbcommand As Data.Odbc.OdbcCommand
        Dim dbdata As Data.Odbc.OdbcDataReader
        Dim dbconsulta As String = ""
        Dim dbinsert As String = ""
        Dim dbresultados As String = ""
        Dim partes_esp() As String
        Dim partes_esp_back() As String
        Dim cadNoasignada As String = "Sin Asignar"
        Dim dbRowCount As String = ""
        Dim index As Integer = 0
        Dim dbControw As Integer = 0

        'If nom = "" Or InStr(1, nom, ".", CompareMethod.Text) = 0 Then GoTo salir


        'dbRowCount = "select COUNT(" & nombre & ") AS contador from [dbo].[" & tblNom & "]  WHERE " & nombre & " = '" & cadNoasignada & "'  "
        'dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
        'dbcommand = New Data.Odbc.OdbcCommand(dbRowCount, dbConexion)
        'dbcommand.CommandType = CommandType.Text
        'dbConexion.Open()
        'dbdata = dbcommand.ExecuteReader

        'If dbdata.HasRows = True Then
        '    dbdata.Read()
        '    dbControw = dbdata.Item(0).ToString

        '    If (dbControw = 0) Then
        '        If opt = 0 Then dbinsert = "insert [dbo].[" & tblNom & "] (jp_cod_ini,jp_nom,jp_pat,jp_mat,jp_als) values('" & cod_iniciativa & "','" & cadNoasignada & "','" & cadNoasignada & "','" & cadNoasignada & "','')"
        '        If opt = 1 Then dbinsert = "insert [dbo].[" & tblNom & "] (ldr_cod_ini,ldr_nom,ldr_pat,ldr_mat,ldr_als) values('" & cod_iniciativa & "','" & cadNoasignada & "','" & cadNoasignada & "','" & cadNoasignada & "','')"
        '        If opt = 2 Then dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom,if_pat,if_mat,if_als) values('" & cod_iniciativa & "','" & cadNoasignada & "','" & cadNoasignada & "','" & cadNoasignada & "','')"
        '        If opt = 3 Then dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom,if_pat,if_mat,if_als) values('" & cod_iniciativa & "','" & cadNoasignada & "','" & cadNoasignada & "','" & cadNoasignada & "','')"

        '        dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
        '        dbcommand = New Data.Odbc.OdbcCommand(dbinsert, dbConexion)
        '        dbcommand.CommandType = CommandType.Text
        '        dbConexion.Open()
        '        dbcommand.ExecuteNonQuery()
        '        dbConexion.Close()
        '        Return dbControw
        '        GoTo salir
        '    Else
        '        GoTo salir
        '    End If
        'End If

        'End If
        partes_esp = Split(nom_esp, ".")
        partes_esp_back = Split(nombre_esp_back, ".")

        If partes_esp.Length = 1 And partes_esp_back.Length = 1 Then GoTo salir


        If UBound(partes_esp) = 2 Or UBound(partes_esp_back) = 2 Then
            dbconsulta = "select if_ide from [dbo].[" & tblNom & "] where if_cod_ini='" & cod_iniciativa & "'"
        End If

        If UBound(partes_esp) = 1 Or UBound(partes_esp_back) = 1 Then
            dbconsulta = "select if_ide from [dbo].[" & tblNom & "] where if_cod_ini='" & cod_iniciativa & "'"
        End If

        If UBound(partes_esp) = 0 Or UBound(partes_esp_back) = 0 Then
            dbconsulta = "select if_ide from [dbo].[" & tblNom & "] where if_cod_ini='" & cod_iniciativa & "'"
        End If

        'Debug.Print(dbconsulta)

        Try
            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(dbconsulta, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader

            If dbdata.HasRows = True Then
                dbdata.Read()
                dbresultados = dbdata.Item(0).ToString

                dbdata.Close()
                dbConexion.Close()

                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing

                Return dbresultados
            Else

                'dbRowCount = "select COUNT(*) AS contador from [dbo].[" & tblNom & "]  WHERE " & nombre & " = '" & partes(0) & "' AND " & apellido & " = '" & partes(1) & "' "
                'dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                'dbcommand = New Data.Odbc.OdbcCommand(dbRowCount, dbConexion)
                'dbcommand.CommandType = CommandType.Text
                'dbConexion.Open()
                'dbdata = dbcommand.ExecuteReader

                'If dbdata.HasRows = True Then
                '    dbdata.Read()
                '    dbRowCount = dbdata.Item(0).ToString
                'If (dbRowCount = 0) Then




                If UBound(partes_esp) = 2 And UBound(partes_esp_back) = 2 Then

                    dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom_esp_if,if_pat_esp_if,if_mat_esp_if,if_nom_esp_if_back,if_pat_esp_if_back,if_mat_esp_if_back) values('" & cod_iniciativa & "','" & partes_esp(0) & "','" & partes_esp(1) & "','" & partes_esp(2) & "','" & partes_esp_back(0) & "','" & partes_esp_back(1) & "','" & partes_esp_back(2) & "')"
                    Log("Se inserto con exito en la tabla: " & tblNom, "exito")
                End If

                If UBound(partes_esp) = 1 And UBound(partes_esp_back) = 1 Then

                    dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom_esp_if,if_pat_esp_if,if_mat_esp_if,if_nom_esp_if_back,if_pat_esp_if_back,if_mat_esp_if_back) values('" & cod_iniciativa & "','" & partes_esp(0) & "','" & partes_esp(1) & "','','" & partes_esp_back(0) & "','" & partes_esp_back(1) & "','')"
                    Log("Se inserto con exito en la tabla: " & tblNom, "exito")
                End If


                If UBound(partes_esp) = 2 And UBound(partes_esp_back) = 0 Then

                    dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom_esp_if,if_pat_esp_if,if_mat_esp_if,if_nom_esp_if_back,if_pat_esp_if_back,if_mat_esp_if_back) values('" & cod_iniciativa & "','" & partes_esp(0) & "','" & partes_esp(1) & "','" & partes_esp(2) & "','','','')"
                    Log("Se inserto con exito en la tabla: " & tblNom, "exito")
                End If


                If UBound(partes_esp) = 2 And UBound(partes_esp_back) = 1 Then

                    dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom_esp_if,if_pat_esp_if,if_mat_esp_if,if_nom_esp_if_back,if_pat_esp_if_back,if_mat_esp_if_back) values('" & cod_iniciativa & "','" & partes_esp(0) & "','" & partes_esp(1) & "','" & partes_esp(2) & "','" & partes_esp_back(0) & "','" & partes_esp_back(1) & "','')"
                    Log("Se inserto con exito en la tabla: " & tblNom, "exito")
                End If



                If UBound(partes_esp) = 1 And UBound(partes_esp_back) = 2 Then

                    dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom_esp_if,if_pat_esp_if,if_mat_esp_if,if_nom_esp_if_back,if_pat_esp_if_back,if_mat_esp_if_back) values('" & cod_iniciativa & "','" & partes_esp(0) & "','" & partes_esp(1) & "','','" & partes_esp_back(0) & "','" & partes_esp_back(1) & "','" & partes_esp_back(2) & "')"
                    Log("Se inserto con exito en la tabla: " & tblNom, "exito")
                End If


                If UBound(partes_esp) = 1 And UBound(partes_esp_back) = 0 Then

                    dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom_esp_if,if_pat_esp_if,if_mat_esp_if,if_nom_esp_if_back,if_pat_esp_if_back,if_mat_esp_if_back) values('" & cod_iniciativa & "','" & partes_esp(0) & "','" & partes_esp(1) & "','','','','')"
                    Log("Se inserto con exito en la tabla: " & tblNom, "exito")
                End If



                'Else
                '    If (partes_esp_back.Length = 2) Then

                '        dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom_esp_if,if_pat_esp_if,if_mat_esp_if,if_nom_esp_if_back,if_pat_esp_if_back,if_mat_esp_if_back) values('" & cod_iniciativa & "','" & partes_esp(0) & "','" & partes_esp(1) & "','','" & partes_esp_back(0) & "','" & partes_esp_back(1) & "','')"

                '        Log("Se inserto con exito en la tabla: " & tblNom, "exito")

                '    Else

                '        dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom_esp_if,if_pat_esp_if,if_mat_esp_if,if_nom_esp_if_back,if_pat_esp_if_back,if_mat_esp_if_back) values('" & cod_iniciativa & "','" & partes_esp(0) & "','" & partes_esp(1) & "','" & partes_esp(2) & "','" & partes_esp_back(0) & "','" & partes_esp_back(1) & "','" & partes_esp_back(2) & "')"

                '    End If
                'End If



                'If UBound(partes_esp) = 2 Then

                'If (partes_esp_back.Length = 1) Then

                '    dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom_esp_if,if_pat_esp_if,if_mat_esp_if,if_nom_esp_if_back,if_pat_esp_if_back,if_mat_esp_if_back) values('" & cod_iniciativa & "','" & partes_esp(0) & "','" & partes_esp(1) & "','" & partes_esp(2) & "','','','')"

                '    Log("Se inserto con exito en la tabla: " & tblNom, "exito")
                'Else
                '    If (partes_esp_back.Length = 2) Then

                '        dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom_esp_if,if_pat_esp_if,if_mat_esp_if,if_nom_esp_if_back,if_pat_esp_if_back,if_mat_esp_if_back) values('" & cod_iniciativa & "','" & partes_esp(0) & "','" & partes_esp(1) & "','','" & partes_esp_back(0) & "','" & partes_esp_back(1) & "','')"

                '        Log("Se inserto con exito en la tabla: " & tblNom, "exito")

                '    Else

                '        dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom_esp_if,if_pat_esp_if,if_mat_esp_if,if_nom_esp_if_back,if_pat_esp_if_back,if_mat_esp_if_back) values('" & cod_iniciativa & "','" & partes_esp(0) & "','" & partes_esp(1) & "','" & partes_esp(2) & "','" & partes_esp_back(0) & "','" & partes_esp_back(1) & "','" & partes_esp_back(2) & "')"

                '    End If
                'End If


                'If UBound(partes_esp) = 1 Then
                '    dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom_esp_if,if_pat_esp_if,if_mat_esp_if,if_nom_esp_if_back,if_pat_esp_if_back,if_mat_esp_if_back) values('" & cod_iniciativa & "','" & partes_esp(0) & "','" & partes_esp(1) & "','" & partes_esp(2) & "','" & partes_esp_back(0) & "','" & partes_esp_back(1) & "','" & partes_esp_back(2) & "')"
                '    Log("Se inserto con exito en la tabla: " & tblNom, "exito")
                'End If

                'If UBound(partes_esp) = 0 Then
                '    dbinsert = "insert [dbo].[" & tblNom & "] (if_cod_ini,if_nom_esp_if,if_pat_esp_if,if_mat_esp_if,if_nom_esp_if_back,if_pat_esp_if_back,if_mat_esp_if_back) values('" & cod_iniciativa & "','" & partes_esp(0) & "','" & partes_esp(1) & "','" & partes_esp(2) & "','" & partes_esp_back(0) & "','" & partes_esp_back(1) & "','" & partes_esp_back(2) & "')"

                '    Log("Se inserto con exito en la tabla: " & tblNom, "exito")
                'End If

                Debug.Print(dbinsert)

                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(dbinsert, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                dbcommand.ExecuteNonQuery()

                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(dbconsulta, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                dbdata = dbcommand.ExecuteReader

                If dbdata.HasRows = True Then
                    dbdata.Read()
                    dbresultados = dbdata.Item(0).ToString

                    dbdata.Close()
                    dbConexion.Close()

                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing

                    Return dbresultados
                Else
                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing

                    Return "0"
                End If

            End If
            'End If
                ' End If
        Catch ex As Exception
            Log("Se ha producido un error en la Función check_JP_LDR_IF  " & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try

salir:
        Return 0

    End Function







    '********************************************************************************************************************************************'
    Private Function check_Gestores_otro_Miembros(ByVal nom As String,
                                                 ByVal codIniciativa As String,
                                            ByVal tblNom As String,
                                           Optional ByVal opt As Integer = 1) As String
        Try
            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim dbconsulta As String = ""
            Dim dbinsert As String = ""
            Dim dbresultados As String = ""
            Dim partes() As String
            Dim nameEnter() As String
            Dim cadNombre() As String
            Dim Rol As String = ""
            Dim FullName As String = ""
            Dim ide_gst As String = ""
            Dim dbinsert_ini_gst As String = ""
            Dim dbinsert_ini_gc As String = ""
            Dim dbresultado_ID As String = ""
            Dim apellido_pat As String = ""
            Dim dbSQL As String = ""


            'Se coloca un enter para separar el o los registros que tenga la columna, y se guarda dentro de un array cada registro'
            nameEnter = nom.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            For i As Integer = 0 To nameEnter.Length - 1
                dbresultados = ""

                cadNombre = Split(nameEnter(i), ":")
                FullName = cadNombre(0) 'Le asigno a la variable la cadena que contiene nombre y apellidos que va a ser la 1era posición del array


                If (cadNombre.Length > 1) Then
                    Rol = cadNombre(1) 'Le asigno a la variable el rol que contiene esa cadena que esta en la 2da posición del array
                Else
                    Rol = "Rol sin asignar"
                End If



                partes = Split(FullName, ".")

                If FullName.Contains(".") Then
                    'Console.WriteLine(FullName.Substring(0, FullName.IndexOf(".")))

                    If partes.Length = 2 Or partes.Length = 1 Then

                        apellido_pat = " "
                    Else
                        apellido_pat = partes(2)
                    End If

                    dbSQL = "select COUNT(*) AS contador from [dbo].[" & tblNom & "] where gst_nom='" & partes(0) & "' AND gst_pat = '" & partes(1) & "' AND gst_cod_ini = '" & codIniciativa & "'"
                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbSQL, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbdata = dbcommand.ExecuteReader

                    If dbdata.HasRows = True Then
                        dbdata.Read()
                        dbresultados = dbdata.Item(0).ToString

                        dbdata.Close()
                        dbConexion.Close()

                        dbConexion = Nothing
                        dbcommand = Nothing
                        dbdata = Nothing
                    End If


                    If (dbresultados = "0") Then
                        'Se inserta en la tabla
                        dbinsert = "INSERT INTO [dbo].[" & tblNom & "] " _
                       & "(gst_ini_ide,gst_cod_ini,gst_nom,gst_pat,gst_mat,gst_als,gst_tip) " _
                       & "values('" & id_iniCod & "','" & codIniciativa & "','" & partes(0) & "','" & partes(1) & "','" & apellido_pat & "', '  ','  ') "

                        dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                        dbcommand = New Data.Odbc.OdbcCommand(dbinsert, dbConexion)
                        dbcommand.CommandType = CommandType.Text
                        dbConexion.Open()
                        dbcommand.ExecuteNonQuery()

                        dbConexion.Close()

                        'Luego de insertados los registros busco en la tabla los ID que se genero con la inserción'
                        'select gst_ide from [dbo].[" & tblNom & "] 

                        id_imp_Gestor = "select gst_ide from [dbo].[" & tblNom & "]  where gst_nom='" & partes(0) & "' and gst_pat='" & partes(1) & "' "
                        dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                        dbcommand = New Data.Odbc.OdbcCommand(id_imp_Gestor, dbConexion)
                        dbcommand.CommandType = CommandType.Text
                        dbConexion.Open()
                        dbdata = dbcommand.ExecuteReader

                        If dbdata.HasRows = True Then
                            dbdata.Read()
                            dbresultado_ID = dbdata.Item(0).ToString

                            'MsgBox(dbresultados)
                            'Luego que recupero el ID del registro que fue insertado, se procede a insertar ese ID en la siguiente tabla

                            dbinsert_ini_gst = "INSERT INTO [dbo].[imp_ini_gst] " _
                            & "(inigst_ini_ide,inigst_cod_ini,inigst_gst_ide,inigst_rol) " _
                            & "values('" & id_iniCod & "', '" & codIniciativa & "','" & dbresultado_ID & "','" & Rol & "' ) "

                            'Debug.Print(dbinsert_ini_gst)

                            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                            dbcommand = New Data.Odbc.OdbcCommand(dbinsert_ini_gst, dbConexion)
                            dbcommand.CommandType = CommandType.Text
                            dbConexion.Open()
                            dbcommand.ExecuteNonQuery()
                            'Console.WriteLine("Se Inserto con Exito en la tabla IMP INI GESTOR... ")
                        End If
                    End If
                End If
            Next
            Log("Se Inserto con Exito en la tabla IMP GESTOR en la Función check_Gestores_otro_Miembros", "exito")
            Console.WriteLine("Se Inserto con Exito en la tabla IMP GESTOR... ")
        Catch ex As Exception
            Log("Se ha producido un error en la Función check_Gestores_otro_Miembros", "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Return "0"
        End Try

salir:
        Return 0
    End Function

    '***********************************************************************************************************************************************'
    Private Function check_ini_esp_if(ByVal tblNom As String) As String

        Try
            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim dbconsulta As String = ""
            Dim dbinsert_imp_esp_if As String = ""
            Dim dbresultados As String = ""
            Dim dbRowCount As String = ""
            Dim strCmnd As String = "SELECT if_ide FROM [dbo].[imp_esp_if]"
            Dim index As Integer = 0
            Dim dbControw As Integer = 0


            dbRowCount = "select COUNT(*) AS contador from [dbo].[imp_esp_if]"
            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(dbRowCount, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader

            If dbdata.HasRows = True Then
                dbdata.Read()
                dbControw = dbdata.Item(0).ToString

                dbdata.Close()
                dbConexion.Close()

                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing
            End If


            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(strCmnd, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader

            'Do While index <= dbresultados
            Do Until index = dbControw

                If dbdata.HasRows = True Then
                    dbdata.Read()
                    dbresultados = dbdata.Item(0).ToString

                    dbinsert_imp_esp_if = "INSERT INTO [dbo].[" & tblNom & "] " _
                    & "(iei_ini_ide,iei_if_ide) " _
                    & "values('" & id_iniCod & "', '" & dbresultados & "' ) "


                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbinsert_imp_esp_if, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbcommand.ExecuteNonQuery()


                Else
                    dbdata.Close()
                    dbConexion.Close()

                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing

                End If
                index += 1

            Loop
            Log("Se Inserto con Exito en la tabla IMP-INI-ESP-IF en la Función check_ini_esp_if ", "exito")
            Console.WriteLine("Se Inserto con Exito en la tabla IMP-INI-ESP-IF. ")
        Catch ex As Exception
            Log("Se ha producido un error en la Función check_ini_esp_if ", "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try

salir:
        Return 0
    End Function

    '*******************************************************************************************************************************************'
    Private Function checkValuespasoProd_comment(cod_iniciativa As String,
                                                ByVal paso_prod_val As String,
                                                 ByVal coment_pap_val As String,
                                                 ByVal estado_pap_val As String,
                                                 ByVal tblNom As String) As String

        Try
            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim dbinsert As String = ""
            Dim dbdelete As String = ""
            Dim cadEnter_paso_prod() As String
            Dim cadEnter_coment_pap() As String
            Dim cadEnter_estado_pap() As String
            Dim fecha_paso_prod As String = ""
            Dim coment_paso_prod As String = ""
            Dim estado_paso_prod As String = ""
            Dim dbinsert_imp_ini_pap As String = ""
            Dim dbsql As String = ""
            Dim dbresult As String = ""

            cadEnter_paso_prod = paso_prod_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_coment_pap = coment_pap_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_estado_pap = estado_pap_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)


            dbsql = "select count(*)  FROM [dbo].[imp_ini_pap]" _
& "where pap_cod_ini = '" & cod_iniciativa & "' AND pap_fec_pap = '" & fecha_paso_prod & "'  AND pap_com_pap = '" & coment_paso_prod & "'"

            'Debug.Print(dbsql)

            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(dbsql, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader

            If dbdata.HasRows = True Then
                dbdata.Read()
                dbresult = dbdata.Item(0).ToString

                dbdata.Close()
                dbConexion.Close()

                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing
            End If

            If (dbresult = "0") Then

                If (cadEnter_paso_prod.Length = 0) And (cadEnter_coment_pap.Length = 0) And (cadEnter_estado_pap.Length = 0) Then

                    dbinsert_imp_ini_pap = "INSERT INTO [dbo].[imp_ini_pap] " _
                  & "(pap_ini_ide, pap_cod_ini,pap_fec_pap,pap_com_pap,pap_estado_pap) " _
                  & "values('" & id_iniCod & "', '" & cod_iniciativa & "','" & fecha_paso_prod & "','" & coment_paso_prod & "','" & estado_paso_prod & "'  ) "


                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbinsert_imp_ini_pap, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbcommand.ExecuteNonQuery()

                    dbConexion.Close()

                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing

                    GoTo salir

                End If
            End If


            For i As Integer = 0 To cadEnter_paso_prod.Length - 1

                fecha_paso_prod = cadEnter_paso_prod(i)

                If coment_pap_val <> "" Then
                    coment_paso_prod = cadEnter_coment_pap(i)
                Else
                    coment_paso_prod = ""
                End If

                If estado_pap_val <> "" Then
                    estado_paso_prod = cadEnter_estado_pap(i)
                Else
                    estado_paso_prod = ""
                End If


                If (dbresult = "0") Then

                    dbinsert_imp_ini_pap = "INSERT INTO [dbo].[imp_ini_pap] " _
                  & "(pap_ini_ide, pap_cod_ini,pap_fec_pap,pap_com_pap,pap_estado_pap) " _
                  & "values('" & id_iniCod & "', '" & cod_iniciativa & "','" & fecha_paso_prod & "','" & coment_paso_prod & "','" & estado_paso_prod & "'  ) "

                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbinsert_imp_ini_pap, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbcommand.ExecuteNonQuery()

                    dbConexion.Close()

                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing
                End If
            Next


            Log("Se Inserto con Exito en la tabla IMP INI PAP en la Función checkValuespasoProd_comment", "exito")
            Console.WriteLine("Se Inserto con Exito en la tabla IMP INI PAP.")
        Catch ex As Exception
            Log("Se ha producido un error en la Función checkValuespasoProd_comment " & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try

salir:
        Return 0

    End Function

    '****************************************Piloto Actual*******************************************
    Private Function checkValuespilotoFields(ByVal cod_ini_val As String,
                                              ByVal fecha_ini_val As String,
                                              ByVal fecha_fin_val As String,
                                              ByVal coment_val As String,
                                              ByVal estado_piloto_val As String,
                                              ByVal tblNom As String) As String
        Try

            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim dbinsert As String = ""
            Dim dbdelete As String = ""
            Dim cadEnter_fecha_ini() As String
            Dim cadEnter_fecha_fin() As String
            Dim cadEnter_coment() As String
            Dim cadEnter_estado_piloto() As String
            Dim dbinsert_ini_piloto_hist As String

            Dim dbinsert_imp_ini_pil As String = ""
            Dim fecha_ini_aux As String = ""
            Dim fecha_fin_aux As String = ""
            Dim coment_aux As String = ""
            Dim estado_aux As String = ""
            Dim stringSeparators() As String = {","c}
            Dim contador As Integer = 0
            Dim query_id_iniciativa As String = ""
            Dim valor_id_iniciativa As String = "0"
            Dim query_SQL As String = ""
            Dim Val As String = ""

            cadEnter_fecha_ini = fecha_ini_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_fecha_fin = fecha_fin_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_coment = coment_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_estado_piloto = estado_piloto_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)


            query_id_iniciativa = "  Select ini_ide FROM [dbo].[imp_iniciativa] WHERE ini_cod = '" & cod_ini_val & "'"
            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(query_id_iniciativa, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader


            If dbdata.HasRows = True Then
                dbdata.Read()
                valor_id_iniciativa = dbdata.Item(0).ToString

                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing

            End If


            dbinsert_ini_piloto_hist = "if exists(select top 1 pil_ini_ide from  [dbo].[imp_ini_pil] where pil_ini_ide='" & valor_id_iniciativa & "') " & _
         "begin " & _
         "declare @hoy datetime; " & _
         "set		@hoy=getdate(); " & _
         "insert [dbo].[imp_ini_pil_hist]( pilH_ini_ide,pilH_ini_cod_ini, pilH_fec_hist_inicio, pilH_fec_hist_fin, pilH_com, pilH_estado," & _
         "pilH_fec_act_pil) " & _
         "select	pil_ini_ide, pil_ini_cod_ini,pil_fec_ini, pil_fec_ter, pil_com, pil_est, @hoy " & _
         "from imp_ini_pil " & _
         "where	pil_ini_ide='" & valor_id_iniciativa & "' " & _
         "ORDER BY pil_ini_ide " &
         " " & _
         "delete from imp_ini_pil where pil_ini_ide='" & valor_id_iniciativa & "' " & _
         "end"

            'Console.WriteLine(dbinsert_ini_piloto_hist)

            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(dbinsert_ini_piloto_hist, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbcommand.ExecuteNonQuery()
            Console.WriteLine("Se Inserto con Exito en la tabla Historica de Piloto... ")
            Log("Se Inserto con Exito en la tabla Historica de Piloto con el codigo iniciativa: " & cod_ini_val, "exito")

            dbConexion.Close()

            dbConexion = Nothing
            dbcommand = Nothing
            dbdata = Nothing


            query_SQL = "  Select count(pil_ini_cod_ini) FROM [dbo].[imp_ini_pil] WHERE pil_ini_cod_ini = '" & cod_ini_val & "'"
            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(query_SQL, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader


            If dbdata.HasRows = True Then
                dbdata.Read()
                Val = dbdata.Item(0).ToString

                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing

            End If

            If (Val = "0") Then

                If (cadEnter_fecha_ini.Length = 0) And (cadEnter_fecha_fin.Length = 0) And (cadEnter_coment.Length = 0) Then

                    dbinsert_imp_ini_pil = "INSERT INTO [dbo].[" & tblNom & "] " _
                           & "(pil_ini_ide,pil_ini_cod_ini,pil_fec_ini,pil_fec_ter,pil_com,pil_est) " _
                           & "values('" & valor_id_iniciativa & "', '" & cod_ini_val & "','" & fecha_ini_aux & "','" & fecha_fin_aux & "', '" & coment_aux & "'," _
                           & " '" & estado_aux & "') "

                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbinsert_imp_ini_pil, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbcommand.ExecuteNonQuery()

                    dbConexion.Close()

                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing

                    GoTo salir

                End If



                For i As Integer = 0 To cadEnter_fecha_ini.Length - 1


                    If (cadEnter_fecha_ini.Length = 1) Then
                        fecha_ini_aux = "Por Confirmar"
                    Else
                        If (cadEnter_fecha_ini.Length > contador) Then

                            fecha_ini_aux = cadEnter_fecha_ini(i).Trim
                        Else
                            fecha_ini_aux = "Sin Datos"

                        End If
                    End If



                    If (cadEnter_fecha_fin.Length = 1) Then
                        fecha_fin_aux = "Por Confirmar"
                    Else
                        If (cadEnter_fecha_fin.Length > contador) Then

                            fecha_fin_aux = cadEnter_fecha_fin(i).Trim
                        Else
                            fecha_fin_aux = "Sin Datos"

                        End If

                    End If


                    'If (coment_val <> "") Then
                    '    coment_aux = cadEnter_coment(i)
                    'Else
                    '    coment_aux = "-"
                    'End If

                    If (cadEnter_coment.Length > contador) Then

                        coment_aux = cadEnter_coment(i).Trim
                    Else
                        coment_aux = ""

                    End If


                    'If (coment_historico_val <> "") Then
                    '    coment_historico_aux = cadEnter_coment_historico(i)

                    'Else
                    '    coment_historico_aux = "-"

                    'End If

                    If (cadEnter_estado_piloto.Length > contador) Then

                        estado_aux = cadEnter_estado_piloto(i).Trim
                    Else
                        estado_aux = ""

                    End If


                    'If (fecha_actual_piloto_val <> "") Then
                    '    fecha_actual_piloto_aux = cadEnter_fecha_actual_piloto(i)
                    'Else
                    '    fecha_actual_piloto_aux = "-"

                    'End If

                    contador += 1

                    dbinsert_imp_ini_pil = "INSERT INTO [dbo].[" & tblNom & "] " _
                            & "(pil_ini_ide,pil_ini_cod_ini,pil_fec_ini,pil_fec_ter,pil_com,pil_est) " _
                            & "values('" & valor_id_iniciativa & "', '" & cod_ini_val & "','" & fecha_ini_aux & "','" & fecha_fin_aux & "', '" & coment_aux & "'," _
                            & " '" & estado_aux & "') "

                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbinsert_imp_ini_pil, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbcommand.ExecuteNonQuery()

                    dbConexion.Close()

                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing

                Next
            End If

            Log("Se Inserto con Exito en la tabla Piloto con el codigo iniciativa: " & cod_ini_val, "exito")

            Console.WriteLine("Se Inserto con Exito en la Tabla Piloto con el codigo iniciativa: " & cod_ini_val)
        Catch ex As Exception
            Log("Se ha producido un error en la Función checkValuespilotoFields " & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try

salir:
        Return 0

    End Function

    '**************************************HISTORICO PILOTO************************************************************'
    Private Function pilotoHistoricoValues(ByVal cod_proyecto_val As String,
                                           ByVal fecha_hist_inicio_val As String,
                                           ByVal fecha_hist_fin_val As String,
                                           ByVal coment_hist_val As String,
                                           ByVal estado_hist_piloto_val As String,
                                           ByVal tblNom As String) As String
        Try

            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim dbinsert As String = ""
            Dim dbdelete As String = ""
            Dim cadEnter_fecha_hist_inicio() As String
            Dim cadEnter_fecha_hist_fin() As String
            Dim cadEnter_coment_hist() As String
            Dim cadEnter_estado_hist() As String
            Dim dbinsert_imp_pil_hist As String = ""
            Dim fecha_hist_ini_aux As String = ""
            Dim fecha_hist_fin_aux As String = ""
            Dim coment_hist_aux As String = ""
            Dim estado_hist_aux As String = ""
            Dim fecha_act_hist_aux As String = ""
            Dim contador As Integer = 0
            Dim insert_piloto As String = ""
            Dim query_id_iniciativa As String = ""
            Dim valor_id_iniciativa As String = ""
            Dim dbdeletepiloto As String = ""

            cadEnter_fecha_hist_inicio = fecha_hist_inicio_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_fecha_hist_fin = fecha_hist_fin_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_coment_hist = coment_hist_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_estado_hist = estado_hist_piloto_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)

            query_id_iniciativa = "  Select ini_ide FROM [dbo].[imp_iniciativa] WHERE ini_cod = '" & cod_proyecto_val & "'"
            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(query_id_iniciativa, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader


            If dbdata.HasRows = True Then
                dbdata.Read()
                valor_id_iniciativa = dbdata.Item(0).ToString

                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing

            End If


            If (valor_id_iniciativa = "0") Then

                If (cadEnter_fecha_hist_inicio.Length = 0) And (cadEnter_fecha_hist_fin.Length = 0) And (cadEnter_coment_hist.Length = 0) Then

                    dbinsert_imp_pil_hist = "INSERT INTO [dbo].[" & tblNom & "] " _
                       & "(pilH_ini_ide,pilH_ini_cod_ini,pilH_fec_hist_inicio,pilH_fec_hist_fin,pilH_com,pilH_estado,pilH_fec_act_pil) " _
                       & "values('" & valor_id_iniciativa & "', '" & cod_proyecto_val & "','" & fecha_hist_ini_aux & "','" & fecha_hist_fin_aux & "', '" & coment_hist_aux & "'," _
                       & " '" & estado_hist_aux & "','" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") & "') "


                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbinsert_imp_pil_hist, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbcommand.ExecuteNonQuery()
                    dbConexion.Close()
                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing

                    GoTo salir

                End If


                For i As Integer = 0 To cadEnter_fecha_hist_inicio.Length - 1

                    If (cadEnter_fecha_hist_inicio.Length = 1) Then
                        fecha_hist_ini_aux = "Por Confirmar"
                    Else
                        If (cadEnter_fecha_hist_inicio.Length > contador) Then

                            fecha_hist_ini_aux = cadEnter_fecha_hist_inicio(i).Trim
                        Else
                            fecha_hist_ini_aux = "Sin Datos"

                        End If
                    End If


                    If (cadEnter_fecha_hist_fin.Length = 1) Then
                        fecha_hist_fin_aux = "Por Confirmar"
                    Else
                        If (cadEnter_fecha_hist_fin.Length > contador) Then

                            fecha_hist_fin_aux = cadEnter_fecha_hist_fin(i).Trim
                        Else
                            fecha_hist_fin_aux = "Sin Datos"

                        End If

                    End If

                    If (cadEnter_coment_hist.Length > contador) Then

                        coment_hist_aux = cadEnter_coment_hist(i).Trim
                    Else
                        coment_hist_aux = ""

                    End If


                    If (cadEnter_estado_hist.Length > contador) Then

                        estado_hist_aux = cadEnter_estado_hist(i).Trim
                    Else
                        estado_hist_aux = ""

                    End If

                    contador += 1

                    dbinsert_imp_pil_hist = "INSERT INTO [dbo].[" & tblNom & "] " _
                            & "(pilH_ini_ide,pilH_ini_cod_ini,pilH_fec_hist_inicio,pilH_fec_hist_fin,pilH_com,pilH_estado,pilH_fec_act_pil) " _
                            & "values('" & valor_id_iniciativa & "', '" & cod_proyecto_val & "','" & fecha_hist_ini_aux & "','" & fecha_hist_fin_aux & "', '" & coment_hist_aux & "'," _
                            & " '" & estado_hist_aux & "','" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") & "') "

                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbinsert_imp_pil_hist, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbcommand.ExecuteNonQuery()

                    dbConexion.Close()

                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing

                Next
            End If
            Log("Se Inserto con Exito los datos de la pestaña Historica Piloto con el codigo: " & cod_proyecto_val, "exito")
            Console.WriteLine("Se Inserto con Exito los datos de la pestaña Historica Piloto")


        Catch ex As Exception
            Log("Se ha producido un error con los datos de la pestaña Historica Piloto con el codigo: " & cod_proyecto_val & ":" & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Return "0"
        End Try

salir:
        Return 0

    End Function

    '********************************************Despliegue************************************************************

    Private Function checkValuesdespliegueFields(ByVal cod_ini_val As String,
                                             ByVal fecha_ini_val As String,
                                             ByVal fecha_fin_val As String,
                                             ByVal coment_val As String,
                                             ByVal estado_piloto_val As String,
                                             ByVal tblNom As String) As String
        Try

            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim dbinsert As String = ""
            Dim dbdelete As String = ""
            Dim cadEnter_fecha_ini() As String
            Dim cadEnter_fecha_fin() As String
            Dim cadEnter_coment() As String
            Dim cadEnter_estado_piloto() As String
            Dim dbResult As String = ""
            Dim query_SQL As String = ""
            Dim dbinsert_despliegue As String = ""
            Dim fecha_ini_aux As String = ""
            Dim fecha_fin_aux As String = ""
            Dim coment_aux As String = ""
            Dim estado_aux As String = ""
            Dim stringSeparators() As String = {","c}
            Dim contador As Integer = 0
            Dim query_id_iniciativa As String = ""
            Dim valor_id_iniciativa As String = "0"

            cadEnter_fecha_ini = fecha_ini_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_fecha_fin = fecha_fin_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_coment = coment_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_estado_piloto = estado_piloto_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)


            query_id_iniciativa = "  Select ini_ide FROM [dbo].[imp_iniciativa] WHERE ini_cod = '" & cod_ini_val & "'"
            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(query_id_iniciativa, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader


            If dbdata.HasRows = True Then
                dbdata.Read()
                valor_id_iniciativa = dbdata.Item(0).ToString

                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing

            End If

            query_SQL = "  Select count(desp_cod_ini) FROM [dbo].[imp_ini_desp] WHERE desp_cod_ini = '" & cod_ini_val & "'"
            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(query_SQL, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader


            If dbdata.HasRows = True Then
                dbdata.Read()
                dbResult = dbdata.Item(0).ToString

                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing

            End If

            If (dbResult = "0") Then


                If (cadEnter_fecha_ini.Length = 0) And (cadEnter_fecha_fin.Length = 0) And (cadEnter_coment.Length = 0) Then

                    dbinsert_despliegue = "INSERT INTO [dbo].[" & tblNom & "] " _
                                  & "(desp_ini_ide,desp_cod_ini,desp_fec_ini,desp_fec_ter,desp_obs,desp_est,desp_fec_act) " _
                                  & "values('" & valor_id_iniciativa & "','" & cod_ini_val & "', '" & fecha_ini_aux & "','" & fecha_fin_aux & "', '" & coment_aux & "'," _
                                  & " '" & estado_aux & "','" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") & "') "


                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbinsert_despliegue, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbcommand.ExecuteNonQuery()
                    dbConexion.Close()
                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing

                    GoTo salir


                End If



                For i As Integer = 0 To cadEnter_fecha_ini.Length - 1


                    If (cadEnter_fecha_ini.Length > contador) Then

                        fecha_ini_aux = cadEnter_fecha_ini(i).Trim
                    Else
                        fecha_ini_aux = ""

                    End If

                    If (cadEnter_fecha_fin.Length > contador) Then

                        fecha_fin_aux = cadEnter_fecha_fin(i).Trim
                    Else
                        fecha_fin_aux = ""

                    End If


                    If (cadEnter_coment.Length > contador) Then

                        coment_aux = cadEnter_coment(i).Trim
                    Else
                        coment_aux = ""

                    End If


                    If (cadEnter_estado_piloto.Length > contador) Then

                        estado_aux = cadEnter_estado_piloto(i).Trim
                    Else
                        estado_aux = ""

                    End If


                    contador += 1

                    dbinsert_despliegue = "INSERT INTO [dbo].[" & tblNom & "] " _
                               & "(desp_ini_ide,desp_cod_ini,desp_fec_ini,desp_fec_ter,desp_obs,desp_est,desp_fec_act) " _
                               & "values('" & valor_id_iniciativa & "','" & cod_ini_val & "', '" & fecha_ini_aux & "','" & fecha_fin_aux & "', '" & coment_aux & "'," _
                               & " '" & estado_aux & "','" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") & "') "


                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbinsert_despliegue, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbcommand.ExecuteNonQuery()
                    dbConexion.Close()
                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing

                Next
            End If

            Log("Se Inserto con Exito en la tabla Despliegue con el codigo iniciativa: " & cod_ini_val, "exito")
            Console.WriteLine("Se Inserto con Exito en la tabla Despliegue. ")
        Catch ex As Exception
            Log("Se ha producido un error en la Función checkValuesdespliegueFields con el codigo iniciativa: " & cod_ini_val & ":" & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try

salir:
        Return 0

    End Function

    '************************CheckValuesHistorico************************************************************************
    Private Function checkValueshistoricodespliegueFields(ByVal fecha_hist_despliegue As String,
                                  ByVal comentario_hist_despliegue As String,
                                  ByVal fecha_act_despliegue As String,
                                  ByVal tblNom As String) As String
        Try
            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim cadEnter_fecha_hist_desp() As String
            Dim cadEnter_coment_hist_desp() As String
            Dim cadEnter_fecha_act_desp() As String
            Dim dbinsert_imp_ini_desp_hist As String = ""
            Dim fecha_hist_desp_aux As String = ""
            Dim coment_hist_desp_aux As String = ""
            Dim fecha_act_desp_aux As String = ""
            Dim stringSeparators() As String = {"\"c}


            cadEnter_fecha_hist_desp = fecha_hist_despliegue.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_coment_hist_desp = comentario_hist_despliegue.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_fecha_act_desp = fecha_act_despliegue.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)


            If (cadEnter_fecha_hist_desp.Length = 0) Then

                fecha_hist_desp_aux = "Sin Asignar"
                coment_hist_desp_aux = "Sin Asignar"
                fecha_act_desp_aux = "Sin Asignar"

                dbinsert_imp_ini_desp_hist = "INSERT INTO [dbo].[" & tblNom & "] " _
                            & "(despH_ini_ide,despH_fec_hist,despH_fec_act_hist,despH_obs) " _
                            & "values('" & id_iniCod & "', '" & fecha_hist_desp_aux & "','" & fecha_act_desp_aux & "', " _
                            & " '" & coment_hist_desp_aux & "') "

                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(dbinsert_imp_ini_desp_hist, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                dbcommand.ExecuteNonQuery()
                Log("Se Inserto con Exito en la tabla IMP INI Desp Historica.", "exito")
                Console.WriteLine("Se Inserto con Exito en la tabla IMP INI Desp Historica.")
                dbConexion.Close()

                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing
            Else


                For i As Integer = 0 To cadEnter_fecha_hist_desp.Length - 1
                    fecha_hist_desp_aux = cadEnter_fecha_hist_desp(i)

                    If (comentario_hist_despliegue <> "") Then
                        coment_hist_desp_aux = cadEnter_coment_hist_desp(i)
                    Else
                        coment_hist_desp_aux = ""
                    End If


                    If (cadEnter_fecha_act_desp.Length > i) Then
                        fecha_act_desp_aux = cadEnter_fecha_act_desp(i)
                    Else

                        fecha_act_desp_aux = ""
                    End If


                    dbinsert_imp_ini_desp_hist = "INSERT INTO [dbo].[" & tblNom & "] " _
                            & "(despH_ini_ide,despH_fec_hist,despH_fec_act_hist,despH_obs) " _
                            & "values('" & id_iniCod & "', '" & fecha_hist_desp_aux & "','" & fecha_act_desp_aux & "', " _
                            & " '" & coment_hist_desp_aux & "') "


                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbinsert_imp_ini_desp_hist, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbcommand.ExecuteNonQuery()
                    dbConexion.Close()
                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing

                Next
                Log("Se Inserto con Exito en la tabla imp_ini_desp_hist", "exito")
                Console.WriteLine("Se Inserto con Exito en la tabla imp_ini_desp_hist. ")
            End If

        Catch ex As Exception
            Log("Se ha producido un error en la Función checkValueshistoricodespliegueFields con el codigo: " & id_iniCod & ":" & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()


            Return "0"
        End Try


salir:
        Return 0

    End Function

    '*****************************************Despliegue Historico****************************************

    Private Function despliegueHistoricoValues(ByVal cod_proyecto_val As String,
                                        ByVal fecha_hist_inicio_val As String,
                                        ByVal fecha_hist_fin_val As String,
                                        ByVal coment_hist_val As String,
                                        ByVal estado_hist_piloto_val As String,
                                        ByVal fecha_act_piloto_val As String,
                                        ByVal tblNom As String) As String
        Try

            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim dbinsert As String = ""
            Dim dbdelete As String = ""
            Dim cadEnter_fecha_hist_inicio() As String
            Dim cadEnter_fecha_hist_fin() As String
            Dim cadEnter_coment_hist() As String
            Dim cadEnter_estado_hist() As String
            Dim cadEnter_fecha_act_piloto() As String

            Dim dbinsert_despliegue_hist As String = ""
            Dim fecha_hist_ini_aux As String = ""
            Dim fecha_hist_fin_aux As String = ""
            Dim coment_hist_aux As String = ""
            Dim estado_hist_aux As String = ""
            Dim fecha_act_hist_aux As String = ""
            Dim contador As Integer = 0
            Dim insert_despliegue_historico As String = ""
            Dim query_id_iniciativa As String = ""
            Dim valor_id_iniciativa As String = ""
            Dim dbdeletedespliegue As String = ""

            If (fecha_hist_inicio_val <> "") Then
                cadEnter_fecha_hist_inicio = fecha_hist_inicio_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_fecha_hist_fin = fecha_hist_fin_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_coment_hist = coment_hist_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_estado_hist = estado_hist_piloto_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_fecha_act_piloto = fecha_act_piloto_val.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)


                query_id_iniciativa = "  Select ini_ide FROM [dbo].[imp_iniciativa] WHERE ini_cod = '" & cod_proyecto_val & "'"
                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(query_id_iniciativa, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                dbdata = dbcommand.ExecuteReader


                If dbdata.HasRows = True Then
                    dbdata.Read()
                    valor_id_iniciativa = dbdata.Item(0).ToString

                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing

                End If

                'where desp_ini_ide = '" & valor_id_iniciativa & "'

                insert_despliegue_historico = "INSERT INTO [dbo].[" & tblNom & "]" _
                & "(despH_ini_ide,despH_cod_ini,despH_fec_inicio,despH_fec_hist_fin," _
                & "despH_coment_hist,despH_estado_hist,despH_fec_act_hist)" _
                & " SELECT desp_ini_ide,desp_cod_ini,desp_fec_ini,desp_fec_ter, desp_obs,desp_est,convert(varchar, getdate(), 120)  " _
                & " FROM [dbo].[imp_ini_desp] "

                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(insert_despliegue_historico, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                dbcommand.ExecuteNonQuery()
                Log("Se Inserto con Exito en la tabla Historica Despliegue con el codigo: " & cod_proyecto_val, "exito")
                Console.WriteLine("Se Inserto con Exito en la tabla Historica Despliegue... ")
                dbConexion.Close()

                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing


            End If


        Catch ex As Exception
            Log("Se ha producido un error en la Función despliegueHistoricoValues con el codigo: " & cod_proyecto_val & ":" & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try

salir:
        Return 0

    End Function

    '**************************************************************************************************************************

    Private Function checkValuesAmbitoFields(ByVal cod_ini_val As String,
                                            ByVal kick_off_val As String,
                                      ByVal requerimiento_val As String,
                                      ByVal proceso_compra_val As String,
                                      ByVal infra_val As String,
                                      ByVal habilitacion_val As String,
                                      ByVal riesgo_ope_val As String,
                                      ByVal seguridad_val As String,
                                      ByVal comun_estud_val As String,
                                      ByVal soport_ope_val As String,
                                      ByVal gest_cambio_val As String,
                                      ByVal monitoreo_val As String,
                                      ByVal normt_proc_val As String,
                                      ByVal coexistencia_val As String,
                                      ByVal gest_reclamos_val As String,
                                      ByVal sist_tecnlogia_val As String,
                                      ByVal roles_val As String,
                                      ByVal gestion_ind_val As String,
                                      ByVal instal_faena_val As String,
                                      ByVal tblNom As String) As String
        Try
            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim dbinsert_imp_ambito As String = ""
            Dim querySQL As String = ""
            Dim valor As String = ""

            'If (kick_off_val <> "") Then
            querySQL = "  Select count(amb_cod_ini) FROM [dbo].[" & tblNom & "]  WHERE amb_cod_ini = '" & cod_ini_val & "'"
            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(querySQL, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader


            If dbdata.HasRows = True Then
                dbdata.Read()
                valor = dbdata.Item(0).ToString

                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing

            End If


            If (valor = "0") Then

                dbinsert_imp_ambito = "INSERT INTO [dbo].[" & tblNom & "] " _
                         & "(amb_ini_ide,amb_cod_ini,amb_kickoff,amb_requerimiento,amb_proc_compra,amb_infraest,amb_habilit,amb_riesg_ope,amb_seguridad,amb_comunic_est, " _
                         & " amb_sop_ope,amb_gst_camb,amb_monitoreo,amb_norm_proc_ctos,amb_coexistencia,amb_gest_recl,amb_sist_tec,amb_roles,amb_gst_indicadores,amb_instal_faena ) " _
                          & "values('" & id_iniCod & "', '" & cod_ini_val & "','" & kick_off_val & "','" & requerimiento_val & "', " _
                          & " '" & proceso_compra_val & "', '" & infra_val & "', '" & habilitacion_val & "','" & riesgo_ope_val & "', '" & seguridad_val & "' , " _
                          & " '" & comun_estud_val & "', '" & soport_ope_val & "', '" & gest_cambio_val & "','" & monitoreo_val & "', '" & normt_proc_val & "', " _
                          & " '" & coexistencia_val & "',  '" & gest_reclamos_val & "', '" & sist_tecnlogia_val & "', '" & roles_val & "','" & gestion_ind_val & "', " _
                          & " '" & instal_faena_val & "' )"

                'Debug.Print(dbinsert_imp_ambito)

                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(dbinsert_imp_ambito, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                Dim count As Integer = dbcommand.ExecuteNonQuery()

                'If (count = 1) Then
                dbConexion.Close()
                dbConexion = Nothing
                dbcommand = Nothing

            End If
            Log("Se Inserto con Exito en la tabla ambito con el codigo:" & id_iniCod, "exito")
            Console.WriteLine("Se Inserto con Exito en la tabla ambito")
        Catch ex As Exception
            Log("Se ha producido un error en la Función checkValuesAmbitoFields con el codigo  :" & id_iniCod & ":" & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try

salir:
        Return 0

    End Function

    '********************************************************************************************************************************'
    Private Function checkValuesTemaFields(ByVal tema_relevante_value As String,
                                  ByVal fecha_tema_relevante As String,
                                  ByVal cod_iniciativa As String,
                                  ByVal tblNom As String) As String
        Try
            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim dbinsert_imp_tema_relevante As String = ""
            Dim query_id_iniciativa As String = ""
            Dim cadEnter_tema_rel() As String
            Dim cadEnter_fecha_tema_rel() As String
            Dim tema_rel_aux As String = ""
            Dim fecha_rel_aux As String = ""
            Dim valor_id_iniciativa As String = ""

            cadEnter_tema_rel = tema_relevante_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_fecha_tema_rel = fecha_tema_relevante.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)

            ' If (tema_relevante_value <> "") Then

            query_id_iniciativa = "  Select ini_ide FROM [dbo].[imp_iniciativa] WHERE ini_cod = '" & cod_iniciativa & "'"
            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(query_id_iniciativa, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader


            If dbdata.HasRows = True Then
                dbdata.Read()
                valor_id_iniciativa = dbdata.Item(0).ToString

                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing

            End If


            dbinsert_imp_tema_relevante = "INSERT INTO [dbo].[" & tblNom & "] " _
                      & "(tema_ini_ide,tema_cod_ini,tema_relevante,fecha_relevante,fecha_actual) " _
                      & "values('" & valor_id_iniciativa & "', '" & cod_iniciativa & "','" & tema_relevante_value & "','" & fecha_tema_relevante & "','" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") & "')"
            'Debug.Print(dbinsert_imp_tema_relevante)

            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(dbinsert_imp_tema_relevante, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbcommand.ExecuteNonQuery()

            Log("Se Inserto con Exito en la tabla tema relevante del Codigo Iniciativa:" & cod_iniciativa, "exito")
            Console.WriteLine("Se Inserto con Exito en la tabla tema relevante del Codigo Iniciativa:" & cod_iniciativa)

            'End If

        Catch ex As Exception
            Log("Se ha producido un error en la Función checkValuesTemaFields con el Codigo Iniciativa:" & cod_iniciativa & ":" & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try

salir:
        Return 0
    End Function

    '***************************************************************************************************************'
    Private Function checkValuesPacknormativoFields(ByVal tiene_pack_value As String,
                                  ByVal fecha_recib_value As String,
                                  ByVal tblNom As String) As String
        Try
            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbinsert_imp_pack_norm As String = ""
            Dim Fecha_recibida As String = ""

            If (fecha_recib_value <> "") Then

                Fecha_recibida = fecha_recib_value
            Else
                Fecha_recibida = ""

            End If

            If (tiene_pack_value <> "") Then
                dbinsert_imp_pack_norm = "INSERT INTO [dbo].[" & tblNom & "] " _
                          & "(pack_normativo_ini_ide,pack_normativo_tiene,pack_normativo_fecha_recib) " _
                          & "values('" & id_iniCod & "', '" & tiene_pack_value & "','" & Fecha_recibida & "')"

                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(dbinsert_imp_pack_norm, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                Dim count As Integer = dbcommand.ExecuteNonQuery()

                If (count = 1) Then

                    Console.WriteLine("Se Inserto con Exito en la tabla pack normativo")
                End If
                Return count
            End If

        Catch ex As Exception
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try


salir:
        Return 0
    End Function

    '***********************************************************************************************************************************************'
    Private Function checkValuesHistoricoPilotFields(ByVal fecha_historico_pil_value As String,
                                  ByVal comentario_historico_pil_value As String,
                                  ByVal tblNom As String) As String

        Try
            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim cadEnter_fecha_hist_pil() As String
            Dim cadEnter_coment_hist_pil() As String
            Dim dbinsert_ini_pil_hist As String = ""
            Dim fecha_hist_pil_aux As String = ""
            Dim coment_hist_pil_aux As String = ""


            cadEnter_fecha_hist_pil = fecha_historico_pil_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_coment_hist_pil = comentario_historico_pil_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)


            For i As Integer = 0 To cadEnter_fecha_hist_pil.Length - 1
                fecha_hist_pil_aux = cadEnter_fecha_hist_pil(i)

                If (comentario_historico_pil_value <> "") Then
                    coment_hist_pil_aux = cadEnter_coment_hist_pil(i)
                Else
                    coment_hist_pil_aux = ""
                End If

                dbinsert_ini_pil_hist = "INSERT INTO [dbo].[" & tblNom & "] " _
                        & "(pilH_ini_ide,pilH_fec_hist,pilH_fec_act_hist,pilH_com) " _
                        & "values('" & id_iniCod & "', '" & fecha_hist_pil_aux & "','" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") & "', " _
                        & " '" & coment_hist_pil_aux & "') "


                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(dbinsert_ini_pil_hist, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                dbcommand.ExecuteNonQuery()
                dbConexion.Close()
                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing

            Next
            Log("Se Inserto con Exito en la tabla imp_ini_pil_hist", "exito")
            Console.WriteLine("Se Inserto con Exito en la tabla imp_ini_pil_hist... ")
        Catch ex As Exception
            Log("Se ha producido un error en la Función checkValuesHistoricoPilotFields" & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try


salir:
        Return 0
    End Function
    '********************************************GET INI INICIATIVA*************************************************************************'
    Public Function getInit_Data(ByVal tblNom As String,
                                  ByVal idCategoria1_val As String,
                                  ByVal idCategoria2_val As String,
                                  ByVal idGantCubo_val As String,
                                  ByVal idJefeProyecto_val As String,
                                  ByVal idLider_val As String,
                                  ByVal idIF_val As String,
                                  ByVal idIFBackup_val As String,
                                  ByVal iniciativa_val As String,
                                  ByVal cod_ini_val As String,
                                  ByVal estado_val As String,
                                  ByVal desc_ejecutiva_val As String,
                                  ByVal entregable_proyecto_val As String,
                                  ByVal canal_impactado_suc_val As String,
                                  ByVal canal_impactado_int_val As String,
                                  ByVal canal_impactado_aut_val As String,
                                  ByVal canal_impactado_otr_val As String,
                                  ByVal responsable_eje_val As String) As String


        Dim dbConexion As Data.Odbc.OdbcConnection
        Dim dbcommand As Data.Odbc.OdbcCommand
        Dim dbconsulta As String = ""
        Dim dbinsert As String = ""
        Dim dbresultados As String = ""
        Dim dbRowCount As String = ""
        Dim dbdata As Data.Odbc.OdbcDataReader


        Try
            If (iniciativa_val <> "") Then

                dbRowCount = "select COUNT(*) AS contador from [dbo].[" & tblNom & "] where ini_nom='" & iniciativa_val & "' "
                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(dbRowCount, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                dbdata = dbcommand.ExecuteReader

                If dbdata.HasRows = True Then
                    dbdata.Read()
                    dbresultados = dbdata.Item(0).ToString

                    dbdata.Close()
                    dbConexion.Close()

                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing

                    If (dbresultados = 0) Then

                        dbinsert = "INSERT [dbo].[" & tblNom & "] " _
                                        & "(ini_cod,ini_adr,ini_nom,ini_des,ini_ent,id_gantcubo,ini_cnl_imp_suc,ini_cnl_imp_int,ini_cnl_imp_aut,ini_cnl_imp_otro,ini_obs,ini_obs_fec,ini_cat2_ide,ini_jp_ide,ini_ldr_ide,ini_est,ini_cat1_ide) " _
                                        & "values('" & cod_iniciativa & "',' ', '" & iniciativa_val & " ', '" & desc_ejecutiva_val & "', '" & entregable_proyecto_val & "', '" & idGantCubo_val & "'," _
                                        & " '" & canal_impactado_suc_val & "', '" & canal_impactado_int_val & "','" & canal_impactado_aut_val & "','" & canal_impactado_otr_val & "' , " _
                                        & " '  ', ' ', '" & idCategoria2_val & "','" & idJefeProyecto_val & "', '" & idLider_val & "', '1', '" & idCategoria1_val & "'  )"

                        'Debug.Print(dbinsert)

                        dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                        dbcommand = New Data.Odbc.OdbcCommand(dbinsert, dbConexion)
                        dbcommand.CommandType = CommandType.Text
                        dbConexion.Open()
                        dbcommand.ExecuteNonQuery()

                        dbConexion.Close()
                        Log("Se Inserto con Exito en la tabla Iniciativa con el codigo: " & id_iniCod, "exito")
                        Console.WriteLine("Se Inserto con Exito en la tabla Iniciativa... ")
                    End If

                    'query_id_iniCod = "select Max(ini_ide) from [dbo].[" & tblNom & "] where ini_cod='" & cod_iniciativa & "' "
                    query_id_iniCod = "select Max(ini_ide) from [dbo].[" & tblNom & "] where ini_cod='" & cod_ini_val & "' "
                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(query_id_iniCod, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbdata = dbcommand.ExecuteReader

                    If dbdata.HasRows = True Then
                        dbdata.Read()
                        id_iniCod = dbdata.Item(0).ToString

                        dbdata.Close()
                        dbConexion.Close()

                        dbConexion = Nothing
                        dbcommand = Nothing
                        dbdata = Nothing

                        Console.WriteLine(" Codigo Iniciativa:" & cod_ini_val & " " & " y su ID es: " & id_iniCod)
                        Return id_iniCod

                    Else
                        dbConexion = Nothing
                        dbcommand = Nothing
                        dbdata = Nothing

                        Return "0"
                    End If

                End If
            End If
            'Log("Se Inserto con Exito en la tabla Iniciativa con el codigo: " & id_iniCod, "exito")
            'Console.WriteLine("Se Inserto con Exito en la tabla Iniciativa... ")
        Catch ex As Exception
            Log("Se produjo un error en la Función getInit_Data con el codigo: " & id_iniCod, "error")
            Console.WriteLine("Error insertando Datos....... " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try

salir:
        Return 0

    End Function

    '*************************************************QUITAR CARACTERES ESPECIALES****************************************'
    Function CleanInput(strIn As String) As String
        ' Replace invalid characters with empty strings.
        Try
            Return Regex.Replace(strIn, "[']", "")
            ' If we timeout when replacing invalid characters, 
            ' we should return String.Empty.
        Catch e As System.TimeoutException
            Return String.Empty
        End Try
    End Function


    '***************************************************MODELO GENERAL*****************************************************************'
    Public Sub ModeloGeneral()

        Dim adoCon As Data.OleDb.OleDbConnection
        Dim adoRs As Data.OleDb.OleDbDataReader
        Dim adoCmd As Data.OleDb.OleDbCommand
        Dim strCmnd As String = "SELECT * FROM [Informacion Global$]"

        Console.WriteLine("Cargando archivo de iniciativas")
        Console.WriteLine("Cadena de conexion : " & GetConnectionString(1))

        adoCon = New Data.OleDb.OleDbConnection(GetConnectionString(1))
        adoCmd = New Data.OleDb.OleDbCommand(strCmnd, adoCon)
        adoCmd.CommandType = CommandType.Text
        adoCon.Open()
        adoRs = adoCmd.ExecuteReader

        If adoRs.HasRows Then
            adoRs.Read()
            Do While adoRs.Read()
                'Do While Convert.ToString(adoRs.Item(0)) <> " "
                If (Convert.ToString(adoRs.Item(0)) <> "Incluir 'Empatía' y texto libre") Then


                    Try
                        iniciativa = Convert.ToString(adoRs.Item(2)).Trim
                        cod_iniciativa = Convert.ToString(adoRs.Item(3)).Trim
                        estado = Convert.ToString(adoRs.Item(4)).Trim
                        desc_ejecutiva = Convert.ToString(adoRs.Item(5)).Trim
                        entregable_proyecto = CleanInput(Convert.ToString(adoRs.Item(6)).Trim)

                        canal_impactado_suc = Convert.ToString(adoRs.Item(8)).Trim
                        canal_impactado_int = Convert.ToString(adoRs.Item(9)).Trim
                        canal_impactado_aut = Convert.ToString(adoRs.Item(10)).Trim
                        canal_impactado_otr = Convert.ToString(adoRs.Item(11)).Trim
                        responsable_eje = Convert.ToString(adoRs.Item(12)).Trim


                        If (cod_iniciativa <> "") Then
                            initData = getInit_Data("imp_iniciativa", idCategoria1, idCategoria2, idGantCubo, idJefeProyecto, idLider, idIF, idIFBackup, iniciativa, cod_iniciativa, estado, desc_ejecutiva, entregable_proyecto, canal_impactado_suc, canal_impactado_int, canal_impactado_aut, canal_impactado_otr, responsable_eje)

                            'Categoria 1
                            idCategoria1 = checkCategoria(Convert.ToString(adoRs.Item(0)), "", cod_iniciativa, "imp_categoria_1", 0)

                            'Categoria 2
                            idCategoria2 = checkCategoria(Convert.ToString(adoRs.Item(1)), Convert.ToString(adoRs.Item(12)), cod_iniciativa, "imp_categoria_2", 1)

                            'Gant Cubo
                            idGantCubo = checkGantCubo(Convert.ToString(adoRs.Item(7)), Convert.ToString(adoRs.Item(3)).Trim, "imp_gc")

                            'jefe proyecto
                            idJefeProyecto = check_JP_LDR_IF(Convert.ToString(adoRs.Item(13)), cod_iniciativa, "jp_nom", "jp_pat", "imp_jp", 0)

                            idLider = check_JP_LDR_IF(Convert.ToString(adoRs.Item(14)), cod_iniciativa, "ldr_nom", "ldr_pat", "imp_ldr", 1)


                            'idIF = check_JP_LDR_IF(Convert.ToString(adoRs.Item(16)), cod_iniciativa, "if_nom", "if_pat", "imp_esp_if", 2)


                            'idIFBackup = check_JP_LDR_IF(Convert.ToString(adoRs.Item(17)), cod_iniciativa, "if_nom", "if_pat", "imp_esp_if", 3)

                            idIF = check_Esp_IF_Back(Convert.ToString(adoRs.Item(16)), Convert.ToString(adoRs.Item(17)), cod_iniciativa, "imp_esp_if")


                            idGestores = check_Gestores_otro_Miembros(Convert.ToString(adoRs.Item(15)), cod_iniciativa, "imp_gst", 1)


                            idEspif = check_ini_esp_if("imp_ini_esp_if")

                            '2da Columna del Archivo cabecera Color Azul Paso Produccion-Comentarios PAP'
                            id_paso_prod_Coment = checkValuespasoProd_comment(cod_iniciativa, Convert.ToString(adoRs.Item(18)), Convert.ToString(adoRs.Item(19)), Convert.ToString(adoRs.Item(20)), "imp_ini_pap")


                            '3era Columna del Archivo cabecera Color Azul (Fecha Inicio Piloto-Fecha Fin Piloto-Comentario Piloto-Fecha Histórica Piloto-Comentario Historico Piloto-Fecha Actualización Piloto)'
                            id_piloto = checkValuespilotoFields(Convert.ToString(adoRs.Item(3)), Convert.ToString(adoRs.Item(21)), Convert.ToString(adoRs.Item(22)), Convert.ToString(adoRs.Item(23)), Convert.ToString(adoRs.Item(24)), "imp_ini_pil")


                            '4ta Columna del Archivo cabecera Color Azul (Fecha Inicio Despliegue-Fecha fin  Despliegue-Comentario Despliegue)'
                            id_despliegue = checkValuesdespliegueFields(Convert.ToString(adoRs.Item(3)), Convert.ToString(adoRs.Item(25)), Convert.ToString(adoRs.Item(26)), Convert.ToString(adoRs.Item(27)), Convert.ToString(adoRs.Item(28)), "imp_ini_desp")

                            ''6ta Columna del Archivo cabecera Color Rojo (KickOff .. Instalación de Faena)
                            id_ambito = checkValuesAmbitoFields(cod_iniciativa, Convert.ToString(adoRs.Item(29)), Convert.ToString(adoRs.Item(30)), Convert.ToString(adoRs.Item(31)), Convert.ToString(adoRs.Item(32)), Convert.ToString(adoRs.Item(33)), Convert.ToString(adoRs.Item(34)), Convert.ToString(adoRs.Item(35)), Convert.ToString(adoRs.Item(36)), Convert.ToString(adoRs.Item(37)), Convert.ToString(adoRs.Item(38)), Convert.ToString(adoRs.Item(39)), Convert.ToString(adoRs.Item(40)), Convert.ToString(adoRs.Item(41)), Convert.ToString(adoRs.Item(42)), Convert.ToString(adoRs.Item(43)), Convert.ToString(adoRs.Item(44)), Convert.ToString(adoRs.Item(45)), Convert.ToString(adoRs.Item(46)), "imp_ambito")

                            'TEMA RELEVANTE HISTORICO***************************************************'
                            id_tema_rel_historico = temaRelevanteHistValuesByVal(cod_iniciativa, "imp_temas_relevantes_hist")

                            '9va Columna del Archivo (Tiene Pack Normativo - Fecha Recibida Pack Normativo)
                            id_pack_norm = checkValuesPacknormativoFields(Convert.ToString(adoRs.Item(51)), Convert.ToString(adoRs.Item(52)), "imp_pack_normativo")


                        End If

                    Catch ex As Exception
                        Log("Error en el Modelo General con el codigo: " & cod_iniciativa & ":" & ex.Message, "error")
                        Console.WriteLine("Error En el Modelo General ....... " & ex.Message)
                        Console.WriteLine("Error con el codigo:" & cod_iniciativa)
                        Console.ReadLine()
                    End Try


                End If

            Loop

        Else
            GoTo salirSinFilas
            Debug.Print("No rows found.")
            Console.ReadLine()
        End If

salirSinFilas:
        adoRs.Close()
        adoCon.Close()

        adoCon = Nothing
        adoRs = Nothing
        adoCmd = Nothing

    End Sub

    '***************************************************TEMA RELEVANTE INSERTAR******************************************************************************************'
    Public Sub InsertTemaRelevante()

        Dim adoCon As Data.OleDb.OleDbConnection
        Dim adoRs As Data.OleDb.OleDbDataReader
        Dim adoCmd As Data.OleDb.OleDbCommand
        Dim strCmnd As String = "SELECT * FROM [Informacion Global$]"

        adoCon = New Data.OleDb.OleDbConnection(GetConnectionString(1))
        adoCmd = New Data.OleDb.OleDbCommand(strCmnd, adoCon)
        adoCmd.CommandType = CommandType.Text
        adoCon.Open()
        adoRs = adoCmd.ExecuteReader


        Try
            If adoRs.HasRows Then
                adoRs.Read()
                Do While adoRs.Read()
                    If (Convert.ToString(adoRs.Item(0)) <> "Incluir 'Empatía' y texto libre") Then
                        ''7ta Columna del Archivo cabecera Color Verde (Temas-Relevamtes - Fecha-tema-Relevante)'
                        id_tema_relevante = checkValuesTemaFields(Convert.ToString(adoRs.Item(53)), Convert.ToString(adoRs.Item(54)), Convert.ToString(adoRs.Item(3)), "imp_temas_relevantes")

                    End If
                Loop
            End If


        Catch ex As Exception
            Log("Se ha producido un error en la funcion InsertTemaRelevante  " & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Log(ex.Message, "error")
            Console.ReadLine()

        End Try


    End Sub

    '******************************INFORMACIÓN GLOBAL EMPATIA (PILOTO HISTORICOS)********************************************************************'

    Public Sub ModeloPilotoHistoricos()

        Dim adoCon As Data.OleDb.OleDbConnection
        Dim adoRs As Data.OleDb.OleDbDataReader
        Dim adoCmd As Data.OleDb.OleDbCommand
        Dim strCmnd As String = "SELECT * FROM [Pilotos Historicos$]"

        adoCon = New Data.OleDb.OleDbConnection(GetConnectionString(1))
        adoCmd = New Data.OleDb.OleDbCommand(strCmnd, adoCon)
        adoCmd.CommandType = CommandType.Text
        adoCon.Open()
        adoRs = adoCmd.ExecuteReader

        If adoRs.HasRows Then
            adoRs.Read()
            Do While adoRs.Read()
                If (Convert.ToString(adoRs.Item(0)) <> "Carácter >=3 y <=6") Then
                    Console.Write(Convert.ToString(adoRs.Item(5)))

                    If (Convert.ToString(adoRs.Item(5)) <> "") Then
                        idHistoricoPiloto = pilotoHistoricoValues(adoRs.Item(0), Convert.ToString(adoRs.Item(1)), Convert.ToString(adoRs.Item(2)), Convert.ToString(adoRs.Item(3)), Convert.ToString(adoRs.Item(4)), "imp_ini_pil_hist")
                    End If
                End If
            Loop

        Else
            GoTo salirSinFilas
            Debug.Print("No rows found.")
            Console.ReadLine()
        End If

salirSinFilas:
        adoRs.Close()
        adoCon.Close()

        adoCon = Nothing
        adoRs = Nothing
        adoCmd = Nothing

    End Sub

    '************************************Actualizar Piloto*******************************
    Public Sub InsertPilotoInicio()

        Dim adoCon As Data.OleDb.OleDbConnection
        Dim adoRs As Data.OleDb.OleDbDataReader
        Dim adoCmd As Data.OleDb.OleDbCommand
        Dim strCmnd As String = "SELECT * FROM [Informacion Global$]"

        adoCon = New Data.OleDb.OleDbConnection(GetConnectionString(1))
        adoCmd = New Data.OleDb.OleDbCommand(strCmnd, adoCon)
        adoCmd.CommandType = CommandType.Text
        adoCon.Open()
        adoRs = adoCmd.ExecuteReader

        If adoRs.HasRows Then
            adoRs.Read()
            Do While adoRs.Read()
                If (Convert.ToString(adoRs.Item(3)) <> "Carácter >=3 y <=6") Then
                    'Console.Write(Convert.ToString(adoRs.Item(5)))

                    If (Convert.ToString(adoRs.Item(21)) <> "") Then
                        id_piloto = checkValuespilotoFields(Convert.ToString(adoRs.Item(3)), Convert.ToString(adoRs.Item(21)), Convert.ToString(adoRs.Item(22)), Convert.ToString(adoRs.Item(23)), Convert.ToString(adoRs.Item(24)), "imp_ini_pil")
                    End If
                End If
            Loop

        Else
            GoTo salirSinFilas
            Debug.Print("No rows found.")
            Console.ReadLine()
        End If

salirSinFilas:
        adoRs.Close()
        adoCon.Close()

        adoCon = Nothing
        adoRs = Nothing
        adoCmd = Nothing

    End Sub


    '************************************Actualizar Despliegue********************************************
    Public Sub InsertDespliegueHistorico()

        Dim adoCon As Data.OleDb.OleDbConnection
        Dim adoRs As Data.OleDb.OleDbDataReader
        Dim adoCmd As Data.OleDb.OleDbCommand
        Dim strCmnd As String = "SELECT * FROM [Informacion Global$]"

        adoCon = New Data.OleDb.OleDbConnection(GetConnectionString(1))
        adoCmd = New Data.OleDb.OleDbCommand(strCmnd, adoCon)
        adoCmd.CommandType = CommandType.Text
        adoCon.Open()
        adoRs = adoCmd.ExecuteReader

        If adoRs.HasRows Then
            adoRs.Read()
            Do While adoRs.Read()
                If (Convert.ToString(adoRs.Item(3)) <> "Carácter >=3 y <=6") Then
                    'Console.Write(Convert.ToString(adoRs.Item(5)))

                    If (Convert.ToString(adoRs.Item(25)) <> "") Then
                        id_piloto = checkValuesdespliegueFields(Convert.ToString(adoRs.Item(3)), Convert.ToString(adoRs.Item(25)), Convert.ToString(adoRs.Item(26)), Convert.ToString(adoRs.Item(27)), Convert.ToString(adoRs.Item(28)), "imp_ini_desp")
                    End If
                End If
            Loop

        Else
            GoTo salirSinFilas
            Debug.Print("No rows found.")
            Console.ReadLine()
        End If

salirSinFilas:
        adoRs.Close()
        adoCon.Close()

        adoCon = Nothing
        adoRs = Nothing
        adoCmd = Nothing

    End Sub


    '********************************Piloto Historico********************************'
    Public Sub ModeloDespliegueHistorico()

        Dim adoCon As Data.OleDb.OleDbConnection
        Dim adoRs As Data.OleDb.OleDbDataReader
        Dim adoCmd As Data.OleDb.OleDbCommand
        Dim strCmnd As String = "SELECT * FROM [Despliegues Historicos$]"
        Dim fechaActDespliegue As String = ""
        Dim fechaHistIni As String = ""

        adoCon = New Data.OleDb.OleDbConnection(GetConnectionString(1))
        adoCmd = New Data.OleDb.OleDbCommand(strCmnd, adoCon)
        adoCmd.CommandType = CommandType.Text
        adoCon.Open()
        adoRs = adoCmd.ExecuteReader

        If adoRs.HasRows Then
            adoRs.Read()
            Do While adoRs.Read()
                If (Convert.ToString(adoRs.Item(0)) <> "Carácter >=3 y <=6") Then
                    'Console.Write(Convert.ToString(adoRs.Item(5)))

                    fechaHistIni = Convert.ToString(adoRs.Item(1)).Trim


                    If (Convert.ToString(adoRs.Item(5)) = "") Then
                        fechaActDespliegue = "-"
                    Else
                        fechaActDespliegue = Convert.ToString(adoRs.Item(5)).Trim

                    End If

                    If (fechaHistIni <> "") Then
                        idHistoricoDespliegue = despliegueHistoricoValues(adoRs.Item(0), Convert.ToString(adoRs.Item(1)), Convert.ToString(adoRs.Item(2)), Convert.ToString(adoRs.Item(3)), Convert.ToString(adoRs.Item(4)), fechaActDespliegue, "imp_ini_desp_hist")

                    End If
                End If
            Loop

        Else
            GoTo salirSinFilas
            Debug.Print("No rows found.")
            Console.ReadLine()
        End If

salirSinFilas:
        adoRs.Close()
        adoCon.Close()

        adoCon = Nothing
        adoRs = Nothing
        adoCmd = Nothing

    End Sub

    '*******************************************************************************************'

    Private Function checkListValuesByVal(fecha_realizacion_value As String,
                                     ByVal codigo_ini_value As String,
                                     ByVal ambito_value As String,
                                     ByVal accion_value As String,
                                     ByVal etapa_value As String,
                                     ByVal hito_value As String,
                                     ByVal preg_value As String,
                                     ByVal resp_value As String,
                                     ByVal obs_value As String,
                                     ByVal tblNom As String) As String

        Try
            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim cadEnter_fecha_realizacion() As String
            Dim cadEnter_codigo_ini() As String
            Dim cadEnter_ambito() As String
            Dim cadEnter_accion() As String
            Dim cadEnter_etapa() As String
            Dim cadEnter_hito() As String
            Dim cadEnter_preg() As String
            Dim cadEnter_resp() As String
            Dim cadEnter_obs() As String
            Dim dbinsert_ini_check_lists As String = ""
            Dim dbinsert_ini_check_lists_hist As String = ""
            Dim fecha_realizacion_aux As String = ""
            Dim codigo_ini_aux As String = ""
            Dim ambito_aux As String = ""
            Dim accion_aux As String = ""
            Dim etapa_aux As String = ""
            Dim hito_aux As String = ""
            Dim preg_aux As String = ""
            Dim resp_aux As String = ""
            Dim obs_aux As String = ""
            Dim cad_null As String = ""
            Dim str_accion As String = ""
            Dim str_etapa As String = ""
            Dim str_obs As String = ""
            Dim dbRowCount As String = ""
            Dim dbresultados As String = ""
            Dim IdResult As String = ""
            Dim dbdeletecheckLists As String = ""
            Dim cont As Integer = 0

            cadEnter_fecha_realizacion = fecha_realizacion_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_codigo_ini = codigo_ini_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_ambito = ambito_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_accion = accion_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_etapa = etapa_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_hito = hito_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_preg = preg_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_resp = resp_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_obs = obs_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)


            'buscar el ID de la iniciativa
            dbRowCount = "select ini_ide  from [dbo].[imp_iniciativa] where ini_cod = '" & codigo_ini_value & "'"
            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(dbRowCount, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader

            If dbdata.HasRows = True Then
                dbdata.Read()
                IdResult = dbdata.Item(0).ToString

            End If


            For i As Integer = 0 To cadEnter_fecha_realizacion.Length - 1
                fecha_realizacion_aux = cadEnter_fecha_realizacion(i)
                codigo_ini_aux = cadEnter_codigo_ini(i)
                ambito_aux = cadEnter_ambito(i)
                If (accion_value <> "") Then
                    accion_aux = cadEnter_accion(i)
                End If

                If (etapa_value <> "") Then
                    etapa_aux = cadEnter_etapa(i)
                End If

                hito_aux = cadEnter_hito(i)
                preg_aux = cadEnter_preg(i)
                resp_aux = cadEnter_resp(i)

                If (obs_value <> "") Then
                    obs_aux = cadEnter_obs(i)
                End If


                dbinsert_ini_check_lists = "INSERT INTO [dbo].[" & tblNom & "] " _
             & "(checkList_ini_ide,checkList_fecha_realizacion,checkList_cod_ini,checkList_ambito,checkList_accion, " _
             & "checkList_etapa,checkList_hito,checkList_pregunta,checkList_respuesta,checkList_observaciones) " _
             & "values('" & IdResult & "', '" & fecha_realizacion_aux & "','" & codigo_ini_aux & "','" & ambito_aux & "',  " _
             & " '" & accion_aux & "','" & etapa_aux & "','" & hito_aux & "','" & preg_aux & "','" & resp_aux & "','" & obs_aux & "')"

                'Debug.Print(dbinsert_ini_check_lists)

                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(dbinsert_ini_check_lists, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                dbcommand.ExecuteNonQuery()

                dbConexion.Close()

                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing

            Next

            Log("Se Inserto con Exito en la tabla CheckLists con el codigo: " & codigo_ini_value, "exito")
            Console.WriteLine("Se Inserto con Exito en la tabla CheckLists.")


        Catch ex As Exception
            Log("Se ha producido un error en la Función checkListValuesByVal con el codigo: " & codigo_ini_value, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Log(ex.Message, "error")
            Console.ReadLine()
            Return "0"
        End Try


salir:
        Return 0
    End Function

    '****************************hist*****************************************
    Private Function checkListHistValuesByVal(fecha_realizacion_value As String,
                                     ByVal codigo_ini_value As String,
                                     ByVal ambito_value As String,
                                     ByVal accion_value As String,
                                     ByVal etapa_value As String,
                                     ByVal hito_value As String,
                                     ByVal preg_value As String,
                                     ByVal resp_value As String,
                                     ByVal obs_value As String,
                                     ByVal tblNom As String) As String

        Try
            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim dbinsert_ini_check_lists As String = ""
            Dim dbinsert_ini_check_lists_hist As String = ""
            Dim fecha_realizacion_aux As String = ""
            Dim codigo_ini_aux As String = ""
            Dim ambito_aux As String = ""
            Dim accion_aux As String = ""
            Dim etapa_aux As String = ""
            Dim hito_aux As String = ""
            Dim preg_aux As String = ""
            Dim resp_aux As String = ""
            Dim obs_aux As String = ""
            Dim cad_null As String = ""
            Dim str_accion As String = ""
            Dim str_etapa As String = ""
            Dim str_obs As String = ""
            Dim dbRowCount As String = ""
            Dim dbresultados As String = ""
            Dim dbdeletecheckLists As String = ""
            Dim cont As Integer = 0

            dbRowCount = "select COUNT(*) AS contador from [dbo].[imp_check_lists] where checkList_cod_ini='" & codigo_ini_value & "'"
            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(dbRowCount, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader

            If dbdata.HasRows = True Then
                dbdata.Read()
                dbresultados = dbdata.Item(0).ToString

                'Si existe Data
                If (dbresultados > 0) Then

                    dbinsert_ini_check_lists_hist = "insert into [dbo].[imp_check_lists_hist] " _
                    & "(checkList_ini_ide,checkList_fecha_realizacion,checkList_cod_ini,checkList_ambito,checkList_accion, " _
                    & "checkList_etapa,checkList_hito, checkList_pregunta, checkList_respuesta, checkList_observaciones,checkList_estado,checkList_fecha_act " _
                    & " )" _
                & " Select checkList_ini_ide, " _
                & " checkList_fecha_realizacion,checkList_cod_ini,checkList_ambito,checkList_accion,checkList_etapa,checkList_hito,checkList_pregunta, " _
                & " checkList_respuesta,checkList_observaciones,1,getdate()  " _
                 & "FROM [dbo].[imp_check_lists] where checkList_cod_ini='" & codigo_ini_value & "'  ORDER BY checkList_ide "

                    'Debug.Print(dbinsert_ini_check_lists_hist)

                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbinsert_ini_check_lists_hist, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbcommand.ExecuteNonQuery()
                    Log("Se Inserto con Exito en la tabla historica CheckLists", "exito")
                    Console.WriteLine("Se Inserto con Exito en la tabla historica CheckLists... ")
                    dbConexion.Close()

                    dbdeletecheckLists = "DELETE FROM [dbo].[imp_check_lists] where checkList_cod_ini='" & codigo_ini_value & "'"
                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbdeletecheckLists, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbcommand.ExecuteNonQuery()
                    Log("Se Elimino con Exito la tabla CheckLists el codigo:" & codigo_ini_value, "exito")
                    Console.WriteLine("Se Elimino con Exito la tabla CheckLists... ")

                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing
                End If
            End If

        Catch ex As Exception
            Log("Se ha producido un error en la funcion checkListHistValuesByVal el codigo:" & codigo_ini_value & ":" & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try


salir:
        Return 0
    End Function
    '************************************************Modelo Check List****************************************************
    Public Sub ModeloCheckList()

        Dim adoCon As Data.OleDb.OleDbConnection
        Dim adoRs As Data.OleDb.OleDbDataReader
        Dim adoCmd As Data.OleDb.OleDbCommand
        Dim strCmnd As String = "SELECT * FROM [CHeckList_ModeloGdeIm$]"



        adoCon = New Data.OleDb.OleDbConnection(GetConnectionString(2))
        adoCmd = New Data.OleDb.OleDbCommand(strCmnd, adoCon)
        adoCmd.CommandType = CommandType.Text
        adoCon.Open()
        adoRs = adoCmd.ExecuteReader

        If adoRs.HasRows Then
            adoRs.Read()
            Do While adoRs.Read()

                '1era Columna del Archivo cabecera Color Naranja'  
                idCheckList = checkListValuesByVal(adoRs.Item(0), adoRs.Item(1), adoRs.Item(2), Convert.ToString(adoRs.Item(3)), Convert.ToString(adoRs.Item(4)), adoRs.Item(5), adoRs.Item(6), adoRs.Item(7), Convert.ToString(adoRs.Item(8)), "imp_check_lists")

                idCheckList_h = checkListHistValuesByVal(adoRs.Item(0), adoRs.Item(1), adoRs.Item(2), Convert.ToString(adoRs.Item(3)), Convert.ToString(adoRs.Item(4)), adoRs.Item(5), adoRs.Item(6), adoRs.Item(7), Convert.ToString(adoRs.Item(8)), "imp_check_lists_hist")

            Loop

        Else
            GoTo salirSinFilas
            Debug.Print("No rows found.")
            Console.ReadLine()

        End If

salirSinFilas:
        adoRs.Close()
        adoCon.Close()

        adoCon = Nothing
        adoRs = Nothing
        adoCmd = Nothing

    End Sub

    '*******************************************************************************************************************************'
    Public Sub ModeloInsertCheckList()

        Dim adoCon As Data.OleDb.OleDbConnection
        Dim adoRs As Data.OleDb.OleDbDataReader
        Dim adoCmd As Data.OleDb.OleDbCommand
        Dim strCmnd As String = "SELECT * FROM [CHeckList_ModeloGdeIm$]"


        adoCon = New Data.OleDb.OleDbConnection(GetConnectionString(2))
        adoCmd = New Data.OleDb.OleDbCommand(strCmnd, adoCon)
        adoCmd.CommandType = CommandType.Text
        adoCon.Open()
        adoRs = adoCmd.ExecuteReader

        If adoRs.HasRows Then
            adoRs.Read()
            Do While adoRs.Read()

                '1era Columna del Archivo cabecera Color Naranja'  
                idCheckList = checkListValuesByVal(adoRs.Item(0), adoRs.Item(1), adoRs.Item(2), Convert.ToString(adoRs.Item(3)), Convert.ToString(adoRs.Item(4)), adoRs.Item(5), adoRs.Item(6), adoRs.Item(7), Convert.ToString(adoRs.Item(8)), "imp_check_lists")

            Loop

        Else
            GoTo salirSinFilas
            Debug.Print("No rows found.")
            Console.ReadLine()
        End If

salirSinFilas:
        adoRs.Close()
        adoCon.Close()

        adoCon = Nothing
        adoRs = Nothing
        adoCmd = Nothing

    End Sub
    '************************************************REUNION PERIODICA**************************************************

    Private Function ReunionPValuesByVal(ByVal id_iniciativa As String,
                                    ByVal cod_Iniciativa As String,
                                    ByVal nomb_archivo_value As String,
                                    ByVal fecha_act_reunion As String,
                                    ByVal id_value As String,
                                    ByVal nivel_esq_value As String,
                                    ByVal nomb_tarea_value As String,
                                    ByVal obs_value As String,
                                    ByVal porc_compl_value As String,
                                    ByVal duracion_value As String,
                                    ByVal comienzo_value As String,
                                    ByVal fin_value As String,
                                    ByVal predec_value As String,
                                    ByVal fecha_act_value As String,
                                    ByVal nomb_actualiz_value As String,
                                    ByVal obj_actualiz_value As String,
                                    ByVal contfilas As Integer,
                                    ByVal tblNom As String) As String


        Try
            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            'Dim cadEnter_cod_ini() As String
            Dim cadEnter_id() As String
            Dim cadEnter_nivel_esq() As String
            'Dim cadEnter_nomb_nivel() As String
            Dim cadEnter_nomb_tarea() As String
            Dim cadEnter_obs() As String
            Dim cadEnter_porc_complet() As String
            Dim cadEnter_duracion() As String
            Dim cadEnter_comienzo() As String
            Dim cadEnter_fin() As String
            Dim cadEnter_predec() As String
            'Dim cadEnter_fecha_act() As String
            Dim cadEnter_nomb_actual() As String
            Dim cadEnter_obj_actual() As String

            Dim dbinsert_ini_planif As String = ""
            Dim dbinsert_ini_planif_hist As String = ""
            Dim cod_ini_aux As String = ""
            Dim id_aux As String = ""
            Dim nivel_esq_aux As String = ""
            Dim nomb_nivel_esq_aux As String = ""
            Dim nomb_tarea_aux As String = ""
            Dim obs_aux As String = ""
            Dim porc_comp_aux As String = ""
            Dim duracion_aux As String = ""
            Dim comienzo_aux As String = ""
            Dim fin_aux As String = ""
            Dim predec_aux As String = ""
            Dim fecha_act_aux As String = ""
            Dim nombre_actualiz_aux As String = ""
            Dim obj_actualiz_aux As String = ""
            Dim dbCodIni As String = ""
            Dim MessageExito As String = ""
            Dim fecha_carga As String = ""


            'cadEnter_cod_ini = cod_inic_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_id = id_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_nivel_esq = nivel_esq_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            'cadEnter_nomb_nivel = nomb_nivel_esq_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_nomb_tarea = nomb_tarea_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_obs = obs_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_porc_complet = porc_compl_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_duracion = duracion_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_comienzo = comienzo_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_fin = fin_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_predec = predec_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            'cadEnter_fecha_act = fecha_act_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_nomb_actual = nomb_actualiz_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_obj_actual = obj_actualiz_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            fecha_carga = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")

            'dbCodIni = "select  count(planif_ini_ide) as valor FROM [dbo].[imp_planificacion_reunion_periodica] where planif_file_arc='" & nomb_archivo_value & "'"
            'Debug.Print(dbCodIni)


            'dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            'dbcommand = New Data.Odbc.OdbcCommand(dbCodIni, dbConexion)
            'dbcommand.CommandType = CommandType.Text
            'dbConexion.Open()
            'dbdata = dbcommand.ExecuteReader

            'If dbdata.HasRows = True Then
            '    dbdata.Read()
            '    cod_ini_aux = dbdata.Item(0).ToString

            'End If

            'If (cod_ini_aux = 0) Then
            For i As Integer = 0 To cadEnter_id.Length - 1

                'cod_ini_aux = cadEnter_cod_ini(i)
                id_aux = cadEnter_id(i)
                nivel_esq_aux = cadEnter_nivel_esq(i)
                'nomb_nivel_esq_aux = cadEnter_nomb_nivel(i)
                nomb_tarea_aux = cadEnter_nomb_tarea(i)
                If (obs_value <> "") Then
                    obs_aux = cadEnter_obs(i)
                End If
                porc_comp_aux = cadEnter_porc_complet(i)
                duracion_aux = cadEnter_duracion(i)
                comienzo_aux = cadEnter_comienzo(i)
                fin_aux = cadEnter_fin(i)

                If (predec_value <> "") Then
                    predec_aux = cadEnter_predec(i)
                End If

                'If (fecha_act_value <> "") Then
                '    fecha_act_aux = cadEnter_fecha_act(i)
                'End If

                If (nomb_actualiz_value <> "") Then
                    nombre_actualiz_aux = cadEnter_nomb_actual(i)
                End If

                If (obj_actualiz_value <> "") Then
                    obj_actualiz_aux = cadEnter_obj_actual(i)
                End If


                dbinsert_ini_planif = "INSERT INTO [dbo].[" & tblNom & "] " _
             & "(planif_ini_ide,planif_cod_inic,planif_id,planif_nivel_esq, " _
             & "planif_nomb_tarea,planif_obs,planif_porct_comp,planif_duracion,planif_comienzo,planif_fin,planif_predecesor,planif_fecha_act,planif_nombre_act,planif_obj_act,planif_file_fec,planif_file_arc) " _
             & "values('" & id_iniciativa & "', '" & cod_Iniciativa & "','" & id_aux & "','" & nivel_esq_aux & "',  " _
             & " '" & nomb_tarea_aux & "','" & obs_aux & "','" & porc_comp_aux & "','" & duracion_aux & "','" & comienzo_aux & "','" & fin_aux & "','" & predec_aux & "','" & fecha_act_reunion & "', " _
             & " '" & nombre_actualiz_aux & "','" & obj_actualiz_aux & "','" & fecha_carga & "','" & nomb_archivo_value & "')"


                'Debug.Print(dbinsert_ini_planif)

                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(dbinsert_ini_planif, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                dbcommand.ExecuteNonQuery()
                dbConexion.Close()
                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing

                MessageExito = "Se Inserto con Exito en la tabla Reunion Periodica el codigo Iniciativa : " & cod_Iniciativa & " - Fila " & contfilas
                Log(MessageExito, "exito")
            Next

            Console.WriteLine("Se Inserto con Exito en la tabla Reunion Periodica el Codigo Iniciativa: " & cod_Iniciativa)

        Catch ex As Exception
            Log("Se ha producido un error  con el codigo iniciativa en la funcion ReunionPValuesByVal: " & cod_Iniciativa & ":" & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try


salir:
        Return 0
    End Function

    '*******************************************************PLANIFICACION*******************************************************************************
    Private Function planifValuesByVal(ByVal id_iniciativa As String,
                                     ByVal cod_Iniciativa As String,
                                     ByVal fecha_carga As String,
                                     ByVal id_value As String,
                                     ByVal nivel_esq_value As String,
                                     ByVal nomb_tarea_value As String,
                                     ByVal obs_value As String,
                                     ByVal porc_compl_value As String,
                                     ByVal duracion_value As String,
                                     ByVal comienzo_value As String,
                                     ByVal fin_value As String,
                                     ByVal predec_value As String,
                                     ByVal fecha_act_value As String,
                                     ByVal nomb_actualiz_value As String,
                                     ByVal obj_actualiz_value As String,
                                     ByVal contfilas As Integer,
                                     ByVal tblNom As String) As String


        Try
            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            'Dim cadEnter_cod_ini() As String
            Dim cadEnter_id() As String
            Dim cadEnter_nivel_esq() As String
            'Dim cadEnter_nomb_nivel() As String
            Dim cadEnter_nomb_tarea() As String
            Dim cadEnter_obs() As String
            Dim cadEnter_porc_complet() As String
            Dim cadEnter_duracion() As String
            Dim cadEnter_comienzo() As String
            Dim cadEnter_fin() As String
            Dim cadEnter_predec() As String
            Dim cadEnter_fecha_act() As String
            Dim cadEnter_nomb_actual() As String
            Dim cadEnter_obj_actual() As String

            Dim dbinsert_ini_planif As String = ""
            Dim dbinsert_ini_planif_hist As String = ""
            Dim cod_ini_aux As String = ""
            Dim id_aux As String = ""
            Dim nivel_esq_aux As String = ""
            Dim nomb_nivel_esq_aux As String = ""
            Dim nomb_tarea_aux As String = ""
            Dim obs_aux As String = ""
            Dim porc_comp_aux As String = ""
            Dim duracion_aux As String = ""
            Dim comienzo_aux As String = ""
            Dim fin_aux As String = ""
            Dim predec_aux As String = ""
            Dim fecha_act_aux As String = ""
            Dim nombre_actualiz_aux As String = ""
            Dim obj_actualiz_aux As String = ""
            Dim dbCodIni As String = ""
            Dim MensajeExito As String = ""

            'cadEnter_cod_ini = cod_inic_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_id = id_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_nivel_esq = nivel_esq_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            'cadEnter_nomb_nivel = nomb_nivel_esq_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_nomb_tarea = nomb_tarea_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_obs = obs_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_porc_complet = porc_compl_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_duracion = duracion_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_comienzo = comienzo_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_fin = fin_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_predec = predec_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_fecha_act = fecha_act_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_nomb_actual = nomb_actualiz_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_obj_actual = obj_actualiz_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)


            dbCodIni = "select ini_cod from [dbo].[imp_iniciativa] where ini_ide='" & id_iniciativa & "'"
            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(dbCodIni, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader

            If dbdata.HasRows = True Then
                dbdata.Read()
                cod_ini_aux = dbdata.Item(0).ToString

            End If


            For i As Integer = 0 To cadEnter_id.Length - 1


                'cod_ini_aux = cadEnter_cod_ini(i)
                id_aux = cadEnter_id(i)
                nivel_esq_aux = cadEnter_nivel_esq(i)
                'nomb_nivel_esq_aux = cadEnter_nomb_nivel(i)
                nomb_tarea_aux = cadEnter_nomb_tarea(i)
                If (obs_value <> "") Then
                    obs_aux = cadEnter_obs(i)
                End If
                porc_comp_aux = cadEnter_porc_complet(i)
                duracion_aux = cadEnter_duracion(i)
                comienzo_aux = cadEnter_comienzo(i)
                fin_aux = cadEnter_fin(i)

                If (predec_value <> "") Then
                    predec_aux = cadEnter_predec(i)
                End If

                If (fecha_act_value <> "") Then
                    fecha_act_aux = cadEnter_fecha_act(i)
                End If

                If (nomb_actualiz_value <> "") Then
                    nombre_actualiz_aux = cadEnter_nomb_actual(i)
                End If

                If (obj_actualiz_value <> "") Then
                    obj_actualiz_aux = cadEnter_obj_actual(i)
                End If


                dbinsert_ini_planif = "INSERT INTO [dbo].[" & tblNom & "] " _
             & "(planif_ini_ide,planif_cod_inic,planif_id,planif_nivel_esq, " _
             & "planif_nomb_tarea,planif_obs,planif_porct_comp,planif_duracion,planif_comienzo,planif_fin,planif_predecesor,planif_fecha_act,planif_nombre_act,planif_obj_act,planif_fecha_realizacion) " _
             & "values('" & id_iniciativa & "', '" & cod_Iniciativa & "','" & id_aux & "','" & nivel_esq_aux & "',  " _
             & " '" & nomb_tarea_aux & "','" & obs_aux & "','" & porc_comp_aux & "','" & duracion_aux & "','" & comienzo_aux & "','" & fin_aux & "','" & predec_aux & "','" & fecha_act_aux & "', " _
             & " '" & nombre_actualiz_aux & "','" & obj_actualiz_aux & "','" & fecha_carga & "')"
                'Debug.Print(dbinsert_ini_check_lists)

                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(dbinsert_ini_planif, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                dbcommand.ExecuteNonQuery()
                dbConexion.Close()
                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing
                MensajeExito = "Se Inserto con Exito en la tabla Planificacion con el Codigo de Iniciativa: " & cod_Iniciativa & " - Fila " & contfilas
                Log(MensajeExito, "exito")
            Next

            Console.WriteLine("Se Inserto con Exito en la tabla Planificacion con el Codigo de Iniciativa: " & cod_Iniciativa)

        Catch ex As Exception
            Log("Se ha producido un error con el Codigo de Iniciativa: " & cod_Iniciativa & ":" & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try


salir:
        Return 0
    End Function

    '******************************************Planificación Historica*******************************************************************************************
    Private Function planifHistValuesByVal(ByVal id_iniciativa As String,
                                           ByVal cod_Iniciativa As String,
                                           ByVal fecha_carga As String,
                                           ByVal tblNom As String) As String

        Try

            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim dbinsert_ini_planif As String = ""
            Dim dbinsert_ini_planif_hist As String = ""
            Dim dbRowCount As String = ""
            Dim dbresultados As String = ""
            Dim dbdeleteplanif As String = ""
            Dim MessageExito As String = ""


            dbinsert_ini_planif_hist = "if exists(select top 1 planif_cod_inic from  imp_planificacion where planif_cod_inic='" & cod_Iniciativa & "') " & _
            "begin " & _
            "declare @hoy datetime; " & _
            "set		@hoy=getdate(); " & _
            "insert imp_planificacion_hist( planif_ini_ide, planif_cod_inic, planif_id, planif_nivel_esq, planif_nomb_tarea, planif_obs, planif_porct_comp, planif_duracion, planif_comienzo, planif_fin, planif_predecesor, planif_fecha_act," & _
            "planif_nombre_act, planif_obj_act, planif_estado, planif_fecha_actual,planif_fecha_realizacion) " & _
            "select	planif_ini_ide, planif_cod_inic, planif_id, planif_nivel_esq, planif_nomb_tarea, planif_obs, planif_porct_comp, planif_duracion, planif_comienzo, planif_fin, planif_predecesor, planif_fecha_act, " & _
            "planif_nombre_act, planif_obj_act,1,@hoy,planif_fecha_realizacion " & _
            "from imp_planificacion " & _
            "where	planif_cod_inic='" & cod_Iniciativa & "' " & _
            "ORDER BY planif_ini_ide " & _
            " " & _
            "delete from imp_planificacion where planif_cod_inic='" & cod_Iniciativa & "' " & _
            "end"

            'Console.WriteLine(dbinsert_ini_planif_hist)

            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(dbinsert_ini_planif_hist, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbcommand.ExecuteNonQuery()
            Console.WriteLine("Se Inserto con Exito en la tabla Historica de Planificacion con el Codigo Iniciativa: " & cod_Iniciativa)
            dbConexion.Close()

            dbConexion = Nothing
            dbcommand = Nothing
            dbdata = Nothing

            MessageExito = "Se Inserto con Exito en la tabla Historica de Planificacion con el Codigo Iniciativa: " & cod_Iniciativa
            Log(MessageExito, "exito")

        Catch ex As Exception
            Log("Se ha producido un error en la funcion planifHistValuesByVal con el codigo iniciativa: " & cod_Iniciativa & ":" & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try

    End Function

    '***************************************Modelo Planificación*******************************************************
    Public Sub ModeloPlanificacion(idIniciativa, codIniciativa, fecha_carga)

        Dim adoCon As Data.OleDb.OleDbConnection
        Dim adoRs As Data.OleDb.OleDbDataReader
        Dim adoCmd As Data.OleDb.OleDbCommand
        Dim strCmnd As String = "SELECT * FROM [Hoja1$]"
        Dim contfilas As Integer = 0

        'Console.Write(value)
        Try

            adoCon = New Data.OleDb.OleDbConnection(GetConnectionString((3)))
            adoCmd = New Data.OleDb.OleDbCommand(strCmnd, adoCon)
            adoCmd.CommandType = CommandType.Text
            adoCon.Open()
            adoRs = adoCmd.ExecuteReader

            If adoRs.HasRows Then
                adoRs.Read()
                Do While adoRs.Read()
                    contfilas = contfilas + 1
                    idplanific = planifValuesByVal(idIniciativa, codIniciativa, fecha_carga, Convert.ToString(adoRs.Item(0)).Trim, Convert.ToString(adoRs.Item(1)).Trim, CleanInput(Convert.ToString(adoRs.Item(2))).Trim, Convert.ToString(adoRs.Item(3)).Trim, Convert.ToString(adoRs.Item(4)).Trim, Convert.ToString(adoRs.Item(5)).Trim, Convert.ToString(adoRs.Item(6)).Trim, Convert.ToString(adoRs.Item(7)), Convert.ToString(adoRs.Item(8)).Trim, Convert.ToString(adoRs.Item(9)).Trim, Convert.ToString(adoRs.Item(10)).Trim, Convert.ToString(adoRs.Item(11)).Trim, contfilas, "imp_planificacion")
                Loop
            Else
                GoTo salirSinFilas
                Debug.Print("No rows found.")
            End If

            adoRs.Close()
            adoCon.Close()
        Catch ex As Exception
            Log("No es posible procesar el archivo de planificacion asociado al código : " & codIniciativa & ":" & ex.Message, "error")
            Console.WriteLine("No es posible procesar el archivo de planificacion asociado al código : " & codIniciativa)
            Console.WriteLine("Error: " & ex.Message)
            Console.WriteLine("")
            Console.ReadLine()
        End Try

salirSinFilas:

        adoCon = Nothing
        adoRs = Nothing
        adoCmd = Nothing

    End Sub

    '*************************************************ModeloPlanificacion_hist**************************************************************
    Public Sub ModeloPlanificacion_hist(idIniciativa, codIniciativa, fecha_carga)
        Try
            idplanific_h = planifHistValuesByVal(idIniciativa, codIniciativa, fecha_carga, "imp_planificacion_hist")
        Catch ex As Exception
            Log("Se ha producido un error en la funcion ModeloPlanificacion_hist " & ex.Message, "error")
            Console.ReadLine()
        End Try



    End Sub

    '******************************************************Modelo Insert Planificación************************************************************
    '    Public Sub ModeloInsertPlanifacion(idIniciativa, codIniciativa)

    '        Dim adoCon As Data.OleDb.OleDbConnection
    '        Dim adoRs As Data.OleDb.OleDbDataReader
    '        Dim adoCmd As Data.OleDb.OleDbCommand
    '        Dim strCmnd As String = "SELECT * FROM [Hoja1$]"

    '        'Console.Write(value)
    '        adoCon = New Data.OleDb.OleDbConnection(GetConnectionString(3))
    '        adoCmd = New Data.OleDb.OleDbCommand(strCmnd, adoCon)
    '        adoCmd.CommandType = CommandType.Text
    '        adoCon.Open()
    '        adoRs = adoCmd.ExecuteReader

    '        If adoRs.HasRows Then
    '            adoRs.Read()
    '            Do While adoRs.Read()

    '                idplanific = planifValuesByVal(idIniciativa, codIniciativa, Convert.ToString(adoRs.Item(0)).Trim, Convert.ToString(adoRs.Item(1)).Trim, CleanInput(Convert.ToString(adoRs.Item(2))).Trim, Convert.ToString(adoRs.Item(3)).Trim, Convert.ToString(adoRs.Item(4)).Trim, Convert.ToString(adoRs.Item(5)).Trim, Convert.ToString(adoRs.Item(6)).Trim, Convert.ToString(adoRs.Item(7)).Trim, Convert.ToString(adoRs.Item(8)).Trim, Convert.ToString(adoRs.Item(9)).Trim, Convert.ToString(adoRs.Item(10)).Trim, Convert.ToString(adoRs.Item(11)).Trim, "imp_planificacion")

    '            Loop
    '        Else
    '            GoTo salirSinFilas
    '            Debug.Print("No rows found.")
    '            Console.ReadLine()
    '        End If

    'salirSinFilas:
    '        adoRs.Close()
    '        adoCon.Close()

    '        adoCon = Nothing
    '        adoRs = Nothing
    '        adoCmd = Nothing

    'End Sub

    '****************************************************************************************************************
    Private Function temaRelevanteHistValuesByVal(ByVal cod_Iniciativa As String,
                                      ByVal tblNom As String) As String

        Try

            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim dbinsert_tema_relevante As String = ""
            Dim dbRowCount As String = ""
            Dim dbresultados As String = ""
            Dim dbdeleteplanif As String = ""
            Dim dbresult As String = ""


            dbinsert_tema_relevante = "if exists(select top 1 tema_ini_ide from  imp_temas_relevantes where tema_cod_ini='" & cod_Iniciativa & "') " & _
            "begin " & _
            "declare @hoy datetime; " & _
            "set		@hoy=getdate(); " & _
            "insert imp_temas_relevantes_hist(tema_ini_ide, tema_cod_ini,tema_relevante, fecha_relevante, fecha_actual) " & _
             "select tema_ini_ide, tema_cod_ini, tema_relevante, fecha_relevante, @hoy " & _
            "from imp_temas_relevantes " & _
            "where	tema_cod_ini='" & cod_Iniciativa & "' " & _
            "ORDER BY tema_ini_ide " & _
            " " & _
            "DELETE FROM imp_temas_relevantes where tema_cod_ini='" & cod_Iniciativa & "' " & _
            "end"

            'Console.WriteLine(dbinsert_ini_planif_hist)

            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(dbinsert_tema_relevante, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()

            dbdata = dbcommand.ExecuteReader

            If dbdata.HasRows = True Then
                dbdata.Read()
                dbresult = dbdata.Item(0).ToString

            End If

            If (dbresult <> "") Then
                dbcommand.ExecuteNonQuery()
                Console.WriteLine("Se Inserto con Exito en la tabla Historica de Temas Relevantes con el Codigo Iniciativa: " & cod_Iniciativa)
                dbConexion.Close()

                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing
            End If

            Log("Se Inserto con Exito en la tabla Historica de Temas Relevantes con el Codigo Iniciativa: " & cod_Iniciativa, "exito")


        Catch ex As Exception
            Log("Se ha producido un error en la funcion temaRelevanteHistValuesByVal con el Codigo Iniciativa: " & cod_Iniciativa & ":" & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try


    End Function

    '*************************************************************REUNIÓN PERIODICA***************************************

    Public Sub ModeloReunionPeriodica(idIniciativa, codIniciativa, nombArchivo, fecha_actualizacion_reunion)

        Dim adoCon As New Data.OleDb.OleDbConnection
        Dim adoRs As Data.OleDb.OleDbDataReader
        Dim adoCmd As Data.OleDb.OleDbCommand
        Dim strCmnd As String = "SELECT * FROM [Hoja1$]"

        Dim dbConexion As New Data.Odbc.OdbcConnection
        Dim dbcommand As Data.Odbc.OdbcCommand
        Dim dbdata As Data.Odbc.OdbcDataReader
        Dim dbConsulta As String
        Dim dbResultado As String
        Dim contfilas As Integer = 0

        ' Revisar que no existe el archivo como registro en la tabla para poder ingresar los datos    

        Try
            dbConsulta = "select  isnull(count(planif_ini_ide),0) [valor] FROM [dbo].[imp_planificacion_reunion_periodica] where planif_file_arc='" & nombArchivo & "'"
            Debug.Print(dbConsulta)

            If dbConexion.State = ConnectionState.Open Then dbConexion.Close()
            If (dbdata) Is Nothing Then
            Else
                If Not dbdata.IsClosed Then dbdata.Close()
            End If

            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(dbConsulta, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader

            If dbdata.HasRows Then
                dbdata.Read()
                If dbdata.Item(0).ToString = "0" Then
                Else
                    Log("Archivo " & nombArchivo & " de reunion periodica asociado al código : " & codIniciativa & ", se encuentra cargado en el sistema", "exito")
                    Console.WriteLine("Archivo " & nombArchivo & " de reunion periodica asociado al código : " & codIniciativa & ", se encuentra cargado en el sistema")

                    dbConexion = Nothing
                    dbdata = Nothing
                    dbcommand = Nothing

                    Exit Sub
                End If
            End If

            dbConexion = Nothing
            dbdata = Nothing
            dbcommand = Nothing

        Catch ex As Exception
            Log("No es posible conectar a la base de datos o encontrar el registros asociados al " & nombArchivo & " de reunion periodica asociado al código : " & codIniciativa, "error")
            Console.WriteLine("No es posible conectar a la base de datos o encontrar el registros asociados al " & nombArchivo & " de reunion periodica asociado al código : " & codIniciativa)
            Console.WriteLine("Error: " & ex.Message)
            Console.WriteLine("")
            Console.ReadLine()

            dbConexion = Nothing
            dbdata = Nothing
            dbcommand = Nothing

            Exit Sub
        End Try

        Try

            'Console.Write(value)
            If adoCon.State = ConnectionState.Open Then adoCon.Close()
            If (adoRs) Is Nothing Then
            Else
                If Not adoRs.IsClosed Then adoRs.Close()
            End If

            adoCon = New Data.OleDb.OleDbConnection(GetConnectionString(8))
            adoCmd = New Data.OleDb.OleDbCommand(strCmnd, adoCon)
            adoCmd.CommandType = CommandType.Text
            adoCon.Open()
            adoRs = adoCmd.ExecuteReader

            If adoRs.HasRows Then
                adoRs.Read()
                Do While adoRs.Read()

                    contfilas = contfilas + 1
                    idReunionP = ReunionPValuesByVal(idIniciativa, codIniciativa, nombArchivo, fecha_actualizacion_reunion, Convert.ToString(adoRs.Item(0)).Trim, Convert.ToString(adoRs.Item(1)).Trim, CleanInput(Convert.ToString(adoRs.Item(2))).Trim, Convert.ToString(adoRs.Item(3)).Trim, Convert.ToString(adoRs.Item(4)).Trim, Convert.ToString(adoRs.Item(5)).Trim, Convert.ToString(adoRs.Item(6)).Trim, Convert.ToString(adoRs.Item(7)).Trim, Convert.ToString(adoRs.Item(8)).Trim, Convert.ToString(adoRs.Item(9)).Trim, Convert.ToString(adoRs.Item(10)).Trim, Convert.ToString(adoRs.Item(11)).Trim, contfilas, "imp_planificacion_reunion_periodica")
                Loop
            Else
                GoTo salirSinFilas
                Debug.Print("No rows found.")
            End If

            adoRs.Close()
            adoCon.Close()
        Catch ex As Exception
            Log("No es posible procesar el archivo " & nombArchivo & " de reunion periodica asociado al código : " & codIniciativa, "error")
            Console.WriteLine("No es posible procesar el archivo " & nombArchivo & " de reunion periodica asociado al código : " & codIniciativa)
            Console.WriteLine("Error: " & ex.Message)
            Console.WriteLine("")
            Console.ReadLine()
        End Try


salirSinFilas:

        adoCon = Nothing
        adoRs = Nothing
        adoCmd = Nothing

    End Sub

    '***********************************************************************************************
    Private Function contactoValuesByVal(nombre_value As String,
                                 ByVal alias_value As String,
                                 ByVal gerencia_value As String,
                                 ByVal sub_gerencia_value As String,
                                 ByVal area_value As String,
                                 ByVal unid_value As String,
                                 ByVal rol_value As String,
                                 ByVal desc_rol_value As String,
                                 ByVal accion_apoyo_value As String,
                                 ByVal ambito_1_value As String,
                                 ByVal ambito_2_value As String,
                                 ByVal ambito_3_value As String,
                                 ByVal posee_hit_asoc As String,
                                 ByVal hitos_if_asoc_value As String,
                                 ByVal palab_claves_value As String,
                                 ByVal tblNom As String) As String


        Try
            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim cadEnter_nombre() As String
            Dim cadEnter_alias() As String
            Dim cadEnter_gerencia() As String
            Dim cadEnter_sub_gerencia() As String
            Dim cadEnter_area() As String
            Dim cadEnter_unid() As String
            Dim cadEnter_rol() As String
            Dim cadEnter_desc_rol() As String
            Dim cadEnter_accion_apoyo() As String
            Dim cadEnter_ambito_1() As String
            Dim cadEnter_ambito_2() As String
            Dim cadEnter_ambito_3() As String
            Dim cadEnter_posee_hit_asoc() As String
            Dim cadEnter_hitos_if_asoc() As String
            Dim cadEnter_palab_claves() As String

            Dim dbinsert_ini_contact As String = ""
            Dim dbinsert_ini_contact_hist As String = ""
            Dim nombre_aux As String = ""
            Dim alias_aux As String = ""
            Dim gerencia_aux As String = ""
            Dim sub_gerencia_aux As String = ""
            Dim area_aux As String = ""
            Dim unidad_aux As String = ""
            Dim rol_aux As String = ""
            Dim desc_rol_aux As String = ""
            Dim accion_apoyo_aux As String = ""
            Dim ambito1_aux As String = ""
            Dim ambito2_aux As String = ""
            Dim ambito3_aux As String = ""
            Dim posee_hitos_aux As String = ""
            Dim hito_if_asoc_aux As String = ""
            Dim palab_claves_aux As String = ""
            Dim stringSeparators() As String = {","c}


            cadEnter_nombre = nombre_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_alias = alias_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_gerencia = gerencia_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_sub_gerencia = sub_gerencia_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_area = area_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_unid = unid_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_rol = rol_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_desc_rol = desc_rol_value.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries)

            cadEnter_accion_apoyo = accion_apoyo_value.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries)
            cadEnter_ambito_1 = ambito_1_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_ambito_2 = ambito_2_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_ambito_3 = ambito_3_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_posee_hit_asoc = posee_hit_asoc.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_hitos_if_asoc = hitos_if_asoc_value.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries)

            cadEnter_palab_claves = palab_claves_value.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries)

            For i As Integer = 0 To cadEnter_nombre.Length - 1

                nombre_aux = cadEnter_nombre(i)

                If (alias_value <> "") Then
                    alias_aux = cadEnter_alias(i)
                End If


                If (gerencia_value <> "") Then
                    gerencia_aux = cadEnter_gerencia(i)
                End If

                If (sub_gerencia_value <> "") Then
                    sub_gerencia_aux = cadEnter_sub_gerencia(i)
                End If


                If (area_value <> "") Then
                    area_aux = cadEnter_area(i)
                End If

                If (unid_value <> "") Then
                    unidad_aux = cadEnter_unid(i)
                End If

                If (rol_value <> "") Then
                    rol_aux = cadEnter_rol(i)
                End If

                If (desc_rol_value <> "") Then
                    desc_rol_aux = cadEnter_desc_rol(i)

                End If

                If (accion_apoyo_value <> "") Then
                    accion_apoyo_aux = cadEnter_accion_apoyo(i)
                End If


                If (ambito_1_value <> "") Then
                    ambito1_aux = cadEnter_ambito_1(i)
                End If

                If (ambito_2_value <> "") Then
                    ambito2_aux = cadEnter_ambito_2(i)
                End If

                If (ambito_3_value <> "") Then
                    ambito3_aux = cadEnter_ambito_3(i)
                End If

                If (posee_hit_asoc <> "") Then
                    posee_hitos_aux = cadEnter_posee_hit_asoc(i)
                End If

                If (hitos_if_asoc_value <> "") Then
                    hito_if_asoc_aux = cadEnter_hitos_if_asoc(i)

                End If

                If (palab_claves_value <> "") Then
                    palab_claves_aux = cadEnter_palab_claves(i)

                End If


                dbinsert_ini_contact = "INSERT INTO [dbo].[" & tblNom & "] " _
             & "(contacto_ini_ide,contacto_nombre,contacto_alias,contacto_gerencia,contacto_subgerencia, " _
             & "contacto_area,contacto_unidad,contacto_rol,contacto_desc_rol,contacto_accion_apoyo,contacto_ambito_1,contacto_ambito_2,contacto_ambito_3,contacto_posee_hitos,contacto_if_asoc,contacto_palab_claves) " _
             & "values('" & id_iniCod & "', '" & nombre_aux & "','" & alias_aux & "','" & gerencia_aux & "','" & sub_gerencia_aux & "',  " _
             & " '" & area_aux & "','" & unidad_aux & "','" & rol_aux & "','" & desc_rol_aux & "','" & accion_apoyo_aux & "','" & ambito1_aux & "','" & ambito2_aux & "','" & ambito3_aux & "', " _
             & " '" & posee_hitos_aux & "','" & hito_if_asoc_aux & "', '" & palab_claves_aux & "')"



                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(dbinsert_ini_contact, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                dbcommand.ExecuteNonQuery()
                dbConexion.Close()
                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing

            Next
            Console.WriteLine("Se Inserto con Exito en la tabla Contactos... ")

        Catch ex As Exception
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try


salir:
        Return 0
    End Function
    '***************************************************************************************************
    Private Function contactoHistValuesByVal(nombre_value As String,
                                 ByVal alias_value As String,
                                 ByVal gerencia_value As String,
                                 ByVal sub_gerencia_value As String,
                                 ByVal area_value As String,
                                 ByVal unid_value As String,
                                 ByVal rol_value As String,
                                 ByVal desc_rol_value As String,
                                 ByVal accion_apoyo_value As String,
                                 ByVal ambito_1_value As String,
                                 ByVal ambito_2_value As String,
                                 ByVal ambito_3_value As String,
                                 ByVal posee_hit_asoc As String,
                                 ByVal hitos_if_asoc_value As String,
                                 ByVal palab_claves_value As String,
                                 ByVal tblNom As String) As String


        Try


            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim dbinsert_ini_contact_hist As String = ""
            Dim dbRowCount As String = ""
            Dim dbresultados As String = ""
            Dim dbdeleteContact As String = ""


            dbRowCount = "select COUNT(*) AS contador from [dbo].[imp_contactosIF]"
            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(dbRowCount, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader

            If dbdata.HasRows = True Then
                dbdata.Read()
                dbresultados = dbdata.Item(0).ToString

                'Si existe Data
                If (dbresultados > 0) Then

                    dbinsert_ini_contact_hist = "insert into [dbo].[" & tblNom & "]" _
                    & "(contacto_ini_ide,contacto_nombre,contacto_alias,contacto_gerencia,contacto_subgerencia, " _
                    & "contacto_area,contacto_unidad,contacto_rol,contacto_desc_rol,contacto_accion_apoyo,contacto_ambito_1,contacto_ambito_2,contacto_ambito_3,contacto_posee_hitos,contacto_if_asoc,contacto_palab_claves,contacto_estado,contacto_fecha_actual" _
                    & " )" _
                & " Select contacto_ini_ide,contacto_nombre, contacto_alias," _
                & " contacto_gerencia,contacto_subgerencia,contacto_area,contacto_unidad,contacto_rol,contacto_desc_rol,contacto_accion_apoyo, " _
                & " contacto_ambito_1,contacto_ambito_2,contacto_ambito_3,contacto_posee_hitos,contacto_if_asoc,contacto_palab_claves,1,getdate()  " _
                 & "FROM [dbo].[imp_contactosIF] "

                    'Debug.Print(dbinsert_ini_contact_hist)

                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbinsert_ini_contact_hist, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbcommand.ExecuteNonQuery()
                    Console.WriteLine("Se Inserto con Exito en la tabla historica Contactos... ")
                    dbConexion.Close()

                    dbdeleteContact = "DELETE FROM [dbo].[imp_contactosIF] where contacto_ini_ide = '" & id_iniCod & "'"
                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbdeleteContact, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbcommand.ExecuteNonQuery()
                    Console.WriteLine("Se Elimino con Exito la tabla Contactos... ")

                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing
                End If
            End If


        Catch ex As Exception
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try

    End Function
    '**************************************************************************************************************
    Public Sub ModeloContacto()

        Dim adoCon As Data.OleDb.OleDbConnection
        Dim adoRs As Data.OleDb.OleDbDataReader
        Dim adoCmd As Data.OleDb.OleDbCommand
        Dim strCmnd As String = "SELECT * FROM [ContactosIF$]"


        adoCon = New Data.OleDb.OleDbConnection(GetConnectionString(4))
        adoCmd = New Data.OleDb.OleDbCommand(strCmnd, adoCon)
        adoCmd.CommandType = CommandType.Text
        adoCon.Open()
        adoRs = adoCmd.ExecuteReader

        If adoRs.HasRows Then
            adoRs.Read()
            Do While adoRs.Read()

                idContactoIF = contactoValuesByVal(adoRs.Item(0), Convert.ToString(adoRs.Item(1)), Convert.ToString(adoRs.Item(2)), Convert.ToString(adoRs.Item(3)), Convert.ToString(adoRs.Item(4)), Convert.ToString(adoRs.Item(5)), Convert.ToString(adoRs.Item(6)), Convert.ToString(adoRs.Item(7)), Convert.ToString(adoRs.Item(8)), Convert.ToString(adoRs.Item(9)), Convert.ToString(adoRs.Item(10)), Convert.ToString(adoRs.Item(11)), Convert.ToString(adoRs.Item(12)), Convert.ToString(adoRs.Item(13)), Convert.ToString(adoRs.Item(14)), "imp_contactosIF")

                idContactoIF_h = contactoHistValuesByVal(adoRs.Item(0), Convert.ToString(adoRs.Item(1)), Convert.ToString(adoRs.Item(2)), Convert.ToString(adoRs.Item(3)), Convert.ToString(adoRs.Item(4)), Convert.ToString(adoRs.Item(5)), Convert.ToString(adoRs.Item(6)), Convert.ToString(adoRs.Item(7)), Convert.ToString(adoRs.Item(8)), Convert.ToString(adoRs.Item(9)), Convert.ToString(adoRs.Item(10)), Convert.ToString(adoRs.Item(11)), Convert.ToString(adoRs.Item(12)), Convert.ToString(adoRs.Item(13)), Convert.ToString(adoRs.Item(14)), "imp_contactosIF_hist")

            Loop

        Else
            GoTo salirSinFilas
            Debug.Print("No rows found.")
            Console.ReadLine()
        End If

salirSinFilas:
        adoRs.Close()
        adoCon.Close()

        adoCon = Nothing
        adoRs = Nothing
        adoCmd = Nothing

    End Sub

    '************************************************************************************************
    Public Sub ModeloInsertContacto()

        Dim adoCon As Data.OleDb.OleDbConnection
        Dim adoRs As Data.OleDb.OleDbDataReader
        Dim adoCmd As Data.OleDb.OleDbCommand
        Dim strCmnd As String = "SELECT * FROM [ContactosIF$]"

        adoCon = New Data.OleDb.OleDbConnection(GetConnectionString(4))
        adoCmd = New Data.OleDb.OleDbCommand(strCmnd, adoCon)
        adoCmd.CommandType = CommandType.Text
        adoCon.Open()
        adoRs = adoCmd.ExecuteReader

        If adoRs.HasRows Then
            adoRs.Read()
            Do While adoRs.Read()

                idContactoIF = contactoValuesByVal(adoRs.Item(0), Convert.ToString(adoRs.Item(1)), Convert.ToString(adoRs.Item(2)), Convert.ToString(adoRs.Item(3)), Convert.ToString(adoRs.Item(4)), Convert.ToString(adoRs.Item(5)), Convert.ToString(adoRs.Item(6)), Convert.ToString(adoRs.Item(7)), Convert.ToString(adoRs.Item(8)), Convert.ToString(adoRs.Item(9)), Convert.ToString(adoRs.Item(10)), Convert.ToString(adoRs.Item(11)), Convert.ToString(adoRs.Item(12)), Convert.ToString(adoRs.Item(13)), Convert.ToString(adoRs.Item(14)), "imp_contactosIF")
            Loop

        Else
            GoTo salirSinFilas
            Debug.Print("No rows found.")
        End If

salirSinFilas:
        adoRs.Close()
        adoCon.Close()

        adoCon = Nothing
        adoRs = Nothing
        adoCmd = Nothing

    End Sub
    '**********************************************************************************************************************'
    Private Function ItemSegValuesByVal(ByVal ideIniciativa As String,
                                         ByVal CodIniciativa As String,
                                   ByVal cat2_value As String,
                                 ByVal fecha_ing_sist_value As String,
                                 ByVal correl_cod_proy_value As String,
                                 ByVal cod_int_value As String,
                                 ByVal compromisos_value As String,
                                 ByVal resp_comp_value As String,
                                 ByVal fecha_venc_value As String,
                                 ByVal fecha_replanif_value As String,
                                 ByVal cantidad_replanif As String,
                                 ByVal fecha_cierre_value As String,
                                 ByVal condicion_value As String,
                                 ByVal edo_comp_value As String,
                                 ByVal coment_item_seg_value As String,
                                 ByVal hoja As String,
                                 ByVal contfilas As Integer,
                             ByVal tblNom As String) As String


        Try
            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim cadEnter_cat2() As String
            Dim cadEnter_fecha_ing_sist() As String
            Dim cadEnter_correl_cod_proy() As String
            Dim cadEnter_cod_int() As String
            'Dim cadEnter_compromisos() As String
            Dim cadEnter_resp_comp() As String
            Dim cadEnter_fecha_venc() As String
            Dim cadEnter_fecha_replanif() As String
            Dim cadEnter_cantidad_replanif() As String
            Dim cadEnter_fecha_cierre() As String
            Dim cadEnter_condicion() As String
            Dim cadEnter_edo_comp() As String
            Dim dbinsert_item_seg As String = ""
            Dim cat2_aux As String = ""
            Dim fecha_ing_aux As String = ""
            Dim correl_cod_proy_aux As String = ""
            Dim cod_int_aux As String = ""
            Dim compromisos_aux As String = ""
            Dim resp_comp_aux As String = ""
            Dim fecha_venc_aux As String = ""
            Dim fecha_replanif_aux As String = ""
            Dim cantidad_replanif_aux As String = ""
            Dim fecha_cierre_comp_aux As String = ""
            Dim condicion_aux As String = ""
            Dim edo_comprom_aux As String = ""
            Dim coment_item_seg_aux As String = ""
            Dim stringSeparators() As String = {"\"c}
            Dim codigo_ini As String = ""
            Dim cadLoop As Integer = 0
            Dim position As Integer = hoja.IndexOf("$")
            Dim name As String = ""
            Dim query As String = ""
            Dim MessageExito As String = ""
            Dim fecha_condicion As String = ""


            'fecha_venc_value.Split("/", "-")
            If (cod_int_value <> "") Then
                'If (fecha_venc_value <> "-") And (fecha_venc_value <> "TBD") And (fecha_venc_value <> "") Then
                '    Dim myDateTime() As String = Regex.Split(fecha_venc_value, "\ - ")
                '    Dim dia As String = myDateTime(0)
                '    Dim mes As String = myDateTime(1)
                '    Dim anio As String = myDateTime(2)
                '    fecha_condicion = String.Concat(anio, "/", mes, "/", dia)

                'End If

                name = hoja.Substring(0, position)

                cadEnter_cat2 = cat2_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_fecha_ing_sist = fecha_ing_sist_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_correl_cod_proy = correl_cod_proy_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_cod_int = cod_int_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                'cadEnter_compromisos = compromisos_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_resp_comp = resp_comp_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_fecha_venc = fecha_venc_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_fecha_replanif = fecha_replanif_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_cantidad_replanif = cantidad_replanif.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries)
                cadEnter_fecha_cierre = fecha_cierre_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_condicion = condicion_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_edo_comp = edo_comp_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                'cadEnter_coment_item_seg = coment_item_seg_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)

                If (fecha_replanif_value = "") Then
                    cadLoop = cadEnter_fecha_ing_sist.Length - 1
                Else
                    cadLoop = cadEnter_fecha_replanif.Length - 1
                End If

                condicion_aux = ""

                For i As Integer = 0 To cadLoop

                    If (cadEnter_fecha_replanif.Length > 1) Then

                        'query = "SELECT DATEDIFF(dd, '" & DateTime.Now.ToString("yyyy-MM-dd") & "', '" & fecha_venc_value & "'); "
                        'dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                        'dbcommand = New Data.Odbc.OdbcCommand(query, dbConexion)
                        'dbcommand.CommandType = CommandType.Text
                        'dbConexion.Open()
                        'dbdata = dbcommand.ExecuteReader

                        'If dbdata.HasRows = True Then
                        '    dbdata.Read()
                        '    dbresultados = dbdata.Item(0).ToString

                        '    If (dbresultados < 1) Then

                        '        condicion_aux = "Rojo"
                        '    Else
                        '        If (dbresultados > 1) And (dbresultados <= 3) Then

                        '            condicion_aux = "Amarillo"
                        '        Else
                        '            If (dbresultados > 3) Then

                        '                condicion_aux = "Verde"

                        '            End If

                        '        End If
                        '    End If
                        'End If

                        If (cat2_value <> "") Then
                            cat2_aux = cadEnter_cat2(0)
                        Else
                            cat2_aux = "-"

                        End If


                        If (fecha_ing_sist_value <> "") Then
                            fecha_ing_aux = cadEnter_fecha_ing_sist(0)
                        Else
                            fecha_ing_aux = ""
                        End If


                        If (correl_cod_proy_value <> "") Then
                            correl_cod_proy_aux = cadEnter_correl_cod_proy(0)
                        Else
                            correl_cod_proy_aux = ""
                        End If

                        If (cod_int_value <> "") Then
                            cod_int_aux = cadEnter_cod_int(0)
                        Else
                            cod_int_aux = ""

                        End If


                        If (compromisos_value <> "") Then
                            compromisos_aux = compromisos_value
                        Else
                            compromisos_aux = ""
                        End If


                        If (resp_comp_value <> "") Then
                            resp_comp_aux = cadEnter_resp_comp(0)
                        Else
                            resp_comp_aux = ""
                        End If


                        If (fecha_venc_value <> "") Then

                            fecha_venc_aux = cadEnter_fecha_venc(0)
                        Else
                            fecha_venc_aux = ""

                        End If


                        If (fecha_replanif_value <> "") Then
                            fecha_replanif_aux = cadEnter_fecha_replanif(i)
                        Else
                            fecha_replanif_aux = ""
                        End If

                        If (cantidad_replanif <> "") Then
                            cantidad_replanif_aux = cadEnter_cantidad_replanif(0)
                        Else
                            cantidad_replanif_aux = ""
                        End If


                        If (fecha_cierre_value <> "") Then
                            fecha_cierre_comp_aux = cadEnter_fecha_cierre(0)
                        Else
                            fecha_cierre_comp_aux = ""
                        End If


                        If (edo_comp_value <> "") Then
                            edo_comprom_aux = cadEnter_edo_comp(0)
                        Else
                            edo_comprom_aux = ""
                        End If

                        If (coment_item_seg_value <> "") Then
                            coment_item_seg_aux = coment_item_seg_value
                        Else
                            coment_item_seg_aux = ""
                        End If
                    Else

                        'query = "SELECT DATEDIFF(dd, '" & DateTime.Now.ToString("yyyy-MM-dd") & "', '" & fecha_condicion & "'); "
                        'query = "SELECT DATEDIFF(dd,  '" & fecha_venc_value & "','" & DateTime.Now.ToString("yyyy-MM-dd") & "'); "
                        'Debug.Print(query)

                        'dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                        'dbcommand = New Data.Odbc.OdbcCommand(query, dbConexion)
                        'dbcommand.CommandType = CommandType.Text
                        'dbConexion.Open()
                        'dbdata = dbcommand.ExecuteReader

                        'If dbdata.HasRows = True Then
                        '    dbdata.Read()
                        '    dbresultados = dbdata.Item(0).ToString

                        '    If (dbresultados < 1) Then

                        '        condicion_aux = "Rojo"
                        '    Else
                        '        If (dbresultados > 1) And (dbresultados <= 3) Then

                        '            condicion_aux = "Amarillo"
                        '        Else
                        '            If (dbresultados > 3) Then

                        '                condicion_aux = "Verde"

                        '            End If

                        '        End If
                        '    End If
                        'End If


                        If (cat2_value <> "") Then
                            cat2_aux = cadEnter_cat2(i)
                        Else
                            cat2_aux = ""

                        End If


                        If (fecha_ing_sist_value <> "") Then
                            fecha_ing_aux = cadEnter_fecha_ing_sist(i)
                        Else
                            fecha_ing_aux = ""
                        End If


                        If (correl_cod_proy_value <> "") Then
                            correl_cod_proy_aux = cadEnter_correl_cod_proy(i)
                        Else
                            correl_cod_proy_aux = ""
                        End If

                        If (cod_int_value <> "") Then
                            cod_int_aux = cadEnter_cod_int(i)
                        Else
                            cod_int_aux = ""

                        End If


                        If (compromisos_value <> "") Then
                            compromisos_aux = compromisos_value
                        Else
                            compromisos_aux = ""
                        End If


                        If (resp_comp_value <> "") Then
                            resp_comp_aux = cadEnter_resp_comp(i)
                        Else
                            resp_comp_aux = ""
                        End If


                        If (fecha_venc_value <> "") Then

                            fecha_venc_aux = cadEnter_fecha_venc(i)
                        Else
                            fecha_venc_aux = ""

                        End If


                        If (fecha_replanif_value <> "") Then
                            fecha_replanif_aux = cadEnter_fecha_replanif(i)
                        Else
                            fecha_replanif_value = ""
                        End If

                        If (cantidad_replanif <> "") Then
                            cantidad_replanif_aux = cadEnter_cantidad_replanif(i)
                        Else
                            cantidad_replanif_aux = ""
                        End If

                        If (fecha_cierre_value <> "") Then
                            fecha_cierre_comp_aux = cadEnter_fecha_cierre(i)
                        Else
                            fecha_cierre_comp_aux = ""
                        End If


                        If (edo_comp_value <> "") Then
                            edo_comprom_aux = cadEnter_edo_comp(i)
                        Else
                            edo_comprom_aux = ""
                        End If

                        If (coment_item_seg_value <> "") Then
                            coment_item_seg_aux = coment_item_seg_value
                        Else
                            coment_item_seg_aux = ""
                        End If

                    End If

                    dbinsert_item_seg = "INSERT INTO [dbo].[" & tblNom & "] " _
                 & "(item_seg_cat,item_seg_cod_ini,item_seg_fecha_ing,item_seg_correl,item_seg_cod_int,item_seg_compromisos, " _
                 & "item_seg_resp_comp,item_seg_fecha_venc,item_seg_fecha_replanif,item_seg_cant_replanif,item_seg_fecha_cierre,item_seg_condicion,item_seg_edo_comp,item_seg_coment_item_seg,item_seg_fecha_ingreso,item_nombre_hoja) " _
                 & "values('" & cat2_aux & "', '" & CodIniciativa & "', '" & fecha_ing_aux & "','" & correl_cod_proy_aux & "','" & cod_int_aux & "','" & compromisos_aux & "',  " _
                 & " '" & resp_comp_aux & "','" & fecha_venc_aux & "','" & fecha_replanif_aux & "','" & cantidad_replanif_aux & "','" & fecha_cierre_comp_aux & "','" & condicion_aux & "','" & edo_comprom_aux & "', " _
                 & " '" & coment_item_seg_aux & "','" & DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") & "','" & hoja & "')"

                    'Debug.Print(dbinsert_item_seg)

                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbinsert_item_seg, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbcommand.ExecuteNonQuery()
                    dbConexion.Close()
                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing

                    MessageExito = "Se Inserto con Exito en la tabla Item Seguimientos con el codigo Iniciativa : " & CodIniciativa & " - Fila " & contfilas
                    Log(MessageExito, "exito")
                Next


                Console.WriteLine("Se Inserto con Exito en la tabla Item Seguimientos... ")

            End If
        Catch ex As Exception
            Log("Se ha producido un error en la función ItemSegValuesByVal con el codigo iniciativa: " & CodIniciativa & ":" & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()

            Return "0"
        End Try


salir:
        Return 0

    End Function
    '****************************************ITEM SEGUIMIENTO ACUERDOS**************************************************************'

    Private Function ItemSegValuesByVal_Acuerdos(ByVal ideIniciativa As String,
                                                ByVal CodIniciativa As String,
                                                ByVal cat2_value As String,
                                                ByVal fecha_ing_sist_value As String,
                                                ByVal correl_cod_proy_value As String,
                                                ByVal cod_int_value As String,
                                                ByVal coment_item_seg_value As String,
                                                ByVal edo_comp_value As String,
                                                ByVal hoja As String,
                                                ByVal contfilas As Integer,
                                                ByVal tblNom As String) As String


        Try
            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim cadEnter_cat2() As String
            Dim cadEnter_fecha_ing_sist() As String
            Dim cadEnter_correl_cod_proy() As String
            Dim cadEnter_cod_int() As String
            'Dim cadEnter_compromisos() As String
            Dim cadEnter_edo_comp() As String
            Dim dbinsert_item_seg As String = ""
            Dim cat2_aux As String = ""
            Dim fecha_ing_aux As String = ""
            Dim correl_cod_proy_aux As String = ""
            Dim cod_int_aux As String = ""
            Dim compromisos_aux As String = ""
            Dim resp_comp_aux As String = ""
            Dim fecha_venc_aux As String = ""
            Dim fecha_replanif_aux As String = ""
            Dim cantidad_replanif_aux As String = ""
            Dim fecha_cierre_comp_aux As String = ""
            Dim condicion_aux As String = ""
            Dim edo_comprom_aux As String = ""
            Dim coment_item_seg_aux As String = ""
            Dim stringSeparators() As String = {"\"c}
            Dim codigo_ini As String = ""
            Dim cadLoop As Integer = 0
            Dim position As Integer = hoja.IndexOf("$")
            Dim name As String = ""
            Dim query As String = ""
            Dim MessageExito As String = ""
            Dim fecha_condicion As String = ""


            'fecha_venc_value.Split("/", "-")
            If (cod_int_value <> "") Then
                'If (fecha_venc_value <> "-") And (fecha_venc_value <> "TBD") And (fecha_venc_value <> "") Then
                '    Dim myDateTime() As String = Regex.Split(fecha_venc_value, "\ - ")
                '    Dim dia As String = myDateTime(0)
                '    Dim mes As String = myDateTime(1)
                '    Dim anio As String = myDateTime(2)
                '    fecha_condicion = String.Concat(anio, "/", mes, "/", dia)

                'End If

                name = hoja.Substring(0, position)

                cadEnter_cat2 = cat2_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_fecha_ing_sist = fecha_ing_sist_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_correl_cod_proy = correl_cod_proy_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_cod_int = cod_int_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_edo_comp = edo_comp_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                'cadEnter_coment_item_seg = coment_item_seg_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)

                'If (fecha_replanif_value = "") Then
                '    cadLoop = cadEnter_fecha_ing_sist.Length - 1
                'Else
                '    cadLoop = cadEnter_fecha_replanif.Length - 1
                'End If


                MessageExito = ""

                For i As Integer = 0 To cadEnter_cat2.Length - 1

                    If (cat2_value <> "") Then
                        cat2_aux = cadEnter_cat2(i)
                    Else
                        cat2_aux = ""

                    End If

                    If (fecha_ing_sist_value <> "") Then
                        fecha_ing_aux = cadEnter_fecha_ing_sist(i)
                    Else
                        fecha_ing_aux = ""
                    End If


                    If (correl_cod_proy_value <> "") Then
                        correl_cod_proy_aux = cadEnter_correl_cod_proy(i)
                    Else
                        correl_cod_proy_aux = ""
                    End If

                    If (cod_int_value <> "") Then
                        cod_int_aux = cadEnter_cod_int(i)
                    Else
                        cod_int_aux = ""
                    End If


                    If (coment_item_seg_value <> "") Then
                        compromisos_aux = coment_item_seg_value
                    Else
                        compromisos_aux = ""
                    End If


                    If (edo_comp_value <> "") Then
                        edo_comprom_aux = cadEnter_edo_comp(i)
                    Else
                        edo_comprom_aux = ""
                    End If


                    dbinsert_item_seg = "INSERT INTO [dbo].[" & tblNom & "] " _
                 & "(item_seg_cat,item_seg_cod_ini,item_seg_fecha_ing,item_seg_correl,item_seg_cod_int,item_seg_compromisos, " _
                 & "item_seg_resp_comp,item_seg_fecha_venc,item_seg_fecha_replanif,item_seg_cant_replanif,item_seg_fecha_cierre,item_seg_condicion,item_seg_edo_comp,item_seg_coment_item_seg,item_seg_fecha_ingreso,item_nombre_hoja) " _
                 & "values('" & cat2_aux & "', '" & CodIniciativa & "', '" & fecha_ing_aux & "','" & correl_cod_proy_aux & "','" & cod_int_aux & "','" & compromisos_aux & "',  " _
                 & " '','','','','','','" & edo_comprom_aux & "', " _
                 & " '','" & DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") & "','" & hoja & "')"

                    'Debug.Print(dbinsert_item_seg)

                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbinsert_item_seg, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbcommand.ExecuteNonQuery()
                    dbConexion.Close()
                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing

                    MessageExito = "Se Inserto con Exito en la tabla Item Seguimientos Acuerdos con el codigo Iniciativa : " & CodIniciativa & " - Fila " & contfilas
                    Log(MessageExito, "exito")

                Next

                Console.WriteLine("Se Inserto con Exito en la tabla Item Seguimientos Acuerdos... ")
            End If


        Catch ex As Exception
            Log("Se ha producido un error en la función ItemSegValuesByVal con el codigo iniciativa: " & CodIniciativa & ":" & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()

            Return "0"
        End Try


salir:
        Return 0

    End Function




    '********************************************ITEM SEGUIMIENTO HISTORICO***********************************'

    Private Function ItemSegValuesByVal_hist(ByVal ideIniciativa As String,
                                         ByVal CodIniciativa As String,
                                          ByVal nombreHoja As String,
                                         ByVal tblNom As String) As String

        Try

            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim dbinsert_ini_planif As String = ""
            Dim dbinsert_item_siguimiento_hist As String = ""
            Dim dbRowCount As String = ""
            Dim dbresultados As String = ""
            Dim dbdeleteseguim As String = ""
            Dim contador As Integer = 0
            Dim MessageExito As String = ""

            dbinsert_item_siguimiento_hist = "if exists(select top 1 item_seg_cod_ini from  imp_item_seguimiento where item_seg_cod_ini='" & CodIniciativa & "' and item_nombre_hoja = '" & nombreHoja & "') " & _
              "begin " & _
              "declare @hoy datetime; " & _
              "set		@hoy=getdate(); " & _
              "insert imp_item_seguimiento_hist( item_seg_cod_ini, item_seg_cat, item_seg_fecha_ing, item_seg_correl, item_seg_cod_int, item_seg_compromisos, item_seg_resp_comp, item_seg_fecha_venc, item_seg_fecha_replanif, item_seg_cant_replanif, item_seg_fecha_cierre, item_seg_condicion," & _
              "item_seg_edo_comp, item_seg_coment_item_seg, item_seg_fecha_ingreso,item_nombre_hoja) " & _
              "select	item_seg_cod_ini, item_seg_cat, item_seg_fecha_ing, item_seg_correl, item_seg_cod_int, item_seg_compromisos,item_seg_resp_comp, item_seg_fecha_venc, item_seg_fecha_replanif,item_seg_cant_replanif, item_seg_fecha_cierre, item_seg_condicion, " & _
              "item_seg_edo_comp, item_seg_coment_item_seg, @hoy,item_nombre_hoja " & _
              "from imp_item_seguimiento " & _
              "where	item_seg_cod_ini='" & CodIniciativa & "'  AND item_nombre_hoja = '" & nombreHoja & "' " & _
              "ORDER BY item_seg_ide " & _
              " " & _
              "delete from imp_item_seguimiento where item_seg_cod_ini='" & CodIniciativa & "' AND item_nombre_hoja = '" & nombreHoja & "' " & _
              "end"

            'Debug.Print(dbinsert_item_siguimiento_hist)

            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(dbinsert_item_siguimiento_hist, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbcommand.ExecuteNonQuery()
            Console.WriteLine("Se Inserto con Exito en la tabla Item Seguimientos... ")
            dbConexion.Close()

            dbConexion = Nothing
            dbcommand = Nothing
            dbdata = Nothing

            MessageExito = "Se Inserto con Exito en la tabla Item Seguimientos con el codigo Iniciativa :" & CodIniciativa
            Log(MessageExito, "exito")

        Catch ex As Exception
            Log("Se ha producido un error en la funcion ItemSegValuesByVal_hist con el codigo Iniciativa: " & CodIniciativa & ":" & ex.Message, "error")
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try


    End Function

    '***************************************************************REPLANIFICACIÓN**********************************************************************
    Private Function ReplanificacionValuesByVal(fecha_rep_comp_value As String,
                         ByVal fecha_ven_hist_value As String,
                         ByVal cant_replanif_value As String,
                         ByVal respons_resol_value As String,
                         ByVal cod_interno_value As String,
                         ByVal columna As Integer,
                         ByVal tblNom As String) As String


        Try
            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim cadEnter_fecha_replanif() As String
            Dim cadEnter_fecha_ven_hist() As String
            Dim dbinsert_replanificacion As String = ""
            Dim dbquery As String = ""
            Dim fecha_replanif_aux As String = ""
            Dim fecha_venc_aux As String = ""
            Dim cant_replanif_aux As String = ""
            Dim resp_resol_aux As String = ""
            Dim dbresultados As String = ""
            Dim stringSeparators() As String = {"\"c}

            If (columna >= 1) Then

                cadEnter_fecha_replanif = fecha_rep_comp_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_fecha_ven_hist = fecha_ven_hist_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)


                'Busco El id que se genero con la inserción  
                dbquery = "SELECT item_seg_ide FROM [dbo].[imp_item_seguimiento] WHERE item_seg_cod_int='" & cod_interno_value & "' "
                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(dbquery, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                dbdata = dbcommand.ExecuteReader

                If dbdata.HasRows = True Then
                    dbdata.Read()
                    dbresultados = dbdata.Item(0).ToString
                    dbdata.Close()
                    dbConexion.Close()

                    For i As Integer = 0 To cadEnter_fecha_replanif.Length - 1

                        If (fecha_rep_comp_value <> "") Then
                            fecha_replanif_aux = cadEnter_fecha_replanif(i)
                        End If


                        If (fecha_ven_hist_value <> "") Then
                            fecha_venc_aux = cadEnter_fecha_ven_hist(i)
                        End If

                        cant_replanif_aux = cant_replanif_value
                        resp_resol_aux = respons_resol_value


                        dbinsert_replanificacion = "INSERT INTO [dbo].[" & tblNom & "] " _
                      & "(replanif_cod_item_seg,replanif_fecha_replanif_comp,replanif_fecha_ven_hist_comp, " _
                      & "replanif_cant_replanif,replanif_respons_resol_comp) " _
                      & "values('" & dbresultados & "', '" & fecha_replanif_aux & "','" & fecha_venc_aux & "','" & cant_replanif_aux & "', " _
                      & "'" & resp_resol_aux & "' )"


                        dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                        dbcommand = New Data.Odbc.OdbcCommand(dbinsert_replanificacion, dbConexion)
                        dbcommand.CommandType = CommandType.Text
                        dbConexion.Open()
                        dbcommand.ExecuteNonQuery()
                        dbConexion.Close()
                        dbConexion = Nothing
                        dbcommand = Nothing
                        dbdata = Nothing

                    Next
                    Console.WriteLine("Se Inserto con Exito en la tabla Replanificacion... ")
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try


salir:
        Return 0

    End Function

    Private Function EvidenciasValuesByVal(nomb_evidcia_value As String,
                     ByVal tipo_evidcia_value As String,
                     ByVal artef_evidcia_value As String,
                     ByVal ruta_evidcia_value As String,
                     ByVal resp_entg_evidcia_value As String,
                     ByVal resp_evidcia_value As String,
                     ByVal obs_evidcia_value As String,
                     ByVal amb_evidcia_value As String,
                     ByVal act_gantt_evidcia_value As String,
                     ByVal cod_interno_value As String,
                     ByVal columna As Integer,
                     ByVal tblNom As String) As String


        Try
            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim cadEnter_nomb_evidcia() As String
            Dim cadEnter_tipo_evidcia() As String
            Dim cadEnter_artef_evidcia() As String
            Dim cadEnter_ruta_evidcia() As String
            Dim cadEnter_resp_entg_evidcia() As String
            Dim cadEnter_resp_evidcia() As String
            Dim cadEnter_obs_evidcia() As String
            Dim cadEnter_amb_evidcia() As String
            Dim cadEnter_act_gantt_evidcia() As String

            Dim dbinsert_evidencia As String = ""
            Dim dbquery As String = ""
            Dim nomb_evidcia_aux As String = ""
            Dim tipo_evidcia_aux As String = ""
            Dim artef_evidcia_aux As String = ""
            Dim ruta_evidcia_aux As String = ""
            Dim resp_entg_evidcia_aux As String = ""
            Dim resp_evidcia_aux As String = ""
            Dim obs_evidcia_aux As String = ""
            Dim amb_evidcia_aux As String = ""
            Dim act_gantt_evidcia_aux As String = ""

            Dim dbresultados As String = ""
            Dim stringSeparators() As String = {"\"c}

            If (columna >= 1) Then

                cadEnter_nomb_evidcia = nomb_evidcia_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_tipo_evidcia = tipo_evidcia_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_artef_evidcia = artef_evidcia_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_ruta_evidcia = ruta_evidcia_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_resp_entg_evidcia = resp_entg_evidcia_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_resp_evidcia = resp_evidcia_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_obs_evidcia = obs_evidcia_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_amb_evidcia = amb_evidcia_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
                cadEnter_act_gantt_evidcia = act_gantt_evidcia_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)

                'Busco El id que se genero con la inserción  
                dbquery = "SELECT item_seg_ide FROM [dbo].[imp_item_seguimiento] WHERE item_seg_cod_int='" & cod_interno_value & "' "
                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(dbquery, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                dbdata = dbcommand.ExecuteReader

                If dbdata.HasRows = True Then
                    dbdata.Read()
                    dbresultados = dbdata.Item(0).ToString
                    dbdata.Close()
                    dbConexion.Close()


                    For i As Integer = 0 To cadEnter_nomb_evidcia.Length - 1

                        If (nomb_evidcia_value <> "") Then
                            nomb_evidcia_aux = cadEnter_nomb_evidcia(i)
                        End If


                        If (tipo_evidcia_value <> "") Then
                            tipo_evidcia_aux = cadEnter_tipo_evidcia(i)
                        End If

                        If (artef_evidcia_value <> "") Then
                            artef_evidcia_aux = cadEnter_artef_evidcia(i)
                        End If

                        If (ruta_evidcia_value <> "") Then
                            ruta_evidcia_aux = cadEnter_ruta_evidcia(i)
                        End If

                        If (resp_entg_evidcia_value <> "") Then
                            resp_entg_evidcia_aux = cadEnter_resp_entg_evidcia(i)
                        End If

                        If (resp_evidcia_value <> "") Then
                            resp_evidcia_aux = cadEnter_resp_evidcia(i)
                        End If

                        If (obs_evidcia_value <> "") Then
                            obs_evidcia_aux = cadEnter_obs_evidcia(i)
                        End If

                        If (amb_evidcia_value <> "") Then
                            amb_evidcia_aux = cadEnter_amb_evidcia(i)
                        End If

                        If (act_gantt_evidcia_value <> "") Then
                            act_gantt_evidcia_aux = cadEnter_act_gantt_evidcia(i)
                        End If

                        dbinsert_evidencia = "INSERT INTO [dbo].[" & tblNom & "] " _
                      & "(evidencia_cod_item_seg,evidencia_nombre,evidencia_tipo,evidencia_artefacto,evidencia_ruta,evidencia_responsable_ent, " _
                      & "evidencia_responsable,evidencia_obs,evidencia_amb,evidencia_actividad_gant) " _
                      & "values('" & dbresultados & "', '" & nomb_evidcia_aux & "','" & tipo_evidcia_aux & "','" & artef_evidcia_aux & "','" & ruta_evidcia_aux & "', " _
                      & "'" & resp_entg_evidcia_aux & "', '" & resp_evidcia_aux & "','" & obs_evidcia_aux & "','" & amb_evidcia_aux & "','" & act_gantt_evidcia_aux & "')"


                        dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                        dbcommand = New Data.Odbc.OdbcCommand(dbinsert_evidencia, dbConexion)
                        dbcommand.CommandType = CommandType.Text
                        dbConexion.Open()
                        dbcommand.ExecuteNonQuery()
                        dbConexion.Close()
                        dbConexion = Nothing
                        dbcommand = Nothing
                        dbdata = Nothing

                    Next
                    Console.WriteLine("Se Inserto con Exito en la tabla Evidencias... ")
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try

salir:
        Return 0

    End Function

    '***********************************************************CHECKLIST INICIAL*******************************************************************************'
    Private Function checkListInicialValuesByVal(ByVal id_iniciativa As String,
                                                 ByVal codigo_iniciativa As String,
                                                 ByVal nombre_ambito As String,
                                                 ByVal fecha_carga As String,
                                                 ByVal ambito_value As String,
                                                 ByVal accion_value As String,
                                                 ByVal etapa_value As String,
                                                 ByVal hito_value As String,
                                                 ByVal preg_value As String,
                                                 ByVal resp_value As String,
                                                 ByVal obs_value As String,
                                                 ByVal contfilas As Integer,
                                                 ByVal tblNom As String) As String

        Try
            Dim dbConexion As Data.Odbc.OdbcConnection
            Dim dbcommand As Data.Odbc.OdbcCommand
            Dim dbdata As Data.Odbc.OdbcDataReader
            Dim cadEnter_ambito() As String
            Dim cadEnter_accion() As String
            Dim cadEnter_etapa() As String
            Dim cadEnter_hito() As String
            Dim cadEnter_preg() As String
            Dim cadEnter_resp() As String
            Dim cadEnter_obs() As String
            Dim dbinsert_check_lists_ini As String = ""
            Dim dbinsert_check_listsInicial As String = ""
            Dim ambito_aux As String = ""
            Dim accion_aux As String = ""
            Dim etapa_aux As String = ""
            Dim hito_aux As String = ""
            Dim preg_aux As String = ""
            Dim resp_aux As String = ""
            Dim obs_aux As String = ""
            Dim cad_null As String = ""
            Dim str_accion As String = ""
            Dim str_etapa As String = ""
            Dim str_obs As String = ""
            Dim dbRowCount As String = ""
            Dim dbresultados As String = ""
            Dim IdResult As String = ""
            Dim dbdeletecheckLists As String = ""
            Dim cont As Integer = 0
            Dim nombre As String = nombre_ambito
            Dim dbquery As String = ""
            Dim position As Integer = nombre.IndexOf("$")
            Dim name As String = ""
            Dim MessageExito As String = ""
            Dim name_ambito As String = ""
            Dim dbSQL As String = ""
            name = nombre.Substring(0, position)
            If (ambito_value = "Convivencia") Then
                name_ambito = "Coexistencia"
            Else
                name_ambito = ambito_value
            End If


            cadEnter_ambito = name_ambito.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_accion = accion_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_etapa = etapa_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_hito = hito_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_preg = preg_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            cadEnter_resp = resp_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)
            'cadEnter_obs = obs_value.Split(ControlChars.CrLf.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)


            For i As Integer = 0 To cadEnter_preg.Length - 1


                If (name_ambito = " ") Then

                    dbquery = " SELECT TOP 1  checkList_ambito FROM [dbo].[" & tblNom & "]  where checkList_ambito = '" & name & "' ORDER BY checkList_ide DESC"
                    Debug.Print(dbquery)


                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbquery, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbdata = dbcommand.ExecuteReader

                    If dbdata.HasRows = True Then
                        dbdata.Read()
                        ambito_aux = dbdata.Item(0).ToString
                        dbdata.Close()
                        dbConexion.Close()

                    End If
                Else

                    ambito_aux = cadEnter_ambito(i)
                End If


                If (accion_value = " ") Then

                    dbquery = " SELECT TOP 1  checkList_accion FROM [dbo].[" & tblNom & "]  where checkList_ambito = '" & name & "' ORDER BY checkList_ide DESC "

                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbquery, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbdata = dbcommand.ExecuteReader

                    If dbdata.HasRows = True Then
                        dbdata.Read()
                        accion_aux = dbdata.Item(0).ToString
                        dbdata.Close()
                        dbConexion.Close()

                    End If
                Else

                    If (accion_value = " ") Then
                        accion_aux = " "

                    Else
                        accion_aux = cadEnter_accion(i)
                    End If
                End If

                If (etapa_value = " ") Then


                    dbquery = "SELECT TOP 1  checkList_etapa FROM [dbo].[" & tblNom & "]   " _
                               & "  where checkList_ambito = '" & name & "' ORDER BY checkList_ide DESC"
                    Debug.Print(dbquery)


                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbquery, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbdata = dbcommand.ExecuteReader

                    If dbdata.HasRows = True Then
                        dbdata.Read()
                        etapa_aux = dbdata.Item(0).ToString
                        dbdata.Close()
                        dbConexion.Close()

                    End If
                Else
                    If (etapa_value = " ") Then
                        etapa_aux = " "
                    Else
                        etapa_aux = cadEnter_etapa(i)

                    End If

                End If

                If (hito_value = " ") Then

                    dbquery = "SELECT TOP 1  checkList_hito FROM [dbo].[" & tblNom & "]   " _
                               & " where checkList_ambito = '" & name & "' ORDER BY checkList_ide DESC"

                    'Debug.Print(dbquery)

                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbquery, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbdata = dbcommand.ExecuteReader

                    If dbdata.HasRows = True Then
                        dbdata.Read()
                        hito_aux = dbdata.Item(0).ToString
                        dbdata.Close()
                        dbConexion.Close()

                    End If
                Else

                    hito_aux = cadEnter_hito(i)
                End If

                preg_aux = cadEnter_preg(i)

                If (resp_value <> " ") Then
                    resp_aux = cadEnter_resp(i)
                Else
                    resp_aux = ""
                End If


                If (obs_value <> " ") Then
                    obs_aux = obs_value
                Else
                    obs_aux = " "
                End If

                dbSQL = "SELECT COUNT(*) AS contador from [dbo].[" & tblNom & "] where checkList_cod_ini='" & codigo_iniciativa & "' " _
                    & "AND checkList_ambito = '" & ambito_aux & "' AND checkList_preguntas = '" & preg_aux & "' AND checkList_hito = '" & hito_aux & "' "
                Debug.Print(dbSQL)

                dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                dbcommand = New Data.Odbc.OdbcCommand(dbSQL, dbConexion)
                dbcommand.CommandType = CommandType.Text
                dbConexion.Open()
                dbdata = dbcommand.ExecuteReader

                If dbdata.HasRows = True Then
                    dbdata.Read()
                    dbresultados = dbdata.Item(0).ToString

                    dbdata.Close()
                    dbConexion.Close()

                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing
                End If


                If (dbresultados = "0") Then

                    dbinsert_check_listsInicial = "INSERT INTO [dbo].[" & tblNom & "] " _
            & "(checkList_ini_ide,checkList_cod_ini,checkList_ambito,checkList_accion,checkList_etapa,checkList_hito,checkList_preguntas, " _
            & "checkList_respuesta,checkList_observaciones,checkList_fecha_realizacion,checkList_fecha_actual) " _
            & "values('" & id_iniciativa & "', '" & codigo_iniciativa & "','" & ambito_aux & "','" & accion_aux & "','" & etapa_aux & "','" & hito_aux & "',  " _
            & " '" & preg_aux & "','" & resp_aux & "','" & obs_aux & "','" & fecha_carga & "','" & DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") & "')"

                    'Debug.Print(dbinsert_check_listsInicial)

                    dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
                    dbcommand = New Data.Odbc.OdbcCommand(dbinsert_check_listsInicial, dbConexion)
                    dbcommand.CommandType = CommandType.Text
                    dbConexion.Open()
                    dbcommand.ExecuteNonQuery()
                    dbConexion.Close()
                    dbConexion = Nothing
                    dbcommand = Nothing
                    dbdata = Nothing
                    MessageExito = "Se Inserto con Exito en la Tabla el Ambito del Checklist: " & codigo_iniciativa & " - Fila " & contfilas
                    Log(MessageExito, "exito")
                    Console.WriteLine("Se Inserto con Exito en la Tabla el Ambito del Checklist: " & codigo_iniciativa)
                End If

            Next



        Catch ex As Exception
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Log("Se ha producido un error en Modelo ChecklistInicial " & ex.Message, "error")
            Console.ReadLine()
            Return "0"
        End Try


salir:
        Return 0
    End Function

    '*****************************************************************************ITEM SEGUIMIENTO (ADMINISTRADOR DE COMPROMISOS)**********************************'
    Public Sub ModeloItemSeguimiento(ideIniciativa, CodIniciativa, nombreHoja)

        Dim adoCon As Data.OleDb.OleDbConnection
        Dim adoRs As Data.OleDb.OleDbDataReader
        Dim adoCmd As Data.OleDb.OleDbCommand
        'Dim strCmnd As String = "SELECT * FROM [MatrizCompromisos - GdeIm$]"
        Dim strCmnd As String = "SELECT * FROM [" & nombreHoja & "]"
        Dim cont As Integer = 0
        Dim Cabecera As String = ""
        Dim columna5 As String = ""
        Dim columna6 As String = ""
        Dim columna7 As String = ""
        Dim columna8 As String = ""
        Dim columna9 As String = ""
        Dim columna10 As String = Nothing
        Dim columna11 As String = Nothing
        Dim columna12 As String = Nothing
        Dim contfilas As Integer = 0
        Dim comentarios As String = ""
        Dim compromisos As String = ""

        Try
            comentarios = "" : compromisos = ""

            'If adoCon.State = ConnectionState.Open Then adoCon.Close()
            'If (adoRs) Is Nothing Then
            'Else
            '    If Not adoRs.IsClosed Then adoRs.Close()
            'End If

            adoCon = New Data.OleDb.OleDbConnection(GetConnectionString(5))
            adoCmd = New Data.OleDb.OleDbCommand(strCmnd, adoCon)
            adoCmd.CommandType = CommandType.Text
            adoCon.Open()
            adoRs = adoCmd.ExecuteReader


            If adoRs.HasRows Then
                adoRs.Read()
                Do While adoRs.Read()

                    If IsDBNull(adoRs.Item(0)) Then
                        Cabecera = ""
                    Else
                        Cabecera = adoRs.Item(0)

                    End If

                    'Console.WriteLine("Cabecera " & adoRs.Item(0))

                    If (Cabecera = "Eje") Then
                        cont += 1
                    End If

                    If (cont >= 1) And (Cabecera <> "Eje") Then

                        If (nombreHoja = "Template_MatrizCompromisos$") Then
                            contfilas = contfilas + 1

                            If IsDBNull(adoRs.Item(4)) Then
                                compromisos = ""
                            Else
                                compromisos = Convert.ToString(adoRs.Item(4))
                            End If

                            If IsDBNull(adoRs.Item(12)) Then
                                comentarios = ""
                            Else
                                comentarios = Convert.ToString(adoRs.Item(12))

                            End If

                            idContactoItemSeg = ItemSegValuesByVal(ideIniciativa, CodIniciativa, Convert.ToString(adoRs.Item(0)).Trim, Convert.ToString(adoRs.Item(1)).Trim, Convert.ToString(adoRs.Item(2)).Trim, Convert.ToString(adoRs.Item(3)).Trim, compromisos, Convert.ToString(adoRs.Item(5)).Trim, Convert.ToString(adoRs.Item(6)).Trim, Convert.ToString(adoRs.Item(7)).Trim, Convert.ToString(adoRs.Item(8)).Trim, Convert.ToString(adoRs.Item(9)).Trim, Convert.ToString(adoRs.Item(10)), Convert.ToString(adoRs.Item(11)).Trim, comentarios, nombreHoja, contfilas, "imp_item_seguimiento")

                        Else

                            If IsDBNull(adoRs.Item(5)) Then
                                columna5 = ""
                            Else
                                columna5 = adoRs.Item(5)
                            End If

                            If IsDBNull(adoRs.Item(6)) Then
                                columna6 = ""
                            Else
                                columna6 = adoRs.Item(6)
                            End If

                            If IsDBNull(adoRs.Item(7)) Then
                                columna7 = ""
                            Else
                                columna7 = adoRs.Item(7)
                            End If

                            If IsDBNull(adoRs.Item(8)) Then
                                columna8 = ""
                            Else
                                columna8 = adoRs.Item(8)
                            End If

                            If IsDBNull(adoRs.Item(9)) Then
                                columna9 = ""
                            Else
                                columna9 = adoRs.Item(9)
                            End If

                            If IsDBNull(adoRs.Item(10)) Then
                                columna10 = ""
                            Else
                                columna10 = adoRs.Item(10)
                            End If

                            If IsDBNull(adoRs.Item(11)) Then
                                columna11 = ""
                            Else
                                columna11 = adoRs.Item(11)
                            End If

                            If IsDBNull(adoRs.Item(12)) Then
                                columna12 = ""
                            Else
                                columna12 = adoRs.Item(12)
                            End If

                            contfilas = contfilas + 1
                            idContactoItemSeg = ItemSegValuesByVal_Acuerdos(ideIniciativa, CodIniciativa, Convert.ToString(adoRs.Item(0)).Trim, Convert.ToString(adoRs.Item(1)).Trim, Convert.ToString(adoRs.Item(2)).Trim, Convert.ToString(adoRs.Item(3)).Trim, Convert.ToString(adoRs.Item(4)).Trim, Convert.ToString(columna5).Trim, nombreHoja, contfilas, "imp_item_seguimiento")


                        End If
                    End If

                Loop

            Else
                GoTo salirSinFilas
                Debug.Print("No rows found.")
                Console.ReadLine()
            End If


salirSinFilas:
            adoRs.Close()
            adoCon.Close()

            adoCon = Nothing
            adoRs = Nothing
            adoCmd = Nothing

        Catch ex As Exception
            Console.WriteLine("Se ha producido un error " & ex.Message)
            Log("Se ha producido un error en Modelo ModeloItemSeguimiento " & ex.Message, "error")
            Console.ReadLine()
        End Try

    End Sub

    '*************************************MODELO ITEM SEGUIMIENTO HISTORICO*****************************************************************************'
    Public Sub ModeloItemSeguimiento_hist(ideIniciativa, CodIniciativa, nombreHoja)

        Try
            idContactoItemSeg_Hist = ItemSegValuesByVal_hist(ideIniciativa, CodIniciativa, nombreHoja, "imp_item_seguimiento_hist")

        Catch ex As Exception
            Log("Se ha producido un error en la funcion ModeloItemSeguimiento_hist:" & ex.Message, "error")
            Console.ReadLine()

        End Try


    End Sub

    '**************************************************INSERT ITEM SEGUIMIENTO *************************************************************'
    Public Sub ModeloInsertItemSeguimiento(ideIniciativa, CodIniciativa, nombreHoja)

        Dim adoCon As Data.OleDb.OleDbConnection
        Dim adoRs As Data.OleDb.OleDbDataReader
        Dim adoCmd As Data.OleDb.OleDbCommand
        'Dim strCmnd As String = "SELECT * FROM [MatrizCompromisos - GdeIm$]"
        Dim strCmnd As String = "SELECT * FROM [" & nombreHoja & "]"
        Dim cont As Integer = 0
        Dim Cabecera As String = ""
        Dim columna5 As String = ""
        Dim columna6 As String = ""
        Dim columna7 As String = ""
        Dim columna8 As String = ""
        Dim columna9 As String = ""
        Dim columna10 As String = Nothing
        Dim columna11 As String = Nothing
        Dim columna12 As String = Nothing
        Dim contfilas As Integer = 0

        adoCon = New Data.OleDb.OleDbConnection(GetConnectionString(5))
        adoCmd = New Data.OleDb.OleDbCommand(strCmnd, adoCon)
        adoCmd.CommandType = CommandType.Text
        adoCon.Open()
        adoRs = adoCmd.ExecuteReader

        If adoRs.HasRows Then
            adoRs.Read()
            Do While adoRs.Read()

                If IsDBNull(adoRs.Item(0)) Then
                    Cabecera = ""
                Else
                    Cabecera = adoRs.Item(0)

                End If

                'Console.WriteLine("Cabecera " & adoRs.Item(0))

                If (Cabecera = "Eje") Then
                    cont += 1
                End If

                If (cont = 1) And (Cabecera <> "Eje") Then

                    If (nombreHoja = "Template_MatrizCompromisos$") Then
                        contfilas = contfilas + 1
                        idContactoItemSeg = ItemSegValuesByVal(ideIniciativa, CodIniciativa, Convert.ToString(adoRs.Item(0)).Trim, Convert.ToString(adoRs.Item(1)).Trim, Convert.ToString(adoRs.Item(2)).Trim, Convert.ToString(adoRs.Item(3)).Trim, Convert.ToString(adoRs.Item(4)).Trim, Convert.ToString(adoRs.Item(5)).Trim, Convert.ToString(adoRs.Item(6)).Trim, Convert.ToString(adoRs.Item(7)).Trim, Convert.ToString(adoRs.Item(8)).Trim, Convert.ToString(adoRs.Item(9)).Trim, Convert.ToString(adoRs.Item(10)).Trim, Convert.ToString(adoRs.Item(11)).Trim, Convert.ToString(adoRs.Item(12)).Trim, nombreHoja, contfilas, "imp_item_seguimiento")

                    Else

                        If IsDBNull(adoRs.Item(5)) Then
                            columna5 = " "
                        Else
                            columna5 = adoRs.Item(5)
                        End If

                        If IsDBNull(adoRs.Item(6)) Then
                            columna6 = " "
                        Else
                            columna6 = adoRs.Item(6)
                        End If

                        If IsDBNull(adoRs.Item(7)) Then
                            columna7 = " "
                        Else
                            columna7 = adoRs.Item(7)
                        End If

                        If IsDBNull(adoRs.Item(8)) Then
                            columna8 = " "
                        Else
                            columna8 = adoRs.Item(8)
                        End If

                        If IsDBNull(adoRs.Item(9)) Then
                            columna9 = " "
                        Else
                            columna9 = adoRs.Item(9)
                        End If

                        If IsDBNull(adoRs.Item(10)) Then
                            columna10 = " "
                        Else
                            columna10 = adoRs.Item(10)
                        End If

                        If IsDBNull(adoRs.Item(11)) Then
                            columna11 = " "
                        Else
                            columna11 = adoRs.Item(11)
                        End If

                        If IsDBNull(adoRs.Item(12)) Then
                            columna12 = " "
                        Else
                            columna12 = adoRs.Item(12)
                        End If

                        contfilas = contfilas + 1
                        idContactoItemSeg = ItemSegValuesByVal(ideIniciativa, CodIniciativa, Convert.ToString(adoRs.Item(0)).Trim, Convert.ToString(adoRs.Item(1)).Trim, Convert.ToString(adoRs.Item(2)).Trim, Convert.ToString(adoRs.Item(3)).Trim, Convert.ToString(adoRs.Item(4)).Trim, Convert.ToString(columna5).Trim, Convert.ToString(columna6).Trim, Convert.ToString(columna7).Trim, Convert.ToString(columna8).Trim, Convert.ToString(columna9).Trim, Convert.ToString(columna10).Trim, Convert.ToString(columna11).Trim, Convert.ToString(columna12).Trim, nombreHoja, contfilas, "imp_item_seguimiento")

                    End If
                End If
            Loop

        Else
            GoTo salirSinFilas
            Debug.Print("No rows found.")
            Console.ReadLine()
        End If

salirSinFilas:
        adoRs.Close()
        adoCon.Close()

        adoCon = Nothing
        adoRs = Nothing
        adoCmd = Nothing

    End Sub

    '*********************************************Ultimo con Cab******************************************************
    Public Sub ModeloCheckListsInicial(idIniciativa, codIniciativa, rutaArchivo, fecha_carga)

        Dim adoCon As New Data.OleDb.OleDbConnection
        Dim adoRs As Data.OleDb.OleDbDataReader
        Dim adoCmd As Data.OleDb.OleDbCommand
        Dim strCuadroResumen As String = "SELECT * FROM [Cuadro Resumen$]"
        'Dim strCmnd As String = "SELECT * FROM [" & nombreHoja & "]"
        Dim strCmnd As String = ""
        Dim columna0 As String = ""
        Dim columna1 As String = ""
        Dim columna2 As String = ""
        Dim columna3 As String = ""
        Dim columna4 As String = ""
        Dim columna5 As String = Nothing
        Dim columna6 As String = Nothing
        Dim contador As Integer = 0
        Dim Cabecera As String = ""
        Dim Arreglo As New List(Of String)
        Dim index As String = ""
        Dim nombreHoja As String = ""
        Dim contAmbito As Integer = 0
        Dim filas As Integer = 0
        Dim filasAvn As Integer = 0
        Dim filaAmbitoCabecera As Integer = 0
        Dim filaAmbito As Integer = 0
        Dim columnaInicio As Integer = 0
        Dim k As Integer
        Dim contfilas As Integer = 0

        If adoCon.State = ConnectionState.Open Then adoCon.Close()
        If (adoRs) Is Nothing Then
        Else
            If Not adoRs.IsClosed Then adoRs.Close()
        End If

        adoCon = New Data.OleDb.OleDbConnection(GetConnectionString((7)))
        adoCmd = New Data.OleDb.OleDbCommand(strCuadroResumen, adoCon)
        adoCmd.CommandType = CommandType.Text
        adoCon.Open()

        adoRs = adoCmd.ExecuteReader

        If adoRs.HasRows Then
            adoRs.Read()
            Arreglo.Clear()

            Do While adoRs.Read()
                If IsDBNull(adoRs.Item(0)) Then
                Else
                    Arreglo.Add((adoRs.Item(0)))
                End If
            Loop

        End If

        For i = 1 To Arreglo.Count - 2
            Try
                strCmnd = "SELECT * FROM [" & Trim(Arreglo(i)) & "$]"

revisarPorCabeceraAmbito:
                If adoCon.State = ConnectionState.Open Then adoCon.Close()
                If Not adoRs.IsClosed Then adoRs.Close()

                adoCon = New Data.OleDb.OleDbConnection(GetConnectionString(7))
                adoCmd = New Data.OleDb.OleDbCommand(strCmnd, adoCon)
                adoCmd.CommandType = CommandType.Text
                adoCon.Open()
                adoRs = adoCmd.ExecuteReader

                '-- REVISAR SI EXISTE CABECERA --
                filas = 0
                contAmbito = 0

                Do While adoRs.Read()
                    filas = filas + 1
                    If Not IsDBNull(adoRs.Item(columnaInicio)) Then
                        If adoRs.Item(columnaInicio) = "Ámbito" Then
                            contAmbito = contAmbito + 1
                            If contAmbito = 1 Then filaAmbitoCabecera = filas
                            If contAmbito = 2 Then filaAmbito = filas
                        End If
                    End If
                Loop

                If contAmbito = 0 Then
                    GoTo salirSinFilas
                End If

                If contAmbito = 1 Then filaAmbito = filaAmbitoCabecera

                '------------------------------------------------------------------

                strCmnd = "SELECT * FROM [" & Trim(Arreglo(i)) & "$]"

                If adoCon.State = ConnectionState.Open Then adoCon.Close()
                If Not adoRs.IsClosed Then adoRs.Close()

                adoCon = New Data.OleDb.OleDbConnection(GetConnectionString(7))
                adoCmd = New Data.OleDb.OleDbCommand(strCmnd, adoCon)
                adoCmd.CommandType = CommandType.Text
                adoCon.Open()
                adoRs = adoCmd.ExecuteReader
                nombreHoja = Arreglo(i) & "$"

                If adoRs.HasRows Then

                    contador = 0

                    For k = 0 To filaAmbito - 1
                        adoRs.Read()
                    Next

                    Do While adoRs.Read()

                        If IsDBNull(adoRs.Item(0)) Then
                            columna0 = " "
                        Else
                            columna0 = adoRs.Item(0)
                        End If

                        If IsDBNull(adoRs.Item(1)) Then
                            columna1 = " "
                        Else
                            columna1 = adoRs.Item(1)
                        End If

                        If IsDBNull(adoRs.Item(2)) Then
                            columna2 = " "
                        Else
                            columna2 = adoRs.Item(2)
                        End If

                        If IsDBNull(adoRs.Item(3)) Then
                            columna3 = " "
                        Else
                            columna3 = adoRs.Item(3)
                        End If

                        If IsDBNull(adoRs.Item(4)) Then
                            columna4 = " "
                        Else
                            columna4 = adoRs.Item(4)
                        End If

                        If IsDBNull(adoRs.Item(5)) Then
                            columna5 = " "
                        Else
                            columna5 = adoRs.Item(5)
                        End If

                        If IsDBNull((adoRs.Item(6))) Then
                            columna6 = " "
                        Else
                            columna6 = adoRs.Item(6)

                        End If

                        If (columna4 <> " ") Then
                            contfilas = contfilas + 1
                            idCheckListInicial = checkListInicialValuesByVal(idIniciativa, codIniciativa, nombreHoja, fecha_carga, columna0, columna1, columna2, columna3, columna4, columna5, columna6, contfilas, "imp_check_list_Inicial")
                        End If
                    Loop

                Else
                    Debug.Print("No rows found.")
                    GoTo salirSinFilas
                End If

            Catch ex As Exception
                Log("No existe el ambito:" & Arreglo(i) & " en la ruta del archivo: " & rutaArchivo, "error")
                Console.WriteLine("No existe el ambito en el archivo: " & Arreglo(i))
                Console.WriteLine("Error: " & ex.Message)
                Console.WriteLine("Comando: " & strCmnd)
                Console.WriteLine("")
                Console.ReadLine()
            End Try

nextHojaChecklist:
        Next

salirSinFilas:
        'adoRs.Close()
        'adoCon.Close()

        adoCon = Nothing
        adoRs = Nothing
        adoCmd = Nothing

    End Sub

    '****************************************************CODIGO DINAMICO LECTURA DIRECTORIO***************************************************'
    Private Sub GetRutas()
        Dim xmlCfg As New System.Xml.XmlDocument
        Dim carpetaBase As String
        Dim carpetas() As String
        Dim carpetaWRK As String


        Dim rutaArchivo As String
        Dim archivos() As String
        Dim partesCarpeta() As String
        Dim partesArchivos() As String
        Dim rutaArchivoTemporal As String
        Dim ideIniciativa As String
        Dim CodIniciativa As String
        Dim numMax As Integer
        Dim numMaxArchivo As String
        Dim fecha_realizacion_check As String

        filePath = "" : appPath = "" : dbStrConexion = ""
        Log("Iniciando archivo de log", "")
        Log("Path de acceso al archivo de configuración en " & appPath, "")

        Try
            Log("Archivo de configuracion: " & appPath & "\xml\config.xml", "")
            'Console.WriteLine("Archivo de configuracion: " & appPath & "\xml\config.xml")
            xmlCfg.Load(appPath & "\xml\config.xml")
        Catch ex As Exception
            Log("Error al cargar el archivo de configuración : " & ex.Message, "Error")
            Console.WriteLine("Se ha producido un error, revise el archivo de LOG")
            Console.ReadLine()
            GoTo salir
        End Try

        If xmlCfg.SelectNodes("content/base/conexion").Count = 1 Then
            dbStrConexion = xmlCfg.SelectNodes("content/base/conexion").Item(0).InnerText
            If dbStrConexion <> "" Then
                Log("Cadena de conexion a base de datos encontrada", "")
            Else
                Log("Error al recuperar la conexion a la base de datos", "Error")
                Console.WriteLine("Se ha producido un error, revise el archivo de LOG")
                Console.ReadLine()
                GoTo salir
            End If
        Else
            Log("Error al recuperar la conexion a la base de datos", "Error")
            Console.WriteLine("Se ha producido un error, revise el archivo de LOG")
            Console.ReadLine()
            GoTo salir
        End If

        'Try
        If xmlCfg.SelectNodes("content/base/carpeta[@tip='padre'][@estado='1']").Count = 1 Then
            carpetaBase = xmlCfg.SelectNodes("content/base/carpeta[@tip='padre'][@estado='1']").Item(0).InnerText
            Log("Carpeta base: " & carpetaBase, "")
            'Console.WriteLine("Carpeta base: " & carpetaBase)
        Else
            Log("Error la carpeta base de los archivos a procesar", "Error")
            Console.WriteLine("Se ha producido un error, revise el archivo de LOG")
            Console.ReadLine()
            GoTo salir
        End If

        If xmlCfg.SelectNodes("content/base/proyectos").Count = 1 Then
            rutaArchivo = carpetaBase & xmlCfg.SelectNodes("content/base/proyectos").Item(0).InnerText
            Log("Archivo de iniciativas: " & rutaArchivo, "")
            'Console.WriteLine("Archivo de iniciativas: " & rutaArchivo)

            archivoExcelCarga = "" & rutaArchivo & ""
            archivoExcelCarga = Chr(34) & archivoExcelCarga & Chr(34)

            Try
                Call ModeloGeneral()
            Catch ex As Exception
                Log("Error Carga Modelo General: " & ex.Message, "error")
                Console.WriteLine("Se ha producido un error, revise el archivo de LOG")
                Console.ReadLine()
            End Try

            Try
                'Tema Relevante
                Call InsertTemaRelevante()

            Catch ex As Exception
                Log("Error Carga funcion InsertTemaRelevante: " & ex.Message, "error")
                Console.WriteLine("Se ha producido un error, revise el archivo de LOG")
                Console.ReadLine()

            End Try


            Try
                'Piloto Historicos
                Call ModeloPilotoHistoricos()

            Catch ex As Exception
                Log("Error Carga funcion ModeloPilotoHistoricos: " & ex.Message, "error")
                Console.WriteLine("Se ha producido un error, revise el archivo de LOG")
                Console.ReadLine()

            End Try


            Try
                '********Piloto*************
                Call InsertPilotoInicio()
            Catch ex As Exception
                Log("Error Carga funcion InsertPilotoInicio: " & ex.Message, "error")
                Console.WriteLine("Se ha producido un error, revise el archivo de LOG")
                Console.ReadLine()
            End Try

            Try
                'Despliegue Historico
                Call ModeloDespliegueHistorico()
            Catch ex As Exception
                Log("Error Carga funcion ModeloDespliegueHistorico: " & ex.Message, "error")
                Console.WriteLine("Se ha producido un error, revise el archivo de LOG")
                Console.ReadLine()

            End Try

            Try
                Call InsertDespliegueHistorico()

            Catch ex As Exception
                Log("Error Carga funcion InsertDespliegueHistorico: " & ex.Message, "error")
                Console.WriteLine("Se ha producido un error, revise el archivo de LOG")
                Console.ReadLine()
            End Try


        Else
            GoTo salir
        End If

        carpetas = System.IO.Directory.GetDirectories(carpetaBase)

        For i = 0 To carpetas.Length - 1
            ' Console.WriteLine(carpetas(i))

            If Not revisarRutas(xmlCfg, carpetas(i)) Then GoTo nextCarpetaIniciativa

            partesCarpeta = Split(Replace(carpetas(i), carpetaBase, ""), "-")
            ' Console.WriteLine(Trim(partesCarpeta(0)))
            ideIniciativa = getIniciativaID(Trim(partesCarpeta(0)))
            CodIniciativa = Trim(partesCarpeta(0))
            'If ideIniciativa = 0 Then Exit For
            'Console.WriteLine(ideIniciativa)



            If xmlCfg.SelectNodes("content/base/carpeta[@tip='hijo'][@estado='1']").Count = 0 Then GoTo salir

            For j = 0 To xmlCfg.SelectNodes("content/base/carpeta[@tip='hijo'][@estado='1']").Count - 1

                carpetaWRK = carpetas(i) & "\" & xmlCfg.SelectNodes("content/base/carpeta[@tip='hijo'][@estado='1']").Item(j).InnerText
                'Console.WriteLine(carpetaWRK)
                archivos = System.IO.Directory.GetFiles(carpetaWRK)

                If archivos.Length = 0 Then
                    If LCase(xmlCfg.SelectNodes("content/base/carpeta[@tip='hijo'][@estado='1']").Item(j).InnerText) = "4.-reunión periódica" Then
                        carpetaWRK = carpetas(i) & "\" & xmlCfg.SelectNodes("content/base/carpeta[@tip='hijo'][@estado='1']").Item(j).InnerText

                        If archivos.Length >= 1 Then
                            For k = 0 To archivos.Length - 1
                                If InStr(1, archivos(k), ".xlsm", CompareMethod.Text) > 0 Then
                                Else
                                    If InStr(1, archivos(k), ".xlsx", CompareMethod.Text) = 0 Then
                                        If InStr(1, archivos(k), ".xls", CompareMethod.Text) = 0 Then
                                            GoTo nextArchivo_1
                                        End If
                                    End If
                                    'Console.WriteLine(archivos(k))
                                End If
nextArchivo_1:
                            Next
                        End If

                        archivosCarpetas(carpetaWRK, ideIniciativa, CodIniciativa)
                    End If
                End If

                If archivos.Length >= 1 Then

                    numMax = 0 : numMaxArchivo = ""

                    For k = 0 To archivos.Length - 1

                        If InStr(1, archivos(k), "~$", CompareMethod.Text) > 0 Then GoTo nextArchivo

                        If InStr(1, archivos(k), ".xlsm", CompareMethod.Text) > 0 Then
                        Else
                            If InStr(1, archivos(k), ".xlsx", CompareMethod.Text) = 0 Then
                                If InStr(1, archivos(k), ".xls", CompareMethod.Text) = 0 Then
                                    GoTo nextArchivo
                                End If
                            End If
                        End If

                        rutaArchivoTemporal = Replace(archivos(k), carpetaWRK & "\", "")

                        Select Case LCase(xmlCfg.SelectNodes("content/base/carpeta[@tip='hijo'][@estado='1']").Item(j).InnerText)
                            Case "2.-check list"
                                If InStr(1, rutaArchivoTemporal, " - ", CompareMethod.Text) < 10 Then
                                    partesArchivos = Split(rutaArchivoTemporal, " - ")
                                    If IsNumeric(Trim(partesArchivos(0))) Then
                                        If numMax < CInt(Trim(partesArchivos(0))) Then
                                            numMax = CInt(Trim(partesArchivos(0)))
                                            numMaxArchivo = rutaArchivoTemporal
                                        End If
                                    End If
                                End If


                            Case "3.-planificación"
                                If InStr(1, rutaArchivoTemporal, " - ", CompareMethod.Text) < 10 Then
                                    partesArchivos = Split(rutaArchivoTemporal, " - ")
                                    If IsNumeric(Trim(partesArchivos(0))) Then
                                        If numMax < CInt(Trim(partesArchivos(0))) Then
                                            numMax = CInt(Trim(partesArchivos(0)))
                                            numMaxArchivo = rutaArchivoTemporal
                                        End If
                                    End If
                                End If


                            Case "5.-administrador_compromisos"
                                numMaxArchivo = rutaArchivoTemporal

                        End Select
nextArchivo:
                    Next


                    If InStr(1, numMaxArchivo, "~$", CompareMethod.Text) > 0 Then GoTo nextCarpeta

                    If numMaxArchivo <> "" Then
                        Select Case LCase(xmlCfg.SelectNodes("content/base/carpeta[@tip='hijo'][@estado='1']").Item(j).InnerText)
                            Case "2.-check list"
                                numMaxArchivo = carpetaWRK & "\" & Replace(numMaxArchivo, "~$", "")

                                Log("2. Procesando archivo : " & numMaxArchivo, "")
                                Console.WriteLine("2. Procesando archivo : " & numMaxArchivo)

                                archivoExcelCheckListInicial = "" & numMaxArchivo & ""
                                archivoExcelCheckListInicial = Chr(34) & archivoExcelCheckListInicial & Chr(34)

                                Dim partes() As String
                                Dim fecha_ultima As String
                                Dim fecha_ult_realizacion As String


                                partes = Split(archivoExcelCheckListInicial, "\")

                                fecha_realizacion_check = Trim(partes(UBound(partes)))
                                fecha_ultima = fecha_realizacion_check.Substring(fecha_realizacion_check.LastIndexOf("_") + 1)
                                fecha_ult_realizacion = fecha_ultima.Substring(0, fecha_ultima.IndexOf(".")).Trim()


                                Try
                                    Call ModeloCheckListsInicial(ideIniciativa, CodIniciativa, archivoExcelCheckListInicial, fecha_ult_realizacion)
                                Catch ex As Exception
                                    Log("Error Carga funcion ModeloCheckListsInicial: " & ex.Message, "error")
                                    Console.WriteLine("ERROR : " & ex.Message)
                                    Console.Write("-------")
                                    Console.ReadLine()

                                End Try

                                Console.WriteLine("2. Archivo procesado : " & numMaxArchivo)


                            Case "3.-planificación"
                                Log("3. Procesando archivo : " & numMaxArchivo, "")

                                numMaxArchivo = carpetaWRK & "\" & Replace(numMaxArchivo, "~$", "")

                                Console.WriteLine("3. Procesando archivo : " & numMaxArchivo)

                                archivoExcelPlanif = "" & numMaxArchivo & ""
                                archivoExcelPlanif = Chr(34) & archivoExcelPlanif & Chr(34)

                                Try

                                    Dim partes_planif() As String
                                    Dim fecha_ultima_planif As String
                                    Dim fecha_ult_realizacion_planificacion As String


                                    partes_planif = Split(archivoExcelPlanif, "\")

                                    fecha_realizacion_check = Trim(partes_planif(UBound(partes_planif)))
                                    fecha_ultima_planif = fecha_realizacion_check.Substring(fecha_realizacion_check.LastIndexOf("_") + 1)
                                    fecha_ult_realizacion_planificacion = fecha_ultima_planif.Substring(0, fecha_ultima_planif.IndexOf(".")).Trim()


                                    'Planificacion Historico
                                    Call ModeloPlanificacion_hist(ideIniciativa, CodIniciativa, fecha_ult_realizacion_planificacion)
                                    'Call ModeloInsertPlanifacion(ideIniciativa, CodIniciativa)

                                    'Planificacion
                                    Call ModeloPlanificacion(ideIniciativa, CodIniciativa, fecha_ult_realizacion_planificacion)

                                Catch ex As Exception
                                    Log("Error Carga funcion ModeloPlanificacion: " & ex.Message, "error")
                                    Console.WriteLine("ERROR : " & ex.Message)
                                    Console.Write("-------")
                                    Console.ReadLine()
                                End Try


                                Console.WriteLine("3. Archivo procesado : " & numMaxArchivo)

                            Case "5.-administrador_compromisos"

                                Log("5. Procesando archivo : " & numMaxArchivo, "")

                                Console.WriteLine("5. Procesando archivo : " & numMaxArchivo)

                                numMaxArchivo = carpetaWRK & "\" & Replace(numMaxArchivo, "~$", "")
                                Console.WriteLine(numMaxArchivo)

                                archivoExcelItemSeguim = "" & numMaxArchivo & ""
                                archivoExcelItemSeguim = Chr(34) & archivoExcelItemSeguim & Chr(34)

                                Try
                                    'ItemSeguimiento Historico
                                    Dim nomb_hoja_matriz_Compromiso As String = "Template_MatrizCompromisos$"
                                    Call ModeloItemSeguimiento_hist(ideIniciativa, CodIniciativa, nomb_hoja_matriz_Compromiso)

                                    'ItemSeguimiento' 
                                    Dim nomb_hoja_matriz_comp As String = "Template_MatrizCompromisos$"
                                    Call ModeloItemSeguimiento(ideIniciativa, CodIniciativa, nomb_hoja_matriz_comp)


                                    'ItemSeguimiento Historico Acuerdos-Ptos
                                    Dim nomb_hoja_matriz_acuerdo As String = "Template_Matriz (Acuerdos-Ptos)$"
                                    Call ModeloItemSeguimiento_hist(ideIniciativa, CodIniciativa, nomb_hoja_matriz_acuerdo)


                                    'ItemSeguimiento Acuerdos-Ptos
                                    Dim nomb_hoja_matriz_acdo As String = "Template_Matriz (Acuerdos-Ptos)$"
                                    Call ModeloItemSeguimiento(ideIniciativa, CodIniciativa, nomb_hoja_matriz_acdo)

                                    Console.WriteLine("5. Archivo procesado : " & numMaxArchivo)

                                Catch ex As Exception
                                    Log("Error Carga funcion ModeloItemSeguimiento: " & ex.Message, "error")
                                    Console.WriteLine("ERROR : " & ex.Message)
                                    Console.Write("-------")
                                    Console.ReadLine()
                                End Try


                        End Select
                    End If

                End If

nextCarpeta:
            Next

nextCarpetaIniciativa:
        Next

        'Catch ex As Exception
        '    Console.WriteLine("ERROR : " & ex.Message)
        '    GoTo salir
        'End Try


salir:
        Log("Terminando archivo de log", "")
        xmlCfg = Nothing

    End Sub

    '***********************************************FUNCION RECURSIVA REUNIÓN PERIODICA*****************************************************
    Private Function archivosCarpetas(ByVal Carpeta As String, ByVal ideIniciativa As String, ByVal CodIniciativa As String)

        Dim carpetasLcl() As String
        Dim archivosLcl() As String
        Dim rutaArchivoTemporal As String
        Dim partesArchivos() As String
        Dim archivoFinal As String
        Dim numMaxArchivo As String
        Dim fecha_act_reunion_period As String

        carpetasLcl = System.IO.Directory.GetDirectories(Carpeta)
        archivosLcl = System.IO.Directory.GetFiles(Carpeta)

        For x = 0 To carpetasLcl.Length - 1

            archivosLcl = System.IO.Directory.GetFiles(carpetasLcl(x))

            If archivosLcl.Length >= 1 Then
                For k = 0 To archivosLcl.Length - 1
                    If InStr(1, archivosLcl(k), ".xlsm", CompareMethod.Text) > 0 Then
                    Else
                        If InStr(1, archivosLcl(k), ".xlsx", CompareMethod.Text) = 0 Then
                            If InStr(1, archivosLcl(k), ".xls", CompareMethod.Text) = 0 Then
                                GoTo nextArchivoLcl_1
                            End If
                        End If

                        rutaArchivoTemporal = Replace(archivosLcl(k), carpetasLcl(x) & "\", "")

                        If InStr(1, rutaArchivoTemporal, " - ", CompareMethod.Text) < 10 Then
                            partesArchivos = Split(rutaArchivoTemporal, " - ")
                            If IsNumeric(Trim(partesArchivos(0))) Then
                                archivoFinal = carpetasLcl(x) & "\" & Replace(rutaArchivoTemporal, "~$", "")

                                'Console.WriteLine(archivoFinal)

                                If InStr(1, archivoFinal, "~$", CompareMethod.Text) = 0 Then

                                    numMaxArchivo = Replace(archivoFinal, "~$", "")

                                    archivoExcelRP = "" & numMaxArchivo & ""
                                    archivoExcelRP = Chr(34) & archivoExcelRP & Chr(34)

                                    Dim partes_reunion_periodica() As String
                                    Dim fecha_reunion As String
                                    Dim fecha_ult_reunion_periodica As String


                                    partes_reunion_periodica = Split(archivoExcelRP, "\")

                                    fecha_act_reunion_period = Trim(partes_reunion_periodica(UBound(partes_reunion_periodica)))
                                    fecha_reunion = fecha_act_reunion_period.Substring(fecha_act_reunion_period.LastIndexOf("_") + 1)
                                    fecha_ult_reunion_periodica = fecha_reunion.Substring(0, fecha_reunion.IndexOf(".")).Trim()


                                    Try
                                        Log("4. Archivo procesado: " & archivoExcelRP, "exito")
                                        Console.WriteLine("4. Procesando archivo : " & archivoExcelRP)

                                        'Reunion Periodica'
                                        Call ModeloReunionPeriodica(ideIniciativa, CodIniciativa, archivoExcelRP, fecha_ult_reunion_periodica)

                                        Console.WriteLine("4. Archivo procesado: " & archivoExcelRP)

                                    Catch ex As Exception
                                        Log("Error Carga funcion ModeloReunionPeriodica: " & ex.Message, "error")
                                        Console.WriteLine("ERROR : " & ex.Message)
                                        Console.Write("-------")
                                        Console.ReadLine()

                                    End Try

                                End If
                            End If
                        End If

                        'Console.WriteLine(archivosLcl(k))
                    End If

nextArchivoLcl_1:
                Next
            End If

            archivosCarpetas(carpetasLcl(x), ideIniciativa, CodIniciativa)
        Next

    End Function

    '****************************************************************************************************************************
    Private Function getIniciativaID(ByVal codIniciativa As String) As String

        Dim dbConexion As Data.Odbc.OdbcConnection
        Dim dbcommand As Data.Odbc.OdbcCommand
        Dim dbdata As Data.Odbc.OdbcDataReader
        Dim dbconsulta As String = ""
        Dim dbresultados As String = ""

        Try
            dbconsulta = "SELECT top 1 ini_ide FROM [dbo].[imp_iniciativa] WHERE ini_cod='" & codIniciativa & "'"

            dbConexion = New Data.Odbc.OdbcConnection(GetConnectionString(0))
            dbcommand = New Data.Odbc.OdbcCommand(dbconsulta, dbConexion)
            dbcommand.CommandType = CommandType.Text
            dbConexion.Open()
            dbdata = dbcommand.ExecuteReader()

            If dbdata.HasRows = True Then
                dbdata.Read()
                dbresultados = dbdata.Item(0).ToString

                dbdata.Close()
                dbConexion.Close()

                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing

                Return dbresultados
            Else
                dbConexion = Nothing
                dbcommand = Nothing
                dbdata = Nothing

                Return "0"
            End If

        Catch ex As Exception
            Console.WriteLine("Se ha producido un error al recuperar el IDE de la iniciativa, ERROR : " & ex.Message)
            Console.ReadLine()
            Return "0"
        End Try

    End Function

    '*********************************************************Sub Main*******************************************************'
    Sub Main()

        '*****************************
        archivoExcelContacto = "C:\Desarrollo SIF\ConstactosIF_Modelo GdeIm.xlsx"
        archivoExcelContacto = Chr(34) & archivoExcelContacto & Chr(34)

        '**************************************************CALL Functions****************************************************************

        GetRutas()

        Console.WriteLine(Now() & "  PROCESO TERMINADO PRESIONE CUALQUIER TECLA PARA CONTINUAR.")

        'ContactosIF
        'Call ModeloContacto()
        'Call ModeloInsertContacto()

        '**********************************************************'
        Console.ReadLine()

    End Sub

    '*****************************************FUNCION PARA PERSONALIZAR LAS CARPETAS A RECORRER DURANTE LA CARGA*******************************'
    Private Function revisarRutas(ByRef xmlCfg As XmlDocument, _
                                  ByVal carpeta As String) As Boolean

        Dim partes() As String
        'Debug.Print(xmlCfg.InnerXml)
        'Debug.Print(xmlCfg.SelectNodes("content/base/carpetas[@estado='0']").Count)

        If xmlCfg.SelectNodes("content/base/carpetas").Count = 0 Then
            Return True
        Else
            If xmlCfg.SelectNodes("content/base/carpetas[@estado='0']").Count = 1 Then
                Return True
            Else
                partes = Split(carpeta, "\")
                'Console.WriteLine("content/base/carpetas[@estado='1'][nom='" & Trim(partes(UBound(partes))) & "']")
                'Console.WriteLine(xmlCfg.SelectNodes("content/base/carpetas[@estado='1'][nom='" & Trim(partes(UBound(partes))) & "']").Count)

                If xmlCfg.SelectNodes("content/base/carpetas[@estado='1'][nom='" & Trim(partes(UBound(partes))) & "']").Count = 1 Then
                    Return True
                Else
                    Return False
                End If
            End If
        End If

    End Function


End Module
