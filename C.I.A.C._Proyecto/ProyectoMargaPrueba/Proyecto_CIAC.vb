
#Region "Librerias de la Aplicación C.I.A.C."

Imports System.IO ' Sistema de Archivos
Imports Word = Microsoft.Office.Interop.Word 'Control de Office
Imports Microsoft.Office.Interop 'Libreria general de Interop para Oficce
Imports System.Data
Imports System
Imports System.Data.SqlClient
Imports System.Globalization

#End Region

Public Class Proyecto_German

    Dim ruta As String
    Dim MainDoc As Word.Document

#Region "Métodos públicos"



    Public KeyAscii As Short

    Dim obj_Word As Word.Application
    Dim obj_Doc As Word.Document
    Dim wd As Word.Application = CreateObject("Word.Application")
    Dim ActiveSheet As Object

    Private Property Panel As Boolean

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'AverDataSet.tablaver' Puede moverla o quitarla según sea necesario.
        Me.TablaverTableAdapter.Fill(Me.AverDataSet.tablaver)
        'quito el valor de posicion poniendolo a cero.
        StartPosition = 0
        'centro el formulario con starposition
        StartPosition = FormStartPosition.CenterScreen

        Etiquetas()
    End Sub

#End Region

#Region "Métodos para Insertar y obtener Datos"

    Private Sub Documento_Inserta_1()

        'Obtengo la Ruta de la carpeta temp, que se crea unicamente para insertar los datos del formulario.
        Dim ruta3 As String
        ruta3 = Application.ExecutablePath.Substring(0, Application.ExecutablePath.LastIndexOf("\")) & "\temp"
        Dim wdDoc As Object = wd.Documents.Open(ruta3 & "\doc_1.doc", Visible:=True)

        '-------------------------------------------------------------------------------------------------------

        wd.ActiveDocument.Bookmarks("SiniestroN").Select()
        wd.Selection.TypeText(TB_Nombre_Archivo.Text)

        wd.ActiveDocument.Bookmarks("NombreAseg").Select()
        wd.Selection.TypeText(TxBx_Nombre_Asegurado.Text)

        wd.ActiveDocument.Bookmarks("Conductor_Vehic_Asegurado_I").Select()
        wd.Selection.TypeText(TB_Nombre_AseguradoII.Text)

        wd.ActiveDocument.Bookmarks("Vehic_Asegurado").Select()
        wd.Selection.TypeText(CB_Marca_Vehiculo_Aseg.Text)

        wd.ActiveDocument.Bookmarks("dominio_vehicu_asegurado_I").Select()
        wd.Selection.TypeText(TB_Dominio_Aseg_Vehiculo.Text)

        wd.ActiveDocument.Bookmarks("Tercero_pasajeros").Select()
        wd.Selection.TypeText(TB_Terceros_conducidos.Text)

        wd.ActiveDocument.Bookmarks("Vehic_Tercero").Select()
        wd.Selection.TypeText(CB_Tipo_Vehiculo_Tercero.Text)

        wd.ActiveDocument.Bookmarks("Marca_vehi_tercero_I").Select()
        wd.Selection.TypeText(CB_Marca_Vehiculo_Tercero.Text)

        wd.ActiveDocument.Bookmarks("Dominio_Vehi_Tecero_I").Select()
        wd.Selection.TypeText(TB_Dominio_Vehiculo_Tercero.Text)

        wd.ActiveDocument.Bookmarks("Fecha_Ocurrencia").Select()
        wd.Selection.TypeText(DTP_Fecha_Ocurrencia.Text)

        wd.ActiveDocument.Bookmarks("Hora_ocurrencia").Select()
        wd.Selection.TypeText(TB_Hora_Ocurrencia.Text)

        wd.ActiveDocument.Bookmarks("Lugar_Ocurrencia").Select()
        wd.Selection.TypeText(TB_Lugar_Ocurrencia.Text)

        wd.ActiveDocument.Bookmarks("Responsabilidad_Aseg").Select()
        wd.Selection.TypeText(CB_Resp_Asegurado.Text)

        wd.ActiveDocument.Bookmarks("Exclusiones").Select()
        wd.Selection.TypeText(CB_Exclusiones.Text)

        wd.ActiveDocument.Bookmarks("Fraude").Select()
        wd.Selection.TypeText(CB_Fraude.Text)

        wd.ActiveDocument.Bookmarks("Abogado_Tercero").Select()
        wd.Selection.TypeText(CB_Abogado_Tercero.Text)

        '--------------------------------------------------------------------------------------------------------

        wdDoc.Save()
        wd.ActiveDocument.Close()
        wdDoc = Nothing

    End Sub

    Private Sub Documento_Inserta_2()

        Dim ruta3 As String
        ruta3 = Application.ExecutablePath.Substring(0, Application.ExecutablePath.LastIndexOf("\")) & "\temp"
        Dim wdDoc As Object = wd.Documents.Open(ruta3 & "\doc_2.doc", Visible:=True)

        '--------------------------------------  doc_2 ----------------------------------------------

        wd.ActiveDocument.Bookmarks("Informe_final_siniestroN").Select()
        wd.Selection.TypeText(TB_Nombre_Archivo.Text)

        wd.ActiveDocument.Bookmarks("MPS_aseg").Select()
        wd.Selection.TypeText(CB_Marca_Vehiculo_Aseg.Text)

        wd.ActiveDocument.Bookmarks("Conductor_vehic_Aseg").Select()
        wd.Selection.TypeText(TB_Nombre_Aseg_Vehiculo.Text)

        wd.ActiveDocument.Bookmarks("Edad_Aseg").Select()
        wd.Selection.TypeText(TB_Edad_Aseg_Vehiculo.Text)

        wd.ActiveDocument.Bookmarks("DNI_Aseg").Select()
        wd.Selection.TypeText(TB_DNI_Aseg_Vehiculo.Text)

        wd.ActiveDocument.Bookmarks("EstadoCivil_Aseg").Select()
        wd.Selection.TypeText(CB_EstCivil_Aseg_Vehiculo.Text)

        wd.ActiveDocument.Bookmarks("Ocupacion_Aseg").Select()
        wd.Selection.TypeText(CB_Ocupacion_Aseg_Vehiculo.Text)

        wd.ActiveDocument.Bookmarks("Domicilio_Aseg").Select()
        wd.Selection.TypeText(TB_Domicilio_Aseg_Vehiculo.Text)

        wd.ActiveDocument.Bookmarks("Telef_Aseg").Select()
        wd.Selection.TypeText(TB_Telf_Aseg_Vehiculo.Text)

        wd.ActiveDocument.Bookmarks("LicDeConducir").Select()
        wd.Selection.TypeText(TB_LicConducir_Aseg_Vehiculo.Text)

        wd.ActiveDocument.Bookmarks("Desde_Aseg").Select()
        wd.Selection.TypeText(TB_Desde_Aseg_Vehiculo.Text)

        wd.ActiveDocument.Bookmarks("Hasta_Aseg").Select()
        wd.Selection.TypeText(TB_Hasta_Aseg_Vehiculo.Text)

        wd.ActiveDocument.Bookmarks("Relacion_Aseg").Select()
        wd.Selection.TypeText(TB_Relacion_Aseg_Vehiculo.Text)

        wd.ActiveDocument.Bookmarks("Tipo_Vehi_Aseg").Select()
        wd.Selection.TypeText(CB_Tipo_Vehiculo_Aseg.Text)

        wd.ActiveDocument.Bookmarks("Marca_Vehi_Aseg").Select()
        wd.Selection.TypeText(CB_Marca_Vehiculo_Aseg.Text)

        wd.ActiveDocument.Bookmarks("Año_Vehi_Aseg").Select()
        wd.Selection.TypeText(TB_Año_Vehiculo_Aseg.Text)

        wd.ActiveDocument.Bookmarks("Dominio_Vehi_Aseg").Select()
        wd.Selection.TypeText(TB_Dominio_Aseg_Vehiculo.Text)

        wd.ActiveDocument.Bookmarks("Uso_Vehi_Aseg").Select()
        wd.Selection.TypeText(CB_Uso_Vehiculo_Aseg.Text)

        wd.ActiveDocument.Bookmarks("Daños_Vehi_Aseg").Select()
        wd.Selection.TypeText(TB_Daños_Aseg_Vehiculo.Text)

        '-------------------------------------------------------------------------------------------

        wdDoc.Save()
        wd.ActiveDocument.Close()
        wdDoc = Nothing

    End Sub

    Private Sub Documento_Inserta_3()

        '--------------------------------------  doc_3 ----------------------------------------------
        Dim ruta3 As String
        ruta3 = Application.ExecutablePath.Substring(0, Application.ExecutablePath.LastIndexOf("\")) & "\temp"
        Dim wdDoc As Object = wd.Documents.Open(ruta3 & "\doc_3.doc", Visible:=True)



        wd.ActiveDocument.Bookmarks("NombreTer").Select()
        wd.Selection.TypeText(TB_Nombre_Tercero_Conductor.Text)

        wd.ActiveDocument.Bookmarks("Edad_Ter").Select()
        wd.Selection.TypeText(TB_Edad_Tercero_Conductor.Text)

        wd.ActiveDocument.Bookmarks("DNI_Ter").Select()
        wd.Selection.TypeText(TB_DNI_Tercero_Conductor.Text)

        wd.ActiveDocument.Bookmarks("EstadoCivil_Ter").Select()
        wd.Selection.TypeText(CB_EstadoCivil_Tercero_Conductor.Text)

        wd.ActiveDocument.Bookmarks("Ocupacion_Ter").Select()
        wd.Selection.TypeText(CB_Ocupacion_Tercero_Conductor.Text)

        wd.ActiveDocument.Bookmarks("Domicilio_Ter").Select()
        wd.Selection.TypeText(TB_Domicilio_Tercero_Conductor.Text)

        wd.ActiveDocument.Bookmarks("Telef_Ter").Select()
        wd.Selection.TypeText(TB_Telf_Tercero_Conductor.Text)

        wd.ActiveDocument.Bookmarks("LicDeConducirTer").Select()
        wd.Selection.TypeText(TB_LicConducir_Tercero_Conductor.Text)

        wd.ActiveDocument.Bookmarks("Desde_Ter").Select()
        wd.Selection.TypeText(TB_Desde_Tercero_Conductor.Text)

        wd.ActiveDocument.Bookmarks("Hasta_Ter").Select()
        wd.Selection.TypeText(TB_Hasta_Tercero_Conductor.Text)

        wd.ActiveDocument.Bookmarks("Relacion_Ter").Select()
        wd.Selection.TypeText(TB_Relacion_Tercero_Conductor.Text)

        wd.ActiveDocument.Bookmarks("Tipo_Vehi_Tercero").Select()
        wd.Selection.TypeText(CB_Tipo_Vehiculo_Tercero.Text)

        wd.ActiveDocument.Bookmarks("Marca_Vehi_Tercero").Select()
        wd.Selection.TypeText(CB_Marca_Vehiculo_Tercero.Text)

        wd.ActiveDocument.Bookmarks("Año_Vehi_Tercero").Select()
        wd.Selection.TypeText(TB_Año_Vehiculo_Tercero.Text)

        wd.ActiveDocument.Bookmarks("Dominio_Vehi_Tercero").Select()
        wd.Selection.TypeText(TB_Dominio_Vehiculo_Tercero.Text)

        wd.ActiveDocument.Bookmarks("Uso_Vehi_Tercero").Select()
        wd.Selection.TypeText(CB_Uso_Vehiculo_Tercero.Text)

        wd.ActiveDocument.Bookmarks("Aseguradora_Vehi_Tercero").Select()
        wd.Selection.TypeText(CB_Aseguradora_Vehiculo_Tercero.Text)

        wd.ActiveDocument.Bookmarks("Titular_Vehi_Tercero").Select()
        wd.Selection.TypeText(TB_Titular_Vehiculo_Tercero.Text)

        wd.ActiveDocument.Bookmarks("DNI_Vehi_Tercero").Select()
        wd.Selection.TypeText(TB_DNI_Vehiculo_Tercero.Text)

        wd.ActiveDocument.Bookmarks("Telefono_Vehi_Tercero").Select()
        wd.Selection.TypeText(TB_Telf_Vehiculo_Tercero.Text)

        wd.ActiveDocument.Bookmarks("Daños_Vehi_Tercero").Select()
        wd.Selection.TypeText(TB_Daños_Vehiculo_Tercero.Text)

        wd.ActiveDocument.Bookmarks("Domicilio_Vehi_Tercero").Select()
        wd.Selection.TypeText(TB_Domicilio_Vehiculo_Tercero.Text)

        wd.ActiveDocument.Bookmarks("Tot_Peritado_Vehi_Tercero").Select()
        wd.Selection.TypeText(TB_TotPeritado_Vehiculo_Tercero.Text)

        wd.ActiveDocument.Bookmarks("Por_Inspector_Vehi_Tercero").Select()
        wd.Selection.TypeText(TB_Inspector_Vehiculo_Tercero.Text)



        wd.ActiveDocument.Bookmarks("Vict_Cantidad").Select()
        wd.Selection.TypeText(CB_Cantidad_Victimas.Text)

        wd.ActiveDocument.Bookmarks("Vict_Nombre").Select()
        wd.Selection.TypeText(TB_Nombre_Victimas.Text)

        wd.ActiveDocument.Bookmarks("Vict_edad").Select()
        wd.Selection.TypeText(TB_Edad_Victimas.Text)

        wd.ActiveDocument.Bookmarks("Vict_DNI").Select()
        wd.Selection.TypeText(TB_DNI_Victimas.Text)

        wd.ActiveDocument.Bookmarks("Vict_Ocupacion").Select()
        wd.Selection.TypeText(CB_Ocupacion_Victimas.Text)

        wd.ActiveDocument.Bookmarks("Vict_Telef").Select()
        wd.Selection.TypeText(TB_Telf_Victimas.Text)

        wd.ActiveDocument.Bookmarks("Vict_Domicilio").Select()
        wd.Selection.TypeText(TB_Domicilio_Victimas.Text)

        wd.ActiveDocument.Bookmarks("Vict_Inca_Futura").Select()
        wd.Selection.TypeText(TB_IncapacidadFutura_Victimas.Text)

        wd.ActiveDocument.Bookmarks("Vict_Inca_Judicial").Select()
        wd.Selection.TypeText(TB_IncapacidadJudicial_Victimas.Text)

        wd.ActiveDocument.Bookmarks("Vict_Lesiones").Select()
        wd.Selection.TypeText(TB_Lesiones_Victimas.Text)

        wd.ActiveDocument.Bookmarks("Vict_Observaciones").Select()
        wd.Selection.TypeText(TB_Observaciones_Victimas.Text)


        '-------------------------------------------------------------------------------------------

        wdDoc.Save()
        wd.ActiveDocument.Close()
        wdDoc = Nothing

    End Sub

    Private Sub Documento_Inserta_4()


        Dim ruta3 As String
        ruta3 = Application.ExecutablePath.Substring(0, Application.ExecutablePath.LastIndexOf("\")) & "\temp"
        Dim wdDoc As Object = wd.Documents.Open(ruta3 & "\doc_4.doc", Visible:=True)

        '--------------------------------------  doc_4 ----------------------------------------------



        wd.ActiveDocument.Bookmarks("IPJ_dep_Interventora").Select()
        wd.Selection.TypeText(TB_DepInterventora_Detalles.Text)

        wd.ActiveDocument.Bookmarks("IPJ_Sumario").Select()
        wd.Selection.TypeText(TB_SumarioNº_Detalles.Text)

        wd.ActiveDocument.Bookmarks("IPJ_Fiscalia").Select()
        wd.Selection.TypeText(TB_Fiscalia_Detalles.Text)



        wd.ActiveDocument.Bookmarks("Inf_Amb_P1").Select()
        wd.Selection.TypeText(CB_InfAmbiental_I_Victimas.Text)

        wd.ActiveDocument.Bookmarks("Inf_Amb_P2").Select()
        wd.Selection.TypeText(CB_InfAmbiental_II_Victimas.Text)

        wd.ActiveDocument.Bookmarks("Inf_Amb_P3").Select()
        wd.Selection.TypeText(CB_InfAmbiental_III_Victimas.Text)

        wd.ActiveDocument.Bookmarks("Inf_Amb_P4").Select()
        wd.Selection.TypeText(CB_InfAmbiental_IV_Victimas.Text)

        wd.ActiveDocument.Bookmarks("Inf_Amb_P5").Select()
        wd.Selection.TypeText(CB_InfAmbiental_V_Victimas.Text)

        wd.ActiveDocument.Bookmarks("Inf_Amb_P6").Select()
        wd.Selection.TypeText(CB_InfAmbiental_VI_Victimas.Text)

        wd.ActiveDocument.Bookmarks("Inf_Amb_P7").Select()
        wd.Selection.TypeText(CB_InfAmbiental_VII_Victimas.Text)

        wd.ActiveDocument.Bookmarks("Analisis_Tec_Cientifico").Select()
        wd.Selection.TypeText(TB_Analisis_Tecnico_Cientifico.Text)

        'aqui faltan los CheckBox ?????????????????????
        'TB_DepInterventora_Detalles.Text = obj_Doc.Bookmarks.Item("").Range.Text
        'TB_DepInterventora_Detalles.Text = obj_Doc.Bookmarks.Item("").Range.Text
        'TB_DepInterventora_Detalles.Text = obj_Doc.Bookmarks.Item("").Range.Text


        '-------------------------------------------------------------------------------------------
        wdDoc.Save()
        wd.ActiveDocument.Close()
        wdDoc = Nothing

    End Sub

    Private Sub Documento_Inserta_5()


        Dim ruta3 As String
        ruta3 = Application.ExecutablePath.Substring(0, Application.ExecutablePath.LastIndexOf("\")) & "\temp"
        Dim wdDoc As Object = wd.Documents.Open(ruta3 & "\doc_5.doc", Visible:=True)

        '--------------------------------------  doc_5 ----------------------------------------------

        wd.ActiveDocument.Bookmarks("Atrib_de_Responsabilidad").Select()
        wd.Selection.TypeText(CB_Atribucion_Responsabilidad.Text)



        wd.ActiveDocument.Bookmarks("Curso_de_Accion").Select()
        wd.Selection.TypeText(CB_Pasivo.Text)

        wd.ActiveDocument.Bookmarks("Curso_de_Acc_Observaciones").Select()
        wd.Selection.TypeText(TB_Observaciones.Text)


        '-------------------------------------------------------------------------------------------
        wdDoc.Save()
        wd.ActiveDocument.Close()
        wdDoc = Nothing

    End Sub


    'Sigue aun en fase de desarrollo...
    Private Sub Documento_Obtiene_1()

        Dim ruta4 As String
        Dim oWrd = New Word.Application
        ruta4 = Application.ExecutablePath.Substring(0, Application.ExecutablePath.LastIndexOf("\")) & "\Siniestros"
        Dim wdDoc As Word.Document = wd.Documents.Open(ruta4 & "\BA" & TB_Nombre_Archivo.Text & ".doc", Visible:=True)
        '--------------------------------------  doc_1 ----------------------------------------------

        If Not (wdDoc Is Nothing) Then
            With wdDoc

                TB_Nombre_Archivo.Text = wdDoc.Bookmarks.Item("SiniestroN").Range.Text

                TxBx_Nombre_Asegurado.Text = wdDoc.Bookmarks.Item("NombreAseg").Range.Text

                TB_Nombre_AseguradoII.Text = wdDoc.Bookmarks.Item("Conductor_Vehic_Asegurado_I").Range.Text

                CB_Marca_Vehiculo_Aseg.Text = wdDoc.Bookmarks.Item("Vehic_Asegurado").Range.Text

                TB_Dominio_Aseg_Vehiculo.Text = wdDoc.Bookmarks.Item("dominio_vehicu_asegurado_I").Range.Text

                TB_Terceros_conducidos.Text = wdDoc.Bookmarks.Item("Tercero_pasajeros").Range.Text

                CB_Tipo_Vehiculo_Tercero.Text = wdDoc.Bookmarks.Item("Vehic_Tercero").Range.Text

                CB_Marca_Vehiculo_Tercero.Text = wdDoc.Bookmarks.Item("Marca_vehi_tercero_I").Range.Text

                TB_Dominio_Vehiculo_Tercero.Text = wdDoc.Bookmarks.Item("Dominio_Vehi_Tecero_I").Range.Text

                DTP_Fecha_Ocurrencia.Text = wdDoc.Bookmarks.Item("Fecha_Ocurrencia").Range.Text

                TB_Hora_Ocurrencia.Text = wdDoc.Bookmarks.Item("Hora_ocurrencia").Range.Text

                TB_Lugar_Ocurrencia.Text = wdDoc.Bookmarks.Item("Lugar_Ocurrencia").Range.Text

                CB_Resp_Asegurado.Text = wdDoc.Bookmarks.Item("Responsabilidad_Aseg").Range.Text

                CB_Exclusiones.Text = wdDoc.Bookmarks.Item("Exclusiones").Range.Text

                CB_Fraude.Text = wdDoc.Bookmarks.Item("Fraude").Range.Text

                CB_Abogado_Tercero.Text = wdDoc.Bookmarks.Item("Abogado_Tercero").Range.Text

            End With
        End If


        'oWrd = New Word.Application
        'odoc = oWrd.Documents.Open(App.Path & "\Datos1.doc") 'Abro Word


        'If Not (odoc Is Nothing) Then
        '    With odoc
        '        'Cargo el dato de ape en el formulario
        '        Text1.Text = .Bookmarks.Item("ape").Range.Words.First
        '        'Esto solo me carga "Sanchez" ponga .First o .Last
        '    End With
        'End If

        ''Cierro word
        ''------------------
        'odoc.Saved = True
        'odoc.Close()
        'odoc = Nothing
        'oWrd.Quit(False)
        'oWord = Nothing



        'wd.ActiveDocument.Bookmarks("NombreAseg").Select()
        'wd.ActiveDocument.Bookmarks("NombreAseg").Range.Select()
        'TxBx_Nombre_Asegurado.Text = wd.ActiveDocument.Bookmarks("NombreAseg").Range.Text


        'wd.selection.TypeText(CB_Atribucion_Responsabilidad.Text)
        'TxBx_Nombre_Asegurado. TxBx_Nombre_Asegurado.Text = wdDoc.Bookmarks.Item("NombreAseg").Range.Text


        'wd.activedocument.Bookmarks("NombreAseg").Select()
        'wd.selection.TypeText(TxBx_Nombre_Asegurado.Text)

        'Dim ruta3 As String
        'ruta3 = Application.ExecutablePath.Substring(0, Application.ExecutablePath.LastIndexOf("\")) & "\Siniestros"
        'Dim wdDoc As Object = wd.documents.Open(ruta3 & "\Abrir_Este.doc", Visible:=True)


        ''--------------------------------------  doc_1 ----------------------------------------------




        'wd.activedocument.Bookmarks("SiniestroN").select()
        'wd.selection.range.text(TB_Nombre_Archivo.Text)


        'wd.activedocument.Bookmarks("NombreAseg").Select()
        'wd.selection.range.text(TxBx_Nombre_Asegurado.Text)




        'obj_Doc.Bookmarks.Item("SiniestroNº").Range.Text = TB_Nombre_Archivo.Text
        'obj_Doc.Bookmarks.Item("NombreAseg").Range.Text = TxBx_Nombre_Asegurado.Text


        '------------------------------------------------------------------------------------------------------------
        'odoc.Saved = True
        'odoc.Close()
        'odoc = Nothing
        'oWrd.Quit(False)
        'oWord = Nothing

        wdDoc.Saved = True
        wdDoc.Close()
        wdDoc = Nothing
        oWrd.Quit(False)
        oWrd = Nothing

        wd.ActiveDocument.Close()
        '  wd = Nothing

    End Sub

    Private Sub Documento_Obtiene_2()

        '--------------------------------------  doc_2 ----------------------------------------------

        obj_Doc.Bookmarks.Item("Informe_final_siniestroN").Range.Text = TB_Nombre_Archivo.Text
        obj_Doc.Bookmarks.Item("MPS_aseg").Range.Text = CB_Marca_Vehiculo_Aseg.Text
        obj_Doc.Bookmarks.Item("Conductor_vehic_Aseg").Range.Text = TB_Nombre_Aseg_Vehiculo.Text
        obj_Doc.Bookmarks.Item("Edad_Aseg").Range.Text = TB_Edad_Aseg_Vehiculo.Text
        obj_Doc.Bookmarks.Item("DNI_Aseg").Range.Text = TB_DNI_Aseg_Vehiculo.Text
        obj_Doc.Bookmarks.Item("EstadoCivil_Aseg").Range.Text = CB_EstCivil_Aseg_Vehiculo.Text
        obj_Doc.Bookmarks.Item("Ocupacion_Aseg").Range.Text = CB_Ocupacion_Aseg_Vehiculo.Text
        obj_Doc.Bookmarks.Item("Domicilio_Aseg").Range.Text = TB_Domicilio_Aseg_Vehiculo.Text
        obj_Doc.Bookmarks.Item("Telef_Aseg").Range.Text = TB_Telf_Aseg_Vehiculo.Text
        obj_Doc.Bookmarks.Item("LicDeConducir").Range.Text = TB_LicConducir_Aseg_Vehiculo.Text
        obj_Doc.Bookmarks.Item("Desde_Aseg").Range.Text = TB_Desde_Aseg_Vehiculo.Text
        obj_Doc.Bookmarks.Item("Hasta_Aseg").Range.Text = TB_Hasta_Aseg_Vehiculo.Text
        obj_Doc.Bookmarks.Item("Relacion_Aseg").Range.Text = TB_Relacion_Aseg_Vehiculo.Text
        obj_Doc.Bookmarks.Item("Tipo_Vehi_Aseg").Range.Text = CB_Tipo_Vehiculo_Aseg.Text
        obj_Doc.Bookmarks.Item("Marca_Vehi_Aseg").Range.Text = CB_Marca_Vehiculo_Aseg.Text
        obj_Doc.Bookmarks.Item("Año_Vehi_Aseg").Range.Text = TB_Año_Vehiculo_Aseg.Text
        obj_Doc.Bookmarks.Item("Dominio_Vehi_Aseg").Range.Text = TB_Dominio_Aseg_Vehiculo.Text
        obj_Doc.Bookmarks.Item("Uso_Vehi_Aseg").Range.Text = CB_Uso_Vehiculo_Aseg.Text
        obj_Doc.Bookmarks.Item("Daños_Vehi_Aseg").Range.Text = TB_Daños_Aseg_Vehiculo.Text

        '-------------------------------------------------------------------------------------------

    End Sub

    Private Sub Documento_Obtiene_3()

        '--------------------------------------  doc_3 ----------------------------------------------

        obj_Doc.Bookmarks.Item("NombreTer").Range.Text = TB_Nombre_Tercero_Conductor.Text
        obj_Doc.Bookmarks.Item("Edad_Ter").Range.Text = TB_Edad_Tercero_Conductor.Text
        obj_Doc.Bookmarks.Item("DNI_Ter").Range.Text = TB_DNI_Tercero_Conductor.Text
        obj_Doc.Bookmarks.Item("EstadoCivil_Ter").Range.Text = CB_EstadoCivil_Tercero_Conductor.Text
        obj_Doc.Bookmarks.Item("Ocupacion_Ter").Range.Text = CB_Ocupacion_Tercero_Conductor.Text
        obj_Doc.Bookmarks.Item("Domicilio_Ter").Range.Text = TB_Domicilio_Tercero_Conductor.Text
        obj_Doc.Bookmarks.Item("Telef_Ter").Range.Text = TB_Telf_Tercero_Conductor.Text
        obj_Doc.Bookmarks.Item("LicDeConducirTer").Range.Text = TB_LicConducir_Tercero_Conductor.Text
        obj_Doc.Bookmarks.Item("Desde_Ter").Range.Text = TB_Desde_Tercero_Conductor.Text
        obj_Doc.Bookmarks.Item("Hasta_Ter").Range.Text = TB_Hasta_Tercero_Conductor.Text
        obj_Doc.Bookmarks.Item("Relacion_Ter").Range.Text = TB_Relacion_Tercero_Conductor.Text

        obj_Doc.Bookmarks.Item("Tipo_Vehi_Tercero").Range.Text = CB_Tipo_Vehiculo_Tercero.Text
        obj_Doc.Bookmarks.Item("Marca_Vehi_Tercero").Range.Text = CB_Marca_Vehiculo_Tercero.Text
        obj_Doc.Bookmarks.Item("Año_Vehi_Tercero").Range.Text = TB_Año_Vehiculo_Tercero.Text
        obj_Doc.Bookmarks.Item("Dominio_Vehi_Tercero").Range.Text = TB_Dominio_Vehiculo_Tercero.Text
        obj_Doc.Bookmarks.Item("Uso_Vehi_Tercero").Range.Text = CB_Uso_Vehiculo_Tercero.Text
        obj_Doc.Bookmarks.Item("Aseguradora_Vehi_Tercero").Range.Text = CB_Aseguradora_Vehiculo_Tercero.Text
        obj_Doc.Bookmarks.Item("Titular_Vehi_Tercero").Range.Text = TB_Titular_Vehiculo_Tercero.Text
        obj_Doc.Bookmarks.Item("DNI_Vehi_Tercero").Range.Text = TB_DNI_Vehiculo_Tercero.Text
        obj_Doc.Bookmarks.Item("Telefono_Vehi_Tercero").Range.Text = TB_Telf_Vehiculo_Tercero.Text
        obj_Doc.Bookmarks.Item("Daños_Vehi_Tercero").Range.Text = TB_Daños_Vehiculo_Tercero.Text
        obj_Doc.Bookmarks.Item("Domicilio_Vehi_Tercero").Range.Text = TB_Domicilio_Vehiculo_Tercero.Text
        obj_Doc.Bookmarks.Item("Tot_Peritado_Vehi_Tercero").Range.Text = TB_TotPeritado_Vehiculo_Tercero.Text
        obj_Doc.Bookmarks.Item("Por_Inspector_Vehi_Tercero").Range.Text = TB_Inspector_Vehiculo_Tercero.Text

        obj_Doc.Bookmarks.Item("Vict_Cantidad").Range.Text = CB_Cantidad_Victimas.Text
        obj_Doc.Bookmarks.Item("Vict_Nombre").Range.Text = TB_Nombre_Victimas.Text
        obj_Doc.Bookmarks.Item("Vict_edad").Range.Text = TB_Edad_Victimas.Text
        obj_Doc.Bookmarks.Item("Vict_DNI").Range.Text = TB_DNI_Victimas.Text
        obj_Doc.Bookmarks.Item("Vict_Ocupacion").Range.Text = CB_Ocupacion_Victimas.Text
        obj_Doc.Bookmarks.Item("Vict_Telef").Range.Text = TB_Telf_Victimas.Text
        obj_Doc.Bookmarks.Item("Vict_Domicilio").Range.Text = TB_Domicilio_Victimas.Text
        obj_Doc.Bookmarks.Item("Vict_Inca_Futura").Range.Text = TB_IncapacidadFutura_Victimas.Text
        obj_Doc.Bookmarks.Item("Vict_Inca_Judicial").Range.Text = TB_IncapacidadJudicial_Victimas.Text
        obj_Doc.Bookmarks.Item("Vict_Lesiones").Range.Text = TB_Lesiones_Victimas.Text
        obj_Doc.Bookmarks.Item("Vict_Observaciones").Range.Text = TB_Observaciones_Victimas.Text

        '-------------------------------------------------------------------------------------------

    End Sub

    Private Sub Documento_Obtiene_4()

        '--------------------------------------  doc_4 ----------------------------------------------

        obj_Doc.Bookmarks.Item("IPJ_dep_Interventora").Range.Text = TB_DepInterventora_Detalles.Text
        obj_Doc.Bookmarks.Item("IPJ_Sumario").Range.Text = TB_SumarioNº_Detalles.Text
        obj_Doc.Bookmarks.Item("IPJ_Fiscalia").Range.Text = TB_Fiscalia_Detalles.Text

        obj_Doc.Bookmarks.Item("Inf_Amb_P1").Range.Text = CB_InfAmbiental_I_Victimas.Text
        obj_Doc.Bookmarks.Item("Inf_Amb_P2").Range.Text = CB_InfAmbiental_II_Victimas.Text
        obj_Doc.Bookmarks.Item("Inf_Amb_P3").Range.Text = CB_InfAmbiental_III_Victimas.Text
        obj_Doc.Bookmarks.Item("Inf_Amb_P4").Range.Text = CB_InfAmbiental_IV_Victimas.Text
        obj_Doc.Bookmarks.Item("Inf_Amb_P5").Range.Text = CB_InfAmbiental_V_Victimas.Text
        obj_Doc.Bookmarks.Item("Inf_Amb_P6").Range.Text = CB_InfAmbiental_VI_Victimas.Text
        obj_Doc.Bookmarks.Item("Inf_Amb_P7").Range.Text = CB_InfAmbiental_VII_Victimas.Text


        TB_Analisis_Tecnico_Cientifico.Text = obj_Doc.Bookmarks.Item("Analisis_Tec_Cientifico").Range.Text

        'aqui faltan los CheckBox ?????????????????????
        'TB_DepInterventora_Detalles.Text = obj_Doc.Bookmarks.Item("").Range.Text
        'TB_DepInterventora_Detalles.Text = obj_Doc.Bookmarks.Item("").Range.Text
        'TB_DepInterventora_Detalles.Text = obj_Doc.Bookmarks.Item("").Range.Text


        '-------------------------------------------------------------------------------------------

    End Sub

    Private Sub Documento_Obtiene_5()

        '--------------------------------------  doc_5 ----------------------------------------------

        obj_Doc.Bookmarks.Item("Atrib_de_Responsabilidad").Range.Text = CB_Atribucion_Responsabilidad.Text

        obj_Doc.Bookmarks.Item("Curso_de_Accion").Range.Text = CB_Pasivo.Text
        obj_Doc.Bookmarks.Item("Curso_de_Acc_Observaciones").Range.Text = TB_Observaciones.Text

        '-------------------------------------------------------------------------------------------

    End Sub


#End Region

#Region "Métodos Etiquetas"

    Private Sub Etiquetas()
        'etiquetas del Proyecto Germán
        ToolTip.SetToolTip(TB_Nombre_Archivo, "Ingrese un número de siniestro.")
        ToolTip.SetToolTip(TC_CIAC, "Ingrese un número de siniestro.")
        ToolTip.SetToolTip(BT_Generar, "Pulse este botón para generar el documento Word.")
        ToolTip.SetToolTip(BT_Comprobar, "Pulse este botón para comprobar si existe el documento Word o si desea insertarlo.")
        ToolTip.SetToolTip(Status, "Barra de estado.")
        ToolTip.SetToolTip(TC_CIAC, "Pestañas del Informe.")

        '---- PESTAÑA ---- PRINCIPAL ------

        ToolTip.SetToolTip(TxBx_Nombre_Asegurado, "Ingrese el nombre del asegurado.")
        ToolTip.SetToolTip(DTP_Fecha_Ocurrencia, "Ingrese la fecha de ocurrencia del siniestro.")
        ToolTip.SetToolTip(TB_Lugar_Ocurrencia, "Ingrese el lugar de ocurrencia del siniestro.")
        ToolTip.SetToolTip(CB_Resp_Asegurado, "Ingrese el tipo de responsabilidad del asegurado.")
        ToolTip.SetToolTip(CB_Exclusiones, "Ingrese una exclusión.")
        ToolTip.SetToolTip(CB_Fraude, "Ingrese si hubo evidencia de fraude.")
        ToolTip.SetToolTip(CB_Abogado_Tercero, "Ingrese si se registró el abogado del tercero.")

    End Sub


#End Region

#Region "Métodos Botón Generar Word"

    Private Sub Generador()
        Try

            'obtenemos las rutas de las carpetas plantilla, temp y siniestros.
            Dim ruta2, ruta3, ruta4 As String
            ruta2 = Application.ExecutablePath.Substring(0, Application.ExecutablePath.LastIndexOf("\")) & "\plantilla"
            ruta3 = Application.ExecutablePath.Substring(0, Application.ExecutablePath.LastIndexOf("\")) & "\temp"
            ruta4 = Application.ExecutablePath.Substring(0, Application.ExecutablePath.LastIndexOf("\")) & "\Siniestros"


            obj_Word = New Word.Application()
            obj_Word.Visible = True

            'Abro el Documento Word

            Dim Files As String()

            'Obtengo los archivos de la carpeta ‘\plantilla’
            Files = IO.Directory.GetFiles(Directory.GetCurrentDirectory() + "\plantilla")
            ruta = Convert.ToString(Files)
            Dim contador = 0


            'copiamos la carpeta "plantilla" en la carpeta "Temp"
            My.Computer.FileSystem.CopyDirectory(ruta2, ruta3, True)


            Dim rng As Microsoft.Office.Interop.Word.Range
            '   Dim MainDoc As Word.Document

            Documento_Inserta_1()
            Documento_Inserta_2()
            Documento_Inserta_3()
            Documento_Inserta_4()
            Documento_Inserta_5()

            Dim strFile As String
            Dim strFolder As String
            strFolder = ruta3

            MainDoc = obj_Word.Documents.Open(ruta2 & "\modelo.doc", Visible:=True)

            strFile = Dir$(strFolder & "\*.doc")

            Do Until strFile = ""
                rng = MainDoc.Range
                rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                rng.InsertFile(FileName:=(strFolder & "\" & strFile))
                strFile = Dir$()
            Loop

            'Guardamos el documento en la ruta Siniestros con el nombre ingresado en el textbox del archivo "Nº de Siniestro:"
            obj_Word.Documents(MainDoc).SaveAs2(ruta4 & "\BA" & TB_Nombre_Archivo.Text & ".doc")

            StatusEstado.Text = ("El siniestro se guardo como: BA" & TB_Nombre_Archivo.Text & ".doc")

            'borramos la carpeta Temp
            My.Computer.FileSystem.DeleteDirectory(ruta3, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

#End Region

#Region "Métodos Base de Datos"

    Private Sub Crear_base_de_datos()

        Dim micomandoInsertarSiniestro As SqlCommand
        Dim insertarIngreso As Integer

        'creamos el comando para Insertar el registro
        micomandoInsertarSiniestro = New SqlCommand("INSERT INTO tablaver (siniestros) VALUES ('" & TB_Nombre_Archivo.Text & "')", SqlConnection1)

        'construimos la conexión a la base de datos
        Dim constructor As New SqlConnectionStringBuilder()
        constructor.DataSource = ".\SQLEXPRESS"
        constructor.AttachDBFilename = Application.ExecutablePath.Substring(0, Application.ExecutablePath.LastIndexOf("\")) & "\aver.mdf"
        constructor.IntegratedSecurity = True
        constructor.ConnectTimeout = 30
        constructor.UserInstance = True
        SqlConnection1.ConnectionString = constructor.ConnectionString
        SqlConnection1.FireInfoMessageEventOnUserErrors = False



        Try
            'abrimos la conexión a la base de datos
            SqlConnection1.Open()

            'insertamos el registro
            insertarIngreso = micomandoInsertarSiniestro.ExecuteNonQuery
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally

            'cerramos la conexión a la base de datos
            SqlConnection1.Close()
        End Try



    End Sub

    Private Sub Comprobar_base_de_datos()

        Dim micomandoComprobar As SqlCommand
        Dim comprobar As Integer

        'comando para comprobar si existe el registro. Contamos los registros que coinciden con el nuestro, si no está el valor será 0, si está, 1 o más si se ha duplicado
        micomandoComprobar = New SqlCommand("SELECT COUNT (*) FROM tablaver WHERE siniestros = '" _
                                             & TB_Nombre_Archivo.Text & "'", SqlConnection1)

        Dim constructor As New SqlConnectionStringBuilder()
        constructor.DataSource = ".\SQLEXPRESS"
        constructor.AttachDBFilename = Application.ExecutablePath.Substring(0, Application.ExecutablePath.LastIndexOf("\")) & "\aver.mdf"
        constructor.IntegratedSecurity = True
        constructor.ConnectTimeout = 30
        constructor.UserInstance = True
        SqlConnection1.ConnectionString = constructor.ConnectionString
        SqlConnection1.FireInfoMessageEventOnUserErrors = False



        Try
            SqlConnection1.Open()

            'comprobamos si está
            comprobar = micomandoComprobar.ExecuteScalar

            If comprobar <> 0 Then
                Dim caption As String = "Atención"
                'preguntamos si quiere continuar, ya que se va a cerrar el expediente
                If (MessageBox.Show("El siniestro ya existe, ¿desea importarlo?", caption, MessageBoxButtons.YesNo) = DialogResult.Yes) Then

                    'aquí pego el código de Obtener los datos del Woooord!!!!
                    Documento_Obtiene_1()
                    StatusEstado.Text = ("Insertando el documento: BA" & TB_Nombre_Archivo.Text & ".doc")

                Else
                    SqlConnection1.Close()
                    Exit Sub
                End If

            Else
                Dim caption As String = "Atención"
                MessageBox.Show("El Siniestro NO EXISTE, si desea Generarlo ingrese los datos en el formulario y pulse el botón 'Generar Word'", caption)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        Finally

            'cerramos la conexión a la base de datos
            SqlConnection1.Close()
        End Try

    End Sub

#End Region

#Region "Métodos relacionados a los controles"

    Private Sub ArchivoGuardar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ArchivoGuardar.Click
        PreguntaSi_o_No()
    End Sub

    Private Sub ArchivoGuardarComo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ArchivoGuardarComo.Click
        Comprobar_base_de_datos()
    End Sub

    Private Sub ArchivoSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ArchivoSalir.Click
        Me.Close()
    End Sub


    Private Sub BT_Generar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Generar.Click
        PreguntaSi_o_No()
    End Sub

    Private Sub BT_Comprobar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Comprobar.Click

        Comprobar_base_de_datos()

    End Sub

    'cierro el formulario pulsando la X y cierro y mato procesos (aun no funciona)
    Private Sub Proyecto_German_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing

        If MsgBox("Estas a punto de salir  ¿deseas continuar?", MsgBoxStyle.YesNo, "Salir de la aplicación C.I.A.C.") = MsgBoxResult.No Then
            e.Cancel = True
            wd = Nothing
            Kill_Word()
        Else

            Kill_Word()
            MainDoc = Nothing
            obj_Word = Nothing
            obj_Doc = Nothing
            wd = Nothing

        End If
    End Sub

    Private Sub PreguntaSi_o_No()

        Try


            Dim ruta4 As String

            ruta4 = Application.ExecutablePath.Substring(0, Application.ExecutablePath.LastIndexOf("\")) & "\Siniestros"

            'Pregunto si el textbox Nombre de Archivo esta vacio, si lo esta, no se puede continuar.
            If (TB_Nombre_Archivo.Text = "") Then

                MsgBox("Error, tiene que ingresar al menos un número de siniestro", MsgBoxStyle.Critical)

            Else

                'Pregunto si existe el fichero en la carpeta siniestro
                If File.Exists(ruta4 & "\BA" & TB_Nombre_Archivo.Text & ".doc") Or TB_Nombre_Archivo.Text = "\BA.doc" Then

                    MsgBox("El número de siniestro ya existe, por favor ingrese un nuevo número de siniestro.", MsgBoxStyle.Information)
                    TB_Nombre_Archivo.Text = ""

                Else

                    ' MsgBox("El número de siniestro no existe", MsgBoxStyle.Information)

                    Crear_base_de_datos()
                    Generador()

                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    'Mato el proceso WINWORD.EXE
    Sub Kill_Word()

        Dim sKillWord As String

        sKillWord = "TASKKILL /F /IM WINWORD.EXE"

        Shell(sKillWord, vbHide)

    End Sub

#End Region

#Region "Métodos Variables"

    Private Sub TxBx_Nombre_Asegurado_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxBx_Nombre_Asegurado.Leave
        TB_Nombre_AseguradoII.Text = TxBx_Nombre_Asegurado.Text
        TB_Nombre_Aseg_Vehiculo.Text = TxBx_Nombre_Asegurado.Text
    End Sub

    Private Sub TB_DNI_Aseg_Vehiculo_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_DNI_Aseg_Vehiculo.Leave

        TB_LicConducir_Aseg_Vehiculo.Text = TB_DNI_Aseg_Vehiculo.Text

    End Sub

    Private Sub TB_DNI_Tercero_Conductor_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_DNI_Tercero_Conductor.Leave

        TB_LicConducir_Tercero_Conductor.Text = TB_DNI_Tercero_Conductor.Text

    End Sub

#End Region

#Region "Solo control de Números y Letras"

    'Solo se pueden digitar números del 0 al 9
    Function SoloNumeros(ByVal Keyascii As Short) As Short
        If InStr("1234567890", Chr(Keyascii)) = 0 Then
            SoloNumeros = 0
        Else
            SoloNumeros = Keyascii
        End If
        Select Case Keyascii
            Case 8
                SoloNumeros = Keyascii
            Case 13
                SoloNumeros = Keyascii
        End Select
    End Function

    'Solo se pueden digitar números del 0 al 9
    Private Sub num(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TB_Nombre_Archivo.KeyPress, TB_Telf_Aseg_Vehiculo.KeyPress, TB_Año_Vehiculo_Aseg.KeyPress, TB_Telf_Tercero_Conductor.KeyPress, TB_Edad_Tercero_Conductor.KeyPress, TB_Año_Vehiculo_Tercero.KeyPress, TB_Telf_Vehiculo_Tercero.KeyPress, TB_Edad_Victimas.KeyPress, TB_Edad_Aseg_Vehiculo.KeyPress, TB_Telf_Victimas.KeyPress, CB_Cantidad_Victimas.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
        KeyAscii = CShort(SoloNumeros(KeyAscii))
        If KeyAscii = 0 Then
            e.Handled = True
        End If

    End Sub

    'Solo se pueden digitar números del 0 al 9 y punto, dos puntos y barra lateral.
    Function SoloNumerosyPuntos(ByVal Keyascii As Short) As Short
        If InStr("1234567890.:/", Chr(Keyascii)) = 0 Then
            SoloNumerosyPuntos = 0
        Else
            SoloNumerosyPuntos = Keyascii
        End If
        Select Case Keyascii
            Case 8
                SoloNumerosyPuntos = Keyascii
            Case 13
                SoloNumerosyPuntos = Keyascii
        End Select
    End Function

    'Solo se pueden digitar números del 0 al 9 y punto, dos puntos y barra lateral.
    Private Sub numYpuntos(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TB_LicConducir_Aseg_Vehiculo.KeyPress, TB_DNI_Aseg_Vehiculo.KeyPress, TB_LicConducir_Tercero_Conductor.KeyPress, TB_DNI_Tercero_Conductor.KeyPress, TB_DNI_Victimas.KeyPress, TB_DNI_Vehiculo_Tercero.KeyPress, TB_Hora_Ocurrencia.KeyPress, TB_Hasta_Tercero_Conductor.KeyPress, TB_Hasta_Aseg_Vehiculo.KeyPress, TB_Desde_Tercero_Conductor.KeyPress, TB_Desde_Aseg_Vehiculo.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
        KeyAscii = CShort(SoloNumerosyPuntos(KeyAscii))
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    'Solo se pueden escribir letras
    Function SoloLETRAS(ByVal KeyAscii As Integer) As Integer
        'Transformar letras minusculas a Mayúsculas
        KeyAscii = Asc(UCase(Chr(KeyAscii)))

        ' Intercepta un código ASCII recibido admitiendo solamente letras, además:
        ' deja pasar sin afectar si recibe tecla de Backspace, Enter o Space.
        If InStr("ABCDEFGHIJKLMNÑOPQRSTUVWXYZ", Chr(KeyAscii)) = 0 Then
            SoloLETRAS = 0
        Else
            SoloLETRAS = KeyAscii
        End If
        ' teclas adicionales permitidas
        If KeyAscii = 8 Then SoloLETRAS = KeyAscii ' Backspace
        If KeyAscii = 13 Then SoloLETRAS = KeyAscii ' Enter
        If KeyAscii = 32 Then SoloLETRAS = KeyAscii ' Space

    End Function

    Private Sub Letras(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TB_LicConducir_Aseg_Vehiculo.KeyPress, TxBx_Nombre_Asegurado.KeyPress, TB_Relacion_Aseg_Vehiculo.KeyPress, TB_Nombre_AseguradoII.KeyPress, TB_Nombre_Aseg_Vehiculo.KeyPress, TB_Lugar_Ocurrencia.KeyPress, TB_Daños_Aseg_Vehiculo.KeyPress, CB_Uso_Vehiculo_Aseg.KeyPress, CB_Tipo_Vehiculo_Aseg.KeyPress, CB_Resp_Asegurado.KeyPress, CB_Ocupacion_Aseg_Vehiculo.KeyPress, CB_Marca_Vehiculo_Aseg.KeyPress, CB_Fraude.KeyPress, CB_Exclusiones.KeyPress, CB_EstCivil_Aseg_Vehiculo.KeyPress, CB_Abogado_Tercero.KeyPress, TB_Titular_Vehiculo_Tercero.KeyPress, TB_Relacion_Tercero_Conductor.KeyPress, TB_Nombre_Tercero_Conductor.KeyPress, TB_Domicilio_Tercero_Conductor.KeyPress, TB_Daños_Vehiculo_Tercero.KeyPress, CB_Uso_Vehiculo_Tercero.KeyPress, CB_Tipo_Vehiculo_Tercero.KeyPress, CB_Ocupacion_Tercero_Conductor.KeyPress, CB_Marca_Vehiculo_Tercero.KeyPress, CB_EstadoCivil_Tercero_Conductor.KeyPress, CB_Aseguradora_Vehiculo_Tercero.KeyPress, TB_Nombre_Victimas.KeyPress, TB_Lesiones_Victimas.KeyPress, TB_Inspector_Vehiculo_Tercero.KeyPress, TB_IncapacidadJudicial_Victimas.KeyPress, TB_IncapacidadFutura_Victimas.KeyPress, TB_Condicion_Victimas.KeyPress, CB_Ocupacion_Victimas.KeyPress, CB_EstadoCivil_Victimas.KeyPress, TB_SumarioNº_Detalles.KeyPress, TB_Observaciones.KeyPress, TB_Fiscalia_Detalles.KeyPress, TB_DepInterventora_Detalles.KeyPress, CB_Pasivo.KeyPress, CB_InfAmbiental_VII_Victimas.KeyPress, CB_InfAmbiental_VI_Victimas.KeyPress, CB_InfAmbiental_V_Victimas.KeyPress, CB_InfAmbiental_IV_Victimas.KeyPress, CB_InfAmbiental_III_Victimas.KeyPress, CB_InfAmbiental_II_Victimas.KeyPress, CB_InfAmbiental_I_Victimas.KeyPress, CB_Atribucion_Responsabilidad.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))
        KeyAscii = CShort(SoloLETRAS(KeyAscii))

        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub


#End Region

#Region "Método Fechas"


    Private Sub TB_Desde_Aseg_Vehiculo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Desde_Aseg_Vehiculo.TextChanged
        TB_Desde_Aseg_Vehiculo.MaxLength = 10
        If Len(TB_Desde_Aseg_Vehiculo.Text) = 2 Then
            TB_Desde_Aseg_Vehiculo.Text = TB_Desde_Aseg_Vehiculo.Text + "/"
            TB_Desde_Aseg_Vehiculo.Select(TB_Desde_Aseg_Vehiculo.Text.Length, 0)
        ElseIf Len(TB_Desde_Aseg_Vehiculo.Text) = 5 Then

            If TB_Desde_Aseg_Vehiculo.Text.Substring(startIndex:=3) = "02" Then
                If TB_Desde_Aseg_Vehiculo.Text.Substring(startIndex:=0, length:=2) > 28 Then
                    MessageBox.Show("Febrero tiene ¡hasta 28 días!")
                    TB_Desde_Aseg_Vehiculo.Text = String.Empty
                    Exit Sub
                End If
                TB_Desde_Aseg_Vehiculo.Text = TB_Desde_Aseg_Vehiculo.Text + "/"
                TB_Desde_Aseg_Vehiculo.Select(TB_Desde_Aseg_Vehiculo.Text.Length, 0)
            Else
                TB_Desde_Aseg_Vehiculo.Text = TB_Desde_Aseg_Vehiculo.Text + "/"
                TB_Desde_Aseg_Vehiculo.Select(TB_Desde_Aseg_Vehiculo.Text.Length, 0)
            End If
        End If
    End Sub

    Private Sub TB_Hasta_Aseg_Vehiculo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Hasta_Aseg_Vehiculo.TextChanged
        TB_Hasta_Aseg_Vehiculo.MaxLength = 10
        If Len(TB_Hasta_Aseg_Vehiculo.Text) = 2 Then
            TB_Hasta_Aseg_Vehiculo.Text = TB_Hasta_Aseg_Vehiculo.Text + "/"
            TB_Hasta_Aseg_Vehiculo.Select(TB_Hasta_Aseg_Vehiculo.Text.Length, 0)
        ElseIf Len(TB_Hasta_Aseg_Vehiculo.Text) = 5 Then
            If TB_Hasta_Aseg_Vehiculo.Text.Substring(startIndex:=3) = "02" Then

                If TB_Hasta_Aseg_Vehiculo.Text.Substring(startIndex:=0, length:=2) > 28 Then
                    MessageBox.Show("Febrero tiene ¡hasta 28 días!")
                    TB_Hasta_Aseg_Vehiculo.Text = String.Empty
                    Exit Sub
                End If
                TB_Hasta_Aseg_Vehiculo.Text = TB_Hasta_Aseg_Vehiculo.Text + "/"
                TB_Hasta_Aseg_Vehiculo.Select(TB_Hasta_Aseg_Vehiculo.Text.Length, 0)
            Else
                TB_Hasta_Aseg_Vehiculo.Text = TB_Hasta_Aseg_Vehiculo.Text + "/"
                TB_Hasta_Aseg_Vehiculo.Select(TB_Hasta_Aseg_Vehiculo.Text.Length, 0)
            End If
        End If
    End Sub

    Private Sub TB_Desde_Tercero_Conductor_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Desde_Tercero_Conductor.TextChanged
        TB_Desde_Tercero_Conductor.MaxLength = 10
        If Len(TB_Desde_Tercero_Conductor.Text) = 2 Then
            TB_Desde_Tercero_Conductor.Text = TB_Desde_Tercero_Conductor.Text + "/"
            TB_Desde_Tercero_Conductor.Select(TB_Desde_Tercero_Conductor.Text.Length, 0)

        ElseIf Len(TB_Desde_Tercero_Conductor.Text) = 5 Then
            If TB_Desde_Tercero_Conductor.Text.Substring(startIndex:=3) = "02" Then

                If TB_Desde_Tercero_Conductor.Text.Substring(startIndex:=0, length:=2) > 28 Then
                    MessageBox.Show("Febrero tiene ¡hasta 28 días!")
                    TB_Desde_Tercero_Conductor.Text = String.Empty
                    Exit Sub
                End If
                TB_Desde_Tercero_Conductor.Text = TB_Desde_Tercero_Conductor.Text + "/"
                TB_Desde_Tercero_Conductor.Select(TB_Desde_Tercero_Conductor.Text.Length, 0)
            Else
                TB_Desde_Tercero_Conductor.Text = TB_Desde_Tercero_Conductor.Text + "/"
                TB_Desde_Tercero_Conductor.Select(TB_Desde_Tercero_Conductor.Text.Length, 0)
            End If
        End If
    End Sub

    Private Sub TB_Hasta_Tercero_Conductor_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_Hasta_Tercero_Conductor.TextChanged
        TB_Hasta_Tercero_Conductor.MaxLength = 10
        If Len(TB_Hasta_Tercero_Conductor.Text) = 2 Then
            TB_Hasta_Tercero_Conductor.Text = TB_Hasta_Tercero_Conductor.Text + "/"
            TB_Hasta_Tercero_Conductor.Select(TB_Hasta_Tercero_Conductor.Text.Length, 0)

        ElseIf Len(TB_Hasta_Tercero_Conductor.Text) = 5 Then
            If TB_Hasta_Tercero_Conductor.Text.Substring(startIndex:=3) = "02" Then

                If TB_Hasta_Tercero_Conductor.Text.Substring(startIndex:=0, length:=2) > 28 Then
                    MessageBox.Show("Febrero tiene ¡hasta 28 días!")
                    TB_Hasta_Tercero_Conductor.Text = String.Empty
                    Exit Sub
                End If
                TB_Hasta_Tercero_Conductor.Text = TB_Hasta_Tercero_Conductor.Text + "/"
                TB_Hasta_Tercero_Conductor.Select(TB_Hasta_Tercero_Conductor.Text.Length, 0)
            Else
                TB_Hasta_Tercero_Conductor.Text = TB_Hasta_Tercero_Conductor.Text + "/"
                TB_Hasta_Tercero_Conductor.Select(TB_Hasta_Tercero_Conductor.Text.Length, 0)
            End If
        End If
    End Sub

#End Region

#Region "Método Impresión"

    Function Imprimir(ByVal Path As String, _
                      Optional ByVal Visible_Word As Boolean = True) As Boolean

        ' variable de objeto para acceder al Word  
        Dim Obj_Word As Object

        ' crea el objeto  
        Obj_Word = CreateObject("Word.Application")

        ' Visible / No visible  
        If Visible_Word Then
            Obj_Word.Visible = True
        Else
            Obj_Word.Visible = False
        End If

        'Abre el documento  
        Obj_Word.Documents.Open(Path)

        ' Imprime el documento activo con Printout  
        Obj_Word.ActiveDocument.Printout()

        ' Cierra el documento  
        Obj_Word.Quit()

        ' Elimina la referencia  
        Obj_Word = Nothing

        ' retorno  
        If Err.Number = 0 Then
            Imprimir = True
        End If



Error_Function:

        ' error  
        MsgBox(Err.Description)
        On Error Resume Next

        Obj_Word = Nothing
        Obj_Word.Quit()

    End Function



    Private Sub ImprimirToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImprimirToolStripMenuItem.Click

        Dim ret As Boolean
        Dim ruta4 As String


        ruta4 = Application.ExecutablePath.Substring(0, Application.ExecutablePath.LastIndexOf("\")) & "\Siniestros"

        ' le pasa el documento de word que se va a imprimir  
        ret = Imprimir(ruta4 & "\BA" & TB_Nombre_Archivo.Text & ".doc", False)

        If ret Then
            MsgBox("Ok", vbInformation)
        End If
    End Sub



#End Region

#Region "Acerca de..."


    Private Sub AyudaAcerca_de_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AyudaAcerca_de.Click
        Dim mensaje As String
        Dim NL As String = Environment.NewLine
        mensaje = "Proyecto C.I.A.C. - 1.0 BETA - 2012 " + NL
        mensaje += "Copyright (c) Germán Jongewaard de Boer, 2012"
        MessageBox.Show(mensaje, "Acerca de Proyecto Germán 2012", _
                     MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

#End Region




   
End Class




