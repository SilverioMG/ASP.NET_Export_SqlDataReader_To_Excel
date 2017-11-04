Protected Sub ExportarToExcel()
    Dim idCliente As Int16 = ComboProyectos.SelectedValue
    Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("EUROPHONESTKConnectionString").ToString())
    Dim cmd As New SqlCommand()
    Dim reader As SqlDataReader
    Dim strFileName As String = ComboProyectos.SelectedItem.ToString()
 
 
    Dim htmlContent As New StringBuilder()
    Dim strContent As String
 
    Try
        conn.Open()
        cmd.Connection = conn
 
        cmd.CommandText = "SPLstControlStock_Proyecto"
        cmd.CommandType = CommandType.StoredProcedure
 
        SqlCommandBuilder.DeriveParameters(cmd)
        cmd.Parameters("@IdProyecto").Value = idCliente
 
        reader = cmd.ExecuteReader()
 
        If reader.HasRows Then
            htmlContent.Append("<TABLE>")
            htmlContent.Append("<TR>")
 
            'Nombres de las columnas o campos:
            For i As Int32 = 0 To reader.FieldCount - 1
                htmlContent.Append("<TD>" + reader.GetName(i) + "</TD>")
            Next
 
            htmlContent.Append("</TR>")
 
            'Filas de la consulta:
            While reader.Read()
                htmlContent.Append("<TR>")
 
                'Valores de las columnas o campos:
                For i As Int32 = 0 To reader.FieldCount - 1
                    htmlContent.Append("<TD>" + reader.GetValue(i).ToString() + "</TD>")
                Next
 
                htmlContent.Append("</TR>")
            End While
 
            htmlContent.Append("</TABLE>")
        End If
 
        conn.Close()
 
        strContent = htmlContent.ToString()
 
        'Descarga del Excel:
        Response.Clear()
        Response.Charset = ""
        Response.AddHeader("content-disposition", "attachment;" + "filename=" + strFileName + ".xls")
        Response.ContentType = "application/vnd.ms-excel"
 
        Response.Write(strContent)
        HttpContext.Current.Response.Flush()
        Response.End()
    Catch ex As Exception
 
    End Try
    
    End Sub
