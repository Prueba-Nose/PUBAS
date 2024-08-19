<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/conecta1.asp" -->
<!--#include file="config.asp"-->
<!--#include file="checklogin.asp"-->
<!--#include file="stylo2.asp" -->
<%
if Request.Form("Submit") <> "" then
    response.Cookies("configsol")("planta") = Request.Form("proyecto")
end if
%>
<%
dim proveedor'Vairable para guardar el proveedor
dim descripcion
dim desc
dim estatus
dim st
dim usuaruio
dim usu
dim id
id = ""
usuario = ""
descripcion = ""

Dim RSusuario__MMColParam
RSusuario__MMColParam = "1"
If (Session("site_nombre") <> "") Then 
  RSusuario__MMColParam = Session("site_nombre")
End If

Dim RSusuario
Dim RSusuario_cmd
Dim RSusuario_numRows

Set RSusuario_cmd = Server.CreateObject ("ADODB.Command")
RSusuario_cmd.ActiveConnection = MM_conecta1_STRING
RSusuario_cmd.CommandText = "SELECT * FROM ksroc.usuarios WHERE Nombre = ?" 
RSusuario_cmd.Prepared = true
RSusuario_cmd.Parameters.Append RSusuario_cmd.CreateParameter("param1", 200, 1, 50, RSusuario__MMColParam) ' adVarChar

Set RSusuario = RSusuario_cmd.Execute
RSusuario_numRows = 0

if (RSusuario.Fields.Item("Tipo").Value = "Administrador") then
    'if para validar que se filtra el proyecto
    If request.Form("proyecto") <> "" Then'proyectooc
        If request.Form("proyecto") <> "0" Then
            proyectooc= " AND proyectore = '" & Request.Form("proyecto") & "'"
        Else
            proyectooc= ""
        End If
    elseif Request.Cookies("configsol")("planta") = "0" OR Request.Cookies("configsol")("planta") = "" then
        proyectooc = ""
    Else
        proyectooc= " AND proyectore = " & Request.Cookies("configsol")("planta") & ""
    END if'proyectooc
Else
    proyectooc = " AND (proyectore = "&Request.Cookies("ksroc")("sucursal_id")&") "
end if


'if para validar que si es administrador ve las pendientes por default
if Session("site_tipo") = "Administrador" then'site_tipo
estatus = " AND status = 'Aprobacion' "
    If request.Form("estatusddl") <> "" Then'estatus
        If request.Form("estatusddl") <> "0" Then
            estatus= " AND status = '" & Request.Form("estatusddl") & "'"
        Else
            estatus= ""
        End If
   End If
end if'site_tipo

    

if Session("site_tipo") <> "Administrador" then
id = " AND idusure = " & Session("site_id")
end if

if Request.Form("usuario") <> "" then
usuario = " AND idusure = "&Request.Form("usuario")
usu = Request.Form("usuario")
end if



    

if Request.Form("descripcion") <> "" then
descripcion = " AND motivore LIKE '%"&Request.Form("descripcion")&"%'"
desc = Request.Form("descripcion")
end if

%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_conecta1_STRING
Recordset1_cmd.CommandText = "SELECT * FROM ksroc.solicitud WHERE idre > 0"&descripcion&estatus&id&usuario&proyectooc&"  ORDER BY idre DESC"
response.Write(Recordset1_cmd.CommandText)
Recordset1_cmd.Prepared = true
Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Recordset2
Dim Recordset2_cmd
Dim Recordset2_numRows

Set Recordset2_cmd = Server.CreateObject ("ADODB.Command")
Recordset2_cmd.ActiveConnection = MM_conecta1_STRING
Recordset2_cmd.CommandText = "SELECT * FROM ksroc.usuarios ORDER BY Nombre ASC" 
Recordset2_cmd.Prepared = true

Set Recordset2 = Recordset2_cmd.Execute
Recordset2_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 15
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim Recordset1_total
Dim Recordset1_first
Dim Recordset1_last

' set the record count
Recordset1_total = Recordset1.RecordCount

' set the number of rows displayed on this page
If (Recordset1_numRows < 0) Then
  Recordset1_numRows = Recordset1_total
Elseif (Recordset1_numRows = 0) Then
  Recordset1_numRows = 1
End If

' set the first and last displayed record
Recordset1_first = 1
Recordset1_last  = Recordset1_first + Recordset1_numRows - 1

' if we have the correct record count, check the other stats
If (Recordset1_total <> -1) Then
  If (Recordset1_first > Recordset1_total) Then
    Recordset1_first = Recordset1_total
  End If
  If (Recordset1_last > Recordset1_total) Then
    Recordset1_last = Recordset1_total
  End If
  If (Recordset1_numRows > Recordset1_total) Then
    Recordset1_numRows = Recordset1_total
  End If
End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = Recordset1
MM_rsCount   = Recordset1_total
MM_size      = Recordset1_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
Recordset1_first = MM_offset + 1
Recordset1_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (Recordset1_first > MM_rsCount) Then
    Recordset1_first = MM_rsCount
  End If
  If (Recordset1_last > MM_rsCount) Then
    Recordset1_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = Server.HTMLEncode(MM_keepMove) & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<%
dim color'Variable para guardar el color de los renglones

'Inicializar variables
color = cgrid2
'////////////////////////////////////////////
%>
<% 
REM Dim Recordset1
REM Dim Recordset1_cmd
REM Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_conecta1_STRING
Recordset1_cmd.CommandText = "SELECT s.*, p.proyecto FROM ksroc.solicitud s INNER JOIN ksroc.proyecto p ON s.Proyectore = p.idpy WHERE idre > 0 "&proyectooc&estatus&"  ORDER BY idre DESC"
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0

REM Dim Repeat1__numRows
REM Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<%
Dim RSProyecto
Dim RSProyecto_cmd
Dim RSProyecto_numRows

Set RSProyecto_cmd = Server.CreateObject ("ADODB.Command")
RSProyecto_cmd.ActiveConnection = MM_conecta1_STRING
RSProyecto_cmd.CommandText = "SELECT * FROM ksroc.proyecto ORDER BY proyecto ASC" 
RSProyecto_cmd.Prepared = true

Set RSProyecto = RSProyecto_cmd.Execute
RSProyecto_numRows = 0
%>
<!DOCTYPE html>

<html lang="en"><!-- InstanceBegin template="/Templates/plantillaksroc.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, shrink-to-fit=no">
    <!-- InstanceBeginEditable name="doctitle" -->
    <title>Arcanet-Compras</title>
    <!-- InstanceEndEditable -->
    <link rel="icon" type="image/x-icon" href="Imagenes/favicon.ico"/>
    <!-- BEGIN GLOBAL MANDATORY STYLES -->
    <link href="https://fonts.googleapis.com/css?family=Quicksand:400,500,600,700&display=swap" rel="stylesheet">
    <link href="Plantilla/bootstrap/css/bootstrap.min.css" rel="stylesheet" type="text/css" />
    <link href="Plantilla/assets/css/plugins.css" rel="stylesheet" type="text/css" />
    <link href="Plantilla/assets/css/structure.css" rel="stylesheet" type="text/css" class="structure" />
    <!-- END GLOBAL MANDATORY STYLES -->
    <!-- InstanceBeginEditable name="CssInicio" -->
    <link rel="stylesheet" type="text/css" href="Plantilla/plugins/table/datatable/datatables.css">
    <link rel="stylesheet" type="text/css" href="Plantilla/plugins/table/datatable/dt-global_style.css">   
    <link href="Plantilla/plugins/bootstrap-select/bootstrap-select.min.css" rel="stylesheet" type="text/css">
    <!-- InstanceEndEditable -->
    <!-- BEGIN PAGE LEVEL PLUGINS/CUSTOM STYLES -->
    <style>
        /*
            The below code is for DEMO purpose --- Use it if you are using this demo otherwise, Remove it
        */
        .navbar .navbar-item.navbar-dropdown {
            margin-left: auto;
        }
        .layout-px-spacing {
            min-height: calc(100vh - 145px)!important;
        }

    </style>
    <style>
        #busqueda {
            display: none;
        }

        #buscar {
            border-top-style: solid;
            border-top-color: #D6D5D5;
        }

            #buscar:hover {
                cursor: pointer;
            }
    </style>
    <script>

    function mostrar(){
        $("#busqueda").toggle("slow", function(){
            if($("#busqueda").css('display') == 'block'){
                $("#buscador").html('Filtrar<svg viewBox="0 0 24 24" width="24" height="24" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round" class="css-i6dzq1"><polyline points="18 15 12 9 6 15"></polyline></svg>');
            }
            else{
                $("#buscador").html('Filtrar<svg viewBox="0 0 24 24" width="24" height="24" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round" class="css-i6dzq1"><polyline points="6 9 12 15 18 9"></polyline></svg>');
            }
        });
    }
</script>
    <!-- END PAGE LEVEL PLUGINS/CUSTOM STYLES -->
</head>

<body class="sidebar-noneoverflow starterkit">
    <!--  BEGIN NAVBAR  -->
   <div class="header-container fixed-top">
        <header class="header navbar navbar-expand-sm">
            <ul class="navbar-item flex-row">
                <li class="nav-item theme-logo">
                  <a href="Default.asp">
                        <img src="Imagenes/mro.svg" class="navbar-logo" alt="logo">
                    </a>
                </li>
            </ul>

            <a href="javascript:void(0);" class="sidebarCollapse" data-placement="bottom"><svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="feather feather-menu"><line x1="3" y1="12" x2="21" y2="12"></line><line x1="3" y1="6" x2="21" y2="6"></line><line x1="3" y1="18" x2="21" y2="18"></line></svg></a>

            <div style="text-align: right; margin-left: auto; font-size: 16px; font-weight: 600; text-transform: uppercase;">
                <%= Session("site_sucursal") %>
            </div>
            <ul class="navbar-item flex-row navbar-dropdown">
                <li class="nav-item dropdown user-profile-dropdown  order-lg-0 order-1">
                    <a href="javascript:void(0);" class="nav-link dropdown-toggle user" id="userProfileDropdown" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
                        <svg viewBox="0 0 24 24" width="24" height="24" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round" class="css-i6dzq1"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path><circle cx="12" cy="7" r="4"></circle></svg>
                    </a>
                    <div class="dropdown-menu position-absolute animated fadeInUp" aria-labelledby="userProfileDropdown">                      
                        <div class="user-profile-section">
                            <div class="media mx-auto">
                                <svg viewBox="0 0 24 24" width="24" height="24" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round" class="css-i6dzq1"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path><circle cx="12" cy="7" r="4"></circle></svg>
                                <div class="media-body">
                                    <h5><%=Session("site_nombre")%></h5>
                                    <p><%=Session("site_area")%></p>
                                </div>
                            </div>
                        </div>
                        <div class="dropdown-item">
                            <a href="user_profile.html">
                                <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="feather feather-user"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path><circle cx="12" cy="7" r="4"></circle></svg> <span>Mi Perfil</span>
                            </a>
                        </div>
                        <div class="dropdown-item">
                            <a href="logout.asp">
                                <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="feather feather-log-out"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"></path><polyline points="16 17 21 12 16 7"></polyline><line x1="21" y1="12" x2="9" y2="12"></line></svg> <span>Log Out</span>
                            </a>
                        </div>
                    </div>
                </li>
            </ul>
        </header>
    </div>
    <!--  END NAVBAR  -->
    <!--  BEGIN MAIN CONTAINER  -->
    <div class="main-container" id="container">
        <div class="overlay"></div>
        <div class="search-overlay"></div>
        <!--  BEGIN SIDEBAR  -->
        <div class="sidebar-wrapper sidebar-theme">            
            <nav id="compactSidebar">
                <ul class="menu-categories">
                    <% 
                    set menu=createobject("ADODB.Recordset")
                    sqltxt= "SELECT * FROM ksroc.Menu ORDER BY Numero ASC"
                    menu.CursorType=1
                    menu.open sqltxt,strCon 
                    
                    While(Not menu.EOF)
                        If  Session("site_tipo") = "Almacen"  Then
                            if menu("tipo") = "Almacen" then 
					   			Response.Write(menu("Imagen1"))
					   		end if 
                        Else
                            If menu("Nivel") = "100" And Session("site_tipo") <> "Administrador" Then
		                       menu.MoveNext()
	                        Else
                                If menu("Nivel") <> 50 then
                                    Response.Write(menu("Imagen1"))
                                end if
                            End If
                        End If
                       menu.MoveNext()
                    Wend
                    %>
                </ul>
            </nav>            
            <div id="compact_submenuSidebar" class="submenu-sidebar">
                <div class="submenu" id="Catalogos">
                   <ul class="submenu-list" data-parent-element="#Catalogos">
                        <%
                        menu.MoveFirst()
       
                        Do While Not menu.EOF
                        
                            If menu("Nivel") = 50 Then
                                Response.Write(menu("Imagen1"))
                            End If
                         
                            menu.MoveNext()
                        Loop
                        %>
                    </ul>
                </div>
            </div>
        </div>
        <!--  END SIDEBAR  -->        
        <!--  BEGIN CONTENT AREA  -->
        <div id="content" class="main-content">
            <div class="layout-px-spacing">
                <!-- CONTENT AREA -->
                <!-- InstanceBeginEditable name="EditRegion1" -->
                <!--Titulo-->
                <div class="page-header">
                    <div class="page-title">
                        <h3>Solicitud de Cotizacion</h3>
                    </div>
                </div>
                <!--Direccion-->
                <nav class="breadcrumb-one" aria-label="breadcrumb">
                    <ol class="breadcrumb">
                        <li class="breadcrumb-item"><a href="javascript:void(0);"></a></li>
                        <li class="breadcrumb-item"><a href="javascript:void(0);"><svg viewBox="0 0 24 24" width="24" height="24" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round" class="css-i6dzq1"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"></path><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"></path></svg> Cat&aacute;logos</a></li>
                        <li class="breadcrumb-item active" aria-current="page"><svg viewBox="0 0 24 24" width="24" height="24" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round" class="css-i6dzq1"><circle cx="18" cy="18" r="3"></circle><circle cx="6" cy="6" r="3"></circle><path d="M13 6h3a2 2 0 0 1 2 2v7"></path><line x1="6" y1="9" x2="6" y2="21"></line></svg> Rutas</li>
                    </ol>
                </nav>
                <!--Tabla-->
                <div class="row layout-top-spacing" id="cancel-row">
                    <div class="col-xl-12 col-lg-12 col-sm-12 layout-spacing">
                        <div class="widget-content widget-content-area br-6">
                            <div class="widget-header">
                                <div class="row">
                                    <div class="col-xl-12 col-md-12 col-sm-12 col-12 text-right">
                                        <a href="solicitudAdd.asp" class="btn btn-primary">
                                            <svg viewBox="0 0 24 24" width="24" height="24" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round" class="css-i6dzq1"><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg>&nbsp;Agregar
                                        </a>
                                    </div>
                                </div>
                                <br>
                                <% If (RSusuario.Fields.Item("Tipo").Value = "Administrador") Then%>
                                <div id="buscar" onclick="mostrar();">
                                    <span id="buscador">Filtrar<svg viewBox="0 0 24 24" width="24" height="24" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round" class="css-i6dzq1"><polyline points="6 9 12 15 18 9"></polyline></svg></span>
                                </div>
                                <div class="row" id="busqueda" style="display: none">
                                    <div class="col-xl-12 col-md-12 col-sm-12 col-12">
                                        <form action="solicitud.asp" method="post" class="email">
                                            <div class="form-row mb-6">
                                                <div class="form-group col-md-3">
                                                    <label for="idre">Planta</label>
                                                    <select name="proyecto" id="proyecto" class="selectpicker mb-4" data-width="100%">
                                                        <option value="0" <%If (Not isNull(Request.Cookies("configsol")("planta"))) Then If (0 = CInt(Request.Cookies("configsol")("planta"))) Then Response.Write("selected=""selected""") : Response.Write("")%>>Todos</option>
                                                        <%While (NOT RSProyecto.EOF)%>
                                                        <option value="<%=(RSProyecto.Fields.Item("idpy").Value)%>" <%If (Not isNull(Request.Cookies("configsol")("planta"))) Then If (CInt(RSProyecto.Fields.Item("idpy").Value) = CInt(Request.Cookies("configsol")("planta"))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(RSProyecto.Fields.Item("proyecto").Value)%></option>
                                                        <%  RSProyecto.MoveNext()
                                                        Wend%>
                                                    </select>
                                                </div>
                                                <div class="form-group col-md-3">
                                                    <label for="status">estatus</label>
                                                    <select name="estatusddl" id="estatusddl" class="selectpicker mb-4" data-width="100%">
                                                        <option value="0">Todos</option>
                                                        <option value="Pendiente">Pendiente</option>
                                                        <option value="Aceptado">Aceptado</option>
                                                        <option value="Aprobacion">Aprobacion</option>
                                                        <option value="Terminado">Terminado</option>
                                                        <option value="Rechazado">Rechazado</option>
                                                    </select>
                                                </div>
                                                <div class="form-group col-md-6"></div>
                                                <div class="form-group col-md-3 text-left">
                                                    <input type="submit" name="Submit" class="btn btn-success mt-3" value="Enviar" />
                                                </div>
                                            </div>
                                        </form>
                                    </div>
                                </div>
                                <% End If %>
                            </div>
                            <div class="table-responsive mb-4 mt-4">
                                <table id="multi-column-ordering" class="table table-hover" style="width:100%">
                                    <thead>
                                        <tr>
                                            <th><strong>No</strong></th>
                                            <th><strong>Motivo</strong></th>
                                            <th><strong>Solicito</strong></th>
                                            <th><strong>Proveedor</strong></th>
                                            <th><strong>Departamento</strong></th>
                                            <th><strong>Planta</strong></th>
                                            <th><strong>Moneda</strong></th>
                                            <th><strong>Requerido</strong></th>
                                            <th><strong>Generada</strong></th>
                                            <th><strong>Estatus</strong></th>
                                            <th><strong>Observaciones</strong></th>
                                            <th><strong>Mod</strong></th>
                                            <th><strong>Imprimir</strong></th>
                                            <th><strong>Requisicion</strong></th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <% While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) %>
                                        <tr>
                                            <td><%=Recordset1.Fields.Item("idre").Value%></td>
                                            <td><%=Recordset1.Fields.Item("motivore").Value%></td>
                                            <td><%=(Recordset1.Fields.Item("usuariore").Value)%></td>
                                            <td><%=(Recordset1.Fields.Item("proveere").Value)%></td>
                                            <td><%=(Recordset1.Fields.Item("Departamentore").Value)%></td>
                                            <td><%=(Recordset1.Fields.Item("Proyecto").Value)%></td>
                                            <td><%=(Recordset1.Fields.Item("moneda").Value)%></td>
                                            <td><%=(Recordset1.Fields.Item("Freqre").Value)%></td>
                                            <td><%=(Recordset1.Fields.Item("Fechare").Value)%></td>
                                            <td><%=(Recordset1.Fields.Item("status").Value)%></td>
                                            <td><%=(Recordset1.Fields.Item("observacionesreq").Value)%></td>
                                            <td><a href="detsolicitud.asp?nr=<%=(Recordset1.Fields.Item("idre").Value)%>"><svg viewBox="0 0 24 24" width="24" height="24" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round" class="css-i6dzq1"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg></a></td>
                                            <td><a href="impSolicitud.asp?id=<%=(Recordset1.Fields.Item("idre").Value)%>" target="_blank"><svg viewBox="0 0 24 24" width="24" height="24" stroke="currentColor" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round" class="css-i6dzq1"><polyline points="9 11 12 14 22 4"></polyline><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"></path></svg></a></td>
                                            <td><% If (Recordset1.Fields.Item("status").Value = "Aceptado") Then %><a href="requiAutoFiltro.aspx?idre=<%=(Recordset1.Fields.Item("idre").Value)%>"><img src="imagenes/Arrow-left.png" width="16" height="16" border="0" /></a><% End If %></td>
                                        </tr>
                                            <% 
                                              Repeat1__index=Repeat1__index+1
                                              Repeat1__numRows=Repeat1__numRows-1
                                              Recordset1.MoveNext()
                                            Wend
                                            %>
                                    </tbody>
                                    <tfoot>
                                        <tr>
                                            <th><strong>No</strong></th>
                                            <th><strong>Motivo</strong></th>
                                            <th><strong>Solicito</strong></th>
                                            <th><strong>Proveedor</strong></th>
                                            <th><strong>Departamento</strong></th>
                                            <th><strong>Planta</strong></th>
                                            <th><strong>Moneda</strong></th>
                                            <th><strong>Requerido</strong></th>
                                            <th><strong>Generada</strong></th>
                                            <th><strong>Estatus</strong></th>
                                            <th><strong>Observaciones</strong></th>
                                            <th><strong>Mod</strong></th>
                                            <th><strong>Imprimir</strong></th>
                                            <th><strong>Requisicion</strong></th>
                                        </tr>
                                    </tfoot>
                                </table>
                            </div>
                        </div>
                    </div>
                </div>
                <!-- InstanceEndEditable -->                
                <!-- CONTENT AREA -->
            </div>
            <div class="footer-wrapper">
                <div class="footer-section f-section-1">
                    <p class="">Copyright © 2020 <a target="_blank" href="https://designreset.com">DesignReset</a>, All rights reserved.</p>
                </div>
                <div class="footer-section f-section-2">
                    <p class="">Coded with <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="feather feather-heart"><path d="M20.84 4.61a5.5 5.5 0 0 0-7.78 0L12 5.67l-1.06-1.06a5.5 5.5 0 0 0-7.78 7.78l1.06 1.06L12 21.23l7.78-7.78 1.06-1.06a5.5 5.5 0 0 0 0-7.78z"></path></svg></p>
                </div>
            </div>
        </div>
        <!--  END CONTENT AREA  -->
    </div>
    <!-- END MAIN CONTAINER -->
    <!-- BEGIN GLOBAL MANDATORY SCRIPTS -->
    <script src="Plantilla/assets/js/libs/jquery-3.1.1.min.js"></script>
    <script src="Plantilla/bootstrap/js/popper.min.js"></script>
    <script src="Plantilla/bootstrap/js/bootstrap.min.js"></script>
    <script src="Plantilla/plugins/perfect-scrollbar/perfect-scrollbar.min.js"></script>
    <script src="Plantilla/assets/js/app.js"></script>    
    <script>
            $(document).ready(function() {
                App.init();
            });
    </script>
    <script src="Plantilla/assets/js/custom.js"></script>
    <!-- END GLOBAL MANDATORY SCRIPTS -->
    <!-- BEGIN PAGE LEVEL PLUGINS/CUSTOM SCRIPTS -->
    <!-- InstanceBeginEditable name="JavaScriptFin" -->
    <script src="Plantilla/plugins/bootstrap-select/bootstrap-select.min.js"></script>
    <script src="Plantilla/plugins/table/datatable/datatables.js"></script>
    <script>
        $('#multi-column-ordering').DataTable({
            "oLanguage": {
                "oPaginate": { "sPrevious": '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="feather feather-arrow-left"><line x1="19" y1="12" x2="5" y2="12"></line><polyline points="12 19 5 12 12 5"></polyline></svg>', "sNext": '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="feather feather-arrow-right"><line x1="5" y1="12" x2="19" y2="12"></line><polyline points="12 5 19 12 12 19"></polyline></svg>' },
                "sInfo": "Mostrando Página _PAGE_ de _PAGES_",
                "sSearch": '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="feather feather-search"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>',
                "sSearchPlaceholder": "Search...",
               "sLengthMenu": "Results :  _MENU_",
            },
            "aaSorting": [0,'desc'],
            "stripeClasses": [],
            "lengthMenu": [10, 20, 50],
            "pageLength": 10,
	        columnDefs: [ { targets: [ 0 ], orderData: [ 0, 1 ] }, 
                          { targets: [ 1 ], orderData: [ 1, 0 ] }
                        ]
	    });
    </script>
    <!-- InstanceEndEditable -->     
    <!-- BEGIN PAGE LEVEL PLUGINS/CUSTOM SCRIPTS -->


</body>
<!-- InstanceEnd --></html>