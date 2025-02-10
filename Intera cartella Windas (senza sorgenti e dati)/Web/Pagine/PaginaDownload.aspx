<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PaginaDownload.aspx.cs" Inherits="NewBfWeb.Pagine.PaginaDownload" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <%
        try
        {
             Response.Clear();
             Response.ClearHeaders();
             Response.ClearContent();
             Response.AddHeader("Content-Disposition", "attachment; filename=" + System.IO.Path.GetFileName(this.Request.Params["file_name_pulito"]));
             Response.AddHeader("Content-Length", new System.IO.FileInfo(this.Request.Params["file_name"]).Length.ToString());
             Response.ContentType = "text/plain";
             Response.Flush();
             Response.TransmitFile(this.Request.Params["file_name"]);
             Response.End();
            //Response.ContentType = "APPLICATION/OCTET-STREAM";
            //this.Response.AppendHeader("Content-Disposition", "Attachment; Filename=\"" + System.IO.Path.GetFileName(this.Request.Params["file_name_pulito"]) + "\"");

            //// transfer the file byte-by-byte to the response object
            //System.IO.FileInfo fileToDownload = new System.IO.FileInfo(this.Request.Params["file_name"]);
            //Response.Flush();
            //Response.WriteFile(fileToDownload.FullName);
        }
        catch (System.Exception e)
        // file IO errors
        {
            //SupportClass.WriteStackTrace(e, Console.Error);
        }
%>
</body>
</html>
