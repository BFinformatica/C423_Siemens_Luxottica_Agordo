<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="CalcolaTag.aspx.cs" Inherits="NewBfWeb.Pagine.CalcolaTag" %>
<%
    Response.ContentType = "application/json";
    Response.Write(this.Result);
%>
