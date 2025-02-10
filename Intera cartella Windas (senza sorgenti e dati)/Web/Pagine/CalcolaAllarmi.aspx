<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="CalcolaAllarmi.aspx.cs" Inherits="NewBfWeb.Pagine.CalcolaAllarmi" %>
<%
    Response.ContentType = "application/json";
    Response.Write(this.Result);
%>
