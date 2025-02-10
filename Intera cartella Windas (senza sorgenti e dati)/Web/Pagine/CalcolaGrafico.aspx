<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="CalcolaGrafico.aspx.cs" Inherits="NewBfWeb.Pagine.CalcolaGrafico" %>
<%
    Response.ContentType = "application/json";
    Response.Write(this.Result);
%>
