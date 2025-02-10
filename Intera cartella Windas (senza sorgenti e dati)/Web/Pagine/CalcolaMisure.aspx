<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="CalcolaMisure.aspx.cs" Inherits="NewBfWeb.Pagine.CalcolaMisure" %>
<%
    Response.ContentType = "application/json";
    Response.Write(this.Result);
%>
