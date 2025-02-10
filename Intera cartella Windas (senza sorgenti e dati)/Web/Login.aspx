<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Login.aspx.cs" Inherits="NewBfWeb.Login" %>
<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <script src="~/script/Jquery.min.js"></script>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <title>Login</title>
    <link rel="stylesheet" type="text/css" href="~/Content/bootstrap/bootstrap.css" />
    <link rel="stylesheet" type="text/css" href="~/Content/Site.css" />
</head>
<body class="text-center">
    <form id="form_login" method="post" class="form-signin">
        <div>
            <img class="mb-4" src="../Img/bf_logo_black.png" alt="Logo bf">
             <% NewBfWeb.classi.Languages lang = new NewBfWeb.classi.Languages(); %>
            <h2 class="h3 mb-3 font-weight-normal"><% Response.Write(lang["login"]); %></h2>
            <label for="username" class="sr-only"><% Response.Write(lang["username"]); %></label>
            <input type="text" id="username" name="username" placeholder="<% Response.Write(lang["placeholder_username"]); %>" class="form-control" required/>
            <label for="password" class="sr-only"><% Response.Write(lang["password"]); %></label>
            <input type="password" id="password" name="password" placeholder="<% Response.Write(lang["placeholder_password"]); %>" class="form-control" required/>
            <asp:Label ID="errore" CssClass="errore" runat="server" Visible="false"/>
            <input type="submit" class="btn btn-lg btn-primary btn-block" value="<% Response.Write(lang["login"]); %>"/>
        </div>
    </form>
</body>
</html>
