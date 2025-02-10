<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="DiarioGrafico_old.aspx.cs" Inherits="NewBfWeb.Pagine.DiarioGrafico_old" %>

<%@ Register Assembly="System.Web.DataVisualization, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" Namespace="System.Web.UI.DataVisualization.Charting" TagPrefix="asp" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <main role="main" class="pl-4 col-md-9 ml-sm-auto col-lg-11">
        <div class="row mt-4 align-items-center">
            <% NewBfWeb.classi.Languages lang = new NewBfWeb.classi.Languages(); %>
            <div class="col-0">
                <div class="input-group ml-1">
                    <div class="input-group-prepend">
                        <div class="input-group-text"><% Response.Write(lang["stazioni"]); %>:</div>
                    </div>
                    <asp:DropDownList CssClass="form-control" AutoPostBack="true" class="custom-select" name="stazioni" ID="stazioni" DataValueField="Code" DataTextField="Description" runat="server"></asp:DropDownList>
                </div>
            </div>
            <div class="col-0 ml-1">
                <div class="input-group ml-1">
                    <div class="input-group-prepend">
                        <div class="input-group-text"><% Response.Write(lang["tipo_dati"]); %>:</div>
                    </div>
                    <asp:DropDownList CssClass="form-control" AutoPostBack="true" class="custom-select" name="tipo" ID="tipo" runat="server"></asp:DropDownList>
                </div>
            </div>
            <div class="col-0 ml-1">
                <div class="input-group ml-1">
                    <div class="input-group-prepend">
                        <div class="input-group-text"><% Response.Write(lang["data_inizio"]); %>:</div>
                    </div>
                    <asp:TextBox CssClass="form-control" ID="inizio" AutoPostBack="true" name="inizio" runat="server" TextMode="Date" />
                </div>
            </div>
            <div class="col-0 ml-1">
                <div class="input-group ml-1">
                    <div class="input-group-prepend">
                        <div class="input-group-text"><% Response.Write(lang["data_fine"]); %>:</div>
                    </div>
                    <asp:TextBox CssClass="form-control" ID="fine" AutoPostBack="true" name="inizio" runat="server" TextMode="Date" />
                </div>
            </div>
            <div class="col-0 ml-2">
                <asp:Button CssClass="btn btn-primary" runat="server" ID="carica" OnClick="carica_Click"/>
            </div>
            <div class="col-0 ml-1">
                <a href="#" class="btn btn-primary" onclick="ResetFiltro()"><% Response.Write(lang["resetta_filtro"]); %></a>
            </div>
        </div>
        <asp:Label ID="errore" runat="server" Visible="false"/>
        <div id="div_filtro" class="row mt-1">
            <asp:Chart ID="Chart1" runat="server" ImageStorageMode="UseImageLocation" Width="1680px" BorderSkin-BorderColor="#333333" BorderSkin-BorderDashStyle="Solid" Height="600px" IsMapAreaAttributesEncoded="False" BackGradientStyle="None" BorderlineColor="#666666" BorderlineDashStyle="Solid">
                <Titles>
                    
                </Titles>
                <Series>
                </Series>
                <ChartAreas>
                    <asp:ChartArea Name="area">
                        <Area3DStyle LightStyle="Simplistic" Enable3D="false" />
                        <AxisX IsLabelAutoFit="true" LineColor="Red" LineWidth="1" IntervalType="Hours" Interval="1">
                            <LabelStyle Font="Arial, 12px" Format="dd/MM - HH"/>
                            <MajorTickMark LineColor="Red" LineWidth="3"/>
                            <MajorGrid LineColor="#c0c0c0" />
                        </AxisX>
                        <AxisY IsLabelAutoFit="true" Maximum="100" Minimum="0" LineColor="Red" LineWidth="1" Interval="10">
                            <LabelStyle Font="Arial, 16px, style=Bold" />
                            <MajorTickMark LineColor="Red" LineWidth="3"/>
                            <MajorGrid LineColor="#c0c0c0" />
                        </AxisY>
                    </asp:ChartArea>
                </ChartAreas>
                <Legends>
                    <asp:Legend Name="legenda" />
                </Legends>
            </asp:Chart>
        </div>
        <%--<script src="../Scripts/grafico.js"></script>--%>
    </main>
</asp:Content>
