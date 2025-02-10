<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="AllarmiStati.aspx.cs" Inherits="NewBfWeb.Pagine.AllarmiStati" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <main role="main" class="pl-4 col-md-9 ml-sm-auto col-lg-11">
        <% NewBfWeb.classi.Languages lang = new NewBfWeb.classi.Languages(); %>
        <div class="row mt-4 align-items-center">
            <div class="col-0">
                <div class="input-group ml-1">
                    <div class="input-group-prepend">
                        <div class="input-group-text"><% Response.Write(lang["stazioni"]); %>:</div>
                    </div>
                    <asp:DropDownList CssClass="form-control" AutoPostBack="true" class="custom-select" name="stazioni" ID="stazioni" DataValueField="Code" DataTextField="Description" runat="server"></asp:DropDownList>
                </div>
            </div>
            <div class="col-2">
                <div class="short-div">
                    <div class="input-group mb-1">
                        <div class="input-group-prepend">
                            <div class="input-group-text label_desc"><% Response.Write(lang["data_inizio"]); %>:</div>
                        </div>
                        <input class="form-control" id="inizio" name="inizio" type="date"/>
                    </div>
                    <div class="input-group">
                        <div class="input-group-prepend">
                            <div class="input-group-text label_desc"><% Response.Write(lang["data_fine"]); %>:</div>
                        </div>
                        <input class="form-control" id="fine" name="inizio" type="date" />
                    </div>
                </div>
            </div>
            <div class="col-0 ml-1">
                <div class="short-div">
                    <label class="switch"><input id="non_riconosciuti" type="checkbox"><span class="slider"></span></label> <% Response.Write(lang["filtra_non_riconosciuti"]); %><br />
                    <label class="switch"><input id="attivo" type="checkbox"><span class="slider"></span></label> <% Response.Write(lang["filtra_attivi"]); %><br />
                    <label class="switch"><input id="storico" type="checkbox"><span class="slider"></span></label> <% Response.Write(lang["filtra_storici"]); %>
                </div>
            </div>
            <div class="col-0 ml-1">
                <div class="short-div">
                    <label class="switch"><input id="allarmi" type="checkbox" checked><span class="slider"></span></label> <i class="fas fa-lightbulb" style="color:red;"></i> <% Response.Write(lang["filtra_allarmi"]); %><br />
                    <label class="switch"><input id="stati" type="checkbox" checked><span class="slider"></span></label> <i class="fas fa-info-circle" style="color:blue;"></i><% Response.Write(lang["filtra_sati"]); %>
                </div>
            </div>
            <div class="col-0 ml-1">
                <div class="input-group input-group-sm">
                    <div class="input-group-prepend">
                      <div class="input-group-text"><i class="fas fa-search"></i></div>
                    </div>
                    <input type="text" id="filtro_allarme" class="form-control" />
                </div>
            </div>
            <div class="col-0 ml-1">
                <button class="btn btn-sm btn-primary" id="riconosci_all"><i class="fas fa-check-double"></i> <% Response.Write(lang["riconosci_tutti"]); %></button>
            </div>
            <div class="col-0 ml-1">
                <table>
                    <tr>
                        <th rowspan="4"><% Response.Write(lang["legenda"]); %>:</th>
                        <th style="text-align:center;"><% Response.Write(lang["stati"]); %>:</th>
                        <th style="text-align:center;"><% Response.Write(lang["allarmi"]); %>:</th>
                    </tr>
                    <tr>
                        <td class="stato_attivo"><i class="fas fa-info-circle"></i> <% Response.Write(lang["attivo_non_riconosciuto"]); %></td>
                        <td class="allarme_attivo"><i class="fas fa-lightbulb"></i> <% Response.Write(lang["attivo_non_riconosciuto"]); %></td>
                    </tr>
                    <tr>
                        <td class="stato_non_attivo"><i class="fas fa-info-circle"></i> <% Response.Write(lang["non_attivo_non_riconosciuto"]); %></td>
                        <td class="allarme_non_attivo"><i class="fas fa-lightbulb"></i> <% Response.Write(lang["non_attivo_non_riconosciuto"]); %></td>
                    </tr>
                    <tr>
                        <td class="stato_riconosciuto"><i class="fas fa-info-circle"></i> <% Response.Write(lang["attivo_riconosciuto"]); %></td>
                        <td class="allarme_riconosciuto"><i class="fas fa-lightbulb"></i> <% Response.Write(lang["attivo_riconosciuto"]); %></td>
                    </tr>
                </table>
            </div>
        </div>
        <div class="row" id="riga_tabella">
            <table class="table table-bordered table-hover table-sm" ID="MainContent_tbl_allarmi">
                <thead>
                    <tr>
                        <th>#</th>
                        <th><% Response.Write(lang["tipo"]); %></th>
                        <th><% Response.Write(lang["descrizione"]); %></th>
                        <th><% Response.Write(lang["data_inizio"]); %></th>
                        <th><% Response.Write(lang["ora_inizio"]); %></th>
                        <th><% Response.Write(lang["data_fine"]); %></th>
                        <th><% Response.Write(lang["ora_fine"]); %></th>
                        <th><% Response.Write(lang["riconosciuto"]); %></th>
                        <th><% Response.Write(lang["data_riconoscimento"]); %></th>
                        <th><% Response.Write(lang["ora_riconoscimento"]); %></th>
                        <th></th>
                    </tr>
                </thead>
                <tbody id="corpo_tbl_allarmi"></tbody>
            </table>
        </div>
        <div class="row" style="position: relative;">
            
        </div>
        <script src="../Scripts/allarmi.js"></script>
    </main>
</asp:Content>
