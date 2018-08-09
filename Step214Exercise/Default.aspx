<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="Step214Exercise._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <h3>Import Excel File to a Database</h3>
    <div>
        <table>
            <tr>
                <td>Select File : </td>
                <td>
                    <asp:FileUpload ID="FileUpload1" runat="server" />
                </td>
                <td>
                    <asp:Button ID="btnImport" runat="server" Text="Import Data" OnClick="btnImport_Click"/>
                </td>
            </tr>
        </table>
        <div>
            <br />
            <asp:Label ID="lblMessage" runat="server" Font-Bold="true" />
            <br />
            <asp:GridView ID="gvData" runat="server" AutoGenerateColumns="false">
                <EmptyDataTemplate>
                    <div style="padding:10px">
                        Data not found!
                    </div>
                </EmptyDataTemplate>
                <Columns>
                    <asp:BoundField HeaderText="Player ID" DataField="PlayerID" />
                    <asp:BoundField HeaderText="First Name" DataField="FirstName" />
                    <asp:BoundField HeaderText="Last Name" DataField="LastName" />
                    <asp:BoundField HeaderText="Team Name" DataField="TeamName" />
                    <asp:BoundField HeaderText="Jersey Number" DataField="JerseyNumber" />
                    <asp:BoundField HeaderText="Points" DataField="Points" />
                </Columns>
            </asp:GridView>
        </div>
    </div>
</asp:Content>
