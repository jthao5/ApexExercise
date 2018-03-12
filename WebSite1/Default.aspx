<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Apex Project Example (Johnny Thao)</title>
</head>
<body>
    <form id="form1" runat="server">
        <h1>Apex Project Example (Johnny Thao)</h1>
        <div>
            <asp:Label ID="Label1" runat="server" Text="Please Select Begin Date" Font-Names="Calibri" Font-Underline="True"></asp:Label>
            <asp:Calendar ID="Calendar1" runat="server" OnSelectionChanged="Calendar1_SelectionChanged" Font-Names="Calibri" TitleStyle-BackColor="#990000" TitleStyle-ForeColor="White"></asp:Calendar>
            <asp:Label ID="Label3" runat="server" Text="Please Select End Date" Font-Names="Calibri" Font-Underline="True"></asp:Label>
            <asp:Calendar ID="Calendar2" runat="server" OnSelectionChanged="Calendar2_SelectionChanged" Font-Names="Calibri" TitleStyle-BackColor="#990000" TitleStyle-ForeColor="White"></asp:Calendar>
            <asp:Label ID="Label5" runat="server" Text="Label" Font-Names="Calibri"></asp:Label><br />
            <asp:Button ID="Button1" runat="server" Text="View the First 15 Records" OnClick="Button1_OnClick" Font-Names="Calibri" />
            <asp:Button ID="Button2" runat="server" Text="Export All Records to Excel" OnClick="Button2_OnClick" Font-Names="Calibri" /><br />
            <asp:Label ID="Label6" runat="server" Text="No Records Found" Visible="False" Font-Names="Calibri" Font-Bold="True" Font-Size="Larger"></asp:Label>
            <asp:GridView ID="GridView1"  runat="server" AutoGenerateColumns="False" Font-Names="Calibri">
                <Columns>
                    <asp:BoundField DataField="Sold_At" HeaderText="Sold At">
                    <ControlStyle Width="50px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="Sold_To" HeaderText="Sold To">
                    <ControlStyle Width="50px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="Account_Number" HeaderText="Account Number">
                    <ControlStyle Width="50px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="Invoice__" HeaderText="Invoice #">
                    <ControlStyle Width="50px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="Customer_PO__" HeaderText="Customer PO #">
                    <ControlStyle Width="50px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="OrderDateText" HeaderText="Order Date">
                    <ControlStyle Width="50px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="DueDateText" HeaderText="Due Date">
                    <ControlStyle Width="50px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="Invoice_Total" HeaderText="Invoice Total" htmlencode="false" DataFormatString = "{0:C2}">
                    <ControlStyle Width="50px" />
                    </asp:BoundField>
                </Columns>
            </asp:GridView>  
        </div>
    </form>
</body>
</html>
