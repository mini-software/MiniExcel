<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="TestIssue223.aspx.cs" Inherits="WebFormTest.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:GridView ID="GridView1" runat="server" OnSelectedIndexChanged="GridView1_SelectedIndexChanged">
            </asp:GridView>
        </div>
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Download Excel" />
        <asp:Button ID="Button2" runat="server" OnClick="Button2_Click" Text="GridViewBind" />
    </form>
</body>
</html>
