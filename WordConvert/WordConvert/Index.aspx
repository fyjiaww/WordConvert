<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Index.aspx.cs" Inherits="WordConvert.Index" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
</head>
<body>
    <form id="form1" runat="server">
        <h4>word转换工具</h4>
        <table border="1" width="700" height="200">
            <tr>
                <td>
                    <asp:Label ID="Label1" runat="server" Text="提取word内容："></asp:Label>
                    <asp:FileUpload ID="FileUpload1" runat="server" /><asp:Button ID="Button1" runat="server" Text="提取" /></td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label2" runat="server" Text="word转html:">
                        <asp:FileUpload ID="FileUpload2" runat="server" /><asp:Button ID="Button2" runat="server" Text="转换" /></asp:Label></td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label3" runat="server" Text="数据库导出：">
                        <asp:DropDownList ID="DropDownList1" runat="server">
                            <asp:ListItem Value="1">转word格式</asp:ListItem>
                            <asp:ListItem Value="2">转excel格式</asp:ListItem>
                            <asp:ListItem Value="3">转pdf格式</asp:ListItem>
                        </asp:DropDownList></asp:Label>
                    <asp:Button ID="Button3" runat="server" Text="导出" /></td>
            </tr>
        </table>
    </form>
</body>
</html>
