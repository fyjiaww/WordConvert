<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Index.aspx.cs" Inherits="WordConvert.Index" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <style type="text/css">
        td {
            padding-left: 20px;
        }
    </style>
</head>
<body height="100%">
    <form id="form1" runat="server" height="100%">
        <h3>word转换工具</h3>
        <table border="1" cellspacing="0" width="700" height="200" style="border: 1px solid beige">
            <tr>                
                <td colspan="2">
                    <asp:Label ID="Label4" runat="server" Text="数据源：" Font-Size="Small" ></asp:Label>                    
                    <asp:DropDownList ID="DropDownList2" runat="server" Width="116">
                        <asp:ListItem Value="1">中国皮革人才网</asp:ListItem>
                        <asp:ListItem Value="2">前程无忧</asp:ListItem>
                        <asp:ListItem Value="3">猎聘网</asp:ListItem>
                        <asp:ListItem Value="4">国际人才</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td width="100">
                    <asp:Label ID="Label1" runat="server" Text="提取word内容：" Font-Size="Small"></asp:Label></td>
                <td>
                    <asp:FileUpload ID="FileUpload1" runat="server" Width="190"/><asp:Button ID="Button1" runat="server" Text="提取" OnClick="Button1_Click" /></td>
            </tr>
            <tr>
                <td width="100">
                    <asp:Label ID="Label2" runat="server" Text="word转html:" Font-Size="Small"></asp:Label></td>
                <td>
                    <asp:FileUpload ID="FileUpload2" runat="server" Width="190"/><asp:Button ID="Button2" runat="server" Text="转换" OnClick="Button2_Click" /></asp:Label></td>
            </tr>
            <tr>
                <td width="100">
                    <asp:Label ID="Label3" runat="server" Text="数据库导出：" Font-Size="Small"></asp:Label></td>
                <td>
                    <asp:TextBox ID="tbName" runat="server" Width="70"></asp:TextBox>
                    <asp:DropDownList ID="DropDownList1" runat="server" Width="100">
                        <asp:ListItem Value="1">转word格式</asp:ListItem>
                        <asp:ListItem Value="2">转excel格式</asp:ListItem>
                        <asp:ListItem Value="3">转pdf格式</asp:ListItem>
                    </asp:DropDownList>
                    <asp:Button ID="Button3" runat="server" Text="导出" OnClick="Button3_Click" /></td>
            </tr>
        </table>
        <iframe id="BodyFrame" runat="server" height="391" width="100%">

        </iframe>
    </form>
</body>
</html>
