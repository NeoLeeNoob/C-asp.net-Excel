<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ECS.aspx.cs" Inherits="ECS.ECS2" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>

    <title></title>
    </head>
    <body>
    <form id="form1" runat="server">
        <div>
            <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
            <asp:TextBox ID="TextBox2" runat="server"></asp:TextBox>
            <asp:TextBox ID="TextBox3" runat="server"></asp:TextBox>
            <asp:TextBox ID="TextBox4" runat="server"></asp:TextBox>
            <asp:TextBox ID="TextBox5" runat="server"></asp:TextBox>
            <br/>
            <asp:Button ID="Button1" runat="server" Text="新增" OnClick="Button1_Click" />
            <asp:Button ID="Button2" runat="server" Text="修改" OnClick="Button2_Click" />
            <asp:Button ID="Button3" runat="server" Text="查詢" OnClick="Button3_Click" />
            <br/>
            <asp:Label ID="Label1" runat="server" Text=""></asp:Label>
            <br/>
            <asp:CheckBox ID="CheckBox1" runat="server" Text="名字" OnCheckedChanged="CheckBox1_CheckedChanged1" Checked="true" AutoPostBack="True" />
            <asp:CheckBox ID="CheckBox2" runat="server" Text="公司" OnCheckedChanged="CheckBox2_CheckedChanged2" Checked="true" AutoPostBack="True" />
            <asp:CheckBox ID="CheckBox3" runat="server" Text="料號" OnCheckedChanged="CheckBox3_CheckedChanged3" Checked="true" AutoPostBack="True" />
            <asp:CheckBox ID="CheckBox4" runat="server" Text="地址" OnCheckedChanged="CheckBox4_CheckedChanged4" Checked="true" AutoPostBack="True" />
            <asp:CheckBox ID="CheckBox5" runat="server" Text="料件" OnCheckedChanged="CheckBox5_CheckedChanged5" Checked="true" AutoPostBack="True" />
            <br/>
            <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False">
                <Columns>
                    <asp:TemplateField>
            <ItemTemplate>
                <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="https://localhost:44345/images/REDATA.JPG" OnClick="ImageButton1_Click" Height="30px" Width="30px" />
                <asp:ImageButton ID="ImageButton2" runat="server" ImageUrl="https://localhost:44345/images/DELETE.JPG" OnClick="ImageButton2_Click" Height="30px" Width="30px" />
            </ItemTemplate>
        </asp:TemplateField>
         <asp:BoundField DataField="ID" HeaderText="ID" SortExpression="ID"   Visible="True" />
        <asp:BoundField DataField="名字" HeaderText="名字" SortExpression="名字" Visible="True"/>
        <asp:BoundField DataField="公司" HeaderText="公司" SortExpression="公司" Visible="True"/>
        <asp:BoundField DataField="料號" HeaderText="料號" SortExpression="料號" Visible="True"/>
        <asp:BoundField DataField="地址" HeaderText="地址" SortExpression="地址" Visible="True"/>
        <asp:BoundField DataField="料件" HeaderText="料件" SortExpression="料件" Visible="True"/>
    </Columns>
</asp:GridView>
        </div>
    </form>
</body>
</html>