<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CallRecordingGUI.aspx.vb" Inherits="Call_Recording_GUI._Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
<style type="text/css">
.loginTable td{font-family:"Arial";font-size:12px; min-width: 150px; width-auto: !important; _width: 150px}
.responseTable td{font-family:"Arial";font-size:12px}
</style>    
<asp:Literal id="styleLiteral" runat="server" />
<title>ReachAll Call Recording Interface</title>
</head>
<body bgcolor="#256BB1" onload="PageLoad();">
<br id="defaultHeaderId"/>
<div id="knowallHeaderId" style="border-bottom:1px dotted #666666; height:90px; position:relative; margin-bottom:40px; display:none;">
  <table width="700" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="320"><img src="assets/login_logo.png" alt="Login logo"/></td>
      <td><h1> ReachAll Call Recording Interface</h1></td>
    </tr>
  </table>
  <img src="assets/logo_swyx_rec_logo.png" style="position:absolute; right:10px; top:20px;" alt="Swyx logo"/>
</div>
<form id="Form1" runat="server">
<asp:Label id="f_offset" Text="" Visible="False" runat="server" />
<asp:Label id="f_loggedIn" Text="No" Visible="False" runat="server" />
<asp:Label id="f_parm" Text="" Visible="False" runat="server" />
<asp:Label ID="f_data" Text="" Visible="False" runat="server" />

<!-- ************************************** -->
<!-- The Login Panel for the legacy version -->
<!-- ************************************** -->
<asp:Panel id="loginPanel" visible="true" runat="server">
    <table class="loginTable" Align="center" CellSpacing="20">
	    <tr>
	        <td align = "Right" style="font-weight: bold">Username</td>
	        <td>
		        <asp:TextBox id="f_userName" runat="server" width="165" />
	        </td>
		</tr>
		
		<tr>
	        <td align = "Right" style="font-weight: bold">Password</td>
	        <td>
		        <asp:TextBox id="f_password" textMode="Password" runat="server" width="165" />
	        </td>
	    </tr>
	    
	    <tr>
	        <td></td>
	        <td align="center">
		        <asp:Button id="loginButton" text = "Login" runat="server" OnClick="LoginPressed" />
	        </td>
	    </tr>
	    
	    <tr>
	        <td></td>
	        <td align="center">
	            <asp:Button id="changePasswordButton" text = "Change Password" runat="server" OnClick="ChangePasswordPressed" visible="false" />
	        </td>
	    </tr>
	
	    <asp:TableRow Visible="false" ID="cpRow_1" runat="server">
            <asp:TableCell HorizontalAlign = "Right" style="font-weight: bold">New Password</asp:TableCell>
		    <asp:TableCell>
			    <asp:TextBox id="newPasswordTextBox" runat="server" width="165" textMode="Password" />
		    </asp:TableCell></asp:TableRow><asp:TableRow Visible="false" ID="cpRow_2" runat="server">
            <asp:TableCell HorizontalAlign = "Right" style="font-weight: bold">Repeat New Password</asp:TableCell><asp:TableCell>
			    <asp:TextBox id="newPasswordAgainTextBox" runat="server" width="165" textMode="Password" />
		    </asp:TableCell></asp:TableRow><asp:TableRow Visible="false" ID="cpRow_3" runat="server">
            <asp:TableCell></asp:TableCell><asp:TableCell HorizontalAlign="Center">
                <asp:Button id="confirmChangePasswordButton" text = "Confirm" runat="server" OnClick="ConfirmChangePasswordPressed" visible="true" />
            </asp:TableCell></asp:TableRow></table><table align="center"><tr><td><asp:Label ID="changePasswordResponseLegacy" Text = "" runat="server" /></td></tr></table>
            </asp:Panel>
 
            <!-- ********************************* -->
            <!-- The Login Panel for New Interface -->
            <!-- ********************************* -->
            <asp:Panel ID="loginPanelNewInterface" visible="false" runat="server">
            <div class="modal-content">
               <div class="imgcontainer">
                    <img src="icons8-person-96.png" alt="Avatar" class="avatar"> </div><div class="container">
                    <label for="userNameNewInterface" class="SectionHeading">Username</label> <asp:TextBox id="userNameNewInterface" placeHolder="Enter Username" cssClass="newInterfaceInputClass" runat="server" />
                    <br />
                    <br />
                    <label for="passwordNewInterface" class="SectionHeading">Password</label> <asp:TextBox id="passwordNewInterface" placeHolder="Enter Password" textMode="Password" cssClass="newInterfaceInputClass" runat="server" />
                    <br />
                    <br />
                    <asp:Button id="loginButtonNewInterface" text="Login" runat="server" OnClick="LoginPressed" cssClass="button" />
                    <br />
                    <br />
		            <asp:Button id="changePasswordButtonNewInterface" text = "Change Password" runat="server" OnClick="ChangePasswordPressed" visible="false" cssClass="button" BackColor="#d35b47" />
                    <asp:Label ID="changePasswordResponseNewInterface" Text = "" runat="server" />
                    <asp:Panel ID="changePasswordPanelNewInterface" visible="false" runat="server">
                        <label for="newPasswordTextBoxNewInterface" class="SectionHeading">New Password</label> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:TextBox 
                            ID="newPasswordTextBoxNewInterface" runat="server" 
                            cssClass="newInterfaceInputClass" placeholder="New Password" 
                            textMode="Password" /><br /><br /><label class="SectionHeading" 
                            for="newPasswordAgainTextBoxNewInterface">Confirm New Password</label>&nbsp;&nbsp;&nbsp;<asp:TextBox id="newPasswordAgainTextBoxNewInterface" placeholder="Confirm New Password" textMode="Password" cssClass="newInterfaceInputClass" runat="server" />
                        <br />
                        <br />
                        <asp:Button id="confirmChangePasswordButtonNewInterface" text ="Confirm" runat="server" OnClick="ConfirmChangePasswordPressed" cssClass="button" />
                    </asp:Panel>
                </div>
            </div>
    </asp:Panel>

<asp:Table id="responseTable" runat="server" class="responseTable" HorizontalAlign="center" CellSpacing="20" Visible="false">
<asp:TableRow>
<asp:TableCell>
<asp:Label ID="changePasswordResponse" Text = "" runat="server" style="font-weight: bold"/>
</asp:TableCell></asp:TableRow></asp:Table><asp:Table id="adminChoiceTable" runat="server" HorizontalAlign="center" CellSpacing="20" Visible="false">
	<asp:TableRow>
	<asp:TableCell></asp:TableCell><asp:TableCell HorizontalAlign="center">
			<asp:Button id="recordingsButton" text = "Recordings" runat="server" OnClick="recordingsButtonPressed" />
		</asp:TableCell><asp:TableCell HorizontalAlign="center">
			<asp:Button id="adminButton" text = "Admin" runat="server" OnClick="AdminButtonPressed"/>
		</asp:TableCell><asp:TableCell HorizontalAlign="center">
			<asp:Button id="closeButton" text = "Close" runat="server" OnClick="CloseButtonPressed"/>
		</asp:TableCell></asp:TableRow></asp:Table><asp:Table id="adminTable" runat="server" HorizontalAlign="center" CellSpacing="20" Visible="false">

<asp:Tablerow>
    <asp:TableCell columnspan="4"><h2>Administrator Logins</h2></asp:TableCell></asp:Tablerow><asp:Tablerow>
<asp:TableCell>
<asp:Literal id="administratorsLit0" runat="server" />
</asp:TableCell><asp:TableCell>
<asp:Literal id="administratorsLit1" runat="server" />
</asp:TableCell><asp:TableCell></asp:TableCell><asp:TableCell></asp:TableCell></asp:Tablerow><asp:TableRow>
<asp:TableCell>
<asp:TextBox id="administratorsUsernameText" runat="server" width="140" />
</asp:TableCell><asp:TableCell>
<asp:TextBox id="administratorsPasswordText" runat="server" width="140" />
</asp:TableCell><asp:TableCell>
<asp:Button id="addAdministratorButton" runat="server" text="Add Administrator" OnClick="AddAdministratorButtonPressed" Width="140" />
</asp:TableCell><asp:TableCell>
</asp:TableCell></asp:TableRow><asp:Tablerow>
<asp:TableCell columnspan="4"><hr /></asp:TableCell></asp:Tablerow>

<asp:Tablerow>
<asp:TableCell columnspan="4"><h2>Supervisor Logins</h2></asp:TableCell></asp:Tablerow><asp:Tablerow>
<asp:TableCell>

<asp:Literal id="supervisorsLit0" runat="server" />
</asp:TableCell><asp:TableCell>
<asp:Literal id="supervisorsLit1" runat="server" />
</asp:TableCell><asp:TableCell ColumnSpan=2>
<asp:Literal id="supervisorsLit2" runat="server" />
</asp:TableCell><asp:TableCell></asp:TableCell></asp:Tablerow><asp:TableRow>
<asp:TableCell>
<asp:TextBox id="supervisorsUsernameText" runat="server" width="140" />
</asp:TableCell><asp:TableCell>
<asp:TextBox id="supervisorsPasswordText" runat="server" width="140" />
</asp:TableCell><asp:TableCell>
<asp:TextBox id="supervisorsDataText" runat="server" width="140" />
</asp:TableCell><asp:TableCell>
<asp:Button id="addSupervisorButton" runat="server" text="Add Supervisor" OnClick="AddSupervisorButtonPressed" Width="140"/>
</asp:TableCell></asp:TableRow><asp:TableRow>
<asp:TableCell></asp:TableCell><asp:TableCell></asp:TableCell><asp:TableCell></asp:TableCell><asp:TableCell>
<asp:Button id="adminOKButton" runat="server" text="Close" onClick = "AdminOKButtonPressed" width="140" />
</asp:TableCell></asp:TableRow></asp:Table><asp:Panel ID="searchPanel" runat="server">
    <table CellSpacing="20" Align="center" style="text-align: center;">
	    <tr HorizontalAlign="center">
		    <td><strong>Start Date (dd/mm/yyyy)yyy)</strong></td><td><strong>End Date (dd/mm/yyyy)</strong></td><td><strong>Extension</strong></td><td><strong>External Number</strong></td><td><strong>Call ID</strong></td></tr><tr HorizontalAlign="center">
		    <td VerticalAlign="Top"><asp:TextBox id="f_startDate" runat="server" /></td>
		    <td VerticalAlign="Top"><asp:TextBox id="f_endDate" runat="server" /></td>
		    <td VerticalAlign="Top"><asp:TextBox id="f_extension" runat="server" /></td>
		    <td VerticalAlign="Top"><asp:TextBox id="f_destination" runat="server" /></td>
		    <td VerticalAlign="Top"><asp:TextBox id="f_callId" runat="server" /></td>
		    <td VerticalAlign="Top"><asp:Button id="TestButton" runat="server" OnClick="TestButtonPressed" visible="False"/></td>
	    </tr>
    	
	    <tr>
		    <td Rowspan="3">
			    <asp:Calendar id="calendar1" BackColor="#e0e5f0" nextPrevFormat="shortmonth" runat="server">
  	 		    <WeekendDayStyle BackColor="#d0d5e0" ForeColor="#ff0000" />
   			    <DayHeaderStyle ForeColor="#0000ff" />
   			    <TodayDayStyle BackColor="#f0f0ff" />
   			    <SelectedDayStyle BackColor="#22aa00" />
			    </asp:Calendar>
		    </td>
    		
		    <td RowSpan="3">
			    <asp:Calendar id="calendar2" BackColor="#e0e5f0" nextPrevFormat="shortmonth" runat="server">
  	 		    <WeekendDayStyle BackColor="#d0d5e0" ForeColor="#ff0000" />
   			    <DayHeaderStyle ForeColor="#0000ff" />
   			    <TodayDayStyle BackColor="#f0f0ff" />
   			    <SelectedDayStyle BackColor="#22aa00" />
			    </asp:Calendar>
		    </td>
    		
		    <td style="text-align: center; vertical-align: bottom;"><strong>Start Time</strong></td><td style="text-align: center; vertical-align: bottom;"><strong>End Time</strong></td><td style="text-align: center; vertical-align: bottom;"><strong>Direction</strong></td></tr><tr>    
		    <td style="text-align: center; vertical-align: top;">
		        <asp:TextBox ID="f_tag" visible="false" runat="server" />
		        <asp:DropDownList verticalalign="Top" ID="startTime" runat="server" Width="40%" visible="true"></asp:DropDownList>
		    </td>
    		
		    <td style="text-align: center; vertical-align: top;">
		        <asp:DropDownList VerticalAlign="Top" ID="endTime" runat="server" Width="40%" Visible="true">
		        </asp:DropDownList>
		    </td>
    		
		    <td style="text-align: center; vertical-align: top;">
		        <asp:DropDownList VerticalAlign="Top" ID="callTypeList" runat="server" Width="70%" Visible="true">
		        </asp:DropDownList>
		    </td>
	    </tr>
    		
	    <tr>
		    <td style="text-align: center; vertical-align: bottom;">
			    <asp:Button id="logoutButton" Text="Logout" runat="server" BackColor="#63a1df" OnClick="LogoutPressed" />
		    </td>
    		
		    <td style="text-align: center; vertical-align: bottom;">
			    <asp:Button id="clearButton" Text="Clear All" runat="server" BackColor="#63a1df" OnClick="Clear" />
		    </td>
    		
		    <td style="text-align: center; vertical-align: bottom;">
			    <asp:Button id="submitButton" Text="Search ..." runat="server" BackColor="#accdee" OnClick="Submit" />
		    </td>
	    </tr>
    	
	    <tr>
	        <td></td>
	        <td></td>
	        <td></td>
	        <td></td>
	        <td HorizontalAlign="center"><asp:CheckBox ID="legacyCheckBox" Text="Use Legacy Data" runat="server" /></td>
	    </tr>
    </table>
</asp:Panel>

    <!-- ************************************** -->
    <!-- The Search Panel for the New Interface -->
    <!-- ************************************** -->
    <asp:Panel ID="newInterfaceSearchPanel" runat="server">
        <table class="myTableClass">
            <tr>
                <td class="SectionHeading">Action</td><td></td><td></td>
            </tr>
            <tr align="center">
                <td style="width: 33.3%">
                    <asp:Button id="newInterfaceLogoutButton" text="Logout" cssClass="ActionButtonClass" onclick="LogoutPressed" runat="server" BackColor="#5c6086" />
                </td>
                <td style="width: 33.3%">
                    <asp:Button ID="newInterfaceClearButton" text="Clear All" cssClass="ActionButtonClass" onclick="Clear" runat="server" BackColor="#d35b47" />
                </td>
                <td style="width: 33.3%">
                    <asp:Button ID="newInterfaceSubmitButton" text="Search ..." cssClass="ActionButtonClass" onClick="NewInterfaceSubmit" runat="server" BackColor="#2e5fa2" />
                </td>
            </tr>
        </table>
        <br />
        <table class="myTableClass">
            <tr>
                <td class="SectionHeading">Time Period</td><td></td>
                <td></td>
            </tr>
            <tr align="center">
                <td style="height: 200px; width: 33.3%;">
                    <!-- <table style="height: 222px;"> -->
    <!--                <table style="height: 100%;"> Does not work in Chrome or IE -->
                    <table style="height: 222px; border-collapse: collapse;">
                        <tr align="center" style="height: 50%">
                            <td style="vertical-align: top; width: 100%"><div class="SmallContainer">
                                    <label for="timeSpanNewInterfaceDropDownList" class="LabelAbove">Time Span</label> <asp:DropDownList id="timeSpanNewInterfaceDropDownList" width = "200" OnSelectedIndexChanged="timeSpanIndexChanged" AutoPostBack="true" runat="server" />
                                </div>
                            </td>
                        </tr>
                        <tr style="vertical-align: bottom;">
                            <td>
                                <div class="SmallContainer">
                                    <table width="100%">
                                        <tr>
                                            <td colspan="2"  style="text-align:center">
                                                <asp:CheckBox id="ignoreTimeNewInterfaceCheckBox" checked="true" text="Ignore Time" cssclass="myClass" onchange="IgnoreTimeChanged();" runat="server" />
                                                <!-- <input type="checkbox" id="ignoreTimeNewInterfaceCheckBox" class="myClass" onchange="IgnoreTimeChanged();" /> -->
                                                <!-- <label for="ignoreTimeCheckBox">Ignore Time</label> -->
                                            </td>
                                        </tr>
                                        <tr style="text-align:center">
                                            <td>
                                                <asp:label id="startTimeNewInterfaceLabel" for="startTimeNewInterfaceDropDownList" cssClass="LabelAbove" enabled="false" runat="server">Start Time</asp:label><asp:DropDownList id="startTimeNewInterfaceDropDownList" enabled="false" width="80" runat="server" />
                                            </td>
                                            <td>
                                                <asp:label id="endTimeNewInterfaceLabel" for="endTimeNewInterfaceDropDownList" cssClass="LabelAbove" enabled="false" runat="server">End Time</asp:label><asp:DropDownList id="endTimeNewInterfaceDropDownList" enabled="false" width="80" runat="server" />
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                            </td>
                        </tr>
                    </table>
                </td>
                <td style="width: 33.3%">
                    <div class="LargeContainer">
                        <asp:Calendar id="calendar3" BorderStyle="None" BorderWidth="0" TitleStyle-BorderStyle="None" TitleStyle-BackColor="White" caption="Start Date" BackColor="#e0e5f0" nextPrevFormat="shortmonth" cssClass="CalendarClass" runat="server">
                            <WeekendDayStyle BackColor="#d0d5e0" ForeColor="#ff0000" />
                            <DayStyle Font-Size="11pt" />
                            <WeekendDayStyle BackColor="#d0d5e0" ForeColor="#ff0000" Font-Size="11pt" />
                            <DayHeaderStyle ForeColor="#0000ff" Font-Size="11pt"/>
                            <TodayDayStyle BackColor="#f0f0ff" />
                            <SelectedDayStyle BackColor="#5c6086" />
                            <NextPrevStyle Font-Size="11pt" ForeColor="#5c6086" />
                        </asp:Calendar>
                    </div>
                </td>
                <td style="width: 33.3%">
                    <div class="LargeContainer">
                        <asp:Calendar id="calendar4" BorderStyle="None" BorderWidth="0" TitleStyle-BorderStyle="None" TitleStyle-BackColor="White" caption="End Date" BackColor="#e0e5f0" nextPrevFormat="shortmonth" cssClass="CalendarClass" runat="server">
                            <WeekendDayStyle BackColor="#d0d5e0" ForeColor="#ff0000" />
                            <DayStyle Font-Size="11pt" />
                            <WeekendDayStyle BackColor="#d0d5e0" ForeColor="#ff0000" Font-Size="11pt" />
                            <DayHeaderStyle ForeColor="#0000ff" Font-Size="11pt"/>
                            <TodayDayStyle BackColor="#f0f0ff" />
                            <SelectedDayStyle BackColor="#5c6086" />
                            <NextPrevStyle Font-Size="11pt" ForeColor="#5c6086" />
                        </asp:Calendar>
                    </div>
                </td>
            </tr>
        </table> 
        <br />
        <table class="myTableClass">
            <tr>
                <td class="SectionHeading">Filter</td><td></td><td></td><td></td>
            </tr>

            <tr align="center">
                <td style="height: 80px; width: 25%; vertical-align: top;">
                    <div class="SmallContainer192">
                        <label for="extensionTextBoxNewInterface" class="LabelAbove">Extension</label> <asp:TextBox id="extensionTextBoxNewInterface" width = "148" runat="server" />
                    </div>
                </td>
                <td style="height: 80px; width: 25%; vertical-align: top;">
                    <div class="SmallContainer192">
                        <label for="externalNumberTextBoxNewInterface" class="LabelAbove">External Number</label> <asp:TextBox id="externalNumberTextBoxNewInterface" width = "148" runat="server" />
                    </div>
                </td>
                <td style="height: 80px; width: 25%; vertical-align: top;">
                    <div class="SmallContainer192">
                        <label for="callIdTextBoxNewInterface" class="LabelAbove">Call Id</label> <asp:TextBox id="callIdTextBoxNewInterface" width = "148" runat="server" />
                    </div>
                </td>
                <td style="height: 80px; width: 25%; vertical-align: top;">
                    <div class="SmallContainer192">
                        <label for="directionDropDownListNewInterface" class="LabelAbove">Call Type</label> <asp:DropDownList id="directionDropDownListNewInterface" width = "148" AutoPostBack="true" runat="server" />
                    </div>
                </td>
            </tr>
        </table>       
    </asp:Panel>

<asp:Panel ID="summaryPanelNewInterface" runat="server">    
<br />
<table id="summaryTable">
    <tr>
        <td>
            <asp:Button id="prevButton" Text="Previous" Visible=False runat="server" OnClick="PrevClicked" />
        </td>
        <td>
            <asp:Label id="recordsLabel" runat="server" />
        </td>
        <td>
            <asp:Button id="nextButton" Text="Next" Visible=False runat="server" OnClick="NextClicked" />
        </td>
    </tr>
</table>
<br id="breakBetweenTables" />
<asp:Repeater id="myRepeater" runat="server">

<HeaderTemplate>
<table id="resultsTable">
<tr>
<th>Time</th><th>Date</th><th>Extension</th><th>External Number</th><th>Direction</th><th>Duration</th><th>Call Id</th><th>Action</th></tr></HeaderTemplate><ItemTemplate>
<tr>
<td><%#Container.DataItem("time")%></td><td><%#Container.DataItem("date")%></td><td><%#Container.DataItem("internalNumber")%></td><td><%#Container.DataItem("externalNumber")%></td><td><%#Container.DataItem("direction")%></td><td><%#Container.DataItem("duration")%></td><td><%#Container.DataItem("callid")%></td><td><%#Container.DataItem("action")%></td></tr></ItemTemplate><AlternatingItemTemplate>
<tr>
<td><%#Container.DataItem("time")%></td><td><%#Container.DataItem("date")%></td><td><%#Container.DataItem("internalNumber")%></td><td><%#Container.DataItem("externalNumber")%></td><td><%#Container.DataItem("direction")%></td><td><%#Container.DataItem("duration")%></td><td><%#Container.DataItem("callid")%></td><td><%#Container.DataItem("action")%></td></AlternatingItemTemplate><FooterTemplate>
</table>
</FooterTemplate>

</asp:Repeater>
<br />
</asp:Panel>
<asp:Literal id="myLiteral" runat="server" />
</form>
<script type="text/javascript">
    function IgnoreTimeChanged() {
        var myFlag = document.getElementById("ignoreTimeNewInterfaceCheckBox").checked;

        document.getElementById("startTimeNewInterfaceDropDownList").disabled = myFlag;
        document.getElementById("endTimeNewInterfaceDropDownList").disabled = myFlag;

        return false;
    }

    function PageLoad() {
        IgnoreTimeChanged();
    }
</script>
</body>
</html>
