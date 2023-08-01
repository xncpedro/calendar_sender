VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim reminderSet As Boolean
Const olAppointmentItem = 1 ' Numeric value for AppointmentItem

Private Sub Application_Startup()
    ' Call your macro right after opening Outlook
    RunMacroWithDelay
End Sub

Sub RunMacroWithDelay()
    ' Set the delay time in seconds
    Dim delayInSeconds As Long
    delayInSeconds = 5 ' Change this value to the desired delay (in seconds)

    ' Set the reminder time
    scheduledTime = Now + TimeSerial(0, 0, delayInSeconds)

    ' Create a reminder
    Dim newReminder As Object ' Use Object type to avoid early binding
    Set newReminder = Application.CreateItem(olAppointmentItem)
    newReminder.reminderSet = True
    newReminder.Start = scheduledTime
    newReminder.Subject = "Envio Calendário" 'can be changed to match your requirements
    newReminder.Body = "O calendário vai ser enviado para o e-mail de destino." 'can be changed to match your requirements
    newReminder.Save

    reminderSet = True
End Sub

Private Sub Application_Reminder(ByVal Item As Object)
    If reminderSet Then
        ' Check if the reminder is the one we created
        If Item.Subject = "Envio Calendário" Then 'if newReminder.Subject is changed be sure to change it here too
            ' Execute your delayed macro here
            Send_Calendar

            ' Delete the reminder
            Item.Delete

            reminderSet = False
        End If
    End If
End Sub
Sub Send_Calendar()

Dim ol As Outlook.Application
    Dim cal As Folder
    Dim exporter As CalendarSharing
    dataInit = Format(Now(), "dd-mm-yyyy")
    dataFim = DateAdd("m", 30, dataInit)
    
    Set ol = Application
    Set cal = ol.Session.GetDefaultFolder(olFolderCalendar)
    Set exporter = cal.GetCalendarExporter
    
    With exporter
        .CalendarDetail = olFullDetails
        .IncludeAttachments = False
        .IncludePrivateDetails = False
        .RestrictToWorkingHours = False
        .IncludeWholeCalendar = False
        .StartDate = dataInit
        .EndDate = dataFim
        
        .SaveAsICal "C:\'Calendar_Save_Path.ics'"
    End With

Dim MyOutlook As Object
Set MyOutlook = CreateObject("Outlook.Application")

Dim MyMail As Object
Set MyMail = MyOutlook.CreateItem(olMailItem)

MyMail.To = "email_to_recive@calendar.com"
MyMail.Subject = "Calendar"

Attached_File = "C:\'Calendar_Save_Path.ics'"
MyMail.Attachments.Add Attached_File

MyMail.Send

Kill ("C:\'Calendar_Save_Path.ics'")

End Sub

