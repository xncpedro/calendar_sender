Attribute VB_Name = "Module1"
Sub Send_Calendar()

Dim ol As Outlook.Application
    Dim cal As Folder
    Dim exporter As CalendarSharing
    dataInit = Format(Now(), "dd-mm-yyyy")
    dataFim = DateAdd("m", 30, dataInit)        'Set here the date range
    
    Set ol = Application
    Set cal = ol.Session.GetDefaultFolder(olFolderCalendar)
    Set exporter = cal.GetCalendarExporter
    
    With exporter
        .CalendarDetail = olFullDetails      'Set here the detail level
        .IncludeAttachments = False
        .IncludePrivateDetails = False
        .RestrictToWorkingHours = False
        .IncludeWholeCalendar = False   'changing this to True will send the whole calendar
        .StartDate = dataInit           'remove this when IncludeWholeCalendar = True
        .EndDate = dataFim              'remove this when IncludeWholeCalendar = True
        
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

