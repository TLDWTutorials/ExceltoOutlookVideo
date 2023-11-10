# ExceltoOutlookVideo
Synchronizing events from Microsoft Excel to Microsoft Outlook using VBA code as shown on my YouTube video: https://youtu.be/Hk_pJ-OWuXQ

# Code
Sub AddAppointmentsToOutlookCalendar()
    Dim olApp As Object ' Outlook.Application
    Dim olNamespace As Object ' Outlook.Namespace
    Dim olFolder As Object ' Outlook.Folder
    Dim olApt As Object ' Outlook.AppointmentItem
    
    ' Create Outlook application object
    Set olApp = CreateObject("Outlook.Application")
    
    ' Get Outlook default namespace
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    ' Get default calendar folder
    Set olFolder = olNamespace.GetDefaultFolder(9) ' olFolderCalendar
    
    ' Excel variables
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    
    ' Set the workbook and worksheet
    Set wb = ThisWorkbook ' or specify the workbook name/path
    Set ws = wb.Worksheets("Sheet1") ' Modify as per your sheet name
    
    ' Find the last non-empty row in column A (Appointment_Name)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Check if there are appointments data
    If lastRow < 2 Then
        MsgBox "No appointments found in the dataset.", vbInformation
        Exit Sub
    End If
    
    ' Set the range based on the longest row that isn't null
    Set rng = ws.Range("A2:H" & lastRow)
    
    ' Loop through each appointment in the range
    For Each cell In rng.Rows
        ' Create a new appointment item
        Set olApt = olFolder.Items.Add(1) ' olAppointmentItem
        
        ' Set appointment properties from Excel cells
        With olApt
            .Subject = cell.Range("A1").Value ' Set subject
            
            ' Set start date/time
            Dim startDate As Date
            Dim startTime As Date
            startDate = cell.Range("B1").Value
            startTime = cell.Range("C1").Value
            .Start = startDate + startTime
            
            ' Set end date/time
            Dim endDate As Date
            Dim endTime As Date
            endDate = cell.Range("D1").Value
            endTime = cell.Range("E1").Value
            .End = endDate + endTime
            
            .Location = cell.Range("F1").Value ' Set location
            .Body = cell.Range("G1").Value ' Set body/description
            .ReminderSet = True ' Set reminder (True/False)
            .ReminderMinutesBeforeStart = 15 ' Set reminder time (if ReminderSet is True)
        End With
        
        ' Save the appointment
        olApt.Save
        
        ' Release object references
        Set olApt = Nothing
    Next cell
    
    ' Release Outlook objects
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
    
    MsgBox "Appointments added to Outlook calendar successfully!", vbInformation
End Sub
