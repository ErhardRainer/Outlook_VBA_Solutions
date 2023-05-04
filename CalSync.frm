VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalSync 
   Caption         =   "Calendar Sync Assistant"
   ClientHeight    =   2715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10920
   OleObjectBlob   =   "CalSync.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "CalSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SourceFolder As Outlook.Folder
Public TargetFolder As Outlook.Folder
Public MyCategory As String

Function GetFolderPath(olFolder As Outlook.Folder) As String
    If Not olFolder Is Nothing Then
        If Not olFolder.Parent Is Nothing And TypeOf olFolder.Parent Is Outlook.NameSpace Then
            GetFolderPath = GetFolderPath(olFolder.Parent) & "\" & olFolder.Name
        Else
            GetFolderPath = olFolder.Name
        End If
    End If
End Function


Private Sub cmd_Cancel_Click()
    Unload CalSync
End Sub

Private Sub cmd_OK_Click()

    ' Declare variables
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olAppointment As Outlook.AppointmentItem
    Dim olTargetAppointment As Outlook.AppointmentItem
    Dim msgResult As VbMsgBoxResult

    ' Get Outlook application and namespace
    Set olApp = Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")

     ' Iterate through shared calendar appointments
    For Each olAppointment In SourceFolder.Items
        If TypeName(olAppointment) = "AppointmentItem" Then
            ' Check if appointment exists in the target calendar
            Set olTargetAppointment = AppointmentExists(TargetFolder, olAppointment.Subject, olAppointment.Start, MyCategory)
            If olTargetAppointment Is Nothing Then
                ' Output the subject and start date to the Immediate window
                Debug.Print "NEW Subject: " & olAppointment.Subject & "; Start: " & olAppointment.Start

                ' Create a new appointment in the target calendar
                Set olTargetAppointment = TargetFolder.Items.Add(olAppointmentItem)

                ' Copy appointment details
                With olTargetAppointment
                    .Subject = olAppointment.Subject
                    .Start = olAppointment.Start
                    .End = olAppointment.End
                    .AllDayEvent = olAppointment.AllDayEvent
                    .Location = olAppointment.Location
                    .Body = olAppointment.Body
                    .Categories = "Uni"
                    .Save
                End With
            Else
                ' Check if appointment subject and category are the same
                If olAppointment.Subject = olTargetAppointment.Subject And olAppointment.Categories = olTargetAppointment.Categories Then
                    ' Delete the old appointment
                    olTargetAppointment.Delete

                    ' Output the subject and start date to the Immediate window
                    Debug.Print "OVERWRITE Subject: " & olAppointment.Subject & "; Start: " & olAppointment.Start

                    ' Create a new appointment in the target calendar
                    Set olTargetAppointment = TargetFolder.Items.Add(olAppointmentItem)
                                    ' Copy appointment details
                With olTargetAppointment
                    .Subject = olAppointment.Subject
                    .Start = olAppointment.Start
                    .End = olAppointment.End
                    .AllDayEvent = olAppointment.AllDayEvent
                    .Location = olAppointment.Location
                    .Body = olAppointment.Body
                    .Categories = MyCategory
                    .Save
                End With
            End If
        End If
    End If
Next olAppointment

MsgBox "Der Online-Kalender wurde erfolgreich importiert und synchronisiert.", vbInformation, "Erfolg"
Unload CalSync
End Sub

Function AppointmentExists(olCalendar As Outlook.Folder, Subject As String, StartDate As Date, Category As String) As Outlook.AppointmentItem

Dim olItems As Outlook.Items
Dim olAppointment As Outlook.AppointmentItem
Dim Filter As String

Filter = "[Subject] = '" & Replace(Subject, "'", "''") & "' AND [Start] = '" & Format(StartDate, "ddddd H:mm") & "' AND [Categories] = '" & Category & "'"
Set olItems = olCalendar.Items
Set olAppointment = olItems.Find(Filter)

If Not olAppointment Is Nothing Then
    Set AppointmentExists = olAppointment
Else
    Set AppointmentExists = Nothing
End If

Set olItems = Nothing
Set olAppointment = Nothing

End Function


Private Sub cmd_Source_Click()
    Dim olNS As Outlook.NameSpace
    Dim olCalendar As Outlook.Folder
    
    Set olNS = Application.GetNamespace("MAPI") ' Initialize olNS
    Set olCalendar = olNS.PickFolder

    ' Check if the selected folder is a calendar
    If olCalendar.DefaultItemType <> olAppointmentItem Then
        MsgBox "Bitte wählen Sie einen Kalenderordner aus.", vbExclamation, "Ungültiger Ordner"
        Exit Sub
    Else
        Set SourceFolder = olCalendar
        Me.txt_Source.Value = GetFolderPath(olCalendar)
    End If
End Sub

Private Sub cmd_target_Click()
    Dim olNS As Outlook.NameSpace
    Dim olCalendar As Outlook.Folder

    Set olNS = Application.GetNamespace("MAPI") ' Initialize olNS
    Set olCalendar = olNS.PickFolder

    ' Check if the selected folder is a calendar
    If olCalendar.DefaultItemType <> olAppointmentItem Then
        MsgBox "Bitte wählen Sie einen Kalenderordner aus.", vbExclamation, "Ungültiger Ordner"
        Exit Sub
    Else
        Set TargetFolder = olCalendar
        Me.txt_Target.Value = GetFolderPath(olCalendar)
    End If
End Sub


Public Sub ListOutlookCategories()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olCategory As Outlook.Category
    Dim olCategories As Outlook.Categories

    Set olApp = New Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    Set olCategories = olNS.Categories

    ' List all available categories
    For Each olCategory In olCategories
        Debug.Print "Name: " & olCategory.Name & "; Color: " & olCategory.Color & "; Shortcuts: " & olCategory.ShortcutKey
    Next olCategory
End Sub



Private Sub UserForm_Initialize()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim olCategory As Outlook.Category
    Dim olCategories As Outlook.Categories

    Set olApp = New Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
    Set olCategories = olNS.Categories

    ' List all available categories
    For Each olCategory In olCategories
        Me.cb_Category.AddItem (olCategory.Name)
    Next olCategory
End Sub
