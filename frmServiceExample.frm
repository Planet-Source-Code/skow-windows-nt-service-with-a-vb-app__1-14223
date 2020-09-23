VERSION 5.00
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "ntsvc.ocx"
Begin VB.Form frmServiceExample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows NT Service Example"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   3255
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmServiceExample.frx":0000
      Top             =   0
      Width           =   6075
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Close Program"
      Default         =   -1  'True
      Height          =   350
      Left            =   2145
      TabIndex        =   0
      Top             =   3360
      Width           =   1815
   End
   Begin NTService.NTService NTService1 
      Left            =   1650
      Top             =   975
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      ServiceName     =   "Simple"
      StartMode       =   3
   End
End
Attribute VB_Name = "frmServiceExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is a basic file to show how to register as a Windows NT Service.
' Key points:
'
'       1. YOU MUST HAVE FULL ADMIN RIGHTS TO INSTALL/MODIFY SERVICES
'       2. A Service is when Windows NT runs a program that a user cannot close
'          it is started on Windows bootup and closes on Windows shutdown.
'       3. Windows NT controls weither the application is visible to the users or not.
'       4. Visual Basic Upto Version 6 (and mabye beyond TBA) DOES NOT nativly support
'          Windows NT Service commands BECAUSE VB is not a multi-thread applicaiton.
'          It is a well known to Microsoft but they don't really mind as it is a huge change
'          to make if they upgraded it.
'
'       5. Microsoft Released a fix for this problem in the way of an OCX file: NTSVC.ocx
'           This can be found on the MSDN Web site: msdn.microsoft.com - Search for NTSVC.OCX
'
'           Asof Dec 2000: HttpRef: http://msdn.microsoft.com/library/techart/msdn_ntsrvocx.htm

Private Sub Command1_Click()
 ' Used because we disabled form_Unload
 If MsgBox("Close Service?", vbYesNo, "Are you sure") = vbYes Then UnloadProgram
 
End Sub

'   This file explains how to use and work with this OCX file.

Private Sub Form_Load()
 ' First things first. We Need to Register this app in the Service Controller
 'But before we can do that, we must setup the NTService1 OCX.
 ' This SHOULD be done pre-runtime but doesn't have to be.
 
With Me.NTService1
    .DisplayName = "Example Service"        ' Displayed on Service List
    .Interactive = True                     ' Can this App be visible with the desktop
    .ServiceName = "Example"                ' Should be ONE WORD no spaces.
    .StartMode = svcStartDisabled           ' We don't really want to start this app.
End With
    

 ' We will do this if the '-install' command line is passed.
 ' We will REMOVE it from the list on '-uninstall' command line.
 
 If Trim$(Command$) <> "" Then
    Select Case UCase$(Trim$(Command$)) ' Upcase the Command Line and remove spaces at both ends.
            Case "-INSTALL"
                    If NTService1.Install Then
                            MsgBox "Result: " & App.Title & " successfully installed as a Windows NT Service." & vbCrLf _
                                & "Service Name: " & NTService1.ServiceName, vbInformation, "Install Complete, Please Re-Start Application"
                    Else
                            MsgBox "Result: " & App.Title & " FAILED to installed as a Windows NT Service." & vbCrLf _
                                & "Service Name: " & NTService1.ServiceName & vbCrLf _
                                & vbCrLf _
                                & "Solutions: Check to see if the service is allready installed. If so, run " & App.EXEName & " -uninstall to remove it." _
                                , vbInformation, "Install Failed, Please Re-Start Application"
                            
                    End If
                    
                    
                    End ' End the program, pass/fail
                    

            
            Case "-UNINSTALL"
            
                    If NTService1.Uninstall Then
                            MsgBox "Result: " & App.Title & " successfully uninstalled as a Windows NT Service." & vbCrLf _
                                & "Removed Service Name: " & NTService1.ServiceName, vbInformation, "UnInstall Complete, Please Re-Start Application"
                    Else
                            MsgBox "Result: " & App.Title & " FAILED to Uninstalled as a Windows NT Service." & vbCrLf _
                                & "Service Name: " & NTService1.ServiceName & vbCrLf _
                                & vbCrLf _
                                & "Solutions: Check to see if the service is installed. If not, run " & App.EXEName & " -install to install it." _
                                , vbInformation, "UnInstall Failed, Please Re-Start Application"
                    End If
                    
                    End ' End the program, pass or fail
            

            
            Case Else
                    ' Unknown Syntax. Inform user
                    MsgBox "Valid Syntax: " & vbCrLf _
                        & vbCrLf _
                        & "-install   To Install " & App.Title & " as a WinNT Service" & vbCrLf _
                        & vbCrLf _
                        & "-uninstall  To UN-INSTALL " & App.Title & " from the WinNT Service List" _
                        , vbInformation, "Invalid Syntax: Aborting Program Launch"
    
    End Select ' Ucase$(Trim$(Command$))
 End If 'Trim$(Command$) <> ""
            

' OK, no command line has been passed.
' We now need to talk to the NT Service Controller and tell them we are open.
NTService1.StartService


' Thats all we need to do for Form_Load
' Now onto:
'  NTService1_Start() - This is called When the service is started by Windows
'  NTService1_Stop()  - This is called By Windows to stop the service.
'  Form_Unload        - To close program as per usual
'  UnloadProgram      - Used like Form_Unload but better

End Sub

Private Sub Form_Unload(Cancel As Integer)
 If Trim$(Me.Tag) = "" Then
    'Windows caused Closure, Ignore this.
    ' I usually send my app to the System Tray here.
    Cancel = 1
    Me.WindowState = vbMinimized
Else
    'Planned Unload, let it pass

End If

End Sub

Private Sub NTService1_Start(Success As Boolean)
    ' Called by Windows upon Form_Load.
    ' It is important to use the NTService1.LogEvent for errors
    ' as it will log to the event log and can make life easy tracking down problems

On Error GoTo Err_Start

    ' Add any code here such as Me.Windowstate = vbMinimized
    ' if you wish this to be a Tray app etc.
    
    Success = True       ' Report success to NT Service Controller

Exit Sub    ' Avoid error logging

Err_Start:  ' error logging via NT Event Log
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & _
         Err.Number & "] " & Err.Description)
    Resume Next ' resume
End Sub

Sub NTService1_Stop()
On Error GoTo Err_Stop
    ' Add code to close files etc here.
    ' I use a UnloadProgram sub to close any open files (ie run from 0-99) and create logs etc
    ' Other then that, nothing is required here.
    
    
    If Trim$(Me.Tag) = "" Then UnloadProgram
    ' This is used so that Logging out does not close this program if it is
    ' set to Active (desktop visible). Because Windows Logout will call
    ' Form_Unload which cannot tell if it is supposed to close or not.
    ' So I use Form.Tag if it is Empty, It is still running, if it is full, close program
    ' In UnloadProgram, I set Me.Tag to "STOP" so that when I unload the service,
    ' it will NOT loop with the above statement (hence the if ()="" then..)
    
    
    
    'End
Exit Sub
Err_Stop:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & _
         Err.Number & "] " & Err.Description)

End Sub


Sub UnloadProgram()
On Error Resume Next
    Me.Tag = "STOP" 'To avoid looping
    If NTService1.Running Then NTService1.StopService ' Stop if not already
    While NTService1.Running
        DoEvents ' let it catch up
    Wend
    ' UNLOAD ALL OTHER FORMS THEN THIS FORM HERE
    ' Doens't matter if window visible or not, better safe then sorry
    
     ' Unload frmAbout
     ' Unload frmSplash
     
     ' Close ALL files here
     'For FileNum = 0 to 99 ' <- use if really ness. Otherwise close files properly
     '  Close FileNum
     'Next
     
    Unload Me
 End ' just to be sure
End Sub

