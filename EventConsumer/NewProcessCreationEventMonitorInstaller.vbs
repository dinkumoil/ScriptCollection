' ===========================================
' Run this script with administrative rights!
' ===========================================


'/////////////////////////////// Configure script's job ///////////////////////////////

'******************** Customize according to your specific needs **********************

'------------------ Set event description and job related variables -------------------

strEventDescription  = "New Process Creation"
strExeName           = "cmd.exe"
strPollingIntervall  = "1"


'----------------------------- Set WMI Event Filter query -----------------------------

strEventFilterQuery  = "SELECT * FROM __InstanceCreationEvent" _
                     & " WITHIN " & strPollingIntervall _
                     & " WHERE TargetInstance ISA 'Win32_Process'" _
                     & " AND TargetInstance.Caption='" & strExeName & "'"


'-------------- Set WMI Namespace where the monitored event will occure ---------------

strEventNamespace    = "root\Cimv2"


'----------- Set parameters which should be passed to Event Handler script ------------

strEventHandlerParam = """" _
                     & "%TargetInstance.ProcessId%" _
                     & """"

'**************************************************************************************



'//////////////////////////////// Set script variables ////////////////////////////////

'- Set Event Handler path, quote it and prepare it and its working directory for WMI --

Set FSO           = CreateObject("Scripting.FileSystemObject")

strEventHandlerWD = Replace(FSO.GetParentFolderName(WScript.ScriptFullName), "\", "\\")

strEventHandler   = Replace(strEventDescription, " ", "") & "EventHandler.vbs"
strEventHandler   = """" & strEventHandlerWD & "\\" & strEventHandler & """"

Set FSO           = Nothing


'----------------------- Test for existing Event Handler script -----------------------

Set FSO = CreateObject("Scripting.FileSystemObject")

If Not FSO.FileExists(Replace(Replace(strEventHandler, "\\", "\"), """", "")) Then
  MsgBox "The Event Handler script " & vbCRLF _
          & Replace(strEventHandler, "\\", "\") & vbCRLF _
          & "does not exist. Please create it.", _
         vbExclamation, _
         "Missing Event Handler script"
End If

Set FSO = Nothing


'-------------------------- Build parameters for CScript.exe --------------------------

strCScriptParam = strEventHandler & " " & strEventHandlerParam


'---------------- Get path of Windows Directory and prepare it for WMI ----------------

Set WshShell    = WScript.CreateObject("WScript.Shell")

strWinDir       = Replace(WshShell.ExpandEnvironmentStrings("%SystemRoot%"), "\", "\\")

Set WshShell    = Nothing


'---------------------- Set path of CScript.exe prepared for WMI ----------------------

strCScriptPath       = strWinDir & "\\System32\\cscript.exe"
strCScriptPathQuoted = """" & strCScriptPath & """"



'////////////////////////// Install permanent Event Consumer //////////////////////////

'------- Get WMI Scripting API object (SWbemServices), Namespace: Subscription --------

strComputer                          = "."
Set objWMIService                    = GetObject("winmgmts:" _
                                                 & "{impersonationLevel=impersonate}!" _
                                                 & "\\" & strComputer & "\root\Subscription")


'------------------------------ Create the Event Filter -------------------------------

Set objFilterClass                   = objWMIService.Get("__EventFilter")
Set objFilter                        = objFilterClass.SpawnInstance_()

objFilter.Name                       = strEventDescription & " Event Filter"
objFilter.EventNamespace             = strEventNamespace
objFilter.QueryLanguage              = "WQL"
objFilter.Query                      = strEventFilterQuery

Set EventFilterPath                  = objFilter.Put_()


'----------------------- Create the Commandline Event Consumer ------------------------

Set objEventConsumerClass            = objWMIService.Get("CommandLineEventConsumer")
Set objEventConsumer                 = objEventConsumerClass.SpawnInstance_()

objEventConsumer.Name                = strEventDescription & " Commandline Event Consumer"
objEventConsumer.CommandLineTemplate = strCScriptPathQuoted & " " & strCScriptParam
objEventConsumer.ExecutablePath      = strCScriptPath
objEventConsumer.WorkingDirectory    = strEventHandlerWD
objEventConsumer.ShowWindowCommand   = 0

Set CommandlineEventConsumerPath     = objEventConsumer.Put_()


'------------------ Bind Event Filter to Commandline Event Consumer -------------------

Set objBindingClass                  = objWMIService.Get("__FilterToConsumerBinding")
Set objBinding                       = objBindingClass.SpawnInstance_()

objBinding.Filter                    = EventFilterPath
objBinding.Consumer                  = CommandlineEventConsumerPath

objBinding.Put_()
