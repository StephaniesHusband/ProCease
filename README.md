# FedEx ProCease

ProCease is a collection of [AutoHotkey](http://www.autohotkey.com) macros/scripts to help with "CEASING the burden of PROCESS" (hence the new name).

## Instructions

### Installation

Clone the git repository and install the files in a folder that's in your PATH.

### Configuration

There is a configuration section at the top of the ProCease.ahk file.

### Executing

You can run ProCease by running the ProCease.ahk file or put a link to it in your start-up folder so that it will start when you boot up.

- - -
# API
ProCease is built upon an API that can be modified to suit your needs.  Here are the existing API (function) calls that can
be made and their descriptions.

####Parameter Type Explanation
AEV - AutoHotkey Expression Variable _(e.g., THIS_VARIABLE) (no percentage signs necessary)_

TV - AutoHotkey Tradtional Variable _(e.g., %SOME_VARIABLE%)_
- - - 
### WaitFor(waitForMe)
Wait for and activate the given window.
>####Parameters
waitForMe = title of window to wait for (AEV)
>####Returns
Nothing

### ClipDefectNumber()
Capture the defect number out of the QC defect dialog and into the Clipboard variable.
>####Parameters
None
>####Returns
Defect Number

### GoToTab(tabName)
Select the given tab
>####Parameters
tabName = Tab Name to go to (AEV). Use the TAB_* constants or any other value.
>####Returns
Nothing

### ClickOk()
Clicks the OK button in the QC Defect Details dialog.
>####Parameters
None
>####Returns
Nothing

### SetAssignedTo(id)
Set the assigned to value to the employee id given.
>####Parameters
id = Employee ID to assign the defect to (AEV)
>####Returns
Nothing

### SetTeamAssigned(team)
Set the team assigned to the team value given.
>####Parameters
team = Team to assign the defect to (AEV). Use the TEAM_* constants or any other value.
>####Returns
Nothing

### SetDefectStatus(status)
Set the defect to the given status. 
>####Parameters
status = Defect status. (AEV) Use the STATUS_* constants or any other value.
>####Returns
Nothing

### SetAssignedToVersion(ver)
Sets the Assigned To Version.
>####Parameters
ver = Software version to assign the defect to. (AEV) Use the VER_ASSIGNED_TO constants or any other value.
>####Returns
Nothing

### SetPlannedClosingVersion(ver)
Set the Planned Closing Version that the defect is fixed in.
>####Parameters
ver = Planned Closing version. (AEV) Use the VER_PLANNED_CLOSING constant or any other value.
>####Returns
Nothing

### SetTargetTestCycle(cycle)
Sets the target test cycle (i.e., Marketing Sprint #)
>####Parameters
cycle = Target test cycle (AEV)
>####Returns
Nothing

### SetDefectType(type)
Set the defect type
>####Parameters
type = Defect Type. use the DEFECT_TYPE constant or any other value.
>####Returns
Nothing

### SetRootCauseTeam(team)
Set the root cause team
>####Parameters
team = Root cause team (AEV). Use the TEAM_ROOT_CAUSE constant or any other value.
>####Returns
Nothing

### SetResolution(res)
Set the resolution to the value in RESOLUTION as a template for resolving defects.
>####Parameters
res = Resolution template value (AEV).  Can use RESOLUTION constant or any other value
>####Returns
Nothing

### FindDefect(id)
Find the given defect id in the QC Defect system. Requires that the QC Defect system be open.
>####Parameters
id = Defect ID (AEV).
>####Returns
Nothing

### FixDefect(pcv, ttc, atv)
Fix the given defect
>####Parameters
pcv = Planned closing version (AEV) - can use the VER_PLANNED_CLOSING constant or any other value.
ttc = Target Test Cycle (AEV) - can use the TARGET_TEST_CYCLE constant or any other value.
atv = Assigned To Version (AEV) - can use the VER_ASSIGNED_TO constant or any other value.
>####Returns
Nothing

### LaunchDefectActionWindow()
Launch the Defect Action window 
>####Parameters
None
>####Returns
Nothing

### LaunchArtifactPicker()
Launch the Artifact Picker dialog so that you can pick an artifact.  Uses FILE_ARTIFACTS constant.
>####Parameters
None
>####Returns
Nothing

### GetClipboard()
Get the contents of the clipboard and trim head and trailing spaces off the value.
>####Parameters
None
>####Returns
Trimmed contents of the clipboard

### SelectAll()
Select everything with a Ctrl-A in whatever context the focus is in at the time of calling.
>####Parameters
None
>####Returns
Nothing
