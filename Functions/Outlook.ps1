function Connect-Outlook {
    $Global:Outlook = New-Object -ComObject Outlook.Application 
}

function Get-OutlookCalendarItems {
    param(
        [scriptblock]$Filter
    )
    if (-not $Global:Outlook) {
        Connect-Outlook
    }

    $Calendar = $Global:Outlook.Session.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)

    if ($Filter) {
        $Appointments = $Calendar.Items | Where-Object -FilterScript $Filter
    }
    else {
        $Appointments = $Calendar.Items | Where-Object { $_.Start -gt (Get-Date).Date -and $_.Start -lt (Get-Date).AddDays(1).Date }
    }

    $Appointments = $Appointments | Sort-Object -Property Start

    $defaultProperties = @('subject', 'start', 'end')

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet("DefaultDisplayPropertySet", [string[]]$defaultProperties)

    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
   
    $Appointments | Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $PSStandardMembers

    $Appointments
}

function New-OutlookCalendarAppointment {
    param(
        [string]$Subject,
        [datetime]$Start = (Get-Date),
        [datetime]$End = (Get-Date).addHours(1)
    )
    if (-not $Global:Outlook) {
        Connect-Outlook
    }
    $Appointment = $Global:Outlook.CreateItem([Microsoft.Office.Interop.Outlook.OlItemType]::olAppointmentItem)
    $Appointment.Start = $Start
    $Appointment.End = $End
    $Appointment.Subject = $Subject
    $Appointment.save()

    $defaultProperties = @('subject', 'start', 'end')

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet("DefaultDisplayPropertySet", [string[]]$defaultProperties)

    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
   
    $Appointment | Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $PSStandardMembers

    $Appointment

}

function New-OutlookTask {
    [CmdletBinding()]
    param ( 
        [parameter(Position = 0)]
        [String]$Description = $(throw "Please specify a description"),
        [switch] $Force 
    )

    DynamicParam {

        if (-not $Global:Outlook) {
            Connect-Outlook
        }

        # Get special folder names for ValidateSet attribute
        $Categories = $Global:Outlook.Session.Categories | Select-Object -ExpandProperty Name
        
        # Create new dynamic parameter
        New-DynamicParameter -Name Category -ValidateSet $Categories -Type ([string[]]) `
            -Position 1
    }

    Process {
        # Bind dynamic parameter to a friendly variable
        New-DynamicParameter -CreateVariables -BoundParameters $PSBoundParameters

        if (-not $Global:Outlook) {
            Connect-Outlook
        }
        ## Create our Outlook and housekeeping variables.  
        ## Note: If you don't have the Outlook wrappers, you'll need 
        ## the commented-out constants instead
    
        $olTaskItem = [Microsoft.Office.Interop.Outlook.OlItemType]::olTaskItem
    
        #$olTaskItem = 3 
        #$olFolderTasks = 13
    
        $Task = $Global:outlook.Application.CreateItem($olTaskItem) 
        $hasError = $false
    
        ## Assign the subject 
        $Task.Subject = $description
    
        ## If they specify a category, then assign it as well. 
        if ($category) { 
            $Task.Categories = $category
        }
    
        ## Save the item if this didn't cause an error, and clean up. 
        if (-not $hasError) { $Task.Save() }

        $Task | Add-Member -MemberType ScriptProperty -Name StatusParsed -Value `
        {
            # Get
            ([Microsoft.Office.Interop.Outlook.OlTaskStatus]$this.Status).ToString() -replace "olTask", ""
        } `
        {
            # Set
            param(
                [ValidateSet("NotStarted", "InProgress", "Complete", "Waiting", "Deferred")]
                [string]$Status
            )
            $this.Status = [Microsoft.Office.Interop.Outlook.OlTaskStatus]"olTask$Status"
            $this.save()
        }

        $Task | Add-Member -MemberType ScriptProperty -Name "Focus Level" -Value `
        {
            # Get
            (($this.Categories -split ',' | Where-Object { $_ -match "Focus Level" }) -split ': ')[1]
        } `
        {
            # Set
            param(
                [ValidateSet(
                    "Low",
                    "Medium",
                    "High"
                )]
                [string]$FocusLevel
            )
    
            if ($this.Categories -match "Focus Level") {
                $this.Categories = ($this.Categories -split ',' | Where-Object { $_ -notmatch "Focus Level" } ) -join ','
            }
    
            $this.Categories += ", Focus Level: $FocusLevel"
            $this.save()
        }
     
        $Task | Add-Member -MemberType ScriptProperty -Name "Location" -Value `
        {
            # Get
            (($this.Categories -split ',' | Where-Object { $_ -match "Location" }) -split ': ')[1]
        } `
        {
            # Set
            param([string]$Location)
            if ($this.Categories -match "Location") {
                $this.Categories = ($this.Categories -split ',' | Where-Object { $_ -notmatch "Location" } ) -join ','
            }
    
            $this.Categories += ", Location: $Location"
            $this.save()
        }
     
        $Task | Add-Member -MemberType ScriptProperty -Name "Estimated Pomodoros" -Value `
        {
            # Get
            (($this.Categories -split ',' | Where-Object { $_ -match "Estimated Pomodoros" }) -split ': ')[1]
        } `
        {
            # Set
            param(
                [ValidateSet(
                    "<1",
                    "1",
                    "2",
                    "3",
                    "4",
                    "5"
                )]
                [string]$Pomodori
            )
    
            if ($this.Categories -match "Estimated Pomodoros") {
                $this.Categories = ($this.Categories -split ',' | Where-Object { $_ -notmatch "Estimated Pomodoros" } ) -join ','
            }
    
            $this.Categories += ", Estimated Pomodoros: $Pomodori"
            $this.save()
        } 
        
        $defaultProperties = @('Subject', 'StatusParsed', 'Categories', 'Focus level', 'Location', 'Estimated Pomodoros')
    
        $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet("DefaultDisplayPropertySet", [string[]]$defaultProperties)
    
        $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
       
        $Task | Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $PSStandardMembers

        $Task

    }
    
    
}

function Get-OutlookTasks {

    param(
        [string]$Folder,

        [ValidateNotNullOrEmpty()]
        [ValidateSet('Open', 'Closed')]
        [string]$State = 'Open'
    )
    

    if (-not $Global:Outlook) {
        Connect-Outlook
    }
    if (-not $Folder) {
        $TasksFolder = $Global:Outlook.Session.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderTasks)
    }
    else {
        $Folders = $Folder -split "\\"
        $TasksFolder = $Global:Outlook.Session.Folders.Item($Folders[0])
        $Folders[1..$Folders.count] | foreach {
            if (-not [string]::IsNullOrEmpty($_)) {
                $TasksFolder = $TasksFolder.Folders.Item($_)   
            }
        }
    }

    if ($State -eq "Closed") {
        $StatusQuery = "[Status] = 2"
    }
    else {
        $StatusQuery = "[Status] <> 2"
    }

    $Tasks = $TasksFolder.Items.Restrict("$StatusQuery And [MessageClass] = 'IPM.Task'") | Sort-Object -Property Subject
    $Tasks | Add-Member -MemberType ScriptProperty -Name StatusParsed -Value `
    {
        # Get
        ([Microsoft.Office.Interop.Outlook.OlTaskStatus]$this.Status).ToString() -replace "olTask", ""
    } `
    {
        # Set
        param(
            [ValidateSet("NotStarted", "InProgress", "Complete", "Waiting", "Deferred")]
            [string]$Status
        )
        $this.Status = [Microsoft.Office.Interop.Outlook.OlTaskStatus]"olTask$Status"
        $this.save()
    }

    $Tasks | Add-Member -MemberType ScriptProperty -Name "Focus Level" -Value `
    {
        # Get
        (($this.Categories -split ',' | Where-Object { $_ -match "Focus Level" }) -split ': ')[1]
    } `
    {
        # Set
        param(
            [ValidateSet(
                "Low",
                "Medium",
                "High"
            )]
            [string]$FocusLevel
        )

        if ($this.Categories -match "Focus Level") {
            $this.Categories = ($this.Categories -split ',' | Where-Object { $_ -notmatch "Focus Level" } ) -join ','
        }

        $this.Categories += ", Focus Level: $FocusLevel"
        $this.save()
    }
 
    $Tasks | Add-Member -MemberType ScriptProperty -Name "Location" -Value `
    {
        # Get
        (($this.Categories -split ',' | Where-Object { $_ -match "Location" }) -split ': ')[1]
    } `
    {
        # Set
        param([string]$Location)
        if ($this.Categories -match "Location") {
            $this.Categories = ($this.Categories -split ',' | Where-Object { $_ -notmatch "Location" } ) -join ','
        }

        $this.Categories += ", Location: $Location"
        $this.save()
    }

    $Tasks | Add-Member -MemberType ScriptProperty -Name "Folder" -Value `
    {
        # Get
        $this.Parent.FolderPath
    } `
    {
        # Set
        param([string]$Folder)
        $Folders = $Folder -split "\\"
        $TasksFolder = $Global:Outlook.Session.Folders.Item($Folders[0])
        $Folders[1..$Folders.count] | foreach {
            if (-not [string]::IsNullOrEmpty($_)) {
                $TasksFolder = $TasksFolder.Folders.Item($_)
            }
        }
        $this.move($TasksFolder)
        $this.save()
    }

    $defaultProperties = @('Subject', 'StatusParsed', 'DueDate', 'Focus level', 'Location', 'Folder')

    @(
        "Completed Pomodori",
        "Internal Interruptions",
        "External Interruptions",
        "Estimated Pomodori"
    ) | ForEach-Object {
        $Tasks | Add-Member -MemberType ScriptProperty -Name $_ -Value `
        ([scriptblock]::Create(@"
            # Get
            if (-not `$this.Userproperties.Item("$($_ -Replace " ", "")")) {
                `$UserProperty = `$this.Userproperties.Add("$($_ -Replace " ", "")", [Microsoft.Office.Interop.Outlook.OlUserPropertyType]::olInteger)
                `$UserProperty.Value = 0
            }
        
            `$this.UserProperties.Item("$($_ -Replace " ", "")").Value
"@)) `
        ([scriptblock]::Create(@"
            # Set
            param(
                [int]`$Pomodori
            )

            `$this.UserProperties.Item("$($_ -Replace " ", "")").Value
            `$this.save()
"@))

        $defaultProperties += $_
    }    
        

    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet("DefaultDisplayPropertySet", [string[]]$defaultProperties)

    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
   
    $Tasks | Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $PSStandardMembers

    [array]$Tasks
}

Function New-OutlookNoteToSelf {
    param(
        $Note
    )

    if (-not $Global:Outlook) {
        Connect-Outlook
    }

    $Inbox = $Outlook.Session.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

    $NoteToSelf = $Inbox.Application.CreateItem([Microsoft.Office.Interop.Outlook.OlItemType]::olMailItem)
    $NoteToSelf.Subject = "Note: $Note"
    $NoteToSelf.Save()
    $NoteToSelf.Move($Inbox) | Out-Null
}

New-Alias -Name todo -Value New-OutlookTask
