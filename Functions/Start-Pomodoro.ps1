Function Start-PomodoroWork {
    Param (
        [int]$Minutes = 25,
        [int]$Seconds = 0,
        [object]
        $Task,
        [int]$EstimatedPomodori = $Task.Userproperties.Item("EstimatedPomodori").Value,
        [int]$BreakDuration
    )
    if (-not $Task.Userproperties.Item("EstimatedPomodori")) {
        $UserProperty = $Task.Userproperties.Add("EstimatedPomodori", [Microsoft.Office.Interop.Outlook.OlUserPropertyType]::olInteger)
        $UserProperty.Value = $EstimatedPomodori
    }

    if (-not $Task.UserProperties.Item("CompletedPomodori")) {
        $UserProperty = $Task.Userproperties.Add("CompletedPomodori", [Microsoft.Office.Interop.Outlook.OlUserPropertyType]::olInteger)
        $CompletedPomodori = 0
    }
    else {
        $CompletedPomodori = $Task.UserProperties.Item("CompletedPomodori").Value     
    }


    if ($EstimatedPomodori -lt 1) {
        $EstimatedPomodori = [int](Read-Host "How many pomodori do you estimate it will take to complete this task?")
        $UserProperty.Value = $EstimatedPomodori
    }


    $Task.Save()

    while (-not $Task.Status -eq 2) {
        $StopWatch = New-Object -TypeName System.Diagnostics.Stopwatch
        Start-ConsoleSong -Song "Mission Impossible"
        $Goal = [timespan]"00:$Minutes`:$Seconds"
        $StopWatch.Start()

        $timer = new-object timers.timer 
    
        $action = {
            function Write-ToPos ([string] $str, [int] $x = 0, [int] $y = 0,
                [string] $bgc = [console]::BackgroundColor,
                [string] $fgc = [Console]::ForegroundColor) {
    
                if ($x -ge 0 -and $y -ge 0 -and $x -le [Console]::WindowWidth -and
    
                    $y -le [Console]::WindowHeight) {
   
                    $saveX = [console]::CursorLeft
                    $saveY = [console]::CursorTop
    
                    $offY = [console]::WindowTop       
    
                    [console]::setcursorposition($x, $offY + $y)
    
                    Write-Host -Object $str -BackgroundColor $bgc -ForegroundColor $fgc -NoNewline
    
                    [console]::setcursorposition($saveX, $saveY)
    
                }
            }

            cls
            Write-ToPos -x 10 -y 9 -str "Working on: $($Event.MessageData.Task.Subject)"
            Write-ToPos -x 10 -y 10 -str "Time Left: $(($Event.MessageData.Goal - $Event.MessageData.StopWatch.Elapsed).ToString())"
            Write-ToPos -x 10 -y 15 -str "Internal interruptions: $($Event.MessageData.PomodoroSyncHash.InternalInterruptions) `t External interruptions: $($Event.MessageData.PomodoroSyncHash.ExternalInterruptions)"
            Write-ToPos -x 10 -y 20 -str "Press [Q] to cancel pomodoro"
            Write-ToPos -x 10 -y 22 -str "Press [-] to log an external interruption"
            Write-ToPos -x 10 -y 24 -str "Press ['] to log an internal interruption"
            Write-ToPos -x 10 -y 26 -str "Pomodori: $( 
                for($i = 1; $i -le $Event.MessageData.EstimatedPomodori -or $i -le $Event.MessageData.CompletedPomodori; $i++) {
                    if ($i -le $Event.MessageData.EstimatedPomodori -and $i -le $Event.MessageData.CompletedPomodori){
                        " [x] "
                    } elseif ($i -gt $Event.MessageData.CompletedPomodori){
                        " [ ] "
                    } else {
                        " x "
                    }
            } )"
        } 
    
        $timer.Interval = 300  
        $PomodoroSynchash = [hashtable]::Synchronized(@{
                "ExternalInterruptions" = 0
                "InternalInterruptions" = 0
            })

        if (-not $Task.Userproperties.Item("ExternalInterruptions")) {
            $UserProperty = $Task.Userproperties.Add("ExternalInterruptions", [Microsoft.Office.Interop.Outlook.OlUserPropertyType]::olInteger)
            $UserProperty.Value = 0
        }
        else {
            $PomodoroSynchash.ExternalInterruptions = $Task.Userproperties.Item("ExternalInterruptions").Value         
        }
        
        if (-not $Task.UserProperties.Item("InternalInterruptions")) {
            $UserProperty = $Task.Userproperties.Add("InternalInterruptions", [Microsoft.Office.Interop.Outlook.OlUserPropertyType]::olInteger)
            $UserProperty.Value = 0
        }
        else {
            $PomodoroSynchash.InternalInterruptions = $Task.Userproperties.Item("InternalInterruptions").Value         
        }

        $MessageData = @{}
        $MessageData.timer = $timer
        $MessageData.task = $Task
        $MessageData.goal = $Goal
        $MessageData.Stopwatch = $StopWatch
        $MessageData.PomodoroSyncHash = $PomodoroSynchas
        $MessageData.EstimatedPomodori = $EstimatedPomodori
        $MessageData.CompletedPomodori = $CompletedPomodori
    
    
        Register-ObjectEvent -InputObject $timer -EventName elapsed `
            -SourceIdentifier thetimer -Action $action -MessageData $MessageData
        
        $timer.start()

        $Interruptions = @()

        [console]::TreatControlCAsInput = $True
        
        $Exit = $False
        while (-not $Exit -and ($Goal - $StopWatch.Elapsed) -gt 0 ) {
            if ([console]::KeyAvailable) {
                switch -regex ([console]::ReadKey().key) {
                    "(Q|C)" { 
                        $Exit = $True 
                        $timer.stop()
                        $StopWatch.Stop() 
                        #cleanup 
                        Unregister-Event thetimer
                        [console]::TreatControlCAsInput = $False
                        cls
                        Write-Host "Giving up eh? This pomodoro will not be logged..."
                        if ($Interruptions) {
                            Complete-Interruptions -Interruptions $Interruptions 
                        }
                        return
                    }
                    "(OemMinus|Subtract)" {
                        cls
                        $Interruptions += Add-Interruption -Type External
                        $PomodoroSynchash.ExternalInterruptions += 1; 
                        $Task.Userproperties.Item("ExternalInterruptions").Value = $PomodoroSynchash.ExternalInterruptions 
                    }
                    "Oem7" {
                        cls 
                        $Interruptions += Add-Interruption -Type Internal
                        $PomodoroSynchash.InternalInterruptions += 1; 
                        $Task.Userproperties.Item("InternalInterruptions").Value = $PomodoroSynchash.InternalInterruptions 
                    }
                    Default {
                        
                    }
                }
            }
        }
    
        $timer.stop()
        $StopWatch.Stop() 
        #cleanup 
        Unregister-Event thetimer
        [console]::TreatControlCAsInput = $False
        Start-ConsoleSong -Song 'Imperial March' 
        if ($Interruptions) {
            Complete-Interruptions -Interruptions $Interruptions 
        }
        $CompletedPomodori += 1
        cls

        $title = "Pomodoro $CompletedPomodori done!"
        $message = "Are you finished with this task?"
        
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
            "We'll mark the task as completed"
        
        $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
            "We'll take a break and then close out of this pomodoro"

        $options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
        
        $result = $host.ui.PromptForChoice($title, $message, $options, 0) 
        
        switch ($result) {
            0 {
                $Task.Status = 2
                $Task.save()
            }
        } 

        cls
        Write-Host "Congrats you successfully completed a pomodoro for the following task: $($Task.Subject). Time for a break.."
        $Task.ActualWork += $Minutes
        $Task.Userproperties.Item("CompletedPomodori").Value = $CompletedPomodori
        $Task.Save()
        Start-PomodoroBreak -BreakDuration $BreakDuration
    }

}

Function Start-ConsoleSong {
    param(
        [ValidateSet("Mission Impossible", "Imperial March")]
        [string]$Song
    )

    switch ($Song) {
        "Imperial March" {
            [console]::beep(440, 500)
            [console]::beep(440, 500) 
            [console]::beep(440, 500) 
            [console]::beep(349, 350) 
            [console]::beep(523, 150) 
            [console]::beep(440, 500) 
            [console]::beep(349, 350) 
            [console]::beep(523, 150)
            [console]::beep(440, 500) 
            [console]::beep(440, 1000) 
            [console]::beep(659, 500) 
            [console]::beep(659, 500) 
            [console]::beep(659, 500) 
            [console]::beep(698, 350) 
            [console]::beep(523, 150) 
            [console]::beep(415, 500) 
            [console]::beep(349, 350) 
            [console]::beep(523, 150)
            Start-Sleep -Milliseconds 2000
            [console]::beep(440, 1000)
        }

        "Mission Impossible" {
            [console]::beep(784, 150) 
            Start-Sleep -m 300 
            [console]::beep(784, 150) 
            Start-Sleep -m 300 
            [console]::beep(932, 150) 
            Start-Sleep -m 150 
            [console]::beep(1047, 150) 
            Start-Sleep -m 150 
            [console]::beep(784, 150) 
            Start-Sleep -m 300 
            [console]::beep(784, 150) 
            Start-Sleep -m 300 
            [console]::beep(699, 150) 
            Start-Sleep -m 150 
            [console]::beep(740, 150) 
            Start-Sleep -m 150 
            [console]::beep(784, 150) 
            Start-Sleep -m 300 
            [console]::beep(784, 150) 
            Start-Sleep -m 300 
            [console]::beep(932, 150) 
            Start-Sleep -m 150 
            [console]::beep(1047, 150) 
            Start-Sleep -m 150 
            [console]::beep(784, 150) 
            Start-Sleep -m 300 
            [console]::beep(784, 150) 
            Start-Sleep -m 300 
            [console]::beep(699, 150) 
            Start-Sleep -m 150 
            [console]::beep(740, 150) 
            Start-Sleep -m 150 
            [console]::beep(932, 150) 
            [console]::beep(784, 150) 
            [console]::beep(587, 1200) 
            Start-Sleep -m 75 
            [console]::beep(932, 150) 
            [console]::beep(784, 150) 
            [console]::beep(554, 1200) 
            Start-Sleep -m 75 
            [console]::beep(932, 150) 
            [console]::beep(784, 150) 
            [console]::beep(523, 1200) 
            Start-Sleep -m 150 
            [console]::beep(466, 150) 
            [console]::beep(523, 150)
        }
    }

    Start-Sleep -Milliseconds 200
}

Function Start-PomodoroBreak {
    param(
        [int]$BreakDuration
    )
    $StopWatch = New-Object -TypeName System.Diagnostics.Stopwatch
    $Goal = [timespan]"00:$BreakDuration`:00"
    $StopWatch.Start()

    
    $timer = new-object timers.timer 
   
    [console]::TreatControlCAsInput = $True
    
    $action = {
        function Write-ToPos ([string] $str, [int] $x = 0, [int] $y = 0,
            [string] $bgc = [console]::BackgroundColor,
            [string] $fgc = [Console]::ForegroundColor) {
    
            if ($x -ge 0 -and $y -ge 0 -and $x -le [Console]::WindowWidth -and
    
                $y -le [Console]::WindowHeight) {
   
                $saveX = [console]::CursorLeft
                $saveY = [console]::CursorTop
    
                $offY = [console]::WindowTop       
    
                [console]::setcursorposition($x, $offY + $y)
    
                Write-Host -Object $str -BackgroundColor $bgc -ForegroundColor $fgc -NoNewline
    
                [console]::setcursorposition($saveX, $saveY)
    
            }
        }

        cls
        Write-ToPos -x 10 -y 9 -str "Take a Break!"
        Write-ToPos -x 10 -y 10 -str "Time Left: $(($Event.MessageData.Goal - $Event.MessageData.StopWatch.Elapsed).ToString())"
        Write-ToPos -x 10 -y 20 -str "Press [Q] to cancel break early"

    } 
    
    $timer.Interval = 300  

    $MessageData.timer = $timer
    $MessageData.goal = $Goal
    $MessageData.Stopwatch = $StopWatch
    
    
    Register-ObjectEvent -InputObject $timer -EventName elapsed `
        -SourceIdentifier thetimer -Action $action -MessageData $MessageData
        
    $timer.start()

        
    $Exit = $False
    while (-not $Exit -and ($Goal - $StopWatch.Elapsed) -gt 0 ) {
        if ([console]::KeyAvailable) {
            switch -regex ([console]::ReadKey().key) {
                "(Q|C)" { 
                    $Exit = $True 
                    $timer.stop()
                    $StopWatch.Stop()
                    [console]::TreatControlCAsInput = $False
                    #cleanup 
                    Unregister-Event thetimer
                    cls
                    Write-Host "Back to work then..."
                    return
                }
                Default {}
            }
        }
    }
    [console]::TreatControlCAsInput = $False 
    $timer.stop()
    $StopWatch.Stop() 
    #cleanup 
    Unregister-Event thetimer

}

Function Complete-Interruptions {
    param(
        [string[]]$Interruptions
    )

    foreach ($Interruption in $Interruptions) {
        $title = "Review Interruption: $Interruption"
        $message = "What do you want to do with this thought?"
        
        $Forget = New-Object System.Management.Automation.Host.ChoiceDescription "&Forget It", `
            "Nothing will be done."
        
        $Todo = New-Object System.Management.Automation.Host.ChoiceDescription "&Todo", `
            "Outlook task will be created"

        $options = [System.Management.Automation.Host.ChoiceDescription[]]($Forget, $Todo)
        
        $result = $host.ui.PromptForChoice($title, $message, $options, 0) 
        
        switch ($result) {
            0 {
                Write-Host "Forgotten..."
            }
            1 {
                New-OutlookTask -Description $Interruption 
            }
        }
    }
}

Function Add-Interruption {
    param(
        [validateset("Internal", "External")]
        [string]$Type,
        [string[]]$Interruptions
    )
   
    if ($Type -eq "Internal") {
        Read-Host "What's on your mind?" 
    }
    else {
        Read-Host "What happened?"
    }
    
}
