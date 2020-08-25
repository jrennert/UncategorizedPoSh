Function Get-Buzzwords {
    return @(
        'Solutionize',
        'Go-Forward Basis',
        'Parking Lot',
        'Leverage',
        'Lets Table That',
        'Synergy',
        'Serverless',
        'Cloud',
        'Streamline',
        'Data-Driven',
        'Blocked (or Blocker)',
        'Deep Dive',
        'Ballpark',
        'Visibility',
        'Return on investment',
        'Ping',
        'Quick Win',
        'Pain Point',
        'Next Gen',
        'Big Data',
        'Bandwidth',
        'On your radar',
        'Out of the Box',
        'Touch base',
        'Circle back',
        'Low-hanging fruit',
        'Agile',
        'Reach out',
        'On the same page',
        'Wheelhouse',
        'Bottom Line',
        '(Digital) Transformation',
        'Buy-in',
        'Single Pane of Glass',
        'All hands on deck',
        'Scalable',
        'Bleeding (or cutting) Edge')
}

Function New-BuzzBingo {
    [cmdletbinding()]
    param(
        [string[]]$Words,
        [switch]$NoFreeSpace
    )
    BEGIN {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.Application]::EnableVisualStyles()

        #Size of each square
        $Height = 150 #Also the width
        #Space between squares
        $Gap = 10

        $Form = New-Object system.Windows.Forms.Form
        $Form.ClientSize = New-Object System.Drawing.Point(810, 810)
        $Form.text = "Buzzword Bingo"
        $Form.TopMost = $false

        $NumWords = 24
        if ($NoFreeSpace) {
            $NumWords += 1
        }
        if ($Words.count -lt $NumWords) {
            throw 'Your list of words is too short.  You need more buzzwords.'
        }
        
        #Choose 24 (or 25 words)
        $i = 0
        $SelectedWords = @()
        while ($i -le $NumWords) {
            #Check if we're on the Center square and if we allowed a free space
            if (($i -ne 12) -or ($NoFreeSpace)){
                #Check if there is more than 1 word in the list before picking a random number
                if ($Words.Count -gt 1){
                    $SelectedWord = $Words[(Get-Random -Minimum 0 -Maximum ($Words.Count - 1))]
                }else{
                    #If there is only one word left pick it
                    $SelectedWord = $Words[0]
                }
                #Remove the Selected word from the pool of options so we don't have duplicates
                $Words = $Words | Where-Object {$_ -ne $SelectedWord}
            }else{
                $SelectedWord = 'FREE SPACE'
            }
            #Add the picked word to the list
            $SelectedWords += $SelectedWord
            $i++
        }
        
    }
    PROCESS {
        $x = $Gap
        $y = $Gap
        $boxNum = 1
        #Loop through all the picked words and make a box and checkbox with the word
        foreach ($Word in $SelectedWords) {
            #Make a panel (the squares)
            $Panel = New-Object system.Windows.Forms.Panel
            $Panel.height = $Height
            $Panel.width = $Height
            $Panel.location = New-Object System.Drawing.Point($x, $y)
            $Panel.BorderStyle = 1
            #Make a checkbox
            $checkbox = New-Object system.Windows.Forms.CheckBox
            $CheckBox.text = $Word
            $CheckBox.AutoSize = $false
            $CheckBox.height = 20
            $CheckBox.width = $Height + $Gap
            if ($Word -eq 'FREE SPACE'){
                $checkbox.checked = $true
            }
            $CheckBox.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
            $CheckBox.location = New-Object System.Drawing.Point(3, 75) #This location is relative to the panel
            
            #Add the checkbox to the panel then add the panel to the form
            $Panel.Controls.AddRange(@($checkbox))
            $Form.Controls.AddRange(@($Panel))
            #Move our x location to the right by the width of the box plus the gap we want between it
            $x += $Height + $Gap
            $boxNum += 1
            #Once we have 5 boxes, move the y location by the height of the box plus the gap and reset the x location (start a new row)
            if ($boxNum -gt 5) {
                $boxNum = 1
                $x = $Gap
                $y += $Height + $Gap
            }
        }
        #Show the form
        $Form.ShowDialog()
    }
    END { }
}

New-BuzzBingo -Words (Get-Buzzwords)