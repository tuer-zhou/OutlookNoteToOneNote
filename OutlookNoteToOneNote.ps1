$wshell = New-Object -ComObject wscript.shell;

$csv = Import-Csv 'PATH_TO_CSV'
$text = "NAME_OF_COLUMN"#need to change to name of the first column
$windowname = "onenote for windows 10" #window name of the Program (in this case onenote for windows 10)

$badChars='+','~','^','%','<','(',')','>'

foreach ($note in $csv){
    $wshell.AppActivate($windowname)#focus to a window
    start-sleep -milliseconds 100

    $wshell.SendKeys("^n")#ctrl+n: new page
    Start-Sleep -Milliseconds 100

    $wshell.SendKeys("~")#enter: skips the name for the page
    Start-Sleep -Milliseconds 150

    

    foreach($letter in $note.$text.ToCharArray()){
        #puts some characters into curly brackets because they have other meaning when not e.g. ~ = enter
        #"[" open the opens the feedback menu in onenote 
        if($letter -eq [Environment]::NewLine){
            $letter = "~"
        }elseif($letter -in $badChars){
            $letter = '{'+$letter+'}'
        }elseif($letter -in ("[","]")){
            if($letter -eq "["){
                $letter = "{(}"
            }elseif ($letter -eq "]"){
                $letter = "{)}"
            }
        }
        
        $wshell.SendKeys($letter)
        
        if($letter -eq "~"){
            start-sleep -Milliseconds 50
        }else{
            start-sleep -Milliseconds 1
        }
    }
    Start-Sleep -Milliseconds 100
    
    
}