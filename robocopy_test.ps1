#https://www.techrepublic.com/blog/windows-and-office/use-robocopys-multi-threaded-feature-to-quickly-back-up-your-data-in-windows-7/
$src = "C:\Users\Administrator\Downloads"
$Dst = "C:\Users\Administrator\Downloads.BACKUP"

if(!(Test-Path $Dst) ){
  New-Item -ItemType Directory $Dst
}

robocopy $src $Dst /MIR /MT:32





