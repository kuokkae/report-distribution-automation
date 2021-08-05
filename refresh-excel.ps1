#get timestamp
$timestamp = get-date -format "yyyyMMdd-HHmm"

#logging
start-transcript -path "C:\where\log\$timestamp.txt" <# -append #>

$path = "C:\where"
$file = "file_name.xlsm"
$fp = Join-Path $path $file

# location of shared folder
$to = "C:\where"
$email = "C:\where\email.py"

#BACKUP START
#shows the file name being back-ed up
Write-Output $fp
$compress = @{
  Path = $fp
  CompressionLevel ="Fastest"
  DestinationPath = $path+"\archive\"+$timestamp+" "+$file+".Zip"
}
Compress-Archive @compress
#BACKUP END

#REFRESH START
#start excel
$excel = new-object -comobject excel.application

#LAUNCH EXCEL
$workbook = $excel.workbooks.open( $fp )

#make it visible ( just to check what is happening )
$excel.visible = $true

#set excel display alerts off
$excel.application.displayalerts = $false

#access the application object and run a macro (not applicable for now)
    # $app = $excel.application
    # $app.run( "refresh_dashboard.refresh_dashboard" )
        # start-sleep -s 30
    # $app.run( "refresh_dashboard.refresh_dashboard" )
        # start-sleep -s 30


#REFRESH THE WHOLE WORKBOOK
$workbook.refreshall( )
        start-sleep -s 10

#set excel display alerts back on
$excel.application.displayalerts = $true
    
#CLOSE AND SAVE
$workbook.close( $true )
$excel.quit( )

#popup box to show completion - you would remove this if using task scheduler
#$wshell = new-object -comobject wscript.shell
#$wshell.popup( "operation completed", 0, "Done", 0x1 )

start-sleep -s 1
#REFRESH END

#COPY TO SHARED FOLDER
Copy-Item -Path $fp -Destination $to

#EMAIL NOTIFICATION
python $email

# LOGGING ENDS
stop-transcript

exit