 cls
function Get-OfficeErrorLogs
{
 param (  
   
    $computername = 'localhost',
    $dayrange = 1
    )

    #get envent logs
    $officeerrorlogs = Get-EventLog -ComputerName $computername -LogName Application -Source "Application Error" -After ( get-date). AddDays(-$dayrange ) | Select-Object Message

    $obj = New-Object -TypeName psobject

    $obj | Add-Member -MemberType NoteProperty `
        -name ComputerName -Value $computername
   
    $obj | Add-Member -MemberType NoteProperty `
        -name Days -value $dayrange

    $obj | Add-Member -MemberType NoteProperty `
        -name Message -Value $officeerrorlogs   `


       Write-Output $obj
      


   
}

Get-OfficeErrorLogs -computername localhost -dayrange 6 | Format-Table -AutoSize
 
