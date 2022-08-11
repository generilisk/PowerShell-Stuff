$List = import-csv "C:\tmp\users.csv"
            
ForEach($User in $List){
     Get-ADUser -filter * -Properties cn,emailaddress | Select-Object cn,emailaddress | export-csv "C:\tmp\users_new.csv" 
     }