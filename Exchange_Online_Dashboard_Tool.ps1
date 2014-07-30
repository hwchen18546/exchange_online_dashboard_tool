Function welcome{
        Clear;
        Write-Host "**********************************************"  -foreground Red 
        Write-Host "        Exchange Online Dashboard Tool          " -foreground Red 
        Write-Host "**********************************************"  -foreground Red 
}

Function Login {

        while(1){
                Welcome;
                Import-Module MsOnline;
                Write-Host "step 1" -ForegroundColor yellow;
                Write-Host " Enter Office365 account : " -nonewline;
                $global:adm_account = Read-Host;
                Write-Host "--------------------------------------------------"-ForegroundColor yellow;

                Write-Host "step 2" -ForegroundColor yellow;
                Write-Host " Please enter your password : " -nonewline;
				
				# Login password
				$global:adm_password_plain = Read-Host;
                $global:adm_password_encrypt = convertto-securestring  $adm_password_plain -asplaintext -force;			
				
                $global:adm_cred=New-Object System.Management.Automation.PSCredential($adm_account,$adm_password_encrypt);
                $report = $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $adm_cred -Authentication Basic ¡VAllowRedirection 2>&1;
                $err = $report | ?{$_.gettype().Name -eq "ErrorRecord"}
				                if($err){
					                Write-Host " $report" -background Black -foreground Red;	
					                Read-Host;
                                    Clear;
				                }
				                else{									
									Import-PSSession $Session;									
					                Write-Host "Login Success!" -background Black -foreground Magenta;	
                                    Write-Host "--------------------------------------------------"-ForegroundColor yellow;
					                Break;
				                }								
        }    
} 

#HTML Format
Function Out_HTML($Input_CSV,$Output_HTML,$Title){
	$header_content = "<style>"
	$header_content = $header_content + "Body{background-color:peachpuff;}"
	$header_content = $header_content + "Table{border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}";
	$header_content = $header_content + "Th{border-width: 1px; padding: 0px; border-style: solid; border-color: black; background-color:thistle;width:200px;}";	
	$header_content = $header_content + "Td{border-width: 1px; padding: 0px;border-style: solid; border-color: black; background-color:palegoldenrod;width:200px;}";
	$header_content = $header_content + "</style>"	
	$day = Get-Date;
	$day_format = $day.Year.ToString()+"."+$day.Month.ToString()+"."+$day.Day.ToString()+" "+$day.Hour.ToString()+"-"+$day.Minute.ToString()+".html";
	$Output_HTML+=$day_format;
	Import-Csv $Input_CSV | ConvertTo-HTML -head $header_content  -body $Title | Out-File $Output_HTML
	Invoke-Item $Output_HTML;
	Remove-Item $Input_CSV;		
}

#1
Function MailboxInfo{
   Get-mailbox -resultsize unlimited | Select-Object WindowsLiveID,DisplayName | Out-Gridview -PassThru -Title "Choose the Mailbox User" |ForEach-Object{
        Get-MailboxStatistics -Identity $_.WindowsLiveID | Select-Object Displayname,TotalItemSize,Itemcount,LastLogonTime,lastlogofftime | FT
   }
}

#2
Function MailboxDetailInfo{
    Get-mailbox -resultsize unlimited | Select-Object WindowsLiveID,DisplayName | Out-Gridview -PassThru -Title "Choose the Mailbox User" |ForEach-Object{
         Write-Host "MailboxStatistics-"$_.WindowsLiveID -foreground Green;    
         Get-MailboxStatistics -Identity $_.WindowsLiveID | FL;  
         Write-Host "MailboxObject-"$_.WindowsLiveID -foreground Green; 
         Get-Mailbox -Identity $_.WindowsLiveID |FL;          
    }
}



#3
Function Check_Active_Mailbox_Report{
	#Check active account
	Get-StaleMailboxReport | Out-GridView -Title "Stale Mailbox Report"
}

#4
Function Check_Inactive_Account_Info{
	$maxdata = Read-Host "Enter the Quantity of Data you want to check(unlimited type 0)"; 
	if($maxdata -eq 0){$maxdata = "unlimited";}
	Get-StaleMailboxDetailReport -ResultSize $maxdata |  Out-GridView -Title "StaleMailboxDetailReport";	
}

#5
Function Create_Login_Info_html{

	$TempCSV = "Login_Info.csv";
	#Prepare Output file with headers
	Out-File -FilePath $TempCSV -InputObject "UserPrincipalName,LastLogonDate,LastLogoffDate" -Encoding UTF8
	
	#Can also use "Get-mailbox -resultsize unlimited| Get-MailboxStatistics | select displayname, lastlogontime"
    Get-mailbox -resultsize unlimited | Select-Object WindowsLiveID,DisplayName | Out-Gridview -PassThru -Title "Choose the Mailbox User" | ForEach-Object{
			try{
				$ObjUser = Get-MailboxStatistics -Identity $_.WindowsLiveID;		
				if($ObjUser.LastLogonTime -eq $null){$ObjUser.LastLogonTime = "No Record"}
				if($ObjUser.LastLogoffTime -eq $null){$ObjUser.LastLogoffTime = "No Record"}
				
				#Format Output
				#"{0,-20} {1,-10} {2,-20} {3,-10} {4,-20}" -f  $_.UserPrincipalName,"LastLogonTime: ",
				#$ObjUser.LastLogonTime," LastLogoffTime: ",$ObjUser.LastLogoffTime ;	
				
				$UserDetails = $_.WindowsLiveID+","+$ObjUser.LastLogonTime+","+$ObjUser.LastLogoffTime;
				Out-File -FilePath $TempCSV -InputObject $UserDetails -Encoding UTF8 -append;
			}
			catch{
			}	
	}
	$Title = "<H2>Exchange Online Mailbox Login Report</H2>";
	Out_HTML $TempCSV "Logon Report " $Title;
}

#6
Function Create_StaleMailbox_Login_Info_html{
	$TempCSV = "StaleMailbox_Info.csv";
	Out-File -FilePath $TempCSV -InputObject "UserPrincipalName,Update,DaysInactive,LastLogin,LastLogonDate,LastLogoffDate" -Encoding UTF8
	$maxdata = Read-Host "Enter the Quantity of Data you want to check(unlimited type 0)"; 
	if($maxdata -eq 0){$maxdata = "unlimited";}
	Get-StaleMailboxDetailReport -ResultSize $maxdata | Out-Gridview -PassThru -Title "Choose the Mailbox User" | ForEach-Object{
			try{
				$ObjUser = Get-MailboxStatistics -Identity $_.WindowsLiveID;		
				if($ObjUser.LastLogonTime -eq $null){$ObjUser.LastLogonTime = "No Record"}
				if($ObjUser.LastLogoffTime -eq $null){$ObjUser.LastLogoffTime = "No Record"}
				if($_.LastLogin -eq $null)	{$_.LastLogin = "No Record"}
				$UserDetails = $_.WindowsLiveID+","+$_.Date+","+$_.DaysInactive+","+$_.LastLogin+","+$ObjUser.LastLogonTime+","+$ObjUser.LastLogoffTime;
				Out-File -FilePath $TempCSV -InputObject $UserDetails -Encoding UTF8 -append;
			}
			catch{
			}	
	}
	$Title = "<H2>Exchange Online StaleMailbox Login Report</H2>";
	Out_HTML $TempCSV "StaleLogon Report " $Title;
}

#7
Function Export_Active_CSV{
    $Filename = Read-Host "Export File name";
	Get-StaleMailboxReport | Export-CSV $Filename"_Active.csv" -Encoding UTF8;
    Invoke-Item $Filename"_Active.csv";
}

#8
Function Export_Inactive_CSV{
    $Filename = Read-Host "Export File name";`
	$maxdata = Read-Host "Enter the Quantity of Data you want to check(unlimited type 0)"; 
	if($maxdata -eq 0){$maxdata = "unlimited";}
	Get-StaleMailboxDetailReport -ResultSize $maxdata | Export-CSV $Filename"_Inactive.csv" -Encoding UTF8;
    Invoke-Item $Filename"_Inactive.csv";
}



<# Main #>
	
		Login;

		while (1) {
				welcome;
				#$ident = Get-MsolUser -UserPrincipalName $adm_account ;
				#Write-Host "$ident "   -BackgroundColor Black -ForegroundColor Magenta ;
				Write-Host "Login : $adm_account "   -BackgroundColor Black  -ForegroundColor Magenta ;
                Write-Host " View Stauts"  -foreground Yellow  ;
				Write-Host "  1.Mailbox Info"  
				Write-Host "  2.Mailbox Detail Info"
                Write-Host " Check Report"  -foreground Yellow  ;
				Write-Host "  3.Active Report"  
				Write-Host "  4.Inactive Account Report"
                Write-Host " Export CSV"  -foreground Yellow  ;
				Write-Host "  5.Active Report CSV "
				Write-Host "  6.Inactive Account Report CSV";
                Write-Host " Export HTML" -foreground Yellow  ;   
                Write-Host "  7.Login Info html";
				Write-Host "  8.Inactive Login Info html";
                Write-Host " Logout or Exit" -foreground Yellow  ;  
                Write-Host "  9.Logout";
				Write-Host "  0.Exit";
				
                Write-Host "Please choose the number: " -foreground Yellow -NoNewline;
				$choose = Read-Host;

				switch ($choose) 
				{ 
						1{ MailboxInfo }
						2{ MailboxDetailInfo }
						3{ Check_Active_Mailbox_Report } 
						4{ Check_Inactive_Account_Info }
						5{ Export_Active_CSV}
						6{ Export_Inactive_CSV}
                        7{ Create_Login_Info_html }
						8{ Create_StaleMailbox_Login_Info_html}
						9{ Get-PSSession | Remove-PSSession;
						    Login;
						}
						default { ; }
				}
			    if($choose -ne -0){
                        Write-Host "Press any key to continue" -ForegroundColor Red;
					    Read-Host;
			    }
                else{
						#Remove session
						Get-PSSession | Remove-PSSession				
                        break;
                } 
		}

<# End Main #>