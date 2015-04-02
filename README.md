############################################
#Author:Lixiaosong
#Email:lixs@ourgame.com;lixiaosong8706@gmail.com
#For:检测Lync照片上传
#Version:1.0
##############################################
$users=Get-ADUser -searchbase "OU=ourgame,DC=GlobalLink,DC=cis,DC=com,DC=cn"  -Properties * -filter * |%{$_.samaccountname}
$userlist=@()
foreach ($user in $users){
        $userphoto=Get-ADUser $user -Properties thumbnailPhoto
        if($userphoto.thumbnailPhoto -like $null){
                $chinesename=Get-ADUser $user  -Properties * | %{$_.Displayname}
                $manager=Get-ADUser $user -Properties manager | %{$_.manager} 
                $manageremail=Get-ADUser $manager -Properties emailaddress | %{$_.emailaddress}  
                echo $manageremail
                $Emailbody=
                "亲爱的 $chinesename 同学：
                     我们检测到到目前为止您还没有在LYNC上传您的照片，请尽快将照片以附件的形式回复此邮件，最好为一寸免冠照，我们会尽快帮您上传。感谢您的支持！"
                
                Send-Mailmessage -from  "it@ourgame.com"  -to "$user@ourgame.com" -cc "$manageremail"  -body $Emailbody -Subject "请尽快上传您的Lync照片" -smtpserver mail.ourgame.com -Encoding ([System.Text.Encoding]::UTF8)
                
                $nophotouser=get-aduser $user -Properties *
                $userobject=New-object psobject
                $userobject | Add-Member -membertype noteproperty -Name 姓名     -value $nophotouser.displayname 
                $userobject | Add-Member -membertype noteproperty -Name 账户名   -Value $nophotouser.samaccountname
                $userobject | Add-Member -membertype noteproperty -Name 部门     -Value $nophotouser.department
                $userlist+=$userobject
             }
}

$EmailbodyHTML=$userlist|sort-object 姓名 |ConvertTo-Html |Out-String
Send-Mailmessage -from  "it@ourgame.com"  -to "Group.LanOper@ourgame.com"   -bodyashtml $EmailbodyHTML  -Subject "管理员通知" -smtpserver mail.ourgame.com -Encoding ([System.Text.Encoding]::UTF8)
