do{
$rnd=Get-Random -minimum 1 -maximum 9
$command='"C:\Program Files (x86)\Winamp\winamp.exe" "E:\videolar\telefon sesleri\Dtmf'+$rnd+'.ogg"'
#echo " $command "
Write-Host "----------------------------------" -foregroundcolor "blue"
Write-Host "----------------------------------" -foregroundcolor "blue"
Write-Host "----------------------------------" -foregroundcolor "blue"
Write-Host "----------------------------------" -foregroundcolor "blue"
Write-Host " şu an çalınan hangisisdir? tekrar dinlemek için +'ya basın" -foregroundcolor "magenta"
iex "& $command"
$iput=Read-Host
if($iput -eq $rnd) 
 {   Write-Host "Bildiniz :) Bravo! " -foregroundcolor "green"}
else
   { Write-Host '"Bilemediniz :( Doğru yanıt ' $rnd ' olacaktı."' -foregroundcolor "Red"}
}until($input -eq 'ESC')