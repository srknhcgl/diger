do{
$iput=Read-Host
$command='"C:\Program Files (x86)\Winamp\winamp.exe" "E:\videolar\telefon sesleri\Dtmf'+$iput+'.ogg"'
#echo " $command "
echo " $iput çalınıyor..."
iex "& $command"
}until($input -eq 'ESC')