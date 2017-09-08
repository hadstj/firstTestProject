$addr_nod32jihuoma = 'http://www.nod32jihuoma.com/'
$addr_nod32Convert = 'https://my.eset.com/convert/'

function cleanIexplore {
    Get-Process | Where-Object {$_.processname -eq "iexplore"} | Stop-Process
}
# 等待加载
function waitForLoad ($ie, $time = 100) {
    while ( $ie.Busy ){ Start-Sleep -Milliseconds $time }
}

function 打开窗口 ($address)
{
	$ie = New-Object -COM InternetExplorer.Application
	
	$ie.Navigate($address)
	
	waitForLoad ($ie)
	#$ie.Visible = $true
	$doc = $ie.Document
	$win = $ie.Document.parentWindow

	return ($ie, $doc, $win)
}

function get-licensekey
{
    Write-Host 'Open nod32jihuoma window'

    $ie, $doc, $win = 打开窗口 $addr_nod32jihuoma
    Start-Sleep -Milliseconds 2000

    $userAndPass = $doc.getElementById('zq_ser').document.getElementsByTagName('dl')[0].innerText

    if ($userAndPass -match '用户名：(.*?) ')
    { 
    $user = $Matches[1]
    }

    if ($userAndPass -match '密 码：(.*?) ')
    { 
    $pass = $Matches[1]
    }

    Start-Sleep -Milliseconds 1000

    $s_closeWindow = 'window.opner=null;
    window.open("","_self");
    window.close();
    '
    $win.execScript($s_closeWindow,'javascript')
    $win.close()
    $ie=$null
    $win=$null
    $doc=$null

    if ($user -and $pass)
    {
        Write-Host 'Open Convert window'

        $ie, $doc, $win = 打开窗口 $addr_nod32Convert

        $s_convert = '
        $("input#body_txtLicKeyUsrn")[0].value="' + $user +'";
        $("input#body_txtLicKeyPss")[0].value="' + $pass +'";
        $("input#body_btnConvert")[0].click();
        '
        Start-Sleep -Milliseconds 3000

        $win.execScript($s_convert,'javascript')
        Start-Sleep -Milliseconds 1000

        $licensekey = $doc.forms['eset'].document.all('body_pnlLicensekey').getElementsByTagName('span')[0].textContent

        $win.execScript($s_closeWindow,'javascript')
        $win.close()

    }

    Write-Host "$user $pass $licensekey"

    return [PSCustomObject]@{
        'UserName' = $user;
        'Password' = $pass;
        'LicenseKey' = $licensekey;
    }

}