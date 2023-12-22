@start powershell.exe -windowstyle hidden -noprofile "$me = '%~f0';. ([scriptblock]::create((gc -li $me|select -skip 1|out-string)))" %*&goto:eof
<#
.SYNOPSIS
クリップボードを監視し、ビットマップ画像がコピーされたときに自動的に保存します。
.DESCRIPTION
クリップボードを監視し、ビットマップ画像がコピーされたときに自動的に保存します。
画像の保存先は、既定ではスクリプト ファイルと同じフォルダーに保存されます。スクリプト実行時の引数で指定するか、XML 設定ファイルで保存先を指定できます。XML 設定ファイルについては、-Full を指定してヘルプを表示してください。

「テキスト」と「画像」の両方を含むデータがコピーされたときは、保存しません。
例えば Excel でセルをコピーしたとき、テキストの描かれた図がコピーされますが、同時に書式付きテキストや書式なしテキスト、Excel 形式のセル参照などもコピーされ、貼り付けする側はそれぞれの形式のデータを取り出することができます。
このように、「画像」と「テキスト」の両方を含む情報は保存しません。
また、ベクトル図（メタファイル形式）も保存しません。

collectbitmap.ps1xml ファイルを作成し、動作をカスタマイズできます。
PS1XML 定義ファイルは、例えば PowerShell にて次の要領で作成できます。
PS> @{
>> SavePath = {Split-Path $me -Parent}
>> FileName = {'{1:yyyyMMdd_HHmmssff}_{0}.png' -f $env:COMPUTERNAME, $captureddatetime}
>> Printing       = $true
>> PrintingFont   = 'Consolas'
>> PrintingSize   = 75
>> PrintingString = {"{1:d} {1:HH:mm:ss.ff}`r`n{0}" -f $env:COMPUTERNAME, $captureddatetime}
>> } | Export-CliXml collectbitmap.ps1xml
SavePath には保存先フォルダーを返すスクリプト ブロックを指定します。これは起動時に一度だけ評価されます。
FileName には保存する画像ファイル名を返すスクリプト ブロックを指定します。これは図を保存する毎に評価されます。
いずれかの値を省略することもできます。省略した値はスクリプトの規定値が用いられます。
.NOTES
Bitmap collector batch version 1.00

MIT License

Copyright (c) 2023 Isao Sato

Permission is hereby granted, free of charge, to any person obtaining a
copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be included
in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

################################################################
# bitmap collector batch
################################
# 2023/12/22
################################################################

param($SaveFolder)

Set-StrictMode -Version 2
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

try {
filter Verify-AuthenticodeSignature([Parameter(Mandatory=$true, ValueFromPipeline=$true)] [string] $LiteralPath, [switch] $Force) {
    [bool] $Result = $false
    $exception = $null
    $private:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
    
    $cert = Get-AuthenticodeSignature -LiteralPath $LiteralPath
    
    if($cert -eq $null) {
        throw (New-Object System.ArgumentException ('ファイル {0} はデジタル署名を検証できません。。' -f $LiteralPath), (New-Object System.UnauthorizedAccessException))
    }
    
    switch($cert) {
        {$cert.Status -eq [System.Management.Automation.SignatureStatus]::Valid} {
            if((Test-Path (Join-Path cert:\CurrentUser\TrustedPublisher ($cert.SignerCertificate.Thumbprint)))) {
                $Result = $true
            } else {
                if((Test-Path (Join-Path cert:\LocalMachine\TrustedPublisher ($cert.SignerCertificate.Thumbprint)))) {
                    $Result = $true
                } elseif((Test-Path (Join-Path cert:\CurrentUser\TrustedPublisher ($cert.SignerCertificate.Thumbprint)))) {
                    $Result = $true
                } else {
                    $exception = New-Object System.Management.Automation.PSSecurityException ('ファイル {0} のデジタル署名の発行元は信頼されてません。このスクリプトはシステムで実行されません。' -f $LiteralPath), (New-Object System.UnauthorizedAccessException)
                }
            }
        }
        {$cert.Status -eq [System.Management.Automation.SignatureStatus]::NotSigned} {
            $exception = New-Object System.Management.Automation.PSSecurityException ('ファイル {0} はデジタル署名されていません。このスクリプトはシステムで実行されません。' -f $LiteralPath), (New-Object System.UnauthorizedAccessException)
        }
        {$cert.Status -eq [System.Management.Automation.SignatureStatus]::UnknownError} {
            $exception = New-Object System.ArgumentException ('ファイル {0} はデジタル署名を検証できません。。' -f $LiteralPath), (New-Object System.UnauthorizedAccessException)
        }
        {$cert.Status -eq [System.Management.Automation.SignatureStatus]::NotSupportedFileFormat} {
            $exception = New-Object System.ArgumentException ('ファイル {0} はデジタル署名を検証できません。。' -f $LiteralPath), (New-Object System.UnauthorizedAccessException)
        }
        default {
            $exception = New-Object System.Management.Automation.PSSecurityException ('ファイル {0} はデジタル署名されていますが無効です。このスクリプトはシステムで実行されません。' -f $LiteralPath), (New-Object System.UnauthorizedAccessException)
        }
    }
    if(-not ($exception -eq $null -or $Force)) {
        throw $exception
    }
    
    $Result
}

filter Verify-ScriptExecution([Parameter(Mandatory=$true, ValueFromPipeline=$true)] [string] $LiteralPath, [switch] $Force) {
    $private:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
    switch((Get-ExecutionPolicy)) {
        {$_ -eq [Microsoft.PowerShell.ExecutionPolicy]::Unrestricted} {
            $true
        }
        {$_ -eq [Microsoft.PowerShell.ExecutionPolicy]::Bypass} {
            $true
        }
        {$_ -eq [Microsoft.PowerShell.ExecutionPolicy]::RemoteSigned} {
            if(([uri] $LiteralPath).IsUnc) {
                Verify-AuthenticodeSignature $LiteralPath -Force:$Force
            }
        }
        {$_ -eq [Microsoft.PowerShell.ExecutionPolicy]::AllSigned} {
            Verify-AuthenticodeSignature $LiteralPath -Force:$Force
        }
        default {
            if(-not $Force) {
                throw New-Object System.Management.Automation.PSSecurityException ('スクリプトの実行がシステムで無効になっているため、ファイル {0} を読み込めません。' -f $LiteralPath), (New-Object System.UnauthorizedAccessException)
            }
            $false
        }
    }
}




################################
# the major logic
################################

function private:Enter-BitmapCapture([System.Collections.Hashtable] $xconf) {
    
    # creating a full path for saving pictures
    
    function Get-SavePath
    {
        Join-Path $xconf['SavePath'] (Invoke-Command ([scriptblock]::Create($xconf['FileName'])))
    }

    # responsing to the event
    
    function Watch-Clipboard_OnClipboardChanged
    {
        $captureddatetime = [datetime]::Now
        [System.Windows.Forms.IDataObject] $dt = [System.Windows.Forms.Clipboard]::GetDataObject()
        if($dt.GetDataPresent([System.Windows.Forms.DataFormats]::Bitmap) -and -not $dt.GetDataPresent([System.Windows.Forms.DataFormats]::Text) -and -not $dt.GetDataPresent([System.Windows.Forms.DataFormats]::MetafilePict))
        {
            $disposepict = $null
            if($pict.Image -ne $null) {
                $disposepict = $pict.Image
                $pict.Image = $null
                $disposepict.Dispose()
            }
            
            $bmp = $dt.GetImage()
            $pict.Image = New-Object System.Drawing.Bitmap $bmp
            $pictsize = $bmp.Size
            $bmp.Dispose()
            
            if($check.Checked) {
                $printingstring = (Invoke-Command ([scriptblock]::Create($xconf['PrintingString'])))
                $fontsize = $xconf['PrintingSize']
                # $fontsize = [Math]::Min($fontsize, $pict.Image.Width /15)
                # $fontsize = [Math]::Min($fontsize, $pict.Image.Height /3)
                $grp = $null
                $fontfamily = $null
                $font = $null
                $pen = $null
                $drawpath = $null
                try {
                    $grp = [System.Drawing.Graphics]::FromImage($pict.Image)
                    $fontfamily = New-Object System.Drawing.FontFamily $xconf['PrintingFont']
                    $font = New-Object System.Drawing.Font $fontfamily, $fontsize
                    $measuredsize = $grp.MeasureString($printingstring, $font)
                    $fontsizescale = 1.0
                    if($measuredsize.Width -gt $pictsize.Width) {
                        $fontsizescale = [Math]::Min($fontsizescale, ($pictsize.Width / $measuredsize.Width))
                    }
                    if($measuredsize.Height -gt $pictsize.Height) {
                        $fontsizescale = [Math]::Min($fontsizescale, ($pictsize.Height / $measuredsize.Height))
                    }
                    $fontsize = [int] ($fontsize * $fontsizescale)
                    $pen = New-Object System.Drawing.Pen ([System.Drawing.Brushes]::Black), 4
                    $drawpath = New-Object System.Drawing.Drawing2D.GraphicsPath
                    $drawpath.AddString(
                        $printingstring,
                        $fontfamily,
                        [int][System.Drawing.FontStyle]::Regular,
                        $fontsize,
                        (New-Object System.Drawing.Point 0, 0),
                        ([System.Drawing.StringFormat]::GenericDefault))
                    $grp.DrawPath($pen, $drawpath)
                    $grp.FillPath([System.Drawing.Brushes]::White, $drawpath)
                }finally{
                    if($drawpath){$drawpath.Dispose()}
                    if($pen){$pen.Dispose()}
                    if($font){$font.Dispose()}
                    if($fontfamily){$fontfamily.Dispose()}
                    if($grp){$grp.Dispose()}
                }
            }
            
            $path = Get-SavePath
            
            $mimetype = "image/png"
            $encparams = $null
            # 画質80/100の JPEG 画像にする場合の例
            # $mimetype = "image/jpeg"
            # $encparams = New-Object System.Drawing.Imaging.EncoderParameters -ArgumentList 1
            # $encparams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter -ArgumentList @([System.Drawing.Imaging.Encoder]::Quality, [System.Int64] 80)
            
            $codecinfo = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object {$_.MimeType -eq $mimetype} | Select-Object -First 1
            $pictext = [System.IO.Path]::GetExtension($codecinfo.FilenameExtension.Split(';')[0])
            
            $pict.Image.Save(
                [System.IO.Path]::ChangeExtension($path, $pictext),
                $codecinfo,
                $encparams)
        }
    }
    
    # main
    
    $check = New-Object System.Windows.Forms.CheckBox
    $check.Text = 'テキスト情報をプリントする'
    $check.Dock = [System.Windows.Forms.DockStyle]::Top
    $check.BackColor = [System.Drawing.Color]::Transparent
    $check.Checked = $xconf['Printing']
    
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "コピーしたビットマップ画像を保存します。`n保存先：`n" +$xconf['SavePath']
    $label.Dock = [System.Windows.Forms.DockStyle]::Fill
    $label.BackColor = [System.Drawing.Color]::Transparent
    
    $pict = New-Object System.Windows.Forms.PictureBox
    $pict.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $pict.Dock = [System.Windows.Forms.DockStyle]::Fill
    
    $pict.Controls.Add($check)
    $pict.Controls.Add($label)
    
    $watcher = New-Object ClipboardWatcher
    $watcher.Text = "ビットマップ収集"
    $watcher.Controls.Add($pict)
    $watcher.Add_ClipboardChanged(${function:Watch-Clipboard_OnClipboardChanged})
    
    [System.Windows.Forms.Application]::Run($watcher)
}


################################
# definitions handlers for window messages
################################

if([psobject].Assembly.GetType('System.Management.Automation.TypeAccelerators')::Get['ClipboardWatcher'] -eq $null) {
    Add-Type -ReferencedAssemblies System.Windows.Forms -TypeDefinition @"
    using System;
    using System.Windows.Forms;
    using System.Runtime.InteropServices;
    
    namespace NASsystems.ClipboardWatching
    {
        public class ClipboardWatcher : Form
        {
            public ClipboardWatcher()
            {
                this.HandleCreated += new EventHandler(this.ClipboardWatcher_OnHandleCreated);
                this.HandleDestroyed += new EventHandler(this.ClipboardWatcher_OnHandleDestroyed);
            }
            
            public event EventHandler ClipboardChanged;
            
            protected void JoinClipboardChain()
            {
                try
                {
                    AddClipboardFormatListener(this.Handle);
                }
                catch
                {
                    // 例外時は AddClipboardFormatListener が存在しなかった（≒ NT60 以前）と仮定する。
                    nextHandle = SetClipboardViewer(this.Handle);
                }
            }
            
            protected void DefectClipboardChain()
            {
                try
                {
                    RemoveClipboardFormatListener(this.Handle);
                }
                catch
                {
                    // 例外時は RemoveClipboardFormatListener が存在しなかった（≒ NT60 以前）と仮定する。
                    bool sts = ChangeClipboardChain(this.Handle, nextHandle);
                }
            }
            
            protected override void WndProc(ref Message msg)
            {
               switch(msg.Msg)
               {
                case WM_CLIPBOARDUPDATE:
                    // for NT6.0 or later
                    RaiseClipboardChanged();
                    break;
                case WM_DRAWCLIPBOARD:
                    // for earlier than NT6.0
                    RaiseClipboardChanged();
                    if(nextHandle != IntPtr.Zero)
                        SendMessage(nextHandle, msg.Msg, msg.WParam, msg.LParam);
                    return;
                case WM_CHANGECBCHAIN:
                    // for earlier than NT6.0
                    if(msg.WParam == nextHandle)
                    {
                        nextHandle = (IntPtr)msg.LParam;
                    }
                    else
                    {
                        if(nextHandle != IntPtr.Zero)
                            SendMessage(nextHandle, msg.Msg, msg.WParam, msg.LParam);
                    }
                    return;
                }
                base.WndProc(ref msg);
            }
            
            protected const int WM_CLIPBOARDUPDATE = 0x031D;
            protected const int WM_DRAWCLIPBOARD   = 0x0308;
            protected const int WM_CHANGECBCHAIN   = 0x030D;
            
            [DllImport("user32.dll", SetLastError=true)]
            protected static extern bool AddClipboardFormatListener(IntPtr hwnd);
            
            [DllImport("user32.dll", SetLastError=true)]
            protected static extern bool RemoveClipboardFormatListener(IntPtr hwnd);
            
            [DllImport("user32")]
            protected static extern IntPtr SetClipboardViewer(IntPtr hWndNewViewer);
            
            [DllImport("user32")]
            protected static extern bool ChangeClipboardChain(IntPtr hWndRemove, IntPtr hWndNewNext);
            
            [DllImport("user32")]
            protected extern static int SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);
            
            private IntPtr nextHandle;
            
            private void ClipboardWatcher_OnHandleCreated(object sender, EventArgs args)
            {
                JoinClipboardChain();
            }
            
            private void ClipboardWatcher_OnHandleDestroyed(object sender, EventArgs args)
            {
                DefectClipboardChain();
            }
            
            private void RaiseClipboardChanged()
            {
                if(ClipboardChanged != null)
                    ClipboardChanged(this, new EventArgs());
            }
        }
    }
"@
    [psobject].Assembly.GetType('System.Management.Automation.TypeAccelerators')::Add('ClipboardWatcher',[NASsystems.ClipboardWatching.ClipboardWatcher])
}


################################
# entry
################################

# スクリプト ファイルと同名の PS1XML ファイルがあったら読み込む
# 設定ファイルが存在しない場合、空の構成情報を構築する。
[System.IO.Path]::ChangeExtension($me, '.ps1xml') |% {
    if(Test-Path $_) {
        Verify-ScriptExecution $_ | Out-Null
        $xconf = Import-CliXml $_
    } else {
        $xconf = @{}
    }
}

# $xconf の未定義の各要素に既定の構成情報を構築する。
if($null -eq $xconf['SavePath']) {
    $xconf['SavePath'] = {Split-Path $me -Parent}
}

if($null -eq $xconf['FileName']) {
    $xconf['FileName'] = {'{1:yyyyMMdd_HHmmssff}_{0}.png' -f $env:COMPUTERNAME, $captureddatetime}
}

if($null -eq $xconf['PrintingString']) {
    $xconf['PrintingString'] = {"{1:d} {1:HH:mm:ss.ff}`r`n{0}" -f $env:COMPUTERNAME, $captureddatetime}
}

if($null -eq $xconf['PrintingFont']) {
    $xconf['PrintingFont'] = 'Consolas'
}

if($null -eq $xconf['PrintingSize']) {
    $xconf['PrintingSize'] = 75
}

if($null -eq $xconf['Printing']) {
    $xconf['Printing'] = $false
}


# 保存先パスを解決する
if($null -eq $SaveFolder -or [string]::IsNullOrEmpty($SaveFolder.ToString())) {
    $xconf['SavePath'] = Invoke-Command ([scriptblock]::Create($xconf['SavePath']))
} else {
    $xconf['SavePath'] = $SaveFolder.ToString()
}

$xconf['SavePath'] |? {-not(Test-Path $_)} |% {mkdir $_} |% {'保存先 {0} を作成しました。' -f $_.FullName}

# 主論理呼び出し
# STA の実行環境を構築して主論理を呼び出す。

$is = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$is.ApartmentState = [System.Threading.ApartmentState]::STA
$is.Variables.Add((New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry 'me', $me, 'the script filename', Constant))

$rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($host, $is)
$rs.ApartmentState = [System.Threading.ApartmentState]::STA
$rs.Open() | Out-Null

$ps = [System.Management.Automation.PowerShell]::Create()
$ps.Runspace = $rs

$ps.AddScript(${function:Enter-BitmapCapture}) | Out-Null
$ps.AddArgument($xconf) | Out-Null

$ps.Invoke()
$ps.Streams.Error
$ps.Dispose()
} catch {
    try {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($me)
    } catch {
        $name = 'collectbitmap'
    }
    [System.Windows.Forms.MessageBox]::Show($_.ToString(), $name)
}
