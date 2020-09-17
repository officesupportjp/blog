# This script imports old TechNet blog data (html & images)
# e.g. Import.ps1 -Source "C:\Users\ryusukef\Desktop\Blog-20181026" -Destination "C:\Users\ryusukef\Desktop\Blog-20181026\Converted"

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    $Source,
    $Destination
)

function Fix-ImgTag {
    [CmdletBinding()]
    param($line)

    $output = [PSCustomObject]@{
        Line = $null
        Images = New-Object System.Collections.Generic.List[string]
    }

    $matchInfo = Select-String -Pattern "<img [^>]*src=`"(?<src>[^`"]+)`"[^>]+>" -AllMatches -InputObject $line          

    foreach ($imgMatch in $matchInfo.Matches) {

        $imgEntire = $imgMatch.Groups[0].Value
        $imgSrc = $imgMatch.Groups['src'].Value
       
        # If src=http..., then ignore.
        if ($imgSrc -like "http*") {
            continue
        }

        # Replace <img> with "![](***.jpg)"
        $fileName = [Io.Path]::GetFileName($imgSrc)
        $line =  [RegEx]::Replace($line, "$imgEntire", "`n`n![]($fileName)`n")        
        $output.Images.Add($imgSrc)        
    }

    $output.Line = $line    
    $output
}


#
# Main
#

if (-not $Destination) {
    $Destination = Join-Path $Source 'Converted'
}

$files= Get-Item (Join-Path $Path '*.html')


# This is for testing.
#$files = @(
    #"C:\Users\ryusukef\Desktop\Blog-20181026\Autodiscover &#12434;&#25391;&#12426;&#36820;&#12427; &#12456;&#12500;&#12477;&#12540;&#12489; 2.html"
#    "C:\Users\ryusukef\Desktop\Blog-20181026\[Office 365] Outlook &#36215;&#21205;&#26178;&#12398;&#35388;&#26126;&#26360;&#12398;&#12475;&#12461;&#12517;&#12522;&#12486;&#12451;&#35686;&#21578;&#12395;&#12388;&#12356;&#12390.html"
#)


foreach ($postFile in $files) {    
    #[IO.StreamReader]$reader = New-Object System.IO.StreamReader($postFile, [System.Text.Encoding]::UTF8)
    [IO.StreamReader]$reader = New-Object System.IO.StreamReader($postFile)
    [PSCustomObject]$post = @{Title = $null; Date = $null; Content = $null}
    try {
        while ($line = $reader.ReadLine()) {
    
            if ($line -match "class='entry-title'>(?<title>.*)<") {        
                $post.Title = [System.Web.HttpUtility]::HtmlDecode($Matches['title']).Replace('\','-').Replace('/','-').Replace(':','').Replace(';','').Replace('"',"'")
            }
            elseif ($line -match "<time>(?<time>.*)</time>") {
                [DateTime]$date = $Matches['time']
                $post.Date = $date.ToString("yyyy-MM-dd")

            }
            elseif ($line -match "id='content'>") {
                # Found the start of content. read till the end.
                $content = New-Object System.Text.StringBuilder

                while ($true) {
                    $fixResult = Fix-ImgTag $line

                    # Copy image files
                    foreach ($image in $fixResult.Images) {
                        $dest = Join-Path $Destination $post.Title
                        if (-not (Test-Path $dest)){
                            New-Item -ItemType directory $dest | Out-Null
                        }
                                            
                        Copy-Item (Join-Path $Source $image) -Destination $dest
                    }

                    $line = $fixResult.Line
                    if ($line -like "*</body>*") {
                        $content.Append($line) | Out-Null
                        break
                    }
                    else {
                        $content.AppendLine($line) | Out-Null                        
                    }

                    $line = $reader.ReadLine()
                }

                $skipCountTop = "<div id='content'>".Length
                $skipCountBottom = "</div>  </div></body>".Length                
                $post.Content = $content.ToString($skipCount, $content.Length - $skipCountTop - $skipCountBottom)
                break
            }
        }

        # Now write to a md file        
        if (-not (Test-Path $Destination)) {
            New-Item -ItemType directory $Destination | Out-Null
        }
        
        try {
            $file = New-Item -ItemType File -Path (Join-Path $Destination "$($post.Title).md") -Force
        }
        catch {
            write-host ryusukef
        }
        $writer = New-Object System.IO.StreamWriter $file
        $writer.WriteLine("---")
        # Title can have "[" and double-quotation marks. Put it in Folded text block
        # https://learn-the-web.algonquindesign.ca/topics/markdown-yaml-cheat-sheet/
        $writer.WriteLine("title: >")
        $writer.WriteLine("  $($post.Title)")
        $writer.WriteLine("date: $($post.Date)")
        $writer.WriteLine("---")
        $writer.Write($post.Content)

    }
    finally {
        if ($writer) {
            $writer.Dispose()
        }

        if ($reader) {
            $reader.Dispose()
        }        
    }
}


