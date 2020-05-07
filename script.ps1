$controlssource = 1
$controlsdest = 1

$objshell = New-Object -ComObject Shell.Application
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$objform = New-Object System.Windows.Forms.Form
$objform.Text = "Filter Music"
$objform.Size = New-Object System.Drawing.Size(300,250)
$objform.StartPosition = "CenterScreen"

$objform.KeyPreview = $True
$objform.Add_KeyDown({if ($_.KeyCode -eq "Escape")
{$objform.Close()}})

$filterbutton = New-Object System.Windows.Forms.Button
$filterbutton.Location = New-Object System.Drawing.Size(20,170)
$filterbutton.Size = New-Object System.Drawing.Size(75,23)
$filterbutton.Text = "Filter"
$filterbutton.Add_Click({$x="filter";$objform.Close()})
$objform.Controls.Add($filterbutton)

$cancelbutton = New-Object System.Windows.Forms.Button
$cancelbutton.Location = New-Object System.Drawing.Size(180,170)
$cancelbutton.Size = New-Object System.Drawing.Size(75,23)
$cancelbutton.Text = "Cancel"
$cancelbutton.Add_Click({$a=0;$x="nil";$objform.Close()})
$objform.Controls.Add($cancelbutton)

if ($controlssource -eq 1){
$objlabel2 = New-Object System.Windows.Forms.Label
$objlabel2.Location = New-Object System.Drawing.Size(10,10)
$objlabel2.Size = New-Object System.Drawing.Size(280,15)
$objlabel2.Text = "Please chose the location of your music library :"
$objform.Controls.Add($objlabel2)}

if ($controlssource -eq 1){
$objtextbox = New-Object System.Windows.Forms.TextBox
$objtextbox.Location = New-Object System.Drawing.Size(10,25)
$objtextbox.Size = New-Object System.Drawing.Size(230,20)
$objform.Controls.Add($objtextbox)}

if ($controlssource -eq 1){
$browsebutton1 = New-Object System.Windows.Forms.Button
$browsebutton1.Location = New-Object System.Drawing.Size(250,24)
$browsebutton1.Size = New-Object System.Drawing.Size(26,22)
$browsebutton1.Text = "..."
$browsebutton1.Add_Click({$fold1 = $objshell.BrowseForFolder(0, "Select Folder", 0, "");$objtextbox.Text = $fold1.self.path})
$objform.Controls.Add($browsebutton1)}

$objform.Add_Shown({$objform.Activate()})
[void] $objform.ShowDialog()

$sFolder = $objtextbox.Text
$objfolder = $objshell.namespace($sFolder)
echo "$sFolder"
if ($x -eq "nil") {exit}

foreach ($strfilename in $objfolder.items())
{
    if($strfilename.IsFileSystem){
        for ($a ; $a -le 266; $a++)
        {
            if ($objfolder.getDetailsOf($objfolder.items, $a) -eq "Contributing artists")
            {
                $artist = $objfolder.getDetailsOf($strfilename, $a)
            }
            if($objfolder.getDetailsOf($objfolder.items, $a) -eq "Album")
            {
                $album = $objfolder.getDetailsOf($strfilename, $a)
            }
        }
        if ($artist -and $album)
        {
            $pathfolder = Split-Path -Path $strfilename.Path()
            $directoryArtistPath = $pathfolder + "\" + $artist
            $directoryAlbumPath = $pathfolder + "\" + $artist + "/" + $album 
            if(!(test-path($directoryArtistPath))){
                echo "hereART"
                new-Item $directoryArtistPath -ItemType directory
            }
            if(test-path($directoryArtistPath)){
                if(!(test-path($directoryAlbumPath))){
                    echo "hereALB"
                    new-Item $directoryAlbumPath -ItemType directory
                }
            }
            $path = $pathfolder + "\" + $artist + "\" + $album + "\" + $strfilename.Name()
            if (!(test-path($path)))
            {
                move-item -Path $strfilename.path() -Destination $path
            }
        }
        clear-variable artist
        clear-Variable album
        $a=0
    }
}