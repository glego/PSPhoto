Function Get-FileDetail () {
    <# 
    .SYNOPSIS
    Get the additional details of the file using the Shell Application COM Object

    .DESCRIPTION
    Details of a file are 

    .PARAMETER FilePath
    Provide the path to the input file.

    .EXAMPLE
    Get-FileDetail -FilePath ".\20170522_193247.jpg"
  
    .OUTPUTS
    Script will return an array of PS Objects, with all detail properties.   

    .LINK

    .NOTES
    Inspired by
    * https://superwidgets.wordpress.com/2014/08/15/powershell-script-to-get-detailed-image-file-information-such-as-datetaken/
    * https://gallery.technet.microsoft.com/scriptcenter/get-file-meta-data-function-f9e8d804

    
    
#>
#Requires -Version 3
[CmdletBinding()] 
Param(
[Parameter(Mandatory=$true,  Position=0)]
    #[ValidateScript({ (Test-Path -Path $_) })]
    [String[]]$FilePath
)

    BEGIN {
        $Shell = New-Object -COMObject Shell.Application
    }
    PROCESS {
        foreach ($Path in $FilePath) {
            
            # To-do: Shell32 Folder - File - GetDetailsOf,  breaks between OS versions. Therefor it is recommended 
            #   to use the IPropertyStore and it is supported since Windows Vista. 
            #
            # References:
            # * https://stackoverflow.com/questions/14439169/is-there-a-correct-way-to-get-file-details-in-windows-since-getdetailsof-column
            # * https://stackoverflow.com/questions/2265759/how-to-read-the-properties-of-files-using-ipropertystorage/2266316#2266316
            # * https://msdn.microsoft.com/en-us/library/windows/desktop/bb761473(v=vs.85).aspx
            # 

            $Directory      = Split-Path -Path $Path                # Get the parent container  
            $FileName       = Split-Path $Path -Leaf                # Display file names
            $ShellDirectory = $Shell.Namespace($Directory)          # https://msdn.microsoft.com/en-us/library/windows/desktop/bb774085(v=vs.85).aspx
            $ShellFile      = $ShellDirectory.ParseName($FileName)  # https://msdn.microsoft.com/en-us/library/windows/desktop/bb787882(v=vs.85).aspx

            # Iterate through the File Details
            $FileDetails = @() # Collection of file details
            for ($i = 0; $i  -le 266; $i++) {
                $Property = @{} # clear object
                $FileDetail = $null # clear object
                
                $PropertyName = $($ShellDirectory.GetDetailsOf($ShellFile.items, $i)) # https://msdn.microsoft.com/en-us/library/windows/desktop/bb787870(v=vs.85).aspx
                $PropertyValue = $($ShellDirectory.GetDetailsOf($ShellFile, $i))

                # To-do: Skip empty Property values
                # To-do: Check for duplicates (Double hash names will result in error)
                $Property.Add($PropertyName,$PropertyValue)

                $FileDetail = New-Object PSObject -Property $Property
                $FileDetail.PSObject.TypeNames.Insert(0, "Glego.PSPhoto.FileDetail")
                
                $FileDetails += $FileDetail
            }
        }
    }
    END {
        Write-Output $FileDetails
    }
}