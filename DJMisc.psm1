using namespace System.Collections;
using namespace System.Text.RegularExpressions;
using namespace System.Management.Automation;
#requires -Module DJHash

Function Convert-DictionarytoText
{
    param(
        [System.Collections.IDictionary]
        $Dictionary
    )

    -join (($Dictionary|ConvertTo-Json -Depth 99) -split "`n" | Foreach-Object {$_.Trim() -replace ':  ', ' : ' -replace ",$", ", "})
}

function Get-DigitSum ([int64]$n)
{
    Return ($n.Tostring() -split "(\d)" -ne "" | Measure-Object -Sum).Sum
}

function Get-ArrayDigitSum ([int[]]$IntArray)
{
    $Return = 0
    ForEach ($n in $IntArray)
    {
        $Return += Get-DigitSum $n
    }

    $Return
}

function Get-PrimeFactorDigitSum
{
    [CmdletBinding()]
    Param(
        # Parameter help description
        [Parameter(Mandatory, ParameterSetName = "SpecifyStartEnd")]
        [int]
        $start,

        [Parameter(Mandatory, ParameterSetName = "SpecifyStartEnd")]
        [int]
        $end,

        [Parameter(Mandatory)]
        [int]
        $target,

        # Parameter help description
        [Parameter(Mandatory, ParameterSetName = "SpecifyList")]
        [int[]]
        $List,

        [Parameter(Mandatory = $false, ParameterSetName = "SpecifyStartEnd")]
        [System.Collections.Generic.Dictionary[[int], [int]]]
        $LocalMaximumTable,

        [switch]
        $NoClear,

        [Parameter(Mandatory = $false, ParameterSetName = "SpecifyStartEnd")]
        [int]
        $PrepopulateMaxSoFar = 0

    )

    Begin
    {
        If (!$NoClear) {$LocalMaximumTable.Clear()}
        If ($List)
        {
            $Measure = $List | Measure-Object -Maximum -Minimum
            $Top = $Measure.Maximum
            $Bottom = $Measure.Minimum
            $Range = $List.Count
        }
        else
        {
            $Top = [math]::Max($end, $start)
            $Bottom = [System.Math]::min($end, $start)
            $Range = $Top - $Bottom
        }
        $MaxSoFar = $PrepopulateMaxSoFar
    }

    Process
    {
        If (!$List)
        {
            $List = ($Bottom..$Top)
        }

        $i = 0
        ForEach ($number in $List)
        {
            Write-Progress -Activity "Getting Factors (MaxsoFar: $MaxSoFar)" -Status $number -PercentComplete ($i++ / $Range) * 100 -ID 0

            $PrimeFactors = (Get-PrimeFactorsRecursive $number)
            $DigitSum = Get-ArrayDigitSum $PrimeFactors

            If ($LocalMaximumTable -And $DigitSum -gt $MaxSoFar)
            {
                Try
                {
                    $LocalMaximumTable.Add($number, $DigitSum)
                }
                Catch
                {
                    Write-Warning "Error adding $number . Current Value: $($LocalMaximumTable[$number])"
                }
            }

            $MaxSoFar = [math]::Max($DigitSum, $MaxSoFar)
            $obj = [pscustomobject]@{
                Number            = $number
                DigitSum          = $DigitSum
                PrimeFactorsArray = $PrimeFactors
            }
            If ($obj.DigitSum -eq $target -And @($obj.PrimeFactorsArray).Count -gt 1)
            {
                $obj
            }
        }
    }
    
}

function Get-PrimeFactorSum
{
    [CmdletBinding()]
    Param(
        # Parameter help description
        [Parameter(Mandatory, ParameterSetName = "SpecifyStartEnd")]        
        [int]
        $start,

        [Parameter(Mandatory, ParameterSetName = "SpecifyStartEnd")]
        [int]
        $end,

        [Parameter(Mandatory)]
        [int]
        $target,

        # Parameter help description
        [Parameter(Mandatory, ParameterSetName = "SpecifyList")]
        [int[]]
        $List,

        [Parameter(Mandatory = $false, ParameterSetName = "SpecifyStartEnd")]
        [System.Collections.Generic.Dictionary[[int], [int]]]
        $LocalMaximumTable,

        [switch]
        $NoClear,

        [Parameter(Mandatory = $false, ParameterSetName = "SpecifyStartEnd")]
        [int]
        $PrepopulateMaxSoFar = 0

    )

    Begin
    {
        If (!$NoClear) {$LocalMaximumTable.Clear()}
        If ($List)
        {
            $Measure = $List | Measure-Object -Maximum -Minimum
            $Top = $Measure.Maximum
            $Bottom = $Measure.Minimum
            $Range = $List.Count
        }
        else
        {
            $Top = [math]::Max($end, $start)
            $Bottom = [System.Math]::min($end, $start)
            $Range = $Top - $Bottom
        }
        $MaxSoFar = $PrepopulateMaxSoFar
    }

    Process
    {
        If (!$List)
        {
            $List = ($Bottom..$Top)
        }

        $i = 0
        ForEach ($number in $List)
        {
            Write-Progress -Activity "Getting Factors (MaxsoFar: $MaxSoFar)" -Status $number -PercentComplete ($i++ / $Range) * 100 -ID 0

            $PrimeFactors = (Get-PrimeFactorsRecursive $number)
            $Sum = ($PrimeFactors | Measure-Object -Sum).Sum

            If ($LocalMaximumTable -And $Sum -gt $MaxSoFar)
            {
                Try
                {
                    $LocalMaximumTable.Add($number, $Sum)
                }
                Catch
                {
                    Write-warning "Error adding $number . Current Value: $($LocalMaximumTable[$number])"
                }
            }

            $MaxSoFar = [math]::Max($Sum, $MaxSoFar)
            $obj = [pscustomobject]@{
                Number            = $number
                Sum               = $Sum
                PrimeFactorsArray = $PrimeFactors
            }
            If ($obj.DigitSum -eq $target -And @($obj.PrimeFactorsArray).Count -gt 1)
            {
                $obj
            }
        }
    }
    
}

Function Get-IndicesOf ($Array, [object[]]$Value)
{
    $i = 0
    $LastIndex = (@(ForEach ($Val in $Value)
            {
                [array]::LastIndexOf($Array, $val)
            }) | Measure-Object -Maximum).Maximum

    ForEach ($el in $Array)
    {
        ForEach ($val in $Value)
        {
            if ($el.Equals($val)) { $i }
        }
        ++$i
        if ($i -gt $LastIndex) { break }
    }
}

function Test-FileDirIsPWD ([string]$FileName)
{
    $IsPathRooted = [io.path]::IsPathRooted($FileName)
    $FullPath = [io.path]::GetFullPath($(
            if (!$IsPathRooted) { [IO.Path]::Combine($PWD, $FileName)}
            else {$FileName}
        ) )
    if (!$IsPathRooted)
    {
        return [io.path]::GetFullPath(( [io.Path]::Combine($PWD, ([io.path]::GetDirectoryName($FileName)) ) )) -eq $PWD
    }
    
    return [io.path]::GetDirectoryName($FullPath) -eq $PWD
}

function Get-NoteProperty
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [System.Object]
        $Object,

        #List of headers not to include in the output
        [string[]]
        $ExcludeProperty
    )

    Begin
    {
        $NoteProperty = [System.Management.Automation.PSMemberTypes]::NoteProperty
    }

    Process
    {
        $AL = New-Object System.Collections.ArrayList
        ForEach ($header in ($Object | Get-Member | Where-Object {$_.MemberType -eq $NoteProperty}).Name)
        {
            if ($ExcludeProperty -notcontains $header)
            {
                [void]$AL.Add($header)
            }
        }

        if ($AL.Count -eq 0)
        {return $null}
        return $AL.ToArray()
    }
}

function Get-CharNum
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [string[]]
        $String
    )

    Process
    {
        ForEach ($str in $String)
        {
            $str.ToCharArray() | ForEach-Object {
                [pscustomobject][ordered]@{
                    Char = $_;
                    Code = [int]$_;
                    Hex  = "{0:X2}" -f [int]$_
                }
            }
        }
    }
    
}

Function Test-IsPath([string]$Path)
{
    try
    {
        return [IO.Path]::GetDirectoryName($Path).Length -gt 0
    }
    catch
    {
        return $false
    }
}

Function Split-String ([string]$String)
{
    <#
    .SYNOPSIS
    Split an input string into space delimited tokens
    
    .DESCRIPTION
    Takes advantage of the built in tokeniser to extract tokens from a supplied input string.
    Any text wrapped in quotes will be treated as a single token
    
    .EXAMPLE
    Split-String '"First String" Second Third'
    
    Will output an array of:
    [0]First String
    [1]Second
    [2]Third
    
    .NOTES
    Commas will also be treated as delimiters and removed from the output
    #>    
    Function SplitTokens { return $args }
    Invoke-Expression "SplitTokens $String"
}

function Convert-PSObjectToHashtable
{
    param (
        [Parameter(ValueFromPipeline)]
        $InputObject
    )

    process
    {
        if ($null -eq $InputObject) { return $null }

        if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string])
        {
            $collection = @(
                foreach ($object in $InputObject) { Convert-PSObjectToHashtable $object }
            )

            Write-Output -NoEnumerate $collection
        }
        elseif ($InputObject -is [psobject])
        {
            $hash = @{}

            foreach ($property in $InputObject.PSObject.Properties)
            {
                $hash[$property.Name] = Convert-PSObjectToHashtable $property.Value
            }

            $hash
        }
        else
        {
            $InputObject
        }
    }
}

Function Get-Params ($text)
{
    <#
    .SYNOPSIS
        Convert a single string into a list of parameter names and values
    
    .DESCRIPTION
        Long description
    
    .PARAMETER text
        The input text string to process
    
    .EXAMPLE
        Get-Params "-login mylogin -FullName 'FirstName LastName'"
        returns a hashtable as per the below
        Name      Value
        ----      ---- -
        login     mylogin
        FullName  FirstName LastName
    #>
    
    
    $split = Split-String $text
    
    $argtable = @(ForEach ($s in $split)
        {
            [PSCustomObject]@{string = $s; IsParam = $s[0] -eq '-' }
        })
    
    $IsParam = $false
    $ParamName = [string]::Empty
    $Params = @{ }
    for ($i = $UnknownParam = 0; $i -lt $argtable.Count; $i++)
    {
        If ($argtable[$i].string -match "^-")
        {
            $IsParam = $true
            $ParamName = ($argtable[$i].string -replace "^-")
            If (!$Params.ContainsKey($ParamName))
            {
                $Params.Add($ParamName, $null)
            }
            Continue
        }
        If ($IsParam)
        {
            $Params.$ParamName = $argtable[$i].string
            $IsParam = $false
            $ParamName = [string]::Empty
            Continue
        }
        if ($ParamName -eq [string]::Empty)
        {
            $ParamName = "UnknownParam_{0}" -f $UnknownParam++
            $Params.$ParamName = $argtable[$i].string
            $ParamName = [string]::Empty
            Continue
        }
        
        $Params.$ParamName = $argtable[$i].string
    
    }
    
    $Params
}

function CompleteTypeName ([string]$TypeString)
{
    $CompletionMatches = [System.Management.Automation.CommandCompletion]::CompleteInput("[$TypeString", $TypeString.Length + 1, @{}).CompletionMatches

    If ($CompletionMatches.Length -gt 0)
    {
        return $CompletionMatches[0].CompletionText
    }
}

function Get-InterfaceTree
{
    <#
    .SYNOPSIS
        Displays a tree of interfaces for a given type
    .DESCRIPTION
        Recursively runs the "GetInterfaces() method on a type and its interfaces to build a picture of all the interfaces that are applied to a type"
    .EXAMPLE
        PS C:\> Get-InterfaceTree hashtable
        
        Hashtable
            IDictionary
                ICollection
                    IEnumerable
                IEnumerable
            ICollection
                IEnumerable
            IEnumerable
            ISerializable
            IDeserializationCallback
            ICloneable
    
    .Example
        PS C:\>Get-InterfaceTree ([System.Collections.Generic.ICollection[string]]) -ShowFullTypeName
        
        System.Collections.Generic.ICollection`1[[System.String]]
            System.Collections.Generic.IEnumerable`1[[System.String]]
                System.Collections.IEnumerable
            System.Collections.IEnumerable

        This example shows use of the ShowFullTypeName switch parameter to show all types including namespaces

    .Example
        PS C:\>Get-InterfaceTree ([System.Collections.Generic.IEnumerable[string]]) -ShowFullTypeName -NoTrimGenericTypeNames
        
        System.Collections.Generic.IEnumerable`1[[System.String, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089]]
            System.Collections.IEnumerable

        This example shows use of the "NoTrimGenericTypeNames" switch parameter when querying a generic type to show the full type details.
        This parameter will only have an effect when used with the "ShowFullTypeName" switch parameter.
    .Parameter Type
        The type, or name of a type on which to get the interface tree. If a string is supplied, the text will be attempted to be matched to a type name. e.g. "hashtab" would resolve to [hashtable]
    .Parameter ShowFullTypeName
        Forces the output to include the full type name in place of the standard type name, e.g. System.Collections.IDictionary instead of "IDictionary"
    .Parameter NoTrimGenericTypeNames 
        By default, this function will trim full generic type information ( e.g. [System.Type, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089] will become [System.Type] )
        Specify this switch parameter to show the full details on generics.
    
    .INPUTS
        A type object or string representation of a type
    .OUTPUTS
        A collection of strings representing the interface name and nesting position within the tree
    .NOTES
        Author: David Johnson
        To pass a type to the -Type parameter, enclose the type name in parentheses, e.g. "-Type ([string])"
    #>
    [cmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory, ValueFromPipeline)]
        [ValidateScript( {$_ -is [type] -or $_ -is [string]})]
        $Type,

        [switch]
        $NoTrimGenericTypeNames,

        [switch]
        $ShowFullTypeName,

        [int]
        $Indent = 0
    )
    Begin
    {
        #$xlr = [psobject].Assembly.GetType("System.Management.Automation.TypeAccelerators")
        function CleanTypeString
        {
            Param(
                # The string on which to do the work
                [Parameter(Mandatory, ValueFromPipeline)]
                [string]$String,
                
                # Skip processing
                [switch]$DoNothing
            )
            if ($DoNothing)
            {
                return $String
            }

            $StringArray = $String -split '\['
            
            $regex = '([\w.\[\]]+?), (.+?), (Version=\d+\.\d+\.\d+\.\d+), (Culture=.+?), (PublicKeyToken=[0-9a-f]+)\]'

            If ($string -notmatch $regex)
            {
                return $String
            }
            
            for ($i = 0; $i -lt $Stringarray.Count; $i++)
            {
                while ($StringArray[$i] -match $regex)
                {
                    $StringArray[$i] = ($StringArray[$i] -replace $regex, '$1]')
                }
            }
            return (CleanTypeString ($StringArray -join '['))
        }
    }

    Process
    {
        $Name = "Name"
        If ($ShowFullTypeName)
        {
            $Name = "FullName"
        }
        
        If ($Type -is [String])
        {
            # Attempt to translate the string to a type object
            try
            {
                $type = [type]$Type
            }
            catch
            {
                $InnerException = $_.Exception
                try
                {
                    $ThrowOnError = $true
                    $IgnoreCase = $true
                    $GuessType = CompleteTypeName -TypeString $Type
                    $type = [type]::GetType($Guesstype, $ThrowOnError, $IgnoreCase)

                }
                catch [System.Management.Automation.MethodInvocationException]
                {
                    throw [System.Management.Automation.MethodInvocationException]::new($_.Exception.Message, $InnerException)
                }
                catch
                {
                    Throw $InnerException
                }
            }
        }

        If ($Indent -eq 0)
        {
            $Type.$Name | CleanTypeString -DoNothing:$NoTrimGenericTypeNames
            return Get-InterfaceTree -Type $Type -ShowFullTypeName:$ShowFullTypeName -NoTrimGenericTypeNames:$NoTrimGenericTypeNames -Indent 1
        }
        $Interfaces = $Type.GetInterfaces()
        If (@($Interfaces).Count -eq 0)
        {
            return
        }
        
        forEach ($Interface in $Interfaces)
        {
            Write-Output (CleanTypeString (("    " * $Indent) + $Interface.$Name) -DoNothing:$NoTrimGenericTypeNames)
            Get-InterfaceTree -Type $InterFace -Indent ($Indent + 1) -ShowFullTypeName:$ShowFullTypeName -NoTrimGenericTypeNames:$NoTrimGenericTypeNames
        }
        
    }
}

function Test-IntegerIsSigned
{
    param(
        $Integer
    )
    Begin
    {
        $SupportedTypes = @(
            [sbyte],
            [int16],
            [int32],
            [int64],
            [byte],
            [uint16],
            [uint32],
            [uint64]
        )
    }
    Process
    {
        If ($SupportedTypes -notcontains $Integer.GetType())
        {
            throw [System.ArgumentException]::new("Input object must be of one of the following types: $($SupportedTypes -join ', ')", "Integer")
        }

        return $Integer.GetType().Name[0] -match "[si]"
    }
}

function GetNumBits ([type]$Type)
{
    <#
        .Synopsis
        Returns the number of bits used for the storage of a supplied number
    #>
    if ($Type.IsValueType)
    {
        return [System.Runtime.InteropServices.Marshal]::SizeOf((1 -as $Type)) * 8
    }
    throw New-Object System.ArgumentException ("Supplied type must be a value type", "Type")

    # Unreachable code. Left for reference
    # Works for integers, messes up with floats
    return ( #In short, this performs:  2^ Ceiling(log2(log2(MaxValue)))
        [math]::pow(2, [math]::ceiling(
                [math]::log( ([math]::log( ($type::MaxValue), 2) ), 2 )
            ))
    ) -as [int]
}

function Convert-IntSign
{
    <#
        .Synopsis
        Converts between the signed and unsigned variants of a supplied integer assuming the same binary representation
        
        .Description
        The most significant bit of a binary number is used to represent the sign (positive or negative) of that number in a signed integer.
        For example, the signed byte type [sbyte] stores the number -100 in binary as "1001 1100"
        In the unsigned [byte] type, the MSB instead represents 2^7 and so the same binary representation would represent 156

        .Example
        Convert-IntSign ([byte]156)
        # Outputs -100

        .Notes
        The method used below is pretty inefficient and was more for pedagogical purposes. Some simple arithmetic based on two's complement would be a better way to achieve it
    #>
    
    Param(
        $Integer
    )

    Begin
    {
        $TypeEquivalents = @{
            [byte]   = [sbyte];
            [sbyte]  = [byte];
            [int16]  = [uint16];
            [uint16] = [int16];
            [int32]  = [uint32];
            [uint32] = [int32];
            [int64]  = [uint64];
            [uint64] = [int64]
        }
    }

    process
    {
        
        $InputType = $Integer.GetType()
        $TargetType = $TypeEquivalents[$InputType]

        $BinaryString = [System.Convert]::ToString($Integer, 2)

        # The ToString method will put out more digits than necessary for negative sbyte values.
        # For the purposes of defensive programming, I'll apply this fix to all outputs.
        $BinaryString = Get-RightString -string $BinaryString -NumChars (GetNumBits -type $InputType)

        # Get the required method for the conversion
        $ConvertMethodToCall = "To" + $TargetType.Name

        [System.Convert].InvokeMember(
            $ConvertMethodToCall,
            [System.Reflection.BindingFlags]::InvokeMethod,
            $null, $null,
            @(
                $BinaryString,
                2 #FromBase
            ),
            $null, $null, $null
        )
    }
}

function Get-RightString ([string]$string, [int]$NumChars = [int]::MaxValue)
{
    # Mimics Excel's RIGHT() string function
    $StringLength = $string.Length
    $StartChar = [math]::Max(0, $StringLength - $NumChars)
    $GetLength = [math]::Min($StringLength, $NumChars)

    $String.Substring($StartChar, $GetLength)
}

function Get-LevenshteinDistance ([string]$StringA, [string]$StringB)
{
    <#
    .SYNOPSIS
    Calculates the Levenshtein Distance for two input strings
    
    .DESCRIPTION
    https://en.wikipedia.com/wiki/Levenshtein_distance
    
    .PARAMETER StringA
    The first reference string
    
    .PARAMETER StringB
    The second reference string
    
    .EXAMPLE
    Get-LevenshteinDistance -StringA "hello" -StrinbB "yellow"

    returns 2
    
    .NOTES
    Adapted from the implementation on Rosetta Code
    #>
    
    if ([String]::IsNullOrEmpty($StringA))
    {
        return $StringB.Length
    }
    if ([string]::IsNullOrEmpty($StringB))
    {
        return $StringA.Length
    }
    if ([string]::Equals($StringA, $StringB))
    {
        return 0
    }

    $ALen = $StringA.Length
    $BLen = $StringB.Length

    $v0 = [int[]]::new($BLen + 1)
    $v1 = [int[]]::new($BLen + 1)

    for ($i = 0; $i -lt $v0.Length; $i++)
    {
        $v0[$i] = $i
    }

    for ($i = 0; $i -lt $ALen ; $i++)
    {
        $v1[0] = $i + 1

        for ($j = 0; $j -lt $BLen; $j++)
        {
            $DeletionCost = $v0[$j + 1] + 1
            $InsertionCost = $v1[$j] + 1
            $SubstitionCost = if ($StringA[$i] -eq $StringB[$j])
            {
                $v0[$j]
            }
            else
            {
                $v0[$j] + 1
            }
            $v1[$j + 1] = [math]::Min([math]::Min($DeletionCost, $InsertionCost), $SubstitionCost)
        }
        $v0, $v1 = $v1, $v0
    }
    return $v0[$Blen]
}

function Add-PsModulePath
{
    [CmdletBinding()]
    Param(
        [string]
        $Path,

        [ValidateSet("Machine", "Process", "User")]
        [string]
        $Target = "Machine"
    )

    Process
    {
        try
        {
            $CurrentPaths = [System.Environment]::GetEnvironmentVariable("PsModulePath", $Target -as [System.EnvironmentVariableTarget])
        }
        catch [System.Management.Automation.MethodException]
        {
            Write-Error "Failed to translate environment variable target"
            return;
        }

        if ($CurrentPaths.Length -eq 0)
        {
            [System.Environment]::SetEnvironmentVariable("PsModulePath", ($Path.Trim() + ";"), $Target -as [System.EnvironmentVariableTarget])
            return
        }

        $CurrentPaths = [System.Collections.ArrayList]::new(($CurrentPaths -split ';'))

        if (($CurrentPaths.ForEach{$_.Trim()} -contains $Path.Trim()))
        {
            Write-Warning "PSModulePath enviroment variable already contains supplied path"
            return
        }

        [void]$CurrentPaths.Add($Path.Trim())

        [System.Environment]::SetEnvironmentVariable("PsModulePath", $CurrentPaths.ToArray() -join ';', $Target -as [System.EnvironmentVariableTarget])

    }
}

function Get-Enumerator
{
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [object]
        $InputObject
    )

    Process
    {
        ForEach ($obj in $InputObject)
        {
            if ($obj -isnot [IEnumerable])
            {
                throw New-Object ArgumentException "Input object is not enumerable"
            }
            $obj.GetEnumerator()
        }
    }
}

function Write-HostColors
{
    <#
    .SYNOPSIS
    Writes to host a square of the available console colors
    
    .DESCRIPTION
    Long description
    
    .EXAMPLE
    Write-HostColors

    Writes the output to screen
    
    .NOTES
    None

    #>
    $colors = [enum]::GetValues([System.ConsoleColor])| % {[int]$_}
    $Fore = "   Foreground    "
    Write-Host "                    Background"
    Write-Host "      00 01 02 03 04 05 06 07 08 09 10 11 12 13 14 15"
    foreach ( $fcolor in $colors )
    {
        $ForeChar = $Fore[[int]$fcolor]
        Write-Host ("{0}  {1:D2} " -f $ForeChar, [int]$fcolor) -NoNewline
        foreach ( $bcolor in $colors )
        {
            
            Write-Host "PS>"-ForegroundColor $fcolor -BackgroundColor $bcolor -NoNewline
        }
        Write-Host
    }
}

function Get-Hue
{
    param(
        [Parameter(ParameterSetName = "FromRGB", Mandatory, Position = 0)]
        [int]
        $Red,
        [Parameter(ParameterSetName = "FromRGB", Mandatory, Position = 1)]
        [int]
        $Green,
        [Parameter(ParameterSetName = "FromRGB", Mandatory, Position = 2)]
        [int]
        $Blue,

        [Parameter(ParameterSetName = "FromColor", Mandatory, Position = 0)]
        [System.Drawing.Color]
        $Color,

        [Parameter(ParameterSetName = "FromConsoleColor", Mandatory, Position = 0)]
        [System.ConsoleColor]
        $ConsoleColor,

        [int]
        $BitsPerChannel = 8
    )

    Begin
    {
        $ConsoleColors = @{
            Black       = [Drawing.Color]"#000000"
            DarkBlue    = [Drawing.Color]"#000080"
            DarkGreen   = [Drawing.Color]"#008000"
            DarkCyan    = [Drawing.Color]"#008080"
            DarkRed     = [Drawing.Color]"#800000"
            DarkMagenta = [Drawing.Color]"#800080"
            DarkYellow  = [Drawing.Color]"#808000"
            Gray        = [Drawing.Color]"#C0C0C0"
            DarkGray    = [Drawing.Color]"#808080"
            Blue        = [Drawing.Color]"#0000FF"
            Green       = [Drawing.Color]"#00FF00"
            Cyan        = [Drawing.Color]"#00FFFF"
            Red         = [Drawing.Color]"#FF0000"
            Magenta     = [Drawing.Color]"#FF00FF"
            Yellow      = [Drawing.Color]"#FFFF00"
            White       = [Drawing.Color]"#FFFFFF"
        }

    }

    process
    {
        switch ($PSCmdlet.ParameterSetName)
        {
            "FromColor"
            {
                Write-Debug "Switch fromcolor"
                $Red = $Color.R
                $Green = $Color.G
                $Blue = $Color.B
            }
            "FromConsoleColor"
            {
                Write-Debug "Switch fromconsolecolor"
                $Color = $ConsoleColors[$ConsoleColor]
                $Red = $Color.R
                $Green = $Color.G
                $Blue = $Color.B
            }
            Default
            {
                Write-Debug "Switch default"
                # Do nothing
            }
        }
        $DivBy = [math]::Pow(2, $BitsPerChannel) - 1 # 255 for 8 bit, 1023 for 10 bit
        [double]$Rval = $red / $DivBy
        [double]$Gval = $green / $DivBy
        [double]$Bval = $blue / $DivBy

        Write-Debug "R:$Rval G:$Gval B:$Bval"

        $Min = [math]::min([math]::min($Rval, $Gval), $Bval)
        $Max = [math]::max([math]::max($Rval, $Gval), $Bval)
        Write-Debug "Min = $Min Max = $Max"
        if ($min -eq $max) {return 0} #Shade of grey
        $Diff = $Max - $Min
        switch ($Max)
        {
            {$_ -eq $Rval}
            {
                Write-Debug "Red is max"
                $hue = ($Gval - $Bval) / $Diff
            }
            {$_ -eq $Gval}
            {
                Write-Debug "Green is Max"
                $hue = 2 + ($Bval - $Rval) / $Diff
            }
            {$_ -eq $Bval}
            {
                Write-Debug "Blue is max"
                $hue = 4 + ($Rval - $Gval) / $Diff
            }
        }
        
        $hue *= 60
        if ($hue -lt 0) {$hue += 360}

        return $hue
        
    }

    end
    {

    }
}

Function Get-StringPermutation
{
    <#
        .SYNOPSIS
            Retrieves the permutations of a given string. Works only with a single word.
 
        .DESCRIPTION
            Retrieves the permutations of a given string Works only with a single word.
       
        .PARAMETER String           
            Single string used to give permutations on
       
        .NOTES
            Name: Get-StringPermutation
            Author: Boe Prox
            DateCreated:21 Feb 2013
            DateModifed:21 Feb 2013
 
        .EXAMPLE
            Get-StringPermutation -String "hat"
            Permutation                                                                          
            -----------                                                                          
            hat                                                                                  
            hta                                                                                  
            ath                                                                                  
            aht                                                                                  
            tha                                                                                  
            tah        

            Description
            -----------
            Shows all possible permutations for the string 'hat'.

        .EXAMPLE
            Get-StringPermutation -String "help" | Format-Wide -Column 4            
            help                  hepl                  hlpe                 hlep                
            hpel                  hple                  elph                 elhp                
            ephl                  eplh                  ehlp                 ehpl                
            lphe                  lpeh                  lhep                 lhpe                
            leph                  lehp                  phel                 phle                
            pelh                  pehl                  plhe                 pleh        

            Description
            -----------
            Shows all possible permutations for the string 'help'.
 
    #>
    [cmdletbinding()]
    Param(
        [parameter(ValueFromPipeline = $True)]
        [string]$String = 'the'
    )
    Begin
    {
        #region Internal Functions
        Function New-Anagram
        { 
            Param([int]$NewSize)              
            If ($NewSize -eq 1)
            {
                return
            }
            For ($i = 0; $i -lt $NewSize; $i++)
            { 
                New-Anagram  -NewSize ($NewSize - 1)
                If ($NewSize -eq 2)
                {
                    New-Object PSObject -Property @{
                        Permutation = $stringBuilder.ToString()
                    }
                }
                Move-Left -NewSize $NewSize
            }
        }
        Function Move-Left
        {
            Param([int]$NewSize)
            $z = 0
            $position = ($Size - $NewSize)
            [char]$temp = $stringBuilder[$position]
            For ($z = ($position + 1); $z -lt $Size; $z++)
            {
                $stringBuilder[($z - 1)] = $stringBuilder[$z]
            }
            $stringBuilder[($z - 1)] = $temp
        }
        #endregion Internal Functions
    }
    Process
    {
        $size = $String.length
        $stringBuilder = New-Object System.Text.StringBuilder -ArgumentList $String
        New-Anagram -NewSize $Size
    }
    End {}
}

Function Get-FunctionsFromScript
{
    param(
        [string]
        $Path
    )
    $tokens = $errors = $null
    $ast = [Language.Parser]::ParseFile(
        $Path,
        [ref]$tokens,
        [ref]$errors)
    
    $ast.FindAll(
        {
            param([Language.Ast] $Ast)
        
            $Ast -is [Language.FunctionDefinitionAst] -and
            # Class methods have a FunctionDefinitionAst under them as well, but we don't want them.
            ($Ast.Parent -isnot [Language.FunctionMemberAst])
        
        }, $false )
}

function ConvertTo-Dictionary
{
    <#
    .SYNOPSIS
    Converts an object to a Dictionary
    
    .DESCRIPTION
    Uses the LINQ ToDictionary method to create a dictionary based on an input IEnumerable object
    
    .PARAMETER Enumerable
    The IEnumerable object from which to create the dictionary
    
    .PARAMETER KeyProperty
    The name of the property to use as the Key of the dictionary. This must be unique throughout all entries of the enumerable
    
    .EXAMPLE
    $ProcDic = ConvertTo-Dictionary -Enumerable (Get-Process) -KeyProperty "Id"
    $ProcDic[1001]

    NPM(K)    PM(M)      WS(M)     CPU(s)      Id  SI ProcessName
    ------    -----      -----     ------      --  -- -----------
        14     4.18       7.61       0.00    1001   0 svchost
    
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [IEnumerable]
        $Enumerable,

        [Parameter(Mandatory = $true)]
        [string]
        $KeyProperty  
    )

    $KeyType = ($Enumerable | Select-Object -First 1).$KeyProperty.GetType()
    
    $DelegateType = [System.Func`2].MakeGenericType(([object]), $KeyType)
    $KeyDelegate = { $args[0].$KeyProperty } -as $DelegateType
    return [Linq.Enumerable]::ToDictionary($Enumerable, $KeyDelegate)
}

function Get-HumanisedTimespan
{
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [timespan]$Timespan
    )
    End
    {
        switch ($Timespan)
            {
                {$_.TotalDays -gt 1} # longer than a day
                {
                    [int]$Days = [math]::floor($_.TotalDays)
                    [int]$Hours = [math]::round($_.TotalHours % 24)
                    return '{0}d{1}h' -f $Days, $Hours
                }
                { $_.TotalHours -gt 1} # Longer than an hour
                {
                    [int]$Hours = [math]::floor($_.TotalHours)
                    [int]$Minutes = [math]::round($_.TotalMinutes % 60)
                    return '{0}h{1}m' -f $Hours, $Minutes
                }
                { $_.totalminutes -ge 1 } # Longer than a minute
                {
                    [int]$min = [math]::floor($_.TotalMinutes)
                    $sec = $_.Seconds
                    return '{0}m{1}s' -f $min, $sec
                }
                { $_.TotalSeconds -ge 10 } #Longer than 10 seconds
                {
                    [decimal]$d = $_.TotalSeconds
                    return '{0:f1}s' -f ($d)
                }
                { $_.TotalSeconds -ge 1 } #Longer than 1 second
                {
                    [decimal]$d = $_.TotalSeconds
                    return '{0:f2}s' -f ($d)
                }
                { $_.TotalMilliseconds -ge 10 } #longer than 10ms
                {
                    [decimal]$d = $_.TotalMilliseconds
                    return '{0:f0}ms' -f ($d)
                }
                { $_.TotalMilliseconds -lt 10 }  # Under 10ms
                {
                    [double]$d = $_.TotalMilliseconds
                    return '{0:f1}ms' -f ($d)
                }
                Default
                {
                    return '[{0:f0}ms]' -f $_.Milliseconds
                }
            }
    }
}

function Convert-Svg2Png
{
    [CmdletBinding()]
    param( 
        # Path to file or directory to convert
        [Alias("FullName")]
        [Parameter(ValueFromPipeLineByPropertyName = $true)]
        [string[]]
        $Path = '.', 
        
        [string]$Exec = 'C:\Program Files\Inkscape\inkscape.exe',
        
        [int]$Width = 64
    ) 
        
    Begin
    {
        if (!(Test-Path $Exec))
        {
            throw [System.IO.FileNotFoundException]::new("Inkspace executable not found in path", $Exec)
        }
    }
    
    Process
    {
        
        if ((Get-Item -Path $Path).Attributes -contains "Directory")
        {
            $PathItems = Get-ChildItem $Path
        }
        else
        {
            $PathItems = Get-Item $Path
        }
    
        Write-Verbose "PAth Count: $($PathItems.count)"
        foreach ($filename in $PathItems)
        { 
            if ($filename.ToString().EndsWith('.svg'))
            { 
                $targetName = $filename.BaseName + ".png"; 
                Write-Verbose "Converting $filename to $targetName ..." 
                $command = "& `"$Exec`" -z -e `"$targetName`" -w $Width `"$filename`"";
                Write-Debug "Command: $command"
                Invoke-Expression $command; 
            } 
        } 
    }
}

function Convert-StringsToExactRegex
{
    <#
    .SYNOPSIS
    Converts an array of strings to a regex that matches any of those strings exactly
    
    .DESCRIPTION
    Long description
    
    .PARAMETER String
    The array of strings used to build the regex
    
    .PARAMETER CaseSensitive
    Forces the resulting regex to be case sensitive. By default, the regex is case insensitive. This parameter has no effect if used with the AsString parameter
    
    .PARAMETER Options
    A RegexOptions object to include within the regex object. This parameter has no effect if used with the AsString parameter
    
    .PARAMETER AsString
    Returns the resultant regex as a string rather than a regex object. This parameter will cause the CaseSensitive and Options parameters to have no effect
    
    .EXAMPLE
    $array = @("hello", "this", "is", "a", "test!")
    Convert-StringstoExactRegex -String $array -AsString
    
    # returns "^hello$|^this$|^is$|^a$|^test!$"

    .Example
    $Opt = [System.Text.RegularExpressions.RegexOptions]::Compiled -bor [System.Text.RegularExpressions.RegexOptions]::Multiline -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
    $regex = Convert-StringsToExactRegex -String "a", "b" -Options $Opt

    # returns a regex object containing "^a$|^b$" with the supplied options, plus "IgnoreCase" as the CaseSensitive parameter was not supplied
    #>
    
    
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [string[]]$String,

        [Parameter()]
        [switch]$CaseSensitive,

        [Parameter()]
        [RegexOptions]$Options,

        [Parameter()]
        [switch]$AsString
    )
    Begin
    {
        $StringList = [Generic.List[string]]::new()
    }
    Process
    {
        if (!$CaseSensitive)
        {
            $Options = $Options -bor ([RegexOptions]::IgnoreCase)
        }
        if ($null -eq $Options)
        {
            $Options = 0
        }
        foreach ($Str in $String)
        {
            $StringList.Add($Str)
        }
        
    }
    End
    {
        $Sb = [System.Text.StringBuilder]::new()
        [void]$Sb.Append("^")
        [void]$Sb.Append($StringList -join '$|^')
        [void]$Sb.Append('$')
        if ($AsString)
        {
            return $Sb.ToString()
        }
        return [regex]::new($Sb.ToString(), $Options)
    
    }
}

function Get-BitmapFromUrl
{
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [string[]]$Url,
        [switch]$UseEmbeddedColorManagement
    )
    Begin
    {
        $Webclient = [System.Net.WebClient]::new()
    }
    Process
    {
        foreach ($path in $Url)
        {
            $ImageBytes = $Webclient.DownloadData($path)
            $MemStream = [System.IO.MemoryStream]::new($ImageBytes)
            Write-Output ([System.Drawing.Bitmap]::FromStream($MemStream))
            $MemStream.Dispose()
        }
        
    }
    End
    {
        $Webclient.Dispose()
    }
}

class CoOrd
{
    [int]$x
    [int]$y

    CoOrd([int]$x, [int]$y)
    {
        $this.x = $x
        $this.y = $y
    }
}

function Get-ImageColourCount
{
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [System.Drawing.Image[]]$Image
    )

    Begin
    {
        # Gets addresses of top left to bottom right: "\"
        function GetLeftObliqueDiagonal ([int]$width, [int]$height)
        {
            $Ratio = $height / $width
            0..($width - 1) | ForEach-Object { [CoOrd]::new($_, [math]::Round($_ * $Ratio)) }
        }

        # Gets addresses of top right to bottom left; "/"
        function GetRightObliqueDiagonal ([int]$width, [int]$height)
        {
            $Ratio = $height / $width
            0..($width - 1) | ForEach-Object { [CoOrd]::new($width - $_, [math]::Round($_ * $Ratio)) }
        }
    }
    
    Process
    {
        $Hashset = [Generic.HashSet[Drawing.Color]]::new()
        
        ForEach ($img in $Image)
        {
            if ($img.Height * $img.Width -eq 0)
            {
                throw "image dimensions are zero"
            }
            $WidthToHeight = $img.Width / $img.Height
            $ThumbHeight = [math]::Min(192, $img.Height)
            $ThumbWidth = [math]::Min(192 * $WidthToHeight, $img.Width)
            Write-Debug "Width: $ThumbWidth Height: $ThumbHeight Ratio: $WidthToHeight"
            $Small = $Image.GetThumbnailImage($ThumbWidth, $ThumbHeight, $null, [IntPtr]::Zero)

            $CrossCoordinates = [Generic.List[CoOrd]]::new(2 * $ThumbWidth)
            $CrossCoordinates.AddRange((GetLeftObliqueDiagonal -width ($ThumbWidth - 1) -height ($ThumbHeight - 1)) -as [CoOrd[]])
            $CrossCoordinates.AddRange((GetRightObliqueDiagonal -width ($ThumbWidth - 1) -height ($ThumbHeight - 1)) -as [CoOrd[]])
            
            foreach ($xy in $CrossCoordinates)
            {
                
                try
                {
                    $Color = $img.GetPixel($xy.x, $xy.y)
                }
                catch
                {
                    Write-Warning "Failed accessing pixel $($xy.x),$($xy.y)"
                    continue
                }
                
                [void]$Hashset.Add($Color)
            }
            return $Hashset.Count
        }


    }

}

class DictionaryUpdateCounter
{
    [int]$Replacement
    [int]$Addition
    [int]$Total
    [Generic.List[string]]$Amendments

    DictionaryUpdateCounter()
    {
        $this.Replacement = $this.Addition = $this.Total = 0
        $this.Amendments = [Generic.List[string]]::new()
    }
    [void]AddAddition()
    {
        $this.Addition += 1
        $this.Total += 1
    }
    [void]AddAddition([string]$amendment)
    {
        $this.Addition += 1
        $this.Total += 1
        $this.Amendments.Add($amendment)
    }
    [void]AddReplacement()
    {
        $this.Replacement += 1
        $this.Total += 1
    }
    [void]AddReplacement([string]$amendment)
    {
        $this.Replacement += 1
        $this.Total += 1
        $this.Amendments.Add($amendment)
    }
}

function New-DictionaryUpdateCounter
{
    return [DictionaryUpdateCounter]::new()
}

function Update-Dictionary
{
    <#
    .SYNOPSIS
    Updates a given Dictionary / hashtable with supplied values, leaving all others untouched. Returns the number of amendments made
    
    .DESCRIPTION
    Long description
    
    .PARAMETER Source
    The input source dictionary to update. Note that the input will be modified in place (passed by reference). The dictionary is not returned to the pipeline.
    
    .PARAMETER Updates
    A dictionary containing the updates to implement

    .PARAMETER CountOnly
    If this switch parameter is specified, only a count of would-be amendments is returned. The input dictionary is not modified
    
    .EXAMPLE
    $StartDictionary = @{key1 = "value1"; key2 = @{key3 = "value3"}}
    $UpdateDictionary = @{key2 = @{key3 = "updated_value"; newkey = "ValueNotPresentBefore"}}
    Update-Dictionary -Source $StartDictionary -Updates $UpdateDictionary

    returns a dictionary of the stucture:
    @{key1 = "value1"; key2 = @{key2 = @{key3 = "updated_value"; newkey = "ValueNotPresentBefore"}}}
    
    .NOTES
    Useful for updating dictionaries of settings objects where you only want to amend a select number of properties, but are obliged to resend the entire dictiory object

    #>
    
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [System.Collections.IDictionary]
        $Source,
        
        [Parameter(Mandatory = $true, Position = 1)]
        [System.Collections.IDictionary]
        $Updates,

        [Parameter()]
        [Alias("noop", "list")]
        [switch]
        $CountOnly,

        # Not for public use. Used internally to track amendment counts when recursing
        [Parameter(Mandatory = $false, DontShow = $true)]
        [ref]
        $UpdateCtr,

        # Not for public use. Used internally to report the path being updated
        [Parameter(Mandatory = $false, DontShow = $true)]
        [string]
        $Root = "[Root]"
    )


    if ($null -eq $UpdateCtr)
    {
        # We make the assumption here that if no counter has been supplied, we are at the initial call and will filter out
        # all but the last counter returned from any recursion
        $StartCounter = [DictionaryUpdateCounter]::new()
        return (Update-Dictionary -Source $Source -Updates $Updates -UpdateCtr ([ref]$StartCounter) -CountOnly:$CountOnly | Select-Object -Last 1)
    }
    foreach ($key in $Updates.Keys)
    {
        $KeyPath = "$Root.$($key.ToString())"
        if ($Source.ContainsKey($key))
        {
            if ($Source[$key] -is [System.Collections.IDictionary])
            {
                Update-Dictionary -Source $Source[$key] -Updates $Updates[$key] -UpdateCtr $UpdateCtr -Root $KeyPath -CountOnly:$CountOnly
            }
            else
            {
                # Need to check this as <array> -compare <array> returns another array, not a bool
                if (($Source[$key] -ne $Updates[$key]) -is [bool])
                {
                    if ($Source[$key] -ne $Updates[$key])
                    {
                        $message = "Replacing $KeyPath"
                        Write-Verbose $message
                        if (!$CountOnly)
                        {
                            $Source[$key] = $Updates[$key]
                        }
                        $UpdateCtr.Value.AddReplacement($message)
                    }
                }
                else
                {
                    if ($null -ne (Compare-Object -ReferenceObject $Source[$key] -DifferenceObject $Updates[$key]))
                    {
                        $message = "Replacing $KeyPath"
                        Write-Verbose $message
                        if (!$CountOnly)
                        {
                            $Source[$key] = $Updates[$key]
                        }
                        $UpdateCtr.Value.AddReplacement($message)
                    }
                }
            }
        }
        else 
        {
            # Update is an addition
            $message = "Adding $Keypath"
            Write-Verbose $message
            if (!$CountOnly)
            {
                $Source[$key] = $Updates[$key]
            }
            $UpdateCtr.Value.AddAddition($message)
        }
    }
    Write-Output $UpdateCtr.Value
}

function New-GenericList
{
    param(
        # The type of the contained object. Pass a type or string, which will be attempted to translate into a type
        [Parameter(Mandatory = $true)]
        $Type
    )

    throw [System.NotImplementedException]::new()
}

function Get-HumanisedBytes 
{
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [int64]
        $Bytes,
        [switch]
        $PowersOfTen,
        [ValidateRange(0, 12)]
        [int]$DecimalPlaces = 2
    )
    $HumanSizes = @(
        "bytes",
        "Kb",
        "Mb",
        "Gb",
        "Tb",
        "Pb",
        "Eb"
    )

    if ($PowersOfTen)
    {
        $Radix = 10
        $Exponent = 3
    }
    else
    { 
        $Radix = 2 
        $Exponent = 10
    }

    $DBytes = $Bytes -as [double]
    
    $Threshold = [math]::pow($Radix, $Exponent)
    for ($i = 0; $i -lt $HumanSizes.Length; $i++)
    {
        if ($DBytes -le $Threshold)
        {
            break
        }
        $DBytes /= $Threshold
    }
    if ($i -eq 0) {$DecimalPlaces = 0}
    return "{0:F$DecimalPlaces} {1}" -f $DBytes, $HumanSizes[$i]
}

function Remove-NullProperties
{
    # For assistance when using ConvertTo-Json (which has no way of turning off null properties, they're all included)
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [object]
        $Object,

        [string[]]
        $Exclude
    )

    if ($Object -is [enum] -Or $Object -is [string] -Or $Object.GetType().BaseType -eq [System.ValueType]) { return $object}
    $AllProps = $Object.psobject.properties.Name

    if ($null -eq $AllProps) { return $Object }

    $NonNull = $AllProps.Where( { $null -ne $Object.$_ -Or $Exclude -contains $_ })

    return $Object | Select-Object $NonNull
}

function Get-WebString
{
    # Function to replace Invoke-RestMethod that seems to mishandle certain UTF8 text
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]
        $Uri,

        [Parameter(Mandatory = $false, Position = 1)]
        [System.Text.Encoding]
        $Encoding = [System.Text.Encoding]::UTF8
    )

    $wr = [System.Net.WebRequest]::Create($uri)
    #$wr.ContentType = "application/json;charset=UTF-8"
    $wr.Method = "Get"
    $response = $wr.GetResponse()
    $stream = $response.GetResponseStream()
    $ms = [System.IO.MemoryStream]::new($response.ContentLength)
    $Buffer = [byte[]]::new(1024)
    $ByteCount = 0
    While (($ByteCount = $stream.Read($Buffer, 0, 1024)) -gt 0)
    {
        $ms.Write($Buffer, 0, $ByteCount)
    }
    $Encoding.GetString($ms.ToArray())
}

function Get-DeviceLocation
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [TimeSpan]
        $TimeOut = [timespan]::FromSeconds(5),

        # The accuracy of the reported location. Default relies on Network location (but can still be quite accurate). "High" requires GPS or similar to produce results
        [Parameter(Mandatory = $false)]
        [string]
        [ValidateSet("Default", "High")]
        $Accuracy = "Default",

        # Opens Google Maps with the resultant location (if found)
        [switch]
        $Map
    )

    Begin
    {
        Add-Type -AssemblyName System.Device
    }
    End
    {
        $Geo = [System.Device.Location.GeoCoordinateWatcher]::new($Accuracy)

        if ($Geo.TryStart($false, $TimeOut))
        {
           $sw = [System.Diagnostics.Stopwatch]::StartNew()
            While ($Geo.Status -ne "Ready" -and $sw.Elapsed -lt $TimeOut)
            {
                Start-sleep -Milliseconds 30
            }
            if ($Map)
            {
                $Location = $Geo.Position.Location
                if ([double]::IsFinite($Location.Latitude * $Location.Longitude) )
                {
                    $LatLng = "{0},{1}" -f $Location.Latitude, $Location.Longitude
                    Start-Process "https://google.com/maps/search/?api=1&query=$($LatLng)"
                }
                else
                {
                    Write-Warning "Unable to retrieve valid values for Latitude and/or Longitude"
                }
            }
            else 
            {
                $Geo.Position.Location
            }
        }
    }
}

function Write-Maze
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [int]
        $Width = $Host.UI.RawUI.WindowSize.Width,

        [Parameter()]
        [int]
        $Height = 10
    )

    End
    {
        $Area = $Width * $Height
        Write-Debug $Area
        $RowBuilder = [System.Text.StringBuilder]::new($Width)
        $RequiredBytes = [System.Math]::Ceiling($Area / 8)
        Write-Debug $RequiredBytes
        $BitArray = [System.Collections.BitArray]::new(
            ([byte[]]$(for ($i = 0; $i++ -lt $RequiredBytes; ) { get-random -Maximum 256 }))
            )
        Write-Debug $BitArray.Length
        for ($row = $Offset = 0; $row -lt $Height; $row++)
        {
            for ($i = 0; $i -lt $Width; $i++)
            {
                try {
                    [void]$RowBuilder.Append(([char](0x2571 + $BitArray[$Offset + $i])))
                }
                catch {
                    
                }
            }
            $Offset += $Width
            $RowBuilder.ToString()
            [void]$RowBuilder.Clear()
        }
    }
}

function Test-NonTerminatingError
{
    [CmdletBinding()]
    param(
        [string]$Message = "Non terminating error",
        [string]$ErrorId = "Custom Error",
        [System.Management.Automation.ErrorCategory]$ErrorCategory = "NotSpecified",
        [object]$TargetObject = $null
    )

    $exception = [System.Exception]::new($Message)
    $errorRecord = [System.Management.Automation.ErrorRecord]::new($exception, $ErrorId, $ErrorCategory, $TargetObject)

    $PSCmdlet.WriteError($errorRecord)
}