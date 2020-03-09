$psd1 = Import-PowerShellDataFile DJMisc.psd1
$Version = $psd1.ModuleVersion
Import-Module DjMisc -RequiredVersion $Version -Force -PassThru | Out-String | Out-Host

Describe "Get-HumanisedTimespan" {
    $SampleData = @{Data = 3549835; Expected = "41d2h" },
                  @{Data = 3549; Expected = "59m9s" }, 
                  @{Data = 354; Expected = "5m54s" },
                  @{Data = 3.548; Expected = "3.55s" },
                  @{Data = 0.3542; Expected = "354ms" },
                  @{Data = 0.034201; Expected = "34ms" },
                  @{Data = 0.00344536; Expected = "3.4ms" },
                  @{Data = 0.000344; Expected = "0.3ms" }

    It "Returns the expected string for a given input timespan" -TestCases $SampleData {
        param(
            [double]$Data,
            [string]$Expected
        )
        [timespan]::FromTicks($Data * [timespan]::TicksPerSecond) | Get-HumanisedTimespan | Should Be $Expected
        Get-HumanisedTimespan -Timespan ([timespan]::FromTicks($Data * [timespan]::TicksPerSecond)) | Should Be $Expected
    }
}

Describe "Convert-StringstoExactRegex" {
    It "Outputs a string with the AsString switch parameter" {
        Convert-StringsToExactRegex -String "a", "b" -AsString | Should BeOfType ([string])
    }

    It "Outputs a regex object when the AsString switch parameter is not supplied" {
        Convert-StringsToExactRegex -String "a", "b" | Should BeOfType ([regex])
    }
    
    It "Correctly joins when passed an array" {
        $testarray = @("hello", "this", "is", "a", "test!")
        $Expected = "^hello$|^this$|^is$|^a$|^test!$"
        Convert-StringstoExactRegex -String $testarray -AsString | Should be $Expected
    }

    It "Correctly joins when values passed through pipeline" {
        $Expected = "^hello$|^this$|^is$|^a$|^test!$"
        @("hello", "this", "is", "a", "test!") | Convert-StringstoExactRegex -AsString | Should be $Expected
    }

    It "Performs as expected with a single string input" {
        Convert-StringsToExactRegex -String "a" -AsString | Should Be '^a$'
    }
    
    It "Sets the IgnoreCase regex option without the CaseSensitive switch parameter" {
        $result = Convert-StringsToExactRegex -String "a", "b"
        $result.Options -band 1 | Should be 1
    }

    It "Does not set the IgnoreCase regex option with the CaseSensitive switch parameter" {
        $result = Convert-StringsToExactRegex -String "a", "b" -CaseSensitive
        $result.Options -band 1 | Should be 0
    }

    It "Sets the specified options in the resultant regex object" {
        $Opt = [System.Text.RegularExpressions.RegexOptions]::Compiled -bor [System.Text.RegularExpressions.RegexOptions]::Multiline -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
        $result = Convert-StringsToExactRegex -String "a", "b" -Options $Opt -CaseSensitive
        $result.Options | Should be $Opt
    }

    It "Is additive with the options and the case sensitive switch parameter" {
        $Opt = [System.Text.RegularExpressions.RegexOptions]::Compiled -bor [System.Text.RegularExpressions.RegexOptions]::Multiline -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
        $result = Convert-StringsToExactRegex -String "a", "b" -Options $Opt
        $result.Options | Should be ($Opt + 1)
    }
}

Describe "Update-Dictionary" {
    It "Returns a DictionaryUpdateCounter" {
        $result = Update-Dictionary -Source @{a = 1 } -Updates @{ }

        $result.GetType().FullName | Should Be "DictionaryUpdateCounter"
    }

    It "Does not change the Source type" {
        $src = @{a = 1 }
        $null = Update-Dictionary -Source $src -Updates @{ }
        $src | Should -BeOfType ($src.GetType())
    }
    
    It "Should permit multiple derived IDictionary types" {
        $srcDic = [System.Collections.Generic.Dictionary[int, string]]::new()
        $srcDic.Add(1, "2")
        $res = Update-Dictionary -Source $srcDic -Updates @{1 = "three" }
        $srcDic | Should BeOfType ($srcDic.GetType())
        $res.Total | Should -Be 1
    }

    It "Returns an unchanged dictionary if the updates dictionary is empty" {
        $src = @{a = 1 }
        $result = Update-Dictionary -Source $src -Updates @{ }
        $src.ContainsKey("a") | Should -BeTrue
        $src["a"] | Should be 1
        $result.Total | Should be 0
    }

    It "Performs a simple substitution on a 1 level dictionary" {
        $src = @{a = 1 ; b = 2 }
        $upd = @{b = 3 }
        $result = Update-Dictionary -Source $src -Updates $upd
        $src.ContainsKey("a") | Should -BeTrue
        $src["a"] | Should be 1
        $src.ContainsKey("b") | Should -BeTrue
        $src["b"] | Should be 3
        $Result.Replacement | Should -Be 1
        $Result.Total | Should -Be 1
    }

    It "Performs a substitution on a nested dictionary" {
        $src = @{a = @{c = 3 } ; b = 2 }
        $upd = @{a = @{c = 6 } }
        $result = Update-Dictionary -Source $src -Updates $upd
        $src.ContainsKey("a") | Should Be $true
        $src["a"] | Should BeOfType [System.Collections.IDictionary]
        $src["a"].ContainsKey("c") | Should Be $true
        $src["a"]["c"] | Should be 6
        $src["b"] | Should be 2
        $result.Replacement | Should be 1
        $result.Total | Should -Be 1
    }

    It "Adds keys present in the updates not in the source" {
        $src = @{a = @{c = 3 } ; b = 2 }
        $upd = @{a = @{d = 10 } }
        $result = Update-Dictionary -Source $src -Updates $upd
        $src["a"].ContainsKey("d") | Should be $true
        $src["a"]["d"] | Should -be 10
        $result.Addition | Should -be 1
        $result.Total | Should -be 1
    }
    It "Correctly Counts replacements and additions seperately" {
        $src = @{a = @{c = 3 } ; b = 2 }
        $upd = @{a = @{d = 10; e = 6 }; b = 3 }
        $result = Update-Dictionary -Source $src -Updates $upd
        $result.Addition | Should -be 2
        $result.Replacement | Should -Be 1
        $result.Total | Should -be 3
    }

    
    It "Makes no changes if using the CountOnly switch" {
        $src = @{a = @{c = 3 } ; b = 2 }
        $upd = @{a = @{d = 10; e = 6 }; b = 3 }
        $result1 = Update-Dictionary -Source $src -Updates $upd -CountOnly
        $src["b"] | Should -Be 2
        $result2 = Update-Dictionary -Source $src -Updates $upd
        $result1.Addition | Should -Be $result2.Addition
        $result1.Replacement | Should -Be $result2.Replacement
    }
    
    It "Makes no changes the 2nd sweep of the same data" {
        $src = @{a = @{c = 3 } ; b = 2 }
        $upd = @{a = @{d = 10; e = 6 }; b = 3 }
        $result1 = Update-Dictionary -Source $src -Updates $upd
        $result1.Total | Should -Not -Be 0
        $result2 = Update-Dictionary -Source $src -Updates $upd
        $result2.Addition | Should -Be 0
        $result2.Replacement | Should -Be 0
    }

    It "Replaces Arrays within a Value" {
        $src = @{a = @(@{b = 2; c = 3 }, @{d = 4; e = 5 }); f = 6 }
        $upd = @{a = @(@{g = 7; h = 8 }) }
        
        $src["a"].Count | Should -Be 2
        Update-Dictionary -Source $src -Updates $upd
        
        $src["a"].Count | Should -Be 1
        $src["a"][0].ContainsKey("g") | Should -BeTrue
        $src["a"][0].ContainsKey("b") | Should -Not -BeTrue
    }

    It "Replaces Lists within a Value" {
        $Hashlist = [System.Collections.Generic.List[hashtable]]::new(2)
        $Hashlist.Add(@{b = 2; c = 3 })
        $HashList.Add(@{d = 4; e = 5 })
        $src = @{a = $HashList; f = 6 }
        $upd = @{a = @(@{g = 7; h = 8 }) }
        
        $src["a"].Count | Should -Be 2
        Update-Dictionary -Source $src -Updates $upd
        
        $src["a"].Count | Should -Be 1
        $src["a"][0].ContainsKey("g") | Should -BeTrue
        $src["a"][0].ContainsKey("b") | Should -Not -BeTrue
    }
}

Describe "New-DictionaryUpdateCounter" {
    It "returns a dictionary counter object" {
        (New-DictionaryUpdateCounter).GetType().FullName | Should -Be "DictionaryUpdateCounter"
    }
}

Describe "DictionaryUpdateCounter" {
    It "Instantiates without issue" {
        $duc = New-DictionaryUpdateCounter
        $duc.Total | Should -Be 0
    }
    It "Correctly instantiates the Amendments list as part of the constructor" {
        $duc = New-DictionaryUpdateCounter
        $duc.Amendments.GetType().FullName | Should -Match "System.Collections.Generic.List"
        $duc.AddAddition("test")
        $duc.Amendments[0] | Should -Be "test"
    }
}

Describe "Remove-NullProperties" {
    It "Removes null properties from a sample pscustomobject" {
        $TestObject = [pscustomobject]@{Property1 = "SomeValue"; Property2 = $null }
        $ShouldObject = [pscustomobject]@{Property1 = "SomeValue" }
        
        $a = $TestObject | Remove-NullProperties | ConvertTo-Json -Compress 
        $b = $ShouldObject | ConvertTo-Json -Compress
        $a | Should -BeExactly $b
    }
    
    It "Removes Null Properties except those provided to Except from a sample PSCustomObject" {
        $TestObject = [pscustomobject]@{Property1 = "SomeValue"; Property2 = $null ; Property3 = $null }
        $ShouldObject = [pscustomobject]@{Property1 = "SomeValue" ; Property2 = $null }

        $TestObject | Remove-NullProperties -Exclude "Property2" | ConvertTo-Json -Compress | Should -BeExactly $($ShouldObject | ConvertTo-Json -Compress)
    }

    It "Returns an object with no null properties as-is" {
        $TestObject = [pscustomobject]@{Property1 = "SomeValue"; Property2 = 4 }
        $ShouldObject = $TestObject

        $TestObject | Remove-NullProperties -Exclude "Property3" | ConvertTo-Json -Compress | Should -BeExactly $($ShouldObject | ConvertTo-Json -Compress)
    }
    
    It "Returns a struct as-is" {
        (1) | Remove-NullProperties | Should Be 1
        ("simple string") | Remove-NullProperties | Should be "simple string"
        ($true) | Remove-NullProperties | Should be $true
        (5.0) | Remove-NullProperties | Should be (5.0)
    }

    It "Returns an enum as-is" {
        ([System.Text.RegularExpressions.RegexOptions]::IgnoreCase | Remove-NullProperties)  -eq ([System.Text.RegularExpressions.RegexOptions]::IgnoreCase )| Should -Be $true
    }
}

Describe "Test-NonTerminatingError" {
    It "Has no mandatory parameters" {
        Test-NonTerminatingError -ErrorAction SilentlyContinue
    }
    It "Populates an error record with given parameters" {
        Test-NonTerminatingError -Message "Pester Message" -ErrorId "1234" -ErrorCategory WriteError -ErrorVariable pestererror -ErrorAction SilentlyContinue

        $pestererror.Count | Should -Be 1
        $pestererror[0].Exception.Message | Should -BeExactly "Pester Message"
        $pestererror[0].CategoryInfo.Category | Should -Be "WriteError"
        $pestererror[0].FullyQualifiedErrorId | Should -Be "1234,Test-NonTerminatingError"
    }
}