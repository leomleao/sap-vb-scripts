
#-Begin-----------------------------------------------------------------

#-Load assembly---------------------------------------------------------
If($PSVersionTable.PSVersion.Major -le 5) {
  Add-Type -AssemblyName "Microsoft.VisualBasic";
  Add-Type -AssemblyName "System.Windows.Forms";
} ElseIf ($PSVersionTable.PSVersion.Major -ge 7) {
  Add-Type -AssemblyName "System.Windows.Forms";
}

#-Function Create-Object------------------------------------------------
Function Create-Object {

  Param(
    [String]$objectName
  )

  Try {
    New-Object -ComObject $objectName;
  } Catch {
    If(($PSVersionTable.PSVersion.Major -le 5) -or `
      ($PSVersionTable.PSVersion.Major -ge 7)) {
      [Void][System.Windows.Forms.MessageBox]::Show(
        "Can't create object", "Important hint", 0);
    }
  }

}

#-Function Get-Object---------------------------------------------------
Function Get-Object {

  Param(
    [String]$class
  )

  If($PSVersionTable.PSVersion.Major -le 5) {
    Try {
      [Microsoft.VisualBasic.Interaction]::GetObject($class);
    } Catch {}
  } ElseIf($PSVersionTable.PSVersion.Major -ge 6) {
    Try {
      $SapROTWr = New-Object -ComObject "SapROTWr.SapROTWrapper";
      $SapROTWr.GetROTEntry($class);
    } Catch {}
  }

}

#-Sub Free-Object-------------------------------------------------------
Function Free-Object {

  Param(
    [__ComObject]$object
  )

  Try {
    [Void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($object);
  } Catch {}

}

#-Function Get-Property-------------------------------------------------
Function Get-Property {

  Param(
    [__ComObject]$object,
    [String]$propertyName, 
    $propertyParameter
  )

  Try {
    $objectType = [System.Type]::GetType($object);
    $objectType.InvokeMember($propertyName,
      [System.Reflection.Bindingflags]::GetProperty,
      $null, $object, $propertyParameter);
  } Catch {}

}

#-Sub Set-Property------------------------------------------------------
Function Set-Property {

  Param(
    [__ComObject]$object,
    [String]$propertyName,
    $propertyValue
  )

  Try {
    $objectType = [System.Type]::GetType($object);
    [Void] $objectType.InvokeMember($propertyName,
      [System.Reflection.Bindingflags]::SetProperty,
      $null, $object, $propertyValue);
  } Catch {}

}

#-Function Invoke-Method------------------------------------------------
Function Invoke-Method {

  Param(
    [__ComObject]$object,
    [String]$methodName,
    $methodParameter
  )

  Try {
    $objectType = [System.Type]::GetType($object);
    $output = $objectType.InvokeMember($methodName,
      [System.Reflection.BindingFlags]::InvokeMethod,
      $null, $object, $methodParameter);
    if ( $output ) { $output }
  } Catch {}

}

#-End-------------------------------------------------------------------
