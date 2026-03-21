Function Get-MemoryMappedFile {
    [CmdletBinding()]Param(
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]
        [System.String]$Name,
        [ValidateSet('Local','Global')]
        [System.String]$Scope  = 'Local',  # Local or Global (Global requires admin privileges)
        [System.Int64]$Size    = 5120,     # 5120 = 5 KB = max. 2560 Characters (UTF8 without Symbols or Emojis)
        [System.Management.Automation.SwitchParameter]$WriteHost,
        [System.Management.Automation.SwitchParameter]$WriteHostResultsOnly
    )

    try {
        # Define variables with initial values
        [System.Text.Encoding]$enc = [System.Text.Encoding]::UTF8
        [System.IntPtr]$handle     = [System.IntPtr]::Zero
        [System.String]$result     = [System.String]::Empty
        [System.String]$MapName    = "$($Scope)\$($Name)"
        
        # Add type for reading existing MemoryMappedFiles
        if (-not (Get-Variable -Name "MMFHelper" -Scope 'Global' -ValueOnly -ErrorAction Ignore)) {
            $Global:MMFHelper = Add-Type -PassThru -TypeDefinition @"
                using System;
                using System.Runtime.InteropServices;
                public class MMFHelper {
                    public const uint FILE_MAP_ALL_ACCESS = 1;
                    public const uint FILE_MAP_EXECUTE    = 2;
                    public const uint FILE_MAP_READ       = 4;
                    public const uint FILE_MAP_WRITE      = 8;
                    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
                    public static extern IntPtr OpenFileMapping(uint dwDesiredAccess, bool bInheritHandle, string lpName);
                    [DllImport("kernel32.dll", SetLastError = true)]
                    public static extern bool CloseHandle(IntPtr hObject);
                }
"@
            if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "Global 'MMFHelper' object created."}
        }
        else {if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "Global 'MMFHelper' object exists."}}
        
        # Check Name value
        if ([System.String]::IsNullOrWhiteSpace($Name)) {throw 'Name parameter value can not be null or whitespaces only!'}

        # Load MemoryMappedFile if already existing
        if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "MMF map name: $MapName"}
        [System.IntPtr]$handle = [MMFHelper]::OpenFileMapping([MMFHelper]::FILE_MAP_READ, $false, $MapName)
        if ($handle -ne [System.IntPtr]::Zero) {
            if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "MMF handle: $handle"}
            
            # Open MemoryMappedFile
            [System.IO.MemoryMappedFiles.MemoryMappedFile]$mmf = [System.IO.MemoryMappedFiles.MemoryMappedFile]::OpenExisting(
                $MapName, 
                [System.IO.MemoryMappedFiles.MemoryMappedFileRights]::Read
            )

            # Create MemoryMappedViewAccessor object
            [System.IO.MemoryMappedFiles.MemoryMappedViewAccessor]$accessor = $mmf.CreateViewAccessor(
                0,
                0,
                [System.IO.MemoryMappedFiles.MemoryMappedFileAccess]::Read
            )
            if ($WriteHost -and -not $WriteHostResultsOnly) {
                Write-Host (
                    "Accessor stats: " + `
                    "CanRead = $($accessor.CanRead) | CanWrite = $($accessor.CanWrite) | " + `
                    "Capacity = $($accessor.Capacity) | PointerOffset = $($accessor.PointerOffset)"
                )
            }

            # Read data from memory space
            [System.Byte[]]$buffer = [System.Byte[]]::new($Size)
            [System.Int64]$arr     = $accessor.ReadArray(0, $buffer, 0, $buffer.Length)
            if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "Content size: $arr"}

            # Free memory space
            if ($null -ne $accessor) {$accessor.Dispose()}
            if ($null -ne $mmf)      {$mmf.Dispose()}
            if ($handle -ne [IntPtr]::Zero) {[System.Boolean]$cHResult = [MMFHelper]::CloseHandle($handle); $handle = [IntPtr]::Zero}
            if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "CloseHandle result: $cHResult"}
            [System.GC]::Collect()

            # Create result value
            [System.Int64]$nullIndex  = 0
            $nullIndex = [System.Array]::IndexOf($buffer, [byte]0)
            if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "Null index: $nullIndex"}
            if ($nullIndex -ge 0) {$result = $enc.GetString($buffer, 0, $nullIndex)}
            else {$result = $enc.GetString($buffer)}
            $result = $result.Trim()
            if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "String length: $($result.Length)"}
            if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "Content: $result"}
            if ($WriteHost -and $WriteHostResultsOnly -and -not ([System.String]::IsNullOrWhiteSpace($result))) {Write-Host $result}
        }
        else {if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "Content not found."}}
    }
    catch {
        # Error handling
        Write-Host $($_ | Out-String).Trim() -ForegroundColor Yellow
    }
    finally {
        # Free memory space
        if ($null -ne $accessor) {$accessor.Dispose()}
        if ($null -ne $mmf)      {$mmf.Dispose()}
        if ($handle -ne [IntPtr]::Zero) {[System.Boolean]$cHResult = [MMFHelper]::CloseHandle($handle); $handle = [IntPtr]::Zero}
        [System.GC]::Collect()
    }

    # Return result from memory space
    return $result
}

Function Set-MemoryMappedFile {
    [CmdletBinding()]Param(
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSReference]$Data,
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]
        [System.String]$Name,
        [ValidateSet('Local','Global')]
        [System.String]$Scope                = 'Local',  # Local or Global (Global requires admin privileges)
        [System.Int64]$Size                  = 5120,     # 5120 = 5 KB = max. 2560 Characters (UTF8 without Symbols or Emojis)
        [System.Int32]$TimeoutInSeconds      = 30,
        [System.Int32]$WritingPauseInSeconds = 3,
        [System.Management.Automation.SwitchParameter]$WriteHost,
        [System.Management.Automation.SwitchParameter]$WriteHostResultsOnly
    )
    
    try {
        # Define variables with initial values
        [System.Text.Encoding]$enc = [System.Text.Encoding]::UTF8
        [System.IntPtr]$handle     = [System.IntPtr]::Zero
        [System.String]$MapName    = "$($Scope)\$($Name)"

        # Add type for reading existing MemoryMappedFiles
        if (-not (Get-Variable -Name "MMFHelper" -Scope 'Global' -ValueOnly -ErrorAction Ignore)) {
            $Global:MMFHelper = Add-Type -PassThru -TypeDefinition @"
                using System;
                using System.Runtime.InteropServices;
                public class MMFHelper {
                    public const uint FILE_MAP_ALL_ACCESS = 1;
                    public const uint FILE_MAP_EXECUTE    = 2;
                    public const uint FILE_MAP_READ       = 4;
                    public const uint FILE_MAP_WRITE      = 8;
                    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
                    public static extern IntPtr OpenFileMapping(uint dwDesiredAccess, bool bInheritHandle, string lpName);
                    [DllImport("kernel32.dll", SetLastError = true)]
                    public static extern bool CloseHandle(IntPtr hObject);
                }
"@
            if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "Global 'MMFHelper' object created."}
        } 
        else {if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "Global 'MMFHelper' object exists."}}

        # Checks
            # Admin permissions required
            if ($Scope -eq 'Global') {
                $isAdmin = ([Security.Principal.WindowsPrincipal]([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                if (-not $isAdmin) {throw 'Admin permissions required to start MemoryMappedFile with scope Global!'}
            }

            # Referenced data value
            if ($null -eq $Data) {throw 'Data parameter value can not be null!'}
            if ($Data.Value.GetType().Name -ne 'String') {throw 'Data parameter value type has to be String!'}

            # Name value
            if ([System.String]::IsNullOrWhiteSpace($Name)) {throw 'Name parameter value can not be null or whitespaces only!'}

        # Load MemoryMappedFile if already existing
        if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "MMF map name: $MapName"}
        [System.IntPtr]$handle = [MMFHelper]::OpenFileMapping([MMFHelper]::FILE_MAP_READ, $false, $MapName)
        if ($handle -ne [System.IntPtr]::Zero) {
            if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "MMF handle: $handle"}
            if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "Opening existing MMf..."}
            
            # Open MemoryMappedFile
            try {
                [System.IO.MemoryMappedFiles.MemoryMappedFile]$mmf = [System.IO.MemoryMappedFiles.MemoryMappedFile]::OpenExisting(
                    $MapName,
                    [System.IO.MemoryMappedFiles.MemoryMappedFileRights]::ReadWrite
                )
            } catch {Write-Host $($_ | Out-String).Trim() -ForegroundColor Gray}
        }

        if (-not $mmf) {
            if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "Creating new MMf..."}
            
            # Set security definitions for MemoryMappedFile access
            $security = [System.IO.MemoryMappedFiles.MemoryMappedFileSecurity]::new()
            $security.SetAccessRuleProtection($true, $false)

                # Everyone (Read)
                $everyone = [System.Security.Principal.SecurityIdentifier]::new([System.Security.Principal.WellKnownSidType]::WorldSid, $null)
                $rule     = [System.Security.AccessControl.AccessRule[System.IO.MemoryMappedFiles.MemoryMappedFileRights]]::new(
                    $everyone, 
                    [System.IO.MemoryMappedFiles.MemoryMappedFileRights]::Read, 
                    [System.Security.AccessControl.AccessControlType]::Allow
                )
                $security.AddAccessRule($rule)

                # Administrators (FullControl and TakeOwnership)
                $admins   = [System.Security.Principal.SecurityIdentifier]::new([System.Security.Principal.WellKnownSidType]::BuiltinAdministratorsSid, $null)
                $rule     = [System.Security.AccessControl.AccessRule[System.IO.MemoryMappedFiles.MemoryMappedFileRights]]::new(
                    $admins, 
                    [System.IO.MemoryMappedFiles.MemoryMappedFileRights]::FullControl, 
                    [System.Security.AccessControl.AccessControlType]::Allow
                )
                $security.AddAccessRule($rule)
                $rule     = [System.Security.AccessControl.AccessRule[System.IO.MemoryMappedFiles.MemoryMappedFileRights]]::new(
                    $admins, 
                    [System.IO.MemoryMappedFiles.MemoryMappedFileRights]::TakeOwnership, 
                    [System.Security.AccessControl.AccessControlType]::Allow
                )
                $security.AddAccessRule($rule)

                # System user (FullControl and TakeOwnership)
                $system   = [System.Security.Principal.SecurityIdentifier]::new([System.Security.Principal.WellKnownSidType]::LocalSystemSid, $null)
                $rule     = [System.Security.AccessControl.AccessRule[System.IO.MemoryMappedFiles.MemoryMappedFileRights]]::new(
                    $system, 
                    [System.IO.MemoryMappedFiles.MemoryMappedFileRights]::FullControl, 
                    [System.Security.AccessControl.AccessControlType]::Allow
                )
                $security.AddAccessRule($rule)
                $rule     = [System.Security.AccessControl.AccessRule[System.IO.MemoryMappedFiles.MemoryMappedFileRights]]::new(
                    $system, 
                    [System.IO.MemoryMappedFiles.MemoryMappedFileRights]::TakeOwnership, 
                    [System.Security.AccessControl.AccessControlType]::Allow
                )
                $security.AddAccessRule($rule)

                # Current user (FullControl and TakeOwnership and SetOwner)
                $cuser    = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
                $rule     = [System.Security.AccessControl.AccessRule[System.IO.MemoryMappedFiles.MemoryMappedFileRights]]::new(
                    $cuser, 
                    [System.IO.MemoryMappedFiles.MemoryMappedFileRights]::FullControl, 
                    [System.Security.AccessControl.AccessControlType]::Allow
                )
                $security.AddAccessRule($rule)
                $rule     = [System.Security.AccessControl.AccessRule[System.IO.MemoryMappedFiles.MemoryMappedFileRights]]::new(
                    $cuser, 
                    [System.IO.MemoryMappedFiles.MemoryMappedFileRights]::TakeOwnership, 
                    [System.Security.AccessControl.AccessControlType]::Allow
                )
                $security.AddAccessRule($rule)
                $security.SetOwner($cuser)
            
            # Create MemoryMappedFile object
            [System.IO.MemoryMappedFiles.MemoryMappedFile]$mmf = [System.IO.MemoryMappedFiles.MemoryMappedFile]::CreateOrOpen(
                $MapName, 
                $Size,
                [System.IO.MemoryMappedFiles.MemoryMappedFileAccess]::ReadWrite,
                [System.IO.MemoryMappedFiles.MemoryMappedFileOptions]::None,
                $security,
                [System.IO.HandleInheritability]::None
            )
        }

        # Create MemoryMappedViewAccessor object
        [System.IO.MemoryMappedFiles.MemoryMappedViewAccessor]$accessor = $mmf.CreateViewAccessor()
        if ($WriteHost -and -not $WriteHostResultsOnly) {
            Write-Host (
                "Accessor stats: " + `
                "CanRead=$($accessor.CanRead) | CanWrite=$($accessor.CanWrite) | " + `
                "Capacity=$($accessor.Capacity) | PointerOffset=$($accessor.PointerOffset)"
            )
        }

        # Run loop for memory space writing
        [System.Int64] $i          = 0
        [System.Int64] $valLength  = 0
        [System.String]$fillVal    = [System.Char][System.Byte]0
        [System.String]$val        = [System.String]::Empty
        [System.Byte[]]$buffer     = [System.Byte[]]::new($Size)
        while ($true) {
            # Check for timeout
            if ($TimeoutInSeconds -gt 0) {if ($i -ge $TimeoutInSeconds) {break}}

            # Check for value
            if ($Data.Value -eq 'StopMmfWriting') {break}

            # Add timestamp to value
            $val = "$(Get-Date -Format 'HH:mm:ss'); $($Data.Value)"

            # Prepare data
            if ($val.Length -lt $valLength) {$val = "$($val)$($fillVal*($valLength - $val.Length))"}; $valLength = $val.Length
            $bytes = $enc.GetBytes($val, 0, $val.Length, $buffer, 0)
            if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "Content size: $($buffer.Count)"}

            # Write data to memory space
            $accessor.WriteArray(0, $buffer, 0, $bytes)

            # Finish loop step
            if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "String length: $($valLength)"}
            if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "Content: $val"}
            if ($WriteHost -and $WriteHostResultsOnly -and -not ([System.String]::IsNullOrWhiteSpace($val))) {Write-Host $val}
            if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "Timeout: $(if ($TimeoutInSeconds -gt 0) {"$($i)/$($TimeoutInSeconds) seconds"} else {"None"})"}
            if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "Waiting $($WritingPauseInSeconds) seconds for next content update..."}
            Start-Sleep -Seconds $(if ($WritingPauseInSeconds -gt 0) {$WritingPauseInSeconds} else {1})
            $i = $i + $WritingPauseInSeconds
        }
    }
    catch {
        # Error handling
        Write-Host $($_ | Out-String).Trim() -ForegroundColor Yellow
    }
    finally {
        # Free memory space
        if ($null -ne $accessor) {$accessor.Dispose()}
        if ($null -ne $mmf)      {$mmf.Dispose()}
        if ($handle -ne [IntPtr]::Zero) {[System.Boolean]$cHResult = [MMFHelper]::CloseHandle($handle); $handle = [IntPtr]::Zero}
        [System.GC]::Collect()
    }
}

Function Set-TrayIconSettingPersistent {
    Param (
        [System.Boolean]$State = $true,
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSReference]$Data,
        [System.String]$MmfName,
        [ValidateSet('Local','Global')]
        [System.String]$MmfScope               = 'Local',
        [System.Int32]$TimeoutInSeconds        = 0,
        [System.Int32]$WritingPauseInSeconds   = 1,
        [System.Management.Automation.SwitchParameter]$WriteHost
    )

    try {
        # Checks
            # Admin permissions required
            if ($Scope -eq 'Global') {
                $isAdmin = ([Security.Principal.WindowsPrincipal]([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                if (-not $isAdmin) {throw 'Admin permissions required to start MemoryMappedFile with scope Global!'}
            }

            # Referenced data value
            if ($null -eq $Data) {throw 'Data parameter value can not be null!'}
            if ($Data.Value.GetType().Name -ne 'String') {throw 'Data parameter value type has to be String!'}

        # Start process
        switch ($State) {
            $true {
                # Check Name value
                if ([System.String]::IsNullOrWhiteSpace($MmfName)) {throw 'MmfName parameter value can not be null or whitespaces only!'}
                
                # Create parameters dictionary
                [System.Collections.IDictionary]$dict = @{
                    Data                  = $Data
                    Name                  = $MmfName
                    Scope                 = $MmfScope
                    TimeoutInSeconds      = $TimeoutInSeconds
                    WritingPauseInSeconds = $WritingPauseInSeconds
                }
                
                # Start Runspace and return info
                [System.String]$sct            = "Set-MemoryMappedFile"
                [System.String]$sctContent     = $(Get-Command -Name $sct).Definition
                if ($WriteHost) {Write-Host "Starting runspace with script '$sct' and parameters (Name=$($MmfName) | Scope=$($MmfScope) | TimeoutInSeconds=$($TimeoutInSeconds) | WritingPauseInSeconds=$($WritingPauseInSeconds))..."}
                $Script:PowershellRunspaceRaw  = $([System.Management.Automation.Powershell]::Create([System.Management.Automation.RunspaceMode]::NewRunspace))
                $Script:PowershellRunspace     = $PowershellRunspaceRaw.AddScript($sctContent).AddParameters($dict)
                $Script:PowershellHandle       = $Script:PowershellRunspace.BeginInvoke()
                if ($WriteHost) {
                    Write-Host (
                        "Runspace stats: " + `
                        "Name=$($Script:PowershellRunspace.Runspace.Name) | " + `
                        "Id=$($Script:PowershellRunspace.Runspace.Id) | " + `
                        "InstanceId=$($Script:PowershellRunspace.Runspace.InstanceId) | " + `
                        "Handle=$($Script:PowershellHandle.AsyncWaitHandle.Handle)"
                    )
                }

                # Return
                return $true
            }
            $false {
                # Stop Runspace and return info
                if (Get-Variable -Name 'PowershellRunspace' -Scope 'Script' -ValueOnly -ErrorAction Ignore) {
                    if ($WriteHost) {Write-Host "Stopping runspace..."}
                    [System.String]$dataMemory      = $Data.Value
                    $Data.Value                    = 'StopMmfWriting'
                    [System.Object]$result         = $Script:PowershellRunspace.EndInvoke($Script:PowershellHandle)
                    $Data.Value                    = [System.String]$dataMemory
                    $Script:PowershellRunspace.Dispose()
                    $Script:PowershellRunspaceRaw.Dispose()
                    $Script:PowershellHandle       = $null
                    $Script:PowershellRunspace     = $null
                    $Script:PowershellRunspaceRaw  = $null
                }

                # Return result
                return $result
            }
        }
    }
    catch {
        # Error handling
        Write-Host $($_ | Out-String).Trim() -ForegroundColor Yellow

        # Unload objects
        if (Get-Variable -Name 'PowershellRunspace' -Scope 'Script' -ValueOnly -ErrorAction Ignore) {
            $Script:PowershellRunspace.Dispose()
            $Script:PowershellRunspaceRaw.Dispose()
            $Script:PowershellHandle       = $null
            $Script:PowershellRunspace     = $null
            $Script:PowershellRunspaceRaw  = $null
        }

        # Return
        return $false
    }
}

Function Show-TrayIconPersistent {
    Param (
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]
        [System.String]$MmfName,
        [ValidateSet('Local','Global')]
        [System.String]$MmfScope                 = 'Local',
        [System.Int32]$NoNewDataTimeoutInSeconds = 15,
        [System.Int32]$ReadingPauseInSeconds     = 1,
        [System.Management.Automation.SwitchParameter]$WriteHost
    )
    
    try {
        # Define initial values
        [System.Boolean] $resultBool     = $false
        [System.String]  $mmfVal         = [System.String]::Empty
        [System.String]  $mmfValOld      = [System.String]::Empty

        # Create syncronizing dictionary object for tray icon updates
        $Script:TrayIconDir = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
        foreach ($key in $Script:TrayIconData.Keys) {$success = $Script:TrayIconDir.TryAdd($key, $Script:TrayIconData.$key)}
        $success = $Script:TrayIconDir.TryAdd('Running', $false)

        # Start tray icon
            # Create TrayIcon script
            $TrayIconScript = {
                Param(
                    [System.Object]$Data,
                    [System.Int32]$UpdateIntervalInMilliseconds = 250
                )

                # Define variables
                [System.String]$tsFormat = 'HH:mm:ss:fff'

                # Set environment
                $ErrorActionPreference = 'Stop'

                try {
                    # Add assemblies
                    Write-Information -MessageData "[$(Get-Date -Format $tsFormat)] Loading assemblies..."
                    Add-Type -AssemblyName @('System.Windows.Forms', 'System.Drawing')

                    $code = @"
                        using System;
                        using System.Drawing;
                        using System.Drawing.Drawing2D;
                        using System.Windows.Forms;

                        public class CustomBlueTable : ProfessionalColorTable {
                            public override Color ImageMarginGradientBegin    { get { return Color.RoyalBlue; } }
                            public override Color ImageMarginGradientMiddle   { get { return Color.CornflowerBlue; } }
                            public override Color ImageMarginGradientEnd      { get { return Color.LightSteelBlue; } }
                            public override Color MenuBorder                  { get { return Color.Transparent; } }
                            public override Color MenuItemBorder              { get { return Color.Transparent; } }
                            public override Color ToolStripDropDownBackground { get { return Color.FromArgb(255, 255, 220); } }
                        }

                        public class RoundedSelectionRenderer : ToolStripProfessionalRenderer {
                            public RoundedSelectionRenderer() : base(new CustomBlueTable()) {
                                this.RoundedEdges = false;
                            }

                            protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e) { }

                            protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e) {
                                if (e.Item.Selected && e.Item.Enabled) {
                                    int marginWidth = 28;
                                    int radius      = 4;

                                    Graphics g      = e.Graphics;
                                    g.SmoothingMode = SmoothingMode.AntiAlias;
                                    float wReduct   = e.Item.Width * 0.1f;
                                    Rectangle rect  = new Rectangle(marginWidth + 2, 2, (int)(e.Item.Width - marginWidth - wReduct), e.Item.Height - 4);

                                    using (GraphicsPath path = GetRoundedRect(rect, radius)) {
                                        using (SolidBrush brush = new SolidBrush(Color.FromArgb(180, 220, 220, 220))) {
                                            g.FillPath(brush, path);
                                        }
                                    }
                                }
                            }

                            private GraphicsPath GetRoundedRect(Rectangle baseRect, int radius) {
                                float diameter    = radius * 2.0f;
                                GraphicsPath path = new GraphicsPath();

                                RectangleF arc = new RectangleF(baseRect.Location, new SizeF(diameter, diameter));
                                path.AddArc(arc, 180, 90); // Top left

                                arc.X = baseRect.Right - diameter;
                                path.AddArc(arc, 270, 90); // Top right

                                arc.Y = baseRect.Bottom - diameter;
                                path.AddArc(arc, 0, 90);   // Bottom right

                                arc.X = baseRect.Left;
                                path.AddArc(arc, 90, 90);  // Bottom left

                                path.CloseFigure();
                                return path;
                            }
                        }
"@
                    Add-Type -TypeDefinition $code -ReferencedAssemblies @('System.Windows.Forms', 'System.Drawing')
                    Write-Information -MessageData "[$(Get-Date -Format $tsFormat)] Loading assemblies finished."

                    # Build notification object for tray icon
                    Write-Information -MessageData "[$(Get-Date -Format $tsFormat)] Building notification object..."
                    
                        # Define variables
                        [System.String]$fontType                     = 'Segoe UI'
                        [System.Int32]$fontSizeTitle                 = 11
                        [System.Int32]$fontSizeNormal                = 10
                        [System.Drawing.Color]$contextMenuBackColor  = [System.Drawing.Color]::WhiteSmoke  # [System.Drawing.Color]::FromArgb(255, 255, 220)

                        # Check files
                        [System.Boolean]$FilePathOverviewExists      = $false
                        if (-not [System.String]::IsNullOrWhiteSpace($Data.FilePathOverview)) {$FilePathOverviewExists = $(Test-Path -Path $Data.FilePathOverview)}
                        [System.Boolean]$FilePathLogExists           = $false
                        if (-not [System.String]::IsNullOrWhiteSpace($Data.FilePathLog))      {$FilePathLogExists      = $(Test-Path -Path $Data.FilePathLog)}

                        # Initialize notification object
                        $notification         = [System.Windows.Forms.NotifyIcon]::new()
                        $notification.Icon    = [System.Drawing.SystemIcons]::Information
                        $notification.Text    = $(if ("$($Data.Title)`n$($Data.Subtitle)".Length -ge 63) {"$("$($Data.Title)`n$($Data.Subtitle)".Substring(0,60))..."} else {"$($Data.Title)`n$($Data.Subtitle)"})
                        $notification.Visible = $true

                        # Modify notification object
                            # Change icon
                            if (-not ([System.String]::IsNullOrWhiteSpace($Data.FilePathImage))) {
                                if (Test-Path -Path $Data.FilePathImage) {
                                    $bitmap = [System.Drawing.Bitmap]::FromFile($Data.FilePathImage)
                                    $hIcon = $bitmap.GetHicon()
                                    [System.Object]$iconObj = [System.Drawing.Icon]::FromHandle($hIcon)
                                } 
                                else {
                                    [System.Object]$iconObj = [System.Drawing.SystemIcons]::Error
                                }
                                $notification.Icon = $iconObj
                            }

                            # Add balloon tip values
                            $notification.BalloonTipTitle = $(if ($Data.Title.Length -ge 63)     {"$($Data.Title.Substring(0,60))..."}     else {$Data.Title})
                            $notification.BalloonTipText  = $(if ($Data.Subtitle.Length -ge 255) {"$($Data.Subtitle.Substring(0,252))..."} else {$Data.Subtitle})
                            $notification.BalloonTipIcon  = [System.Windows.Forms.ToolTipIcon]::Info
 
                            # Add events
                            $notification.add_DoubleClick({
                                $notification.ShowBalloonTip(5)
                            })

                            # Add context menu
                                # Create and modify context menu object
                                $contextMenu              = [System.Windows.Forms.ContextMenuStrip]::new()
                                $contextMenu.Visible      = $true
                                $contextMenu.Enabled      = $true
                                $contextMenu.Renderer     = [RoundedSelectionRenderer]::new()
                                $contextMenu.Padding      = [System.Windows.Forms.Padding]::new(0, 2, 0, 2)
                                $contextMenu.Margin       = [System.Windows.Forms.Padding]::Empty
                                $contextMenu.AutoSize     = $true
                                $contextMenu.MinimumSize  = [System.Drawing.Size]::new(200, 30)
                                $contextMenu.MaximumSize  = [System.Drawing.Size]::new(0, 0)
                                $contextMenu.BackColor    = $contextMenuBackColor
                                $contextMenu.LayoutStyle  = [System.Windows.Forms.ToolStripLayoutStyle]::VerticalStackWithOverflow
                                $contextMenu.VerticalScroll.Visible = $false
                                $contextMenu.AllowTransparency = $true
                                $contextMenu.ShowCheckMargin   = $false
                                $contextMenu.ShowImageMargin   = $true
                                $contextMenu.ShowItemToolTips  = $false

                                # Create and modify objects
                                    # Labels
                                    $itemTitle                = [System.Windows.Forms.ToolStripLabel]::new()
                                    $itemTitle.ForeColor      = [System.Drawing.Color]::FromArgb(0, 0, 255)
                                    $itemTitle.Font           = [System.Drawing.Font]::new($fontType, $fontSizeTitle, [System.Drawing.FontStyle]::Bold)
                                    $itemTitle.Text           = $Data.Title
                                    if (-not ([System.String]::IsNullOrWhiteSpace($Data.FilePathTitleImage))) {
                                        if (Test-Path -Path $Data.FilePathTitleImage) {
                                            $bitmap = [System.Drawing.Bitmap]::FromFile($Data.FilePathTitleImage)
                                            $hIcon = $bitmap.GetHicon()
                                            [System.Object]$iconObj = [System.Drawing.Icon]::FromHandle($hIcon)
                                        } 
                                        else {
                                            [System.Object]$iconObj = [System.Drawing.SystemIcons]::Error
                                        }
                                        $itemTitle.Image = $iconObj
                                        $itemTitle.ImageAlign = [System.Drawing.ContentAlignment]::MiddleRight
                                    }
                                    $contextMenu.Items.Add($itemTitle) | Out-Null

                                    $itemSubtitle             = [System.Windows.Forms.ToolStripLabel]::new()
                                    $itemSubtitle.ForeColor   = [System.Drawing.Color]::FromArgb(0, 0, 175)
                                    $itemSubtitle.Font        = [System.Drawing.Font]::new($fontType, $fontSizeNormal, [System.Drawing.FontStyle]::Regular)
                                    $itemSubtitle.Text        = $Data.Subtitle
                                    $contextMenu.Items.Add($itemSubtitle) | Out-Null

                                    # Separator
                                    $itemSep1                 = [System.Windows.Forms.ToolStripSeparator]::new()
                                    $contextMenu.Items.Add($itemSep1)  | Out-Null
                            
                                    # Menu item (button)
                                    $itemButton1              = $contextMenu.Items.Add($Data.MenuTextOverviewFile)
                                    $itemButton1.Visible      = $FilePathOverviewExists
                                    $itemButton1.Add_Click({
                                        try {Start-Process -FilePath $Data.FilePathOverview -WindowStyle Maximized -ErrorAction Stop}
                                        catch {[System.Windows.Forms.MessageBox]::Show("Error:`n$(($_ | Out-String).Trim())")}
                                    })

                                    # Menu item (button)
                                    $itemButton2              = $contextMenu.Items.Add($Data.MenuTextLogFile)
                                    $itemButton2.Visible      = $FilePathLogExists
                                    $itemButton2.Add_Click({
                                        try {Start-Process -FilePath $Data.FilePathLog -WindowStyle Maximized -ErrorAction Stop}
                                        catch {[System.Windows.Forms.MessageBox]::Show("Error:`n$(($_ | Out-String).Trim())")}
                                    })

                                    # Separator
                                    $itemSep2                 = [System.Windows.Forms.ToolStripSeparator]::new()
                                    $contextMenu.Items.Add($itemSep2)  | Out-Null

                                    # Menu item (button)
                                    $itemButtonExit           = $contextMenu.Items.Add($Data.MenuTextExit)
                                    $itemButtonExit.Visible   = $false
                                    $itemButtonExit.Add_Click({
                                        $Data.Running = $false
                                    })

                                # Modify objects in loop process
                                    # Separator
                                    foreach ($sep in @($itemSep1, $itemSep2)) {
                                        $sep.Visible     = $false
                                    }

                                    # Labels
                                    foreach ($lbl in @($itemTitle, $itemSubtitle)) {
                                        $lbl.Visible     = $true
                                        $lbl.Enabled     = $true
                                        $lbl.Margin      = [System.Windows.Forms.Padding]::new(0, 0, 0 , 4)
                                        $lbl.Padding     = [System.Windows.Forms.Padding]::new(0)
                                        $lbl.AutoSize    = $true
                                        $lbl.TextAlign   = [System.Drawing.ContentAlignment]::MiddleLeft
                                    }

                                    # Menu items (buttons)
                                    foreach($item in @($itemButton1, $itemButton2, $itemButtonExit)) {
                                        $item.Enabled   = $true
                                        $item.Margin    = [System.Windows.Forms.Padding]::Empty
                                        $item.Padding   = [System.Windows.Forms.Padding]::new(10, 3, 15, 3)
                                        $item.AutoSize  = $true
                                        $item.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                                        $item.BackColor = [System.Drawing.Color]::Transparent
                                        $item.ForeColor = [System.Drawing.Color]::Black
                                        $item.Font      = [System.Drawing.Font]::new($fontType, $fontSizeNormal, [System.Drawing.FontStyle]::Regular)
                                    }

                                # Finalize context menu object
                                $notification.ContextMenuStrip = $contextMenu
                                
                    Write-Information -MessageData "[$(Get-Date -Format $tsFormat)] Building notification object finished."

                    # Set Timer for updates
                    Write-Information -MessageData "[$(Get-Date -Format $tsFormat)] Creating timer..."
                    $timer = [System.Windows.Forms.Timer]::new()
                    $timer.Interval = $UpdateIntervalInMilliseconds
                    $timer.add_Tick({
                        # Define default values
                        [System.Boolean]$checkVal   = $false
                        [System.Boolean]$menuUpdate = $false
                        
                        # Hide menu elements
                            # Menu items (buttons)
                            if ($Data.FilePathOverviewUpdated) {
                                $Data.FilePathOverviewUpdated = $false
                                if (-not ([System.String]::IsNullOrWhiteSpace($Data.FilePathOverview))) {
                                    $checkVal = $(Test-Path -Path $Data.FilePathOverview)
                                    if ($itemButton1.Visible -ne $checkVal) {
                                        $itemButton1.Visible = $checkVal
                                        $menuUpdate = $true
                                    }
                                }
                            }
                            if ($Data.FilePathLogUpdated) {
                                $Data.FilePathLogUpdated = $false
                                if (-not ([System.String]::IsNullOrWhiteSpace($Data.FilePathLog))) {
                                    $checkVal = $(Test-Path -Path $Data.FilePathLog)
                                    if ($itemButton2.Visible -ne $checkVal) {
                                        $itemButton2.Visible = $checkVal
                                        $menuUpdate = $true
                                    }
                                }
                            }

                            $checkVal = $Data.UserExitAllowed
                            if ($itemButtonExit.Visible -ne $checkVal) {
                                $itemButtonExit.Visible = $checkVal
                                $menuUpdate = $true
                            }

                            # Separators
                            $checkVal = $($itemButton1.Visible -or $itemButton2.Visible)
                            if ($itemSep1.Visible -ne $checkVal) {
                                $itemSep1.Visible = $checkVal
                                $menuUpdate = $true
                            }
                            
                            $checkVal = $itemButtonExit.Visible
                            if ($itemSep2.Visible -ne $checkVal) {
                                $itemSep2.Visible = $checkVal
                                $menuUpdate = $true
                            }

                        # Update Menu elements
                        if ($itemButton1.Text -ne $Data.MenuTextOverviewFile) {
                            $itemButton1.Text = $Data.MenuTextOverviewFile
                            $menuUpdate = $true
                        }

                        if ($itemButton2.Text -ne $Data.MenuTextLogFile) {
                            $itemButton2.Text = $Data.MenuTextLogFile
                            $menuUpdate = $true
                        }

                        if ($itemButtonExit.Text -ne $Data.MenuTextExit) {
                            $itemButtonExit.Text = $Data.MenuTextExit
                            $menuUpdate = $true
                        }
                        
                        # Update Title
                        if ($itemTitle.Text -ne $Data.Title) {
                            $itemTitle.Text               = $Data.Title
                            $notification.BalloonTipTitle = $(if ($Data.Title.Length -ge 63) {"$($Data.Title.Substring(0,60))..."} else {$Data.Title})
                            $menuUpdate = $true
                        }

                        if ($Data.FilePathTitleImageUpdated) {
                            $Data.FilePathTitleImageUpdated = $false
                            if (-not ([System.String]::IsNullOrWhiteSpace($Data.FilePathTitleImage))) {
                                if (Test-Path -Path $Data.FilePathTitleImage) {
                                    $bitmap = [System.Drawing.Bitmap]::FromFile($Data.FilePathTitleImage)
                                    $hIcon = $bitmap.GetHicon()
                                    [System.Object]$iconObj = [System.Drawing.Icon]::FromHandle($hIcon)
                                }
                                else {
                                    [System.Object]$iconObj = [System.Drawing.SystemIcons]::Error
                                }
                                $itemTitle.Image = $iconObj
                            }
                        }

                        # Update Subtitle
                        if ($itemTitle.Text -ne $Data.Subtitle) {
                            $itemSubtitle.Text            = $Data.Subtitle
                            $notification.BalloonTipText  = $(if ($Data.Subtitle.Length -ge 255) {"$($Data.Subtitle.Substring(0,252))..."} else {$Data.Subtitle})
                            $menuUpdate = $true
                        }

                        # Update MouseOverText
                        if ($notification.Text -ne "$($Data.Title)`n$($Data.Subtitle)") {
                            $notification.Text            = $(if ("$($Data.Title)`n$($Data.Subtitle)".Length -ge 63) {"$("$($Data.Title)`n$($Data.Subtitle)".Substring(0,60))..."} else {"$($Data.Title)`n$($Data.Subtitle)"})
                            $menuUpdate = $true
                        }

                        # Update context menu
                        if ($contextMenu.Width -lt $itemTitle.Width) {
                            $contextMenu.MinimumSize = [System.Drawing.Size]::new($itemTitle.Width, 0)
                            $menuUpdate = $true
                        }

                        if ($contextMenu.Width -lt $itemSubtitle.Width) {
                            $contextMenu.MinimumSize = [System.Drawing.Size]::new($itemSubtitle.Width, 0)
                            $menuUpdate = $true
                        }

                        if ($menuUpdate -eq $true) {
                            $contextMenu.PerformLayout()
                            $contextMenu.Size = $contextMenu.PreferredSize
                            $contextMenu.Refresh()
                        }

                        # Update Icon
                        if ($Data.FilePathImageUpdated) {
                            $Data.FilePathImageUpdated = $false
                            if (-not ([System.String]::IsNullOrWhiteSpace($Data.FilePathImage))) {
                                if (Test-Path -Path $Data.FilePathImage) {
                                    $bitmap = [System.Drawing.Bitmap]::FromFile($Data.FilePathImage)
                                    $hIcon = $bitmap.GetHicon()
                                    [System.Object]$iconObj = [System.Drawing.Icon]::FromHandle($hIcon)
                                }
                                else {
                                    [System.Object]$iconObj = [System.Drawing.SystemIcons]::Error
                                }
                                $notification.Icon = $iconObj
                            }
                        }

                        # Exit
                        if ($Data.Running -eq $false) {
                            # Unload objects
                            if ($notification) {
                                $notification.Visible = $false
                                $notification.Dispose()
                            }
                    
                            if ($timer) {
                                Write-Information -MessageData "[$(Get-Date -Format $tsFormat)] Timer stopped."
                                $timer.Stop()
                                $timer.Dispose()
                            }
                            
                            [System.Windows.Forms.Application]::Exit()
                            Write-Information -MessageData "[$(Get-Date -Format $tsFormat)] Application stopped."
                        }
                    })
                    Write-Information -MessageData "[$(Get-Date -Format $tsFormat)] Creating timer finished."

                    # Start timer
                    Write-Information -MessageData "[$(Get-Date -Format $tsFormat)] Timer starting..."
                    $timer.Start()

                    # Start Application
                    Write-Information -MessageData "[$(Get-Date -Format $tsFormat)] Application starting..."
                    $appContext = [System.Windows.Forms.ApplicationContext]::new()
                    [System.Windows.Forms.Application]::Run($appContext)
                    Write-Output "[$(Get-Date -Format $tsFormat)] TrayIcon process successfully executed."
                }
                catch {
                    # Write error object as string
                    Write-Output "[$(Get-Date -Format $tsFormat)] TrayIcon process error:`n$($($_ | Out-String).Trim())"

                    # Unload objects
                    if ($notification) {
                        $notification.Visible = $false
                        $notification.Dispose()
                    }
                    
                    if ($timer) {
                        $timer.Stop()
                        $timer.Dispose()
                    }
                }
            }

            # Start Runspace and return info
            $Script:TrayIconDir.Running    = $true
            $Script:PowershellRunspaceRaw  = $([System.Management.Automation.Powershell]::Create([System.Management.Automation.RunspaceMode]::NewRunspace))
            $Script:PowershellRunspace     = $PowershellRunspaceRaw.AddScript($TrayIconScript).AddArgument($Script:TrayIconDir)
            $Script:PowershellHandle       = $Script:PowershellRunspace.BeginInvoke()
            [System.Object]$rsInfo         = $Script:PowershellRunspace.Streams.Information
            Write-Output $rsInfo

        # Loop for tray icon
        [System.DateTime]$resultDateTime = Get-Date
        while ($true) {
            # Reset check values
            [System.Boolean]$timestampKey   = [System.String]::Empty
            [System.Boolean]$timestampVal   = [System.String]::Empty
            [System.Boolean]$timestampValid = $true
            
            # Get info from sender
            [System.String]$mmfVal = Get-MemoryMappedFile -Name $MmfName -Scope $MmfScope
            if ($WriteHost -and -not ([System.String]::IsNullOrWhiteSpace($mmfVal))) {Write-Host $mmfVal}

            # Check if info is empty
            if (-not ([System.String]::IsNullOrWhiteSpace($mmfVal))) {
                # Check if info is new
                if ($mmfVal -ne $mmfValOld) {
                    # Read values from info string
                    [System.String[]]$mmfArr = $mmfVal.Split(';')
                    foreach ($item in $mmfArr) {
                        if (-not [System.String]::IsNullOrWhiteSpace($item)) {
                            if ($item.Contains('=')) {
                                $key = ($item.Split('=',2)[0]).Trim()
                                $val = ($item.Split('=',2)[1]).Trim()
                                switch ($key) {
                                    'Timestamp'                    {
                                        $timestampKey = $key
                                        $timestampVal = $val
                                        if (-not [System.DateTime]::TryParse($val,[Ref]$resultDateTime)) {$timestampValid = $false; break}
                                    }
                                    'Title'                        {$TrayIconDir.$key = $val; break}
                                    'Subtitle'                     {$TrayIconDir.$key = $val; break}
                                    'MenuTextOverviewFile'         {$TrayIconDir.$key = $val; break}
                                    'MenuTextLogFile'              {$TrayIconDir.$key = $val; break}
                                    'MenuTextExit'                 {$TrayIconDir.$key = $val; break}
                                    'FilePathOverview'             {$TrayIconDir.$key = $val; break}
                                    'FilePathLog'                  {$TrayIconDir.$key = $val; break}
                                    'FilePathImage'                {$TrayIconDir.$key = $val; break}
                                    'FilePathTitleImage'           {$TrayIconDir.$key = $val; break}
                                    'FilePathOverviewUpdated'      {if ([System.Boolean]::TryParse($val,[Ref]$resultBool)) {$TrayIconDir.$key = $resultBool}; break}
                                    'FilePathLogUpdated'           {if ([System.Boolean]::TryParse($val,[Ref]$resultBool)) {$TrayIconDir.$key = $resultBool}; break}
                                    'FilePathImageUpdated'         {if ([System.Boolean]::TryParse($val,[Ref]$resultBool)) {$TrayIconDir.$key = $resultBool}; break}
                                    'FilePathTitleImageUpdated'    {if ([System.Boolean]::TryParse($val,[Ref]$resultBool)) {$TrayIconDir.$key = $resultBool}; break}
                                    'UserExitAllowed'              {if ([System.Boolean]::TryParse($val,[Ref]$resultBool)) {$TrayIconDir.$key = $resultBool}; break}
                                }
                            }
                            else {
                                $key = "<EMPTY>"
                                $val = $item.Trim()
                                $timestampKey = $key
                                $timestampVal = $val
                                if (-not [System.DateTime]::TryParse($val,[Ref]$resultDateTime)) {$timestampValid = $false}
                            }
                        }
                    }

                    # Set info current state
                    $mmfValOld = $mmfVal
                }
            }

            # Break loop if Timestamp is not valid
            if ($timestampValid -ne $true) {Write-Output "Stop - Timestamp is not valid (Key: $timestampKey | Value: $timestampVal)"; break}

            # Break loop if Timestamp is too old
            [System.DateTime]$nowDateTime    = $(Get-Date)
            [System.TimeSpan]$resultTimeSpan = $($nowDateTime - $resultDateTime)
            if ($resultTimeSpan.Seconds -gt $NoNewDataTimeoutInSeconds) {
                [System.String]$nowTimeStamp     = $nowDateTime.ToString('HH:mm:ss')
                [System.String]$lastTimeStamp    = $resultDateTime.ToString('HH:mm:ss')
                Write-Output -InputObject "Stop - Timestamp is too old ($($NoNewDataTimeoutInSeconds) seconds | Now: $nowTimeStamp | Last timestamp: $lastTimeStamp)"
                break
            }

            # Break loop if tray icon is not running anymore (manual exit)
            if ($Script:TrayIconDir.Running -ne $true) {Write-Output "Stop - Tray icon is not running anymore"; break}

            # Wait for next step
            if ($WriteHost -and -not $WriteHostResultsOnly) {Write-Host "Waiting $($ReadingPauseInSeconds) seconds for next content update..."}
            Start-Sleep -Seconds $ReadingPauseInSeconds
        }
    }
    catch {
        # Error handling
        Write-Host $($_ | Out-String).Trim() -ForegroundColor Yellow
        Write-Output $($_ | Out-String).Trim()
    }
    finally {
        # Stop tray icon
            # Stop Runspace and collect information
            $Script:TrayIconDir.Running    = $false
            [System.Object]$rsInfo         = $Script:PowershellRunspace.Streams.Information
            [System.Object]$result         = $Script:PowershellRunspace.EndInvoke($Script:PowershellHandle)
            Write-Output $rsInfo
            
            # Unload objects
            if (Get-Variable -Name 'PowershellRunspace' -Scope 'Script' -ValueOnly -ErrorAction Ignore) {
                $Script:PowershellRunspace.Dispose()
                $Script:PowershellRunspaceRaw.Dispose()
                $Script:PowershellRunspace     = $null
                $Script:PowershellRunspaceRaw  = $null
                $Script:PowershellHandle       = $null
            }
    }

    # Return tray icon process result
    return $result
}

Function Set-TrayIconState {
    <#
    .SYNOPSIS
        Displaying a tray icon with updating values.
    .DESCRIPTION
        Displaying a tray icon in the windows tray icon area (right side of the taskbar). 
        The tray icon runs via a separate script (created by this function). Dynamic updates will be invoked asynchronously via a Runspace and MemoryMappedFile reading process.
        This function also acts as a sender writing asyncrouniously a SettingString in a MemoryMappedFile to update the values of the tray icon.
    .PARAMETER ActionType
        Aliases:                     Enabled
        DataType:                    Boolean
    .PARAMETER TrayIconTitle
        Aliases:                     Title
        DataType:                    String
    .PARAMETER TrayIconSubtitle
        Aliases:                     Subtitle
        DataType:                    String
    .PARAMETER TrayIconMenuTextOverviewFile
        Aliases:                     MenuTextOverviewFile, MenuTextOverview, ButtonOverviewFile, ButtonOverview
        DataType:                    String
    .PARAMETER TrayIconMenuTextLogFile
        Aliases:                     MenuTextLogFile, MenuTextLog, ButtonLogFile, ButtonLog
        DataType:                    String
    .PARAMETER TrayIconMenuTextExit
        Aliases:                     MenuTextExit, ButtonExit
        DataType:                    String
    .PARAMETER TrayIconFilePathOverview
        Aliases:                     FilePathOverview
        DataType:                    String
        Notes:                       Full path to the file
                                     If the path do not exist the tray icon will hide the button.
    .PARAMETER TrayIconFilePathLog
        Aliases:                     FilePathLog
        DataType:                    String
        Notes:                       Full path to the file
                                     If the path do not exist the tray icon will hide the button.
    .PARAMETER TrayIconFilePathImage
        Aliases:                     FilePathImage
        DataType:                    String
        Notes:                       Full path to the file
                                     If the path do not exist the tray icon will hide the button.
    .PARAMETER TrayIconFilePathTitleImage
        Aliases:                     FilePathTitleImage
        DataType:                    String
        Notes:                       Full path to the file
                                     If the path do not exist the tray icon will hide the button.
    .PARAMETER TrayIconFilePathOverviewUpdated
        Aliases:                     FilePathOverviewUpdated, OverviewUpdated
        DataType:                    Boolean
        Notes:                       Tray icon will only update the file path if this value is set to true.
    .PARAMETER TrayIconFilePathLogUpdated
        Aliases:                     FilePathLogUpdated, LogUpdated
        DataType:                    Boolean
        Notes:                       Tray icon will only update the file path if this value is set to true.
    .PARAMETER TrayIconFilePathImageUpdated
        Aliases:                     FilePathImageUpdated, ImageUpdated
        DataType:                    Boolean
        Notes:                       Tray icon will only update the file path if this value is set to true.
    .PARAMETER TrayIconFilePathTitleImageUpdated
        Aliases:                     FilePathTitleImageUpdated, TitleImageUpdated
        DataType:                    Boolean
        Notes:                       Tray icon will only update the file path if this value is set to true.
    .PARAMETER TrayIconUserExitAllowed
        Aliases:                     UserExitAllowed, UserExitEnabled
        DataType:                    Boolean
        Notes:                       Tray icon will only update the file path if this value is set to true.
    .PARAMETER SettingString
        Aliases:                     Setting
        DataType:                    System.Management.Automation.PSReference
        Notes:                       Must be defined as reference, e. g. [ref]$Global:TrayIconSettingString
                                     Reference object must be a string.
    .PARAMETER MmfName
        Aliases:                     Name
        DataType:                    String
        Notes:                       Name of the MemoryMappedFile
    .PARAMETER MmfScope
        Aliases:                     Scope
        DataType:                    String
        Notes:                       Scope of the MemoryMappedFile (Global or Local)
    .PARAMETER NoNewDataTimeoutInSeconds
        Aliases:                     TimeoutInSeconds, Timeout
        DataType:                    Int32
    .PARAMETER ReadingPauseInSeconds
        Aliases:                     PauseInSeconds, Pause
        DataType:                    Int32
    .PARAMETER TempPath
        Aliases:                     TempDirectory, TempDir
        DataType:                    String
        Notes:                       Path to create the client part script to show the tray icon
    .PARAMETER PassThru
        DataType:                    Switch
        Notes:                       Returning the result of the MemoryMappedFile writer.
    .INPUTS
        No pipeline input or default value accepted (will be arguments order).
        SettingString and MmfName are mandatory.
    .OUTPUTS
        Boolean for success
    .NOTES
        Created by Steffen Spanknebel, 34134 Kassel, Germany, at 2026-03-21.
    .LINK
        None
    #>

    [CmdletBinding()]Param (
        [Alias('Enabled')]
        [System.Boolean]$State                               = $true,

        [Alias('Title')]
        [System.String]$TrayIconTitle                        = "<No title defined>",

        [Alias('Subtitle')]
        [System.String]$TrayIconSubtitle                     = "<No subtitle defined>",

        [Alias('MenuTextOverviewFile', 'MenuTextOverview', 'ButtonOverviewFile', 'ButtonOverview')]
        [System.String]$TrayIconMenuTextOverviewFile         = "Open overview file",

        [Alias('MenuTextLogFile', 'MenuTextLog', 'ButtonLogFile', 'ButtonLog')]
        [System.String]$TrayIconMenuTextLogFile              = "Open log file",

        [Alias('MenuTextExit', 'ButtonExit')]
        [System.String]$TrayIconMenuTextExit                 = "Exit",

        [Alias('FilePathOverview')]
        [System.String]$TrayIconFilePathOverview             = [System.String]::Empty,

        [Alias('FilePathLog')]
        [System.String]$TrayIconFilePathLog                  = [System.String]::Empty,

        [Alias('FilePathImage')]
        [System.String]$TrayIconFilePathImage                = [System.String]::Empty,

        [Alias('FilePathTitleImage')]
        [System.String]$TrayIconFilePathTitleImage           = [System.String]::Empty,

        [Alias('FilePathOverviewUpdated', 'OverviewUpdated')]
        [System.Boolean]$TrayIconFilePathOverviewUpdated     = $false,

        [Alias('FilePathLogUpdated', 'LogUpdated')]
        [System.Boolean]$TrayIconFilePathLogUpdated          = $false,

        [Alias('FilePathImageUpdated', 'ImageUpdated')]
        [System.Boolean]$TrayIconFilePathImageUpdated        = $false,

        [Alias('FilePathTitleImageUpdated', 'TitleImageUpdated')]
        [System.Boolean]$TrayIconFilePathTitleImageUpdated   = $false,

        [Alias('UserExitAllowed', 'UserExitEnabled')]
        [System.Boolean]$TrayIconUserExitAllowed             = $false,

        [Parameter(Mandatory=$true)]
        [Alias('Setting')]
        [System.Management.Automation.PSReference]$SettingString,

        [Parameter(Mandatory=$true)]
        [Alias('Name')]
        [ValidateNotNullOrEmpty()]
        [System.String]$MmfName,

        [Alias('Scope')]
        [ValidateSet('Local','Global')]
        [System.String]$MmfScope                             = 'Local',

        [Alias('TimeoutInSeconds', 'Timeout')]
        [System.Int32]$NoNewDataTimeoutInSeconds             = 15,

        [Alias('PauseInSeconds', 'Pause')]
        [System.Int32]$ReadingPauseInSeconds                 = 1,

        [Alias('TempDirectory', 'TempDir')]
        [System.String]$TempPath                             = $env:TEMP,

        [System.Management.Automation.SwitchParameter]$PassThru
    )

    try {
        # Check whether admin permissions are required
        if ($MmfScope -eq 'Global') {
            $isAdmin = ([Security.Principal.WindowsPrincipal]([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            if (-not $isAdmin) {throw 'Admin permissions required to start Set-TrayIconState with scope Global!'}
        }

        # Check referenced data type
        if ($null -eq $SettingString)                         {throw 'SettingString parameter value can not be null!'}
        if ($SettingString.Value.GetType().Name -ne 'String') {throw 'SettingString parameter value type has to be String!'}

        # Check if temp path is available
        if (-not (Test-Path -Path $TempPath)) {throw 'Defined temp path is not available!'}

        switch ($State) {
            $true{
                # Client part
                    # Build script
                        # Section for functions
                        [System.String]$fcnTmp = $(Get-Command -Name @('Show-TrayIconPersistent','Get-MemoryMappedFile') -CommandType Function) | `
                            ForEach-Object -Process {"Function $($_.Name) {$($_.Definition)}`n`n"}
    
                        # Section for invocation
                        [System.String]$invTmp = @"
                            `$ErrorActionPreference = 'Stop'
                            Clear-Host

                            `$TrayIconData                   = [Hashtable]@{
                                Title                        = `"$($TrayIconTitle)`"
                                Subtitle                     = `"$($TrayIconSubtitle)`"
                                MenuTextOverviewFile         = `"$($TrayIconMenuTextOverviewFile)`"
                                MenuTextLogFile              = `"$($TrayIconMenuTextLogFile)`"
                                MenuTextExit                 = `"$($TrayIconMenuTextExit)`"
                                FilePathOverview             = `"$($TrayIconFilePathOverview)`"
                                FilePathLog                  = `"$($TrayIconFilePathLog)`"
                                FilePathImage                = `"$($TrayIconFilePathImage)`"
                                FilePathTitleImage           = `"$($TrayIconFilePathTitleImage)`"
                                FilePathOverviewUpdated      = $(if ($TrayIconFilePathOverviewUpdated)   {'$true'} else {'$false'})
                                FilePathLogUpdated           = $(if ($TrayIconFilePathLogUpdated)        {'$true'} else {'$false'})
                                FilePathImageUpdated         = $(if ($TrayIconFilePathImageUpdated)      {'$true'} else {'$false'})
                                FilePathTitleImageUpdated    = $(if ($TrayIconFilePathTitleImageUpdated) {'$true'} else {'$false'})
                                UserExitAllowed              = $(if ($TrayIconUserExitAllowed)           {'$true'} else {'$false'})
                            }

                            [System.String]`$logFile         = "`$([System.IO.Path]::GetDirectoryName(`$MyInvocation.MyCommand.Definition))\`$([System.IO.Path]::GetFileNameWithoutExtension(`$MyInvocation.MyCommand.Definition)).log"
                            `$result = Show-TrayIconPersistent -MmfName `"$MmfName`" -MmfScope `"$MmfScope`" -NoNewDataTimeoutInSeconds $NoNewDataTimeoutInSeconds -ReadingPauseInSeconds $ReadingPauseInSeconds -WriteHost
                            `$(`$result | Out-String).Trim() | Out-File -FilePath `$logFile -Force
                            `$(`$error  | Out-String).Trim() | Out-File -FilePath `$logFile -Force -Append
"@
                        [System.String]$invTmpNew    = [System.String]::Empty
                        [System.String[]]$invTmpArr  = $invTmp.Split("`n")
                        [System.Int32]$invTmpIndent  = $($($invTmpArr[0]).Length - $($invTmpArr[0].Trim()).Length)
                        foreach ($line in $invTmpArr) {if ($line.Length -gt $invTmpIndent) {$invTmpNew += "$($line.Substring($invTmpIndent))`n"} else {$invTmpNew += "$($line)`n"}}; $invTmpNew = $invTmpNew.Trim()

                    # Create script file (30 seconds tolerance)
                    [System.String]$TempScriptPath = "$($TempPath.TrimEnd('\'))\$($MyInvocation.MyCommand.Name).ps1"
                    [System.Int32]$iMax            = 30
                    [System.Int32]$i               = 0
                    while ($i -lt $iMax) {
                        try {"$($fcnTmp)`n$($invTmpNew)" | Out-File -FilePath $TempScriptPath -Encoding utf8 -Force; break} 
                        catch {Start-Sleep -Seconds 5; $i = ($i + 5)}
                    }
                    if ($i -ge $iMax) {throw "Can not create/overwrite client script at '$TempScriptPath'!"}
    
                    # Start script
                    [System.String]$psExe    = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
                    [System.String]$psParams = "-NoLogo -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy ByPass -File `"$TempScriptPath`""
                    $result = Start-Process -FilePath $psExe -ArgumentList $psParams -WindowStyle Hidden -PassThru

                # Server part
                    # Start process
                    $result = Set-TrayIconSettingPersistent -State $State -Data $SettingString -MmfName $MmfName -MmfScope $MmfScope -TimeoutInSeconds 0 -WritingPauseInSeconds 1 -WriteHost
                    if ($result -eq $false) {return $false}

                # Return process result
                if ($PassThru) {return $result}
                return $true
            }
            $false {
                # Server part
                    # Stop process
                    $result = Set-TrayIconSettingPersistent -State $State -Data $SettingString -MmfName $MmfName
                    if ($result -eq $false) {return $false}
                    return $true
            }
        }
    }
    catch {
        # Error handling
        Write-Host $($_ | Out-String -Width 1024).Trim() -ForegroundColor Yellow
        return $false
    }
}


# ---  Main  ---
$ErrorActionPreference = 'Stop'
Clear-Host

[System.Int32] $pauseDur                     = 5
[System.String]$Global:TrayIconSettingString = [System.String]::Empty
[System.Collections.Hashtable]$Params        = @{
    SettingString                            = $([ref]$Global:TrayIconSettingString)
    MmfName                                  = 'SPKTrayIcon'
    MmfScope                                 = 'Local'
    TrayIconTitle                            = "accompio Software Packaging Kit"
    TrayIconSubtitle                         = "Initializing ..."
    TrayIconFilePathOverview                 = "C:\Temp\PackageOverview.txt"
    TrayIconFilePathLog                      = "C:\Temp\Package.log"
    TrayIconFilePathImage                    = "C:\Temp\Logo.png"
    TrayIconFilePathTitleImage               = "C:\Temp\Logo.png"
    TrayIconMenuTextOverviewFile             = 'Übersicht anzeigen'
    TrayIconMenuTextLogFile                  = 'Protokoll anzeigen'
    TrayIconMenuTextExit                     = 'Beenden'
    NoNewDataTimeoutInSeconds                = 15
}

if (-not (Set-TrayIconState -State $true @Params)) {return}

Start-Sleep -Seconds $pauseDur
$Global:TrayIconSettingString = "Title=First TITLE; Subtitle = First Subtitle"
Start-Sleep -Seconds $pauseDur
$Global:TrayIconSettingString = "Title=Install Google Chrome 1.2.3; Subtitle = Phase: Install application"
Start-Sleep -Seconds $pauseDur
$Global:TrayIconSettingString = "FilePathImage=C:\Temp\Hello.png;FilePathImageUpdated=true; UserExitAllowed=true"
Start-Sleep -Seconds $pauseDur
$Global:TrayIconSettingString = "Title=This TITLE is loooooooooooooooooooooooooooooooooooooooooooooog; FilePathTitleImage=C:\Temp\Hello.png;FilePathTitleImageUpdated=true"
Start-Sleep -Seconds $pauseDur
$Global:TrayIconSettingString = "Title=The last TITLE; Subtitle = Subtitle as end`nwith this new line`nand with this new line; FilePathTitleImage=C:\Temp\Hello.png;FilePathTitleImageUpdated=true"
Start-Sleep -Seconds $pauseDur

if (-not (Set-TrayIconState -State $false @Params)) {return}


<#TEST
Title: Software package name
Subtitle: Install status and defer history of the package
#>