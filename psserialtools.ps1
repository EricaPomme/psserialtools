function Write-SerialPort {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]$InputObject,
        [Parameter()][string]$PortName = 'COM1',
        [Parameter()][Int32]$BaudRate = 115200,
        [Parameter()][Int32]$DataBits = 8,
        [Parameter()][IO.Ports.Parity]$Parity = 'None',
        [Parameter()][IO.Ports.StopBits]$StopBits = 1,
        [Parameter()][string]$Prefix = "",
        [Parameter()][string]$Suffix = "`r`n"
)

<#
    .SYNOPSIS
    Writes a string (or strings, if fed a collection) to a serial port.

    .PARAMETER InputObject
    String or string-like object. May be omitted in favour of reading from pipeline.

    .PARAMETER PortName
    Name of the port to send to. Omit colon. (Default: 'COM1')

    .PARAMETER BaudRate
    Baud rate for port. (Default: 115200)

    .PARAMETER DataBits
    Number of data bits. (Default: 8)

    .PARAMETER Parity
    Parity setting. (Default: 'None')

    .PARAMETER StopBits
    Number of stop bits. (Default: 1)

    .INPUTS
    String or string-like object or collection of objects.

    .OUTPUTS
    None.

    .EXAMPLE
    "Hello World!" | Write-SerialPort

    Would write "Hello World!" to COM1:, with the settings 115200 baud, 8 data bits, No parity, 1 stop bit.

    .EXAMPLE
    Write-SerialPort -String (Get-Date) -PortName COM4

    Would write the current date to COM4:, with the settings 115200 baud, 8 data bits, No parity, 1 stop bit.
#>

    BEGIN {
        $port = [IO.Ports.SerialPort]::new($PortName, $BaudRate, $Parity, $DataBits, $StopBits)
        Write-Debug "Opening $($port.PortName) ($($port.BaudRate), $($port.Parity), $($port.DataBits))."
        $port.Open()
    }
    PROCESS {
        foreach ( $obj in $InputObject ) {
            $msg = "$($PREFIX)$obj$($SUFFIX)"
            Write-Verbose "Sending $msg to $($port.PortName)."
            $port.Write($msg)
        }
    }
    END {
        Write-Debug "Closing $($port.PortName)."
        $port.Close()
    }
}
