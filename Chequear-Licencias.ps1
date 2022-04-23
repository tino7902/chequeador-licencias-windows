#
#requires -version 4
<#
.SYNOPSIS
    Chequea el estado de activación de las licencias de Office y Windows.
.DESCRIPTION
    El script genera una lista de los datos de licencias del sistema operativo y otros productos de Microsoft como Office
.PARAMETER ComputerName
    El parámetro ComputerName pide el nombre de la computadora a chequear o su dirección IP, si no se completa, usará la computadora en la que se ejecute el comando como valor
.PARAMETER TextFile
    Este parámetro pide obligatoriamente el nombre de un archivo de texto. El archivo de texto debe incluir un listado de nombres de computadoras o direcciones IP las cuales seran chequeadas.
.PARAMETER Protocol
    Este parámetro es opcional. Elige cual de los dos protocolos usar para conectarse remotamente a las computadoras. El default es DCOM que funciana en cualquier caso. La otra opción es WSman que requiere que PS remoting se encuentre activado.
.PARAMETER Credential
    Este parámetro opcional pide credenciales para conectarse como un usuario distinto al que inicio sesión. Si no se incluye se usará el usuario que haya iniciado sesión.
.INPUTS
    [System.String]
.OUTPUTS
    [System.Object]
.NOTES
    Version:        1.0
    Author:         Kunal Udapi
    Creation Date:  20 September 2017
    Purpose/Change: Get windows office and OS licensing information.
    Useful URLs: http://vcloud-lab.com
.EXAMPLE
    PS C:\>.\Chequear-Licencias -ComputerName computadora1,computadora2,computadora3

    Este ejemplo hace un listado de los detalles de licencia de las computadoras ingresadas, usando el protocolo DCom. Usa el mismo usuario que haya iniciado sesión
.EXAMPLE
    PS C:\>.\Chequear-Licencias -ComputerName Server01 -Protocol Wsman -Crdential

    Este ejemplo hace un listado de los detalles de licencia de las computadoras ingresadas, usando el protocolo WSmam. al agregar el parámetro credential pedirá un usuario y contraseña para conectarse a las computadoras remotamente.
.EXAMPLE
    PS C:\>.\Chequear-Licencias -TextFile C:\Temp\list.txt

    Text file has computer name list, information is collected using wmi (DCom) protocol, this will try to connect remote computers with currently logged in user account.
    El archivo de texto contiene una lista de computadoras que serán chequeadas. La información se consigue con el protocolo DCom, se intentará conectarse a las computadoras utilizando el mismo usuario que haya iniciado sesión.
.EXAMPLE
    PS C:\>.\Chequear-Licencias -TextFile C:\Temp\list.txt -Protocol Wsman -Crdential

    El archivo de texto contiene una lista de computadoras que serán chequeadas. La información se consigue con el protocolo WSman, pedirá un usuario y contraseña.

#>
[CmdletBinding(SupportsShouldProcess=$True,
    ConfirmImpact='Medium',
    HelpURI='http://vcloud-lab.com',
    DefaultParameterSetName='CN')]
Param
(
    [parameter(Position=0, Mandatory=$True, ParameterSetName='File', ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage='Ingresa la dirección de un archivo de texto válido')]
    [ValidateScript({
        If(Test-Path $_){$true}else{throw "Invalid path given: $_"}
    })]
    [alias('File')]
    [string]$TextFile,
    
    [parameter(Position=0, Mandatory=$false, ParameterSetName='CN', ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage='Ingresa la dirección de un archivo de texto válido')]
    [alias('CN', 'Name')]
    [String[]]$ComputerName = $env:COMPUTERNAME,

    [parameter(ParameterSetName = 'File', Position=1, Mandatory=$false, HelpMessage='Ingresa la dirección de un archivo de texto válido')]
    [parameter(ParameterSetName = 'CN', Position=0, Mandatory=$false)]
    [ValidateSet('Dcom','Default','Wsman')]
    [String]$Protocol = 'Dcom',

    [parameter(ParameterSetName = 'File', Position=2, Mandatory=$false)]
    [parameter(ParameterSetName = 'CN', Position=2, Mandatory=$false)]    
    [Switch]$Credential
)
Begin {
    #[String[]]$ComputerName = $env:COMPUTERNAME
    if ($Credential.IsPresent -eq $True) {
        $Cred = Get-Credential -Message 'Type domain credentials to connect remote Server' -UserName (WhoAmI)
    }
    $CimSessionOptions = New-CimSessionOption -Protocol $Protocol
    $Query = "Select * from  SoftwareLicensingProduct Where PartialProductKey LIKE '%'"
}
Process {
    switch ($PsCmdlet.ParameterSetName) {
        'CN' {
            Break
        }
        'File' {
            $ComputerName = Get-Content $TextFile
            Break
        }
    }
    foreach ($Computer in $ComputerName) {
        if (-not(Test-Connection -ComputerName $Computer -Count 2 -Quiet)) {
            Write-Host -BackgroundColor DarkYellow ([char]8734) -NoNewline
            # Write-Host " $Computer is not reachable, ICMP may be disabled...."
            Write-Host " No se puede conectar a $Computer , ICMP puede estar desactivado...."
            #Break
        }
        else {
            Write-Host -BackgroundColor DarkGreen ([char]8730) -NoNewline
            # Write-Host " $Computer is reachable connecting...."
            Write-Host " Conectando a $Computer ...."
        }
        try {
            if ($Credential.IsPresent -eq $True) {
                $Cimsession = New-CimSession -Name $Computer -ComputerName $Computer -SessionOption $CimSessionOptions -Credential $Cred -ErrorAction Stop
            }
            else {
                $Cimsession = New-CimSession -Name $Computer -ComputerName $Computer -SessionOption $CimSessionOptions  -ErrorAction Stop
            }
            $LicenseInfo = Get-CimInstance -Query $Query -CimSession $Cimsession -ErrorAction Stop 
            Switch ($LicenseInfo.LicenseStatus) {
                0 {$LicenseStatus = 'Unlicensed'; Break}
                1 {$LicenseStatus = 'Licensed'; Break}
                2 {$LicenseStatus = 'OOBGrace'; Break}
                3 {$LicenseStatus = 'OOTGrace'; Break}
                4 {$LicenseStatus = 'NonGenuineGrace'; Break}
                5 {$LicenseStatus = 'Notification'; Break}
                6 {$LicenseStatus = 'ExtendedGrace'; Break}
            } 
            $LicenseInfo | Select-Object PSComputerName, Name, @{N = 'LicenseStatus'; E={$LicenseStatus}},AutomaticVMActivationLastActivationTime, Description, GenuineStatus, GracePeriodRemaining, LicenseFamily, PartialProductKey, RemainingSkuReArmCount, IsKeyManagementServiceMachine #, ApplicationID
            }
        catch {
            Write-Host -BackgroundColor DarkRed ([char]215) -NoNewline
            # Write-Host " Cannot fetch information from $Computer" 
            Write-Host " No se puede conseguir información de $Computer" 
        }
    }
}