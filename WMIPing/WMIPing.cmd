@echo off & setlocal

set "PingCnt=1"
set "TimeOut=1000"
set "PingLoop="
set "TTL=80"
set "DontFragment="
set "ResolveIPAddress="
set "HostAddress="
set "StatusCode="
set "AdressResolutionStatusCode="


:ParseCommandLine
  if /i "%~1" equ "-n" (
    set /a "PingCnt=%~2"
    shift & shift

  ) else if /i "%~1" equ "/n" (
    set /a "PingCnt=%~2"
    shift & shift

  ) else if /i "%~1" equ "-i" (
    set /a "TTL=%~2"
    shift & shift

  ) else if /i "%~1" equ "/i" (
    set /a "TTL=%~2"
    shift & shift

  ) else if /i "%~1" equ "-w" (
    set /a "TimeOut=%~2"
    shift & shift

  ) else if /i "%~1" equ "/w" (
    set /a "TimeOut=%~2"
    shift & shift

  ) else if /i "%~1" equ "-t" (
    set "PingLoop=1"
    shift

  ) else if /i "%~1" equ "/t" (
    set "PingLoop=1"
    shift

  ) else if /i "%~1" equ "-a" (
    set "ResolveIPAddress=1"
    shift

  ) else if /i "%~1" equ "/a" (
    set "ResolveIPAddress=1"
    shift

  ) else if /i "%~1" equ "-f" (
    set "DontFragment=1"
    shift

  ) else if /i "%~1" equ "/f" (
    set "DontFragment=1"
    shift

  ) else if /i "%~1" equ "-?" (
    call :PrintHelp
    exit /b 0

  ) else if /i "%~1" equ "/?" (
    call :PrintHelp
    exit /b 0

  ) else (
    set "HostAddress=%~1"
    shift
  )
if "%~1" neq "" goto :ParseCommandLine

if not defined HostAddress (
  echo Error: No host address provided
  exit /b 1
)

if %PingCnt% lss 1 (
  echo Error: Parameter PingCount has to be greater 0
  exit /b 2
)

if %TimeOut% lss 0 (
  echo Error: Parameter TimeOut has to be greater 0
  exit /b 2
)


set "WMINameSpace=/namespace:\\root\cimv2"
set "WMIClass=Win32_PingStatus"

set "WMIQueryFilter=Address='%HostAddress%'"
set "WMIQueryFields=PrimaryAddressResolutionStatus^,StatusCode"

if defined ResolveIPAddress (
  set "WMIQueryFilter=%WMIQueryFilter% and ResolveAddressNames='True'"
  set "WMIQueryFields=%WMIQueryFields%^,ProtocolAddressResolved"
) else (
  set "WMIQueryFields=%WMIQueryFields%^,ProtocolAddress"
)

if defined TTL (
  set "WMIQueryFilter=%WMIQueryFilter% and TimeToLive=%TTL%"
)

if defined DontFragment (
  set "WMIQueryFilter=%WMIQueryFilter% and NoFragmentation='True'"
)

set "WMIQueryFilter=%WMIQueryFilter% and Timeout=%TimeOut%"


set "WMICommand=path %WMIClass% where "%WMIQueryFilter%" get %WMIQueryFields%"


:PingLoop
for /l %%i in (1, 1, %PingCnt%) do (
  for /f "usebackq tokens=1,2 delims==" %%a in (`wmic %WMINameSpace% %WMICommand% /value 2^>NUL`) do (
    if /i "%%a" equ "PrimaryAddressResolutionStatus" (
      for %%z in (%%b) do set "AdressResolutionStatusCode=%%z"
    )

    if /i "%%a" equ "StatusCode" (
      for %%z in (%%b) do set "StatusCode=%%z"
    )

    if /i "%%a" equ "ProtocolAddress" (
      for %%z in (%%b) do set "Host=%%z"
    )

    if /i "%%a" equ "ProtocolAddressResolved" (
      for %%z in (%%b) do set "Host=%%z"
    )
  )

  call :PrintStatusMessage
  if not defined StatusCode exit /b 3
)

if defined PingLoop goto :PingLoop

if %StatusCode% equ 0 exit /b 0
exit /b 4



:PrintStatusMessage
  if not defined StatusCode (
    if "%AdressResolutionStatusCode%" neq "0" (
      echo Error: Address resolution failed
    ) else (
      echo Error: Internal error
    )
    
    exit /b
  )

  if %StatusCode% equ 0 (
    echo Ping to host address %HostAddress% [%Host%] successfull

  ) else if %StatusCode% equ 11001 (
    echo Error: Buffer too small

  ) else if %StatusCode% equ 11002 (
    echo Error: Destination net unreachable

  ) else if %StatusCode% equ 11003 (
    echo Error: Destination host unreachable

  ) else if %StatusCode% equ 11004 (
    echo Error: Destination protocol unreachable

  ) else if %StatusCode% equ 11005 (
    echo Error: Destination port unreachable

  ) else if %StatusCode% equ 11006 (
    echo Error: No resources

  ) else if %StatusCode% equ 11007 (
    echo Error: Bad option

  ) else if %StatusCode% equ 11008 (
    echo Error: Hardware error

  ) else if %StatusCode% equ 11009 (
    echo Error: Packet too big

  ) else if %StatusCode% equ 11010 (
    echo Error: Request timed out

  ) else if %StatusCode% equ 11011 (
    echo Error: Bad request

  ) else if %StatusCode% equ 11012 (
    echo Error: Bad route

  ) else if %StatusCode% equ 11013 (
    echo Error: TimeToLive expired transit

  ) else if %StatusCode% equ 11014 (
    echo Error: TimeToLive expired reassembly

  ) else if %StatusCode% equ 11015 (
    echo Error: Parameter problem

  ) else if %StatusCode% equ 11016 (
    echo Error: Source quench

  ) else if %StatusCode% equ 11017 (
    echo Error: Option too big

  ) else if %StatusCode% equ 11018 (
    echo Error: Bad destination

  ) else if %StatusCode% equ 11032 (
    echo Error: Negotiating IPSEC

  ) else if %StatusCode% equ 11050 (
    echo Error: General failure

  ) else (
    echo Error: Unknown error
  )
exit /b



:PrintHelp
  echo(
  echo(Syntax: wmiping [-t] [-a] [-f] [-n Count] [-i TTL] [-w Time limit] Destination
  echo(
  echo(Options:
  echo(    -t              Send continuously ping packets to the destination host.
  echo(                    Press CTRL+C or CTRL+Break to cancel the operation.
  echo(    -a              Resolve IP address to host name.
  echo(    -f              Set "Don't fragment" bit.
  echo(    -n Count        Number of ping packets to send. Default: 1
  echo(    -i TTL          Life span of a ping packet (in number of hops). Default: 80
  echo(    -w Time limit   Time limit in milliseconds for a response. Default: 1000
  echo(    Destination     Host name or IP address to send ping packets to.
exit /b
