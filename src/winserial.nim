#
## Serial library for Windows.
##
#
## MIT Licence:
##  Permission is hereby granted, free of charge, to any person obtaining a copy of
##  this software and associated documentation files (the "Software"), to deal in
##  the Software without restriction, including without limitation the rights to
##  use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
##  the Software, and to permit persons to whom the Software is furnished to do so,
##  subject to the following conditions:
##  - The above copyright notice and this permission notice shall be included in
##    all copies or substantial portions of the Software.
##  - The Software is provided "as is", without warranty of any kind, express or
##    implied, including but not limited to the warranties of merchantability,
##    fitness for a particular purpose and noninfringement. In no event shall the
##    authors or copyright holders be liable for any claim, damages or other
##    liability, whether in an action of contract, tort or otherwise, arising from,
##    out of or in connection with the Software or the use or other dealings in the
##    Software.
##
##
##

{.deadCodeElim: on.}

import winim/lean
import os

# Public Constants

# Output Control Lines (CommSetLine)
const
  LINE_BREAK* = 1
  LINE_DTR* = 2
  LINE_RTS* = 3

# Input Control Lines (CommGetLines)
const
  LINE_CTS* = 0x10
  LINE_DSR* = 0x20
  LINE_RING* = 0x40
  LINE_RLSD* = 0x80
  LINE_CD* = 0x80

# System Constants
const
  ERROR_IO_INCOMPLETE* = 996
  ERROR_IO_PENDING* = 997
  GENERIC_READ* = 0x80000000
  GENERIC_WRITE* = 0x40000000
  FILE_ATTRIBUTE_NORMAL* = 0x80
  FILE_FLAG_OVERLAPPED* = 0x40000000
  FORMAT_MESSAGE_FROM_SYSTEM* = 0x1000
  OPEN_EXISTING* = 3

# COMM Functions
const
  MS_CTS_ON* = 0x10
  MS_DSR_ON* = 0x20
  MS_RING_ON* = 0x40
  MS_RLSD_ON* = 0x80
  PURGE_RXABORT* = 0x2
  PURGE_RXCLEAR* = 0x8
  PURGE_TXABORT* = 0x1
  PURGE_TXCLEAR* = 0x4

# COMM Escape Functions
const
  CLRBREAK* = 9
  CLRDTR* = 6
  CLRRTS* = 4
  SETBREAK* = 8
  SETDTR* = 5
  SETRTS* = 3

#-------------------------------------------------------------------------------
# System Functions
#-------------------------------------------------------------------------------


proc BuildCommDCB*(lpDef: LPCSTR, lpDCB: LPDCB): WINBOOL {.stdcall, dynlib: "kernel32", importc: "BuildCommDCBA".}
  ## Fills a specified DCB structure with values specified in a device-control string.


proc ClearCommError*(hFile: HANDLE, lpErrors: LPDWORD, lpStat: LPCOMSTAT): WINBOOL {.stdcall, dynlib: "kernel32", importc.}
  ## Retrieves information about a communications error and reports
  ## the current status of a communications device. The function is
  ## called when a communications error occurs, and it clears the
  ## device's error flag to enable additional input and output
  ## (I/O) operations.

proc CloseHandle*(hObject: HANDLE): WINBOOL {.stdcall, dynlib: "kernel32", importc.}
  ## Closes an open communications device or file handle.

proc CreateFile*(lpFileName: LPCSTR, dwDesiredAccess: DWORD, dwShareMode: DWORD, lpSecurityAttributes: LPSECURITY_ATTRIBUTES, dwCreationDisposition: DWORD, dwFlagsAndAttributes: DWORD, hTemplateFile: HANDLE): HANDLE {.stdcall, dynlib: "kernel32", importc: "CreateFileA".}
  ## Creates or opens a communications resource and returns a handle that can be used to access the resource.


proc EscapeCommFunction*(hFile: HANDLE, dwFunc: DWORD): WINBOOL {.stdcall, dynlib: "kernel32", importc.}
  ## Directs a specified communications device to perform a function.


proc FormatMessage*(dwFlags: DWORD, lpSource: LPCVOID, dwMessageId: DWORD, dwLanguageId: DWORD, lpBuffer: LPSTR, nSize: DWORD, Arguments: ptr va_list): DWORD {.stdcall, dynlib: "kernel32", importc: "FormatMessageA".}
  ## Formats a message string such as an error string returned by anoher function.


proc GetCommModemStatus*(hFile: HANDLE, lpModemStat: LPDWORD): WINBOOL {.stdcall, dynlib: "kernel32", importc.}
  ## Retrieves modem control-register values.

proc GetCommState*(hFile: HANDLE, lpDCB: LPDCB): WINBOOL {.stdcall, dynlib: "kernel32", importc.}
  ## Retrieves the current control settings for a specified communications device.


proc GetLastError*(): DWORD {.stdcall, dynlib: "kernel32", importc.}
  ## Retrieves the calling thread's last-error code value.


proc GetOverlappedResult*(hFile: HANDLE, lpOverlapped: LPOVERLAPPED, lpNumberOfBytesTransferred: LPDWORD, bWait: WINBOOL): WINBOOL {.stdcall, dynlib: "kernel32", importc.}
  ## Retrieves the results of an overlapped operation on the specified file, named pipe, or communications device.


proc PurgeComm*(hFile: HANDLE, dwFlags: DWORD): WINBOOL {.stdcall, dynlib: "kernel32", importc.}
  ## Discards all characters from the output or input buffer of a specified communications resource. It can also terminate pending read or write operations on the resource.


proc ReadFile*(hFile: HANDLE, lpBuffer: LPVOID, nNumberOfBytesToRead: DWORD, lpNumberOfBytesRead: LPDWORD, lpOverlapped: LPOVERLAPPED): WINBOOL {.stdcall, dynlib: "kernel32", importc.}
  ## Reads data from a file, starting at the position indicated by the
  ## file pointer. After the read operation has been completed, the
  ## file pointer is adjusted by the number of bytes actually read,
  ## unless the file handle is created with the overlapped attribute.
  ## If the file handle is created for overlapped input and output
  ## (I/O), the application must adjust the position of the file pointer
  ## after the read operation.


proc SetCommState*(hFile: HANDLE, lpDCB: LPDCB): WINBOOL {.stdcall, dynlib: "kernel32", importc.}
  ## Configures a communications device according to the specifications
  ## in a device-control block (a DCB structure). The function
  ## reinitializes all hardware and control settings, but it does not
  ## empty output or input queues.


proc SetCommTimeouts*(hFile: HANDLE, lpCommTimeouts: LPCOMMTIMEOUTS): WINBOOL {.stdcall, dynlib: "kernel32", importc.}
  ## Sets the time-out parameters for all read and write operations on a specified communications device.


proc SetupComm*(hFile: HANDLE, dwInQueue: DWORD, dwOutQueue: DWORD): WINBOOL {.stdcall, dynlib: "kernel32", importc.}
  ## Initializes the communications parameters for a specified communications device.


proc WriteFile*(hFile: HANDLE, lpBuffer: LPCVOID, nNumberOfBytesToWrite: DWORD, lpNumberOfBytesWritten: LPDWORD, lpOverlapped: LPOVERLAPPED): WINBOOL {.stdcall, dynlib: "kernel32", importc.}
  ## Writes data to a file and is designed for both synchronous and a
  ## synchronous operation. The function starts writing data to the file
  ## at the position indicated by the file pointer. After the write
  ## operation has been completed, the file pointer is adjusted by the
  ## number of bytes actually written, except when the file is opened with
  ## FILE_FLAG_OVERLAPPED. If the file handle was created for overlapped
  ## input and output (I/O), the application must adjust the position of
  ## the file pointer after the write operation is finished.

#-------------------------------------------------------------------------------
# Program Constants
#-------------------------------------------------------------------------------

const MAX_PORTS = 4

type
  COMM_ERROR = object
    lngErrorCode: DWORD
    strFunction: string
    strErrorMessage: string

type 
  COMM_PORT = object
    lngHandle: HANDLE
    blnPortOpen: bool
    udtDCB: DCB

#-------------------------------------------------------------------------------
# Program Storage
#-------------------------------------------------------------------------------
var udtCommOverlap*:OVERLAPPED
var udtCommError*:COMM_ERROR
var udtPorts*: array[MAX_PORTS, COMM_PORT]


proc GetSystemMessage*(lngErrorCode: DWORD): string =
  ## Gets system error text for the specified error code.
  var strMsgBuff = cast[LPSTR](newWString(int 256))
  discard FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, nil, lngErrorCode, 0, strMsgBuff, 255, nil)
  return $(strMsgBuff)


proc SetCommError*(strFunction:string): DWORD {.discardable.} =
  udtCommError.lngErrorCode = GetLastError()
  udtCommError.strFunction = strFunction
  udtCommError.strErrorMessage = GetSystemMessage(udtCommError.lngErrorCode)
  return udtCommError.lngErrorCode
  
proc SetCommErrorEx*(strFunction:string, lngHnd: HANDLE): DWORD {.discardable.}= 
  var lngErrorFlags: DWORD
  var udtCommStat: COMSTAT

  udtCommError.lngErrorCode = GetLastError()
  udtCommError.strFunction = strFunction
  udtCommError.strErrorMessage = GetSystemMessage(udtCommError.lngErrorCode)

  discard ClearCommError(lngHnd, &lngErrorFlags, &udtCommStat)
  udtCommError.strErrorMessage = udtCommError.strErrorMessage & "  COMM Error Flags = " & toHex(lngErrorFlags)
  
  return udtCommError.lngErrorCode

proc CommOpen*(intPortID: WORD, strPort: string, strSettings:string): DWORD {.discardable.}=
  ## Opens/Initializes serial port.
  ##
  ##
  ## Parameters:
  ##   - intPortID   - Port ID used when port was opened.
  ##   - strPort     - COM port name. (COM1, COM2, COM3, COM4)
  ##   - strSettings - Communication settings. Example: "baud=9600 parity=N data=8 stop=1"
  ##
  ## Returns:
  ##   - Error Code  - 0 = No Error.
  ##
  var lngStatus:DWORD
  var udtCommTimeOuts:COMMTIMEOUTS

  if udtPorts[intPortID].blnPortOpen:
    lngStatus = -1
    udtCommError.lngErrorCode = lngStatus
    udtCommError.strFunction = "CommOpen"
    udtCommError.strErrorMessage = "Port in use."
    return lngStatus
  
  # Open serial port.
  udtPorts[intPortID].lngHandle = CreateFile(strPort, cast[DWORD](GENERIC_READ or GENERIC_WRITE), 0, nil, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

  if udtPorts[intPortID].lngHandle == -1:
    SetCommError("CommOpen (CreateFile)")
    echo udtCommError
    return udtCommError.lngErrorCode

  udtPorts[intPortID].blnPortOpen = true

  # Setup device buffers (1K each).
  lngStatus = SetupComm(udtPorts[intPortID].lngHandle, 1024, 1024)
  if lngStatus == 0: 
    SetCommError("CommOpen (SetupComm)")
    echo udtCommError
    return udtCommError.lngErrorCode
  
  # Purge buffers.
  lngStatus = PurgeComm(udtPorts[intPortID].lngHandle, PURGE_TXABORT or PURGE_RXABORT or PURGE_TXCLEAR or PURGE_RXCLEAR)
  if lngStatus == 0: 
    SetCommError("CommOpen (PurgeComm)")
    echo udtCommError
    return udtCommError.lngErrorCode

  # Set serial port timeouts.
  udtCommTimeOuts.ReadIntervalTimeout = -1
  udtCommTimeOuts.ReadTotalTimeoutMultiplier = 0
  udtCommTimeOuts.ReadTotalTimeoutConstant = 1000
  udtCommTimeOuts.WriteTotalTimeoutMultiplier = 0
  udtCommTimeOuts.WriteTotalTimeoutMultiplier = 1000

  lngStatus = SetCommTimeouts(udtPorts[intPortID].lngHandle, &udtCommTimeOuts)
  if lngStatus == 0: 
    SetCommError("CommOpen (SetCommTimeouts)")
    echo udtCommError
    return udtCommError.lngErrorCode

  # Get the current state (DCB).
  lngStatus = GetCommState(udtPorts[intPortID].lngHandle, &udtPorts[intPortID].udtDCB)
  if lngStatus == 0: 
    SetCommError("CommOpen (GetCommState)")
    echo udtCommError
    return udtCommError.lngErrorCode

  # Modify the DCB to reflect the desired settings.
  lngStatus = BuildCommDCB(strSettings, &udtPorts[intPortID].udtDCB)
  if lngStatus == 0: 
    SetCommError("CommOpen (BuildCommDCB)")
    echo udtCommError
    return udtCommError.lngErrorCode

  # Set the new state.
  lngStatus = SetCommState(udtPorts[intPortID].lngHandle, &udtPorts[intPortID].udtDCB)
  if lngStatus == 0: 
    SetCommError("CommOpen (SetCommState)")
    echo udtCommError
    return udtCommError.lngErrorCode

  # no errors, return 0
  return 0

proc CommSet*(intPortID:WORD, strSettings:string):DWORD {.discardable.} =
  ## Modifies the serial port settings.
  ##
  ## Parameters:
  ##   - intPortID   - Port ID used when port was opened.
  ##   - strSettings - Communication settings.
  ##                 Example: "baud=9600 parity=N data=8 stop=1"
  ##
  ## Returns:
  ##   - Error Code  - 0 = No Error.
  var lngStatus: DWORD

  lngStatus = GetCommState(udtPorts[intPortID].lngHandle, &udtPorts[intPortID].udtDCB)
  if lngStatus == 0:
    SetCommError("CommSet (GetCommState)")
    echo udtCommError
    return udtCommError.lngErrorCode

  lngStatus = BuildCommDCB(strSettings, &udtPorts[intPortID].udtDCB)
  if lngStatus == 0:
    SetCommError("CommSet (BuildCommDCB)")
    echo udtCommError
    return udtCommError.lngErrorCode

  lngStatus = SetCommState(udtPorts[intPortID].lngHandle, &udtPorts[intPortID].udtDCB)
  if lngStatus == 0:
    SetCommError("CommSet (SetCommState)")
    echo udtCommError
    return udtCommError.lngErrorCode

  return 0

proc CommClose*(intPortID: WORD):DWORD {.discardable.} =
  ## Close the serial port.
  ##
  ## Parameters:
  ##   - intPortID   - Port ID used when port was opened.
  ##
  ## Returns:
  ##   - Error Code  - 0 = No Error.
  var lngStatus: DWORD

  if udtPorts[intPortID].blnPortOpen:
    lngStatus = CloseHandle(udtPorts[intPortID].lngHandle)
    if lngStatus == 0:
      SetCommError("CommClose (CloseHandle)")
      echo udtCommError
      return udtCommError.lngErrorCode
    
    udtPorts[intPortID].blnPortOpen = false

  return 0

proc CommFlush*(intPortID: WORD):DWORD {.discardable.} =
  ## Flush the send and receive serial port buffers.
  ##
  ## Parameters:
  ##   - intPortID   - Port ID used when port was opened.
  ##
  ## Returns:
  ##   - Error Code  - 0 = No Error.
  var lngStatus: DWORD

  lngStatus = PurgeComm(udtPorts[intPortID].lngHandle, PURGE_TXABORT or PURGE_RXABORT or PURGE_TXCLEAR or PURGE_RXCLEAR)
  if lngStatus == 0:
    SetCommError("CommFlush (PurgeComm)")
    echo udtCommError
    return udtCommError.lngErrorCode
  
  return 0

proc CommRead*(intPortID:WORD, strData:var string, lngSize: DWORD):DWORD {.discardable.} =
  ## Read serial port input buffer.
  ##
  ## Parameters:
  ##   - intPortID   - Port ID used when port was opened.
  ##   - strData     - Data buffer.
  ##   - lngSize     - Maximum number of bytes to be read.
  ##
  ## Returns:
  ##   - Error Code  - 0 = No Error.
  var lngStatus: DWORD
  var lngRdSize: DWORD
  var lngBytesRead: DWORD
  var lngRdStatus: DWORD
  var strRdBuffer = newString(int 1024)
  var lngErrorFlags: DWORD
  var udtCommStat: COMSTAT

   # Clear any previous errors and get current status.
  lngStatus = ClearCommError(udtPorts[intPortID].lngHandle, &lngErrorFlags, &udtCommStat)

  if lngStatus == 0:
    SetCommError("CommRead (ClearCommError)")
    echo udtCommError
    return udtCommError.lngErrorCode
  
  if udtCommStat.cbInQue > 0:
    if udtCommStat.cbInQue > lngSize:
        lngRdSize = udtCommStat.cbInQue
    else:
        lngRdSize = lngSize
  else:
    lngRdSize = 0

  if lngRdSize > 0:

    lngRdStatus = ReadFile(udtPorts[intPortID].lngHandle, &strRdBuffer, lngRdSize, &lngBytesRead, &udtCommOverlap)

    if lngRdStatus == 0:
      lngStatus = GetLastError()
      if lngStatus == ERROR_IO_PENDING:
        # Wait for read to complete.
        # This function will timeout according to the
        # COMMTIMEOUTS.ReadTotalTimeoutConstant variable.
        # Every time it times out, check for port errors.
        
        # Loop until operation is complete.
        while GetOverlappedResult(udtPorts[intPortID].lngHandle, &udtCommOverlap, &lngBytesRead, true) == 0:

          lngStatus = GetLastError()
          if lngStatus != ERROR_IO_INCOMPLETE:
            lngBytesRead = -1
            lngStatus = SetCommErrorEx("CommRead (GetOverlappedResult)", udtPorts[intPortID].lngHandle)
            echo udtCommError
            return lngStatus
      else:
        # Some other error occurred.
        lngBytesRead = -1
        lngStatus = SetCommErrorEx("CommRead (ReadFile)", udtPorts[intPortID].lngHandle)
        echo udtCommError
        return lngStatus

    echo substr(strRdBuffer, 0, lngRdSize)

proc CommWrite*(intPortID: WORD, strData:string):DWORD {.discardable.} =
  ## Output data to the serial port.
  ##
  ## Parameters:
  ##   - intPortID   - Port ID used when port was opened.
  ##   - strData     - Data to be transmitted.
  ##
  ## Returns:
  ##   - Error Code  - 0 = No Error.
  var lngStatus: DWORD
  var lngSize: DWORD
  var lngWrSize: DWORD
  var lngWrStatus: DWORD
  
  # Get the length of the data.
  lngSize = DWORD len(strData)

  # Output the data.
  #proc WriteFile*(hFile: HANDLE, lpBuffer: LPCVOID, nNumberOfBytesToWrite: DWORD, lpNumberOfBytesWritten: LPDWORD, lpOverlapped: LPOVERLAPPED): WINBOOL {.stdcall, dynlib: "kernel32", importc.}
  lngWrStatus = WriteFile(udtPorts[intPortID].lngHandle,  &strData, lngSize, &lngWrSize, &udtCommOverlap)

  # Note that normally the following code will not execute because the driver
  # caches write operations. Small I/O requests (up to several thousand bytes)
  # will normally be accepted immediately and WriteFile will return true even
  # though an overlapped operation was specified.

  if lngWrStatus == 0:
    lngStatus = GetLastError()
  
    if lngStatus == 0:
        return lngWrSize
    elif lngStatus == ERROR_IO_PENDING:
        # We should wait for the completion of the write operation so we know
        # if it worked or not.
        #
        # This is only one way to do this. It might be beneficial to place the
        # writing operation in a separate thread so that blocking on completion
        # will not negatively affect the responsiveness of the UI.
        #
        # If the write takes long enough to complete, this function will timeout
        # according to the CommTimeOuts.WriteTotalTimeoutConstant variable.
        # At that time we can check for errors and then wait some more.

        # Loop until operation is complete.
        while GetOverlappedResult(udtPorts[intPortID].lngHandle, &udtCommOverlap, &lngWrSize, true) == 0:

            lngStatus = GetLastError()

            if lngStatus != ERROR_IO_INCOMPLETE:
                lngStatus = SetCommErrorEx("CommWrite (GetOverlappedResult)", udtPorts[intPortID].lngHandle)
                echo udtCommError
                return lngWrSize
    else:
      # Some other error occurred.
      lngWrSize = -1

      lngStatus = SetCommErrorEx("CommWrite (WriteFile)", udtPorts[intPortID].lngHandle)
      echo udtCommError
      return lngWrSize
  
  
  sleep(5)
  
  return lngWrSize

proc CommGetLine*(intPortID: WORD, intLine: WORD, blnState:var bool): DWORD {.discardable.} =
  ## Get the state of selected serial port control lines.
  ##
  ## Parameters:
  ##   - intPortID   - Port ID used when port was opened.
  ##   - intLine     - Serial port line. CTS, DSR, RING, RLSD (CD)
  ##   - blnState    - Returns state of line (Cleared or Set).
  ##
  ## Returns:
  ##   - Error Code  - 0 = No Error.
  var lngStatus: DWORD
  var lngModemStatus: DWORD

  lngStatus = GetCommModemStatus(udtPorts[intPortID].lngHandle, &lngModemStatus)
  if lngStatus == 0:
    SetCommError("CommGetLine (GetCommModemStatus)")
    echo udtCommError
    return udtCommError.lngErrorCode
  
  if lngModemStatus != 0 and intLine != 0:
    blnState = true
  else:
    blnState = false

  return 0

proc CommSetLine*(intPortID: WORD, intLine: WORD, blnState: bool): DWORD {.discardable.} =
  ## Set the state of selected serial port control lines.
  ##
  ## Parameters:
  ##   - intPortID   - Port ID used when port was opened.
  ##   - intLine     - Serial port line. BREAK, DTR, RTS
  ##                 Note: BREAK actually sets or clears a "break" condition on
  ##                 the transmit data line.
  ##   - blnState    - Sets the state of line (Cleared or Set).
  ##
  ## Returns:
  ##   - Error Code  - 0 = No Error.
  var lngStatus: DWORD
  var lngNewState: DWORD

  if intLine == LINE_BREAK:
  
    if blnState:
        lngNewState = SETBREAK
    else:
        lngNewState = CLRBREAK
  
  elif intLine == LINE_DTR:
  
    if blnState:
      lngNewState = SETDTR
    else:
      lngNewState = CLRDTR
  
  elif intLine == LINE_RTS:

    if blnState:
      lngNewState = SETRTS
    else:
      lngNewState = CLRRTS

  lngStatus = EscapeCommFunction(udtPorts[intPortID].lngHandle, lngNewState)
  if lngStatus == 0:
    SetCommError("CommSetLine (EscapeCommFunction)")
    echo udtCommError
    return udtCommError.lngErrorCode

  return 0

proc CommGetError(strMessage:var string): DWORD {.discardable.} =
  ## Get the last serial port error message.
  ##
  ## Parameters:
  ##   - strMessage  - Error message from last serial port error.
  ##
  ## Returns:
  ##   - Error Code  - Last serial port error code.
  strMessage = "Error (" & $(udtCommError.lngErrorCode) & "): " & udtCommError.strFunction & " - " & udtCommError.strErrorMessage
  return udtCommError.lngErrorCode




proc SendComm*(intPortID: WORD, strData:string): DWORD {.discardable.} =
  ## convinience proc to send data to specified COM port
  var lngStatus: DWORD
  var strError: string
  var lngSize: DWORD
  var lngWrSize: DWORD

  # Initialize Communications
  lngStatus = CommOpen(intPortID, "COM" & $(intPortID), "baud=9600 parity=N data=8 stop=1")
  
  # Handle Error.
  if lngStatus != 0:
    CommGetError(strError)
    SetCommError("SendComm (CommOpen) - " & strError)
    echo udtCommError
    #return udtCommError.lngErrorCode


  # Set modem control lines.
  lngStatus = CommSetLine(intPortID, LINE_RTS, true)
  lngStatus = CommSetLine(intPortID, LINE_DTR, true)

  # Write data to serial port.
  lngSize = DWORD len(strData)
  lngWrSize = CommWrite(intPortID, strData)
  echo "lngWrSize: ", lngWrSize
  echo "lngSize: ", lngSize
  if lngWrSize != lngSize:
      # Handle error.
      SetCommError("SendComm (CommWrite)")
      echo udtCommError
      return udtCommError.lngErrorCode
  
  # Reset modem control lines.
  lngStatus = CommSetLine(intPortID, LINE_RTS, false)
  lngStatus = CommSetLine(intPortID, LINE_DTR, false)
  
  # Close communications.
  CommClose(intPortID)