Attribute VB_Name = "Common"
'*********************************************************************'
'* Copyright = 5763-XA1 (C) Copyright IBM Corp 1994, 1995.           *'
'*   All rights reserved.                                            *'
'*   Licensed Material - Program property of IBM                     *'
'*   Refer to copyright instructions form number G120-2083.          *'
'*********************************************************************'
'*********************************************************************'
'*                                                                   *'
'* Module:                                                           *'
'*   CWB.BAS                                                         *'
'*                                                                   *'
'* Purpose:                                                          *'
'*   Common declarations for Client Access/400 APIs                  *'
'*                                                                   *'
'* Usage Notes:                                                      *'
'*                                                                   *'
'*********************************************************************'

'*********************************************************************'
'* Common definitions                                                *'
'*********************************************************************'


Public Const CWB_FALSE = 0
Public Const CWB_TRUE = 1



'*********************************************************************'
'* Client Access/400 return codes fall into the following            *'
'* categories:                                                       *'
'* - Global return codes corresponding to operating system errors    *'
'* - Global return codes unique to Client Access/400                 *'
'* - Specific return codes for each Client Access/400 function       *'
'*                                                                   *'
'* The global return codes are defined in this file.  The function   *'
'* specific return codes are defined in the function specific        *'
'* header files.                                                     *'
'*********************************************************************'

'*********************************************************************'
'* Global return codes corresponding to operating system errors      *'
'*********************************************************************'

  Public Const CWB_OK = 0
  Public Const CWB_INVALID_FUNCTION = 1
  Public Const CWB_FILE_NOT_FOUND = 2
  Public Const CWB_PATH_NOT_FOUND = 3
  Public Const CWB_TOO_MANY_OPEN_FILES = 4
  Public Const CWB_ACCESS_DENIED = 5
  Public Const CWB_INVALID_HANDLE = 6
  Public Const CWB_NOT_ENOUGH_MEMORY = 8
  Public Const CWB_INVALID_DRIVE = 15
  Public Const CWB_NO_MORE_FILES = 18
  Public Const CWB_DRIVE_NOT_READY = 21
  Public Const CWB_GENERAL_FAILURE = 31
  Public Const CWB_SHARING_VIOLATION = 32
  Public Const CWB_LOCK_VIOLATION = 33
  Public Const CWB_END_OF_FILE = 38
  Public Const CWB_NOT_SUPPORTED = 50
  Public Const CWB_BAD_NETWORK_PATH = 53
  Public Const CWB_NETWORK_BUSY = 54
  Public Const CWB_DEVICE_NOT_EXIST = 55
  Public Const CWB_UNEXPECTED_NETWORK_ERROR = 59
  Public Const CWB_NETWORK_ACCESS_DENIED = 65
  Public Const CWB_FILE_EXISTS = 80
  Public Const CWB_ALREADY_ASSIGNED = 85
  Public Const CWB_INVALID_PARAMETER = 87
  Public Const CWB_NETWORK_WRITE_FAULT = 88
  Public Const CWB_OPEN_FAILED = 110
  Public Const CWB_BUFFER_OVERFLOW = 111
  Public Const CWB_DISK_FULL = 112
  Public Const CWB_PROTECTION_VIOLATION = 115
  Public Const CWB_INVALID_LEVEL = 124
  Public Const CWB_BUSY_DRIVE = 142
  Public Const CWB_INVALID_FSD_NAME = 252
  Public Const CWB_INVALID_PATH = 253

'*********************************************************************'
'* Global return codes unique to Client Access/400                   *'
'*********************************************************************'
  Public Const CWB_START = 4000
  Public Const CWB_LAST = 5999

  Public Const CWB_USER_CANCELLED_COMMAND = CWB_START
  Public Const CWB_CONFIG_ERROR = CWB_START + 1
  Public Const CWB_LICENSE_ERROR = CWB_START + 2
  Public Const CWB_PROD_OR_COMP_NOT_SET = CWB_START + 3
  Public Const CWB_SECURITY_ERROR = CWB_START + 4
  Public Const CWB_GLOBAL_CFG_FAILED = CWB_START + 5
  Public Const CWB_PROD_RETRIEVE_FAILED = CWB_START + 6
  Public Const CWB_COMP_RETRIEVE_FAILED = CWB_START + 7
  Public Const CWB_COMP_CFG_FAILED = CWB_START + 8
  Public Const CWB_COMP_FIX_LEVEL_UPDATE_FAILED = CWB_START + 9
  Public Const CWB_INVALID_API_HANDLE = CWB_START + 10
  Public Const CWB_INVALID_API_PARAMETER = CWB_START + 11
  Public Const CWB_HOST_NOT_FOUND = CWB_START + 12
  Public Const CWB_NOT_COMPATIBLE = CWB_START + 13
  Public Const CWB_INVALID_POINTER = CWB_START + 14
  Public Const CWB_SERVER_PROGRAM_NOT_FOUND = CWB_START + 15
  Public Const CWB_API_ERROR = CWB_START + 16
  Public Const CWB_CA_NOT_STARTED = CWB_START + 17
  Public Const CWB_FILE_IO_ERROR = CWB_START + 18
  Public Const CWB_COMMUNICATIONS_ERROR = CWB_START + 19
  Public Const CWB_RUNTIME_CONSTRUCTOR_FAILED = CWB_START + 20
  Public Const CWB_DIAGNOSTIC = CWB_START + 21
  Public Const CWB_COMM_VERSION_ERROR = CWB_START + 22
  Public Const CWB_NO_VIEWER = CWB_START + 23
  Public Const CWB_MODULE_NOT_LOADABLE = CWB_START + 24
  Public Const CWB_ALREADY_SETUP = CWB_START + 25
  Public Const CWB_CANNOT_START_PROCESS = CWB_START + 26

