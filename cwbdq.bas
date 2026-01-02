Attribute VB_Name = "DataQueue"
'********************************************************************/
'* Copyright = 5763-XA1 (C) Copyright IBM Corp 1994, 1995.          */
'*  All rights reserved.                                            */
'*  Licensed Material - Program property of IBM                     */
'*  Refer to copyright instructions form number G120-2083.          */
'********************************************************************/
'********************************************************************/
'*                                                                  */
'* Module:                                                          */
'*   CWBDQ.bas                                                        */
'*                                                                  */
'* Purpose:                                                         */
'*   The functions listed in this file can be used to access AS/400 */
'*   data queues.  They are useful for developing client/server     */
'*   applications without requiring the use of communications APIs. */
'*                                                                  */
'*   The APIs can be broken down into three groups:                 */
'*                                                                  */
'*   The following APIs can be used to access the AS/400 data queue.*/
'*   After the cwbDQ_Open() API is used to create a connection to a */
'*   specific data queue the other APIs can be used to utilize it.  */
'*   Use the cwbDQ_Close() API when the connection is no longer     */
'*   needed.                                                        */
'*      cwbDQ_Create()                                              */
'*      cwbDQ_Delete()                                              */
'*                                                                  */
'*      cwbDQ_AsyncRead()                                           */
'*      cwbDQ_Cancel()                                              */
'*      cwbDQ_CheckData()                                           */
'*      cwbDQ_Clear()                                               */
'*      cwbDQ_Close()                                               */
'*      cwbDQ_GetLibName()                                          */
'*      cwbDQ_GetQueueAttr()                                        */
'*      cwbDQ_GetQueueName()                                        */
'*      cwbDQ_GetSysName()                                          */
'*      cwbDQ_Open()                                                */
'*      cwbDQ_Peek()                                                */
'*      cwbDQ_Read()                                                */
'*      cwbDQ_Write()                                               */
'*                                                                  */
'*   Following are the declarations for the attributes of a data    */
'*   queue.  The attribute object is used when creating a data      */
'*   queue or when get the data queue attributes.                   */
'*      cwbDQ_CreateAttr()                                          */
'*      cwbDQ_SetMaxRecLen()                                        */
'*      cwbDQ_SetOrder()                                            */
'*      cwbDQ_SetAuthority()                                        */
'*      cwbDQ_SetForceToStorage()                                   */
'*      cwbDQ_SetSenderID()                                         */
'*      cwbDQ_SetKeySize()                                          */
'*      cwbDQ_SetDesc()                                             */
'*      cwbDQ_GetMaxRecLen()                                        */
'*      cwbDQ_GetOrder()                                            */
'*      cwbDQ_GetAuthority()                                        */
'*      cwbDQ_GetForceToStorage()                                   */
'*      cwbDQ_GetSenderID()                                         */
'*      cwbDQ_GetKeySize()                                          */
'*      cwbDQ_GetDesc()                                             */
'*      cwbDQ_DeleteAttr()                                          */
'*                                                                  */
'*   After creation an attribute object will have the default       */
'*   values of:                                                     */
'*      - maximum record length - 0                                 */
'*      - order - FIFO                                              */
'*      - public authority - *LIBCRTAUT                             */
'*      - force to auxiliary storage - FALSE                        */
'*      - sender ID - FALSE                                         */
'*      - key size - 0                                              */
'*      - description - NONE                                        */
'*                                                                  */
'*   Following are the declarations for the functions that use the  */
'*   data object for writing to and reading from a data queue:      */
'*      cwbDQ_CreateData()                                          */
'*      cwbDQ_DeleteData()                                          */
'*      cwbDQ_GetConvert()                                          */
'*      cwbDQ_GetData()                                             */
'*      cwbDQ_GetDataAddr()                                         */
'*      cwbDQ_GetDataLen()                                          */
'*      cwbDQ_GetKey()                                              */
'*      cwbDQ_GetKeyLen()                                           */
'*      cwbDQ_GetRetDataLen()                                       */
'*      cwbDQ_GetRetKey()                                           */
'*      cwbDQ_GetRetKeyLen()                                        */
'*      cwbDQ_GetSearchOrder()                                      */
'*      cwbDQ_GetSenderInfo()                                       */
'*      cwbDQ_SetConvert()                                          */
'*      cwbDQ_SetData()                                             */
'*      cwbDQ_SetDataAddr()                                         */
'*      cwbDQ_SetKey()                                              */
'*      cwbDQ_SetSearchOrder()                                      */
'*                                                                  */
'*   After creation a data object will have the default values of:  */
'*      - data - NULL and length 0                                  */
'*      - key - NULL and length 0                                   */
'*      - sender ID info - NULL                                     */
'*      - search order - NONE                                       */
'*      - convert - NO                                              */
'*                                                                  */
'* Usage Notes:                                                     */
'*   Link with CWB.LIB import library.                              */
'*   This module is to be used in conjunction with CWBDQ.DLL.       */
'********************************************************************/
 
'*------------------------------------------------------------------*/
'* TYPEDEFS:                                                        */
'*------------------------------------------------------------------*/
'*------------------------------------------------------------------*/
'*                                                                  */
'* Definitions for data queue constants for authority               */
'*                                                                  */
'*------------------------------------------------------------------*/
Public Const CWBDQ_ALL = 0
Public Const CWBDQ_EXCLUDE = 1
Public Const CWBDQ_CHANGE = 2
Public Const CWBDQ_USE = 3
Public Const CWBDQ_LIBCRTAUT = 4
 
'*------------------------------------------------------------------*/
'*                                                                  */
'* Definitions for data queue constants for order                   */
'*                                                                  */
'*------------------------------------------------------------------*/
Public Const CWBDQ_SEQ_LIFO = 0
Public Const CWBDQ_SEQ_FIFO = 1
Public Const CWBDQ_SEQ_KEYED = 2
 
'*------------------------------------------------------------------*/
'*                                                                  */
'* Definitions for data queue constants for search order            */
'*                                                                  */
'*------------------------------------------------------------------*/
Public Const CWBDQ_NONE = 0
Public Const CWBDQ_EQUAL = 1
Public Const CWBDQ_NOT_EQUAL = 2
Public Const CWBDQ_GT_OR_EQUAL = 3
Public Const CWBDQ_GREATER = 4
Public Const CWBDQ_LT_OR_EQUAL = 5
Public Const CWBDQ_LESS = 6
 
'*------------------------------------------------------------------*/
'* Component DQ errors based on CWB_LAST (in CWB.H)                 */
'*------------------------------------------------------------------*/
 
 Public Const CWBDQ_START = CWB_LAST + 1
 
'*------------------------------------------------------------------*/
'* Invalid attributes handle                                        */
'*------------------------------------------------------------------*/
Public Const CWBDQ_INVALID_ATTRIBUTE_HANDLE = CWBDQ_START
 
'*------------------------------------------------------------------*/
'* Invalid data handle                                              */
'*------------------------------------------------------------------*/
Public Const CWBDQ_INVALID_DATA_HANDLE = CWBDQ_START + 1
 
'*------------------------------------------------------------------*/
'* Invalid queue handle                                             */
'*------------------------------------------------------------------*/
Public Const CWBDQ_INVALID_QUEUE_HANDLE = CWBDQ_START + 2
 
'*------------------------------------------------------------------*/
'* Invalid data queue read handle                                   */
'*------------------------------------------------------------------*/
Public Const CWBDQ_INVALID_READ_HANDLE = CWBDQ_START + 3
 
'*------------------------------------------------------------------*/
'* Invalid maximum record length for a data queue                   */
'*------------------------------------------------------------------*/
Public Const CWBDQ_INVALID_QUEUE_LENGTH = CWBDQ_START + 4
 
'*------------------------------------------------------------------*/
'* Invalid key length                                               */
'*------------------------------------------------------------------*/
Public Const CWBDQ_INVALID_KEY_LENGTH = CWBDQ_START + 5
 
'*------------------------------------------------------------------*/
'* Invalid queue order                                              */
'*------------------------------------------------------------------*/
Public Const CWBDQ_INVALID_ORDER = CWBDQ_START + 6
 
'*------------------------------------------------------------------*/
'* Invalid queue authority                                          */
'*------------------------------------------------------------------*/
Public Const CWBDQ_INVALID_AUTHORITY = CWBDQ_START + 7
 
'*------------------------------------------------------------------*/
'* Queue title (description) is too long or cannot be converted     */
'*------------------------------------------------------------------*/
Public Const CWBDQ_INVALID_QUEUE_TITLE = CWBDQ_START + 8
 
'*------------------------------------------------------------------*/
'* Queue name is too long or cannot be converted                    */
'*------------------------------------------------------------------*/
Public Const CWBDQ_BAD_QUEUE_NAME = CWBDQ_START + 9
 
'*------------------------------------------------------------------*/
'* Library name is too long or cannot be converted                  */
'*------------------------------------------------------------------*/
Public Const CWBDQ_BAD_LIBRARY_NAME = CWBDQ_START + 10
 
'*------------------------------------------------------------------*/
'* System name is too long or cannot be converted                   */
'*------------------------------------------------------------------*/
Public Const CWBDQ_BAD_SYSTEM_NAME = CWBDQ_START + 11
 
'*------------------------------------------------------------------*/
'* Length of key is not correct for this data queue                 */
'* or key length is greater than 0 for a LIFO or FIFO data queue    */
'*------------------------------------------------------------------*/
Public Const CWBDQ_BAD_KEY_LENGTH = CWBDQ_START + 12
 
'*------------------------------------------------------------------*/
'* Length of data is not correct for this data queue.  Either the   */
'* data length is zero or it is greater than the maximum allowed    */
'* of 31744 bytes.                                                  */
'*------------------------------------------------------------------*/
Public Const CWBDQ_BAD_DATA_LENGTH = CWBDQ_START + 13
 
'*------------------------------------------------------------------*/
'* Wait time is not correct                                         */
'*------------------------------------------------------------------*/
Public Const CWBDQ_INVALID_TIME = CWBDQ_START + 14
 
'*------------------------------------------------------------------*/
'* Search order is not correct                                      */
'*------------------------------------------------------------------*/
Public Const CWBDQ_INVALID_SEARCH = CWBDQ_START + 15
 
'*------------------------------------------------------------------*/
'* Returned data was truncated                                      */
'*------------------------------------------------------------------*/
Public Const CWBDQ_DATA_TRUNCATED = CWBDQ_START + 16
 
'*------------------------------------------------------------------*/
'* Wait time has expired and no data has been returned              */
'*------------------------------------------------------------------*/
Public Const CWBDQ_TIMED_OUT = CWBDQ_START + 17
 
'*------------------------------------------------------------------*/
'* Command rejected by user exit program                            */
'*------------------------------------------------------------------*/
Public Const CWBDQ_REJECTED_USER_EXIT = CWBDQ_START + 18
 
'*------------------------------------------------------------------*/
'* Error in user exit program or invalid number of exit programs    */
'*------------------------------------------------------------------*/
Public Const CWBDQ_USER_EXIT_ERROR = CWBDQ_START + 19
 
'*------------------------------------------------------------------*/
'* Library not found on system                                      */
'*------------------------------------------------------------------*/
Public Const CWBDQ_LIBRARY_NOT_FOUND = CWBDQ_START + 20
 
'*------------------------------------------------------------------*/
'* Queue not found on system                                        */
'*------------------------------------------------------------------*/
Public Const CWBDQ_QUEUE_NOT_FOUND = CWBDQ_START + 21
 
'*------------------------------------------------------------------*/
'* No authority to library or data queue                            */
'*------------------------------------------------------------------*/
Public Const CWBDQ_NO_AUTHORITY = CWBDQ_START + 22
 
'*------------------------------------------------------------------*/
'* Data queue is in an unusable state                               */
'*------------------------------------------------------------------*/
Public Const CWBDQ_DAMAGED_QUEUE = CWBDQ_START + 23
 
'*------------------------------------------------------------------*/
'* Data queue already exists                                        */
'*------------------------------------------------------------------*/
Public Const CWBDQ_QUEUE_EXISTS = CWBDQ_START + 24
 
'*------------------------------------------------------------------*/
'* Invalid message length - exceeds queue maximum record length     */
'*------------------------------------------------------------------*/
Public Const CWBDQ_INVALID_MESSAGE_LENGTH = CWBDQ_START + 25
 
'*------------------------------------------------------------------*/
'* Queue destroyed while waiting to read or peek a record           */
'*------------------------------------------------------------------*/
Public Const CWBDQ_QUEUE_DESTROYED = CWBDQ_START + 26
 
'*------------------------------------------------------------------*/
'* No data was received                                             */
'*------------------------------------------------------------------*/
Public Const CWBDQ_NO_DATA = CWBDQ_START + 27
 
'*------------------------------------------------------------------*/
'* Data cannot be converted for this data queue.  The data queue    */
'* can be used but data cannot be converted between ASCII and       */
'* EBCDIC.  The convert flag on the data object will be ignored.    */
'*------------------------------------------------------------------*/
Public Const CWBDQ_CANNOT_CONVERT = CWBDQ_START + 28
 
'*------------------------------------------------------------------*/
'* Syntax of the data queue name is incorrect.  Queue name must     */
'* follow AS/400 object syntax.   First character must be           */
'* alphabetic and all following characters alphanumeric             */
'*------------------------------------------------------------------*/
Public Const CWBDQ_QUEUE_SYNTAX = CWBDQ_START + 29
 
'*------------------------------------------------------------------*/
'* Syntax of the library name is incorrect.  Library name must      */
'* follow AS/400 object syntax.   First character must be           */
'* alphabetic and all following characters alphanumeric             */
'*------------------------------------------------------------------*/
Public Const CWBDQ_LIBRARY_SYNTAX = CWBDQ_START + 30
 
'*------------------------------------------------------------------*/
'* Address not set.  The data object was not set with with          */
'* cwbDQ_SetDataAddr(), so the address cannot be retrieved.         */
'* Use cwbDQ_GetData() instead of cwbDQ_GetDataAddr().              */
'*------------------------------------------------------------------*/
Public Const CWBDQ_ADDRESS_NOT_SET = CWBDQ_START + 31
 
'*------------------------------------------------------------------*/
'* Host error occurred for which no return is defined.              */
'* See the error handle for the message text.                       */
'*------------------------------------------------------------------*/
Public Const CWBDQ_HOST_ERROR = CWBDQ_START + 32
 
'*------------------------------------------------------------------*/
'* Unexpected error                                                 */
'*------------------------------------------------------------------*/
Public Const CWBDQ_UNEXPECTED_ERROR = CWBDQ_START + 99
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_AsyncRead                                                */
'*                                                                  */
'* Purpose:                                                         */
'*   Read a record from the AS/400 data queue object that is        */
'*   identified by the specified handle.  The AsyncRead will return */
'*   control to the caller immediately.  This call is used in       */
'*   conjunction with the CheckData API.  When a record is read     */
'*   from a data queue it is removed from the data queue.  You may  */
'*   wait for a record if the data queue is empty by specifying a   */
'*   wait time from 0 to 99,999 or forever (-1).  A wait time of    */
'*   zero will return immediately if there is no data in the data   */
'*   queue.                                                         */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_QueueHandle queueHandle - input                          */
'*     Handle that was returned by a previous call to the           */
'*     cwbDQ_Open function.  This identifies the AS/400 data        */
'*     queue object.                                                */
'*                                                                  */
'*   cwbDQ_Data data - input                                        */
'*     The data object to be read from the AS/400 data queue.       */
'*                                                                  */
'*   signed long waitTime - input                                   */
'*     Length of time in seconds to wait for data, if the data      */
'*     queue is empty.  A wait time of -1 indicates to wait forever.*/
'*                                                                  */
'*   cwbDQ_ReadHandle * readHandle - output                         */
'*     Pointer to where the cwbDQ_ReadHandle will be written.       */
'*     This handle will be used in subsequent calls to the          */
'*     cwbDQ_CheckData API.                                         */
'*                                                                  */
'*   cwbSV_ErrHandle errorHandle - output                           */
'*     Any returned messages will be written to this object.  It    */
'*     is created with the cwbSV_CreateErrHandle API.  The          */
'*     messages may be retrieved through the cwbSV_GetErrText API.  */
'*     If the parameter is set to zero, no messages will be         */
'*     retrieved.                                                   */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_TIME - Invalid wait time.                        */
'*   CWBDQ_INVALID_QUEUE_HANDLE - Invalid queue handle.             */
'*   CWBDQ_INVALID_SEARCH - Invalid search order.                   */
'*                                                                  */
'* Usage Notes:                                                     */
'*   This function requires that you have previously issued:        */
'*     cwbDQ_Open()                                                 */
'*     cwbDQ_CreateData()                                           */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_AsyncRead Lib "CWBDQ" (queueHandle As Long, data As Long, ByVal WaitTime As Long, readHandle As Long, errorHandle As Long) As Integer

'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_Cancel                                                   */
'*                                                                  */
'* Purpose:                                                         */
'*   Cancel a previously issued AsyncRead.  This will end the read  */
'*   on the AS/400 data queue.                                      */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_ReadHandle readHandle - input                            */
'*     The handle that was returned by the AsyncRead API.           */
'*                                                                  */
'*   cwbSV_ErrHandle errorHandle - output                           */
'*     Any returned messages will be written to this object.  It    */
'*     is created with the cwbSV_CreateErrHandle API.  The          */
'*     messages may be retrieved through the cwbSV_GetErrText API.  */
'*     If the parameter is set to zero, no messages will be         */
'*     retrieved.                                                   */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_READ_HANDLE - Invalid read handle.               */
'*                                                                  */
'* Usage Notes:                                                     */
'*   This function requires that you have previously issued:        */
'*     cwbDQ_Open()                                                 */
'*     cwbDQ_CreateData()                                           */
'*     cwbDQ_AsyncRead()                                            */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_Cancel Lib "CWBDQ" (readHandle As Long, errorHandle As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_CheckData                                                */
'*                                                                  */
'* Purpose:                                                         */
'*   Check if data has returned from a previously issued AsyncRead  */
'*   API.  This API can be issued multiple times for a single       */
'*   AsyncRead call.  It will return 0 when the data has actually   */
'*   been returned.                                                 */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_ReadHandle readHandle - input                            */
'*     The handle that was returned by the AsyncRead API.           */
'*                                                                  */
'*   cwbSV_ErrHandle errorHandle - output                           */
'*     Any returned messages will be written to this object.  It    */
'*     is created with the cwbSV_CreateErrHandle API.  The          */
'*     messages may be retrieved through the cwbSV_GetErrText API.  */
'*     If the parameter is set to zero, no messages will be         */
'*     retrieved.                                                   */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_READ_HANDLE - Invalid read handle.               */
'*   CWBDQ_DATA_TRUNCATED - Data truncated.                         */
'*   CWBDQ_TIMED_OUT - Wait time expired and no data returned.      */
'*   CWBDQ_REJECTED_USER_EXIT - Command rejected by user exit       */
'*                              program.                            */
'*   CWBDQ_QUEUE_DESTROYED - Queue was destroyed.                   */
'*   CWBDQ_NO_DATA - No data.                                       */
'*   CWBDQ_CANNOT_CONVERT - Unable to convert data.                 */
'*                                                                  */
'* Usage Notes:                                                     */
'*   This function requires that you have previously issued:        */
'*     cwbDQ_Open()                                                 */
'*     cwbDQ_CreateData()                                           */
'*     cwbDQ_AsyncRead()                                            */
'*   If a time limit was specified on the AsyncRead, this API will  */
'*   return CWBDQ_NO_DATA until data is returned (return code will  */
'*   be CWB_OK) or the time limit expires (return code will be      */
'*   CWBDQ_TIMED_OUT).                                              */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_CheckData Lib "CWBDQ" (readHandle As Long, errorHandle As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_Clear                                                    */
'*                                                                  */
'* Purpose:                                                         */
'*   Remove all messages from the AS/400 data queue object that is  */
'*   identified by the specified handle.  If the queue is keyed,    */
'*   messages for a particular key may be removed by specifying     */
'*   the key and and key length.  These values should be set to     */
'*   NULL and zero, respectively, if you want to clear all messages */
'*   from the queue.                                                */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_QueueHandle queueHandle - input                          */
'*     Handle that was returned by a previous call to the           */
'*     cwbDQ_Open function.  This identifies the AS/400 data        */
'*     queue object.                                                */
'*                                                                  */
'*   unsigned char * key - input                                    */
'*     Pointer to the key.  The key may contain embedded NULLs,     */
'*     so it is not an ASCIIZ string.                               */
'*                                                                  */
'*   unsigned short keyLength - input                               */
'*     Length of the key in bytes.                                  */
'*                                                                  */
'*   cwbSV_ErrHandle errorHandle - output                           */
'*     Any returned messages will be written to this object.  It    */
'*     is created with the cwbSV_CreateErrHandle API.  The          */
'*     messages may be retrieved through the cwbSV_GetErrText API.  */
'*     If the parameter is set to zero, no messages will be         */
'*     retrieved.                                                   */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_QUEUE_HANDLE - Invalid queue handle.             */
'*   CWBDQ_BAD_KEY_LENGTH - Length of key is not correct.           */
'*   CWBDQ_REJECTED_USER_EXIT - Command rejected by user exit       */
'*                              program.                            */
'*                                                                  */
'* Usage Notes:                                                     */
'*   This function requires that you have previously issued:        */
'*     cwbDQ_Open()                                                 */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_Clear Lib "CWBDQ" (ByVal queueHandle As Long, key As Byte, ByVal keyLength As Integer, ByVal errorHandle As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_Close                                                    */
'*                                                                  */
'* Purpose:                                                         */
'*   End the connection with the AS/400 data queue object that is   */
'*   identified by the specified handle.  This will end the         */
'*   conversation with the AS/400 system.                           */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_QueueHandle queueHandle - input                          */
'*     Handle that was returned by a previous call to the           */
'*     cwbDQ_Open function.  This identifies the AS/400 data        */
'*     queue object.                                                */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_QUEUE_HANDLE - Invalid queue handle.             */
'*                                                                  */
'* Usage Notes:                                                     */
'*   This function requires that you have previously issued:        */
'*     cwbDQ_Open()                                                 */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_Close Lib "CWBDQ" (ByVal queueHandle As Long) As Integer
 
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_Create                                                   */
'*                                                                  */
'* Purpose:                                                         */
'*   Create an AS/400 data queue object.  After the object is       */
'*   created it can be opened using the cwbDQ_Open API.  It will    */
'*   have the attributes that you specify in the attributes handle. */
'*                                                                  */
'* Parameters:                                                      */
'*   char * queue - input                                           */
'*     Pointer to the data queue name contained in an ASCIIZ string.*/
'*                                                                  */
'*   char * library - input                                         */
'*     Pointer to the library name contained in an ASCIIZ string.   */
'*     If this pointer is NULL then the current library will be     */
'*     used (set library to "*CURLIB").                             */
'*                                                                  */
'*   char * systemName - input                                      */
'*     Pointer to the system name contained in an ASCIIZ string.    */
'*                                                                  */
'*   cwbDQ_Attr queueAttributes - input                             */
'*     Handle to the attributes for the data queue.                 */
'*                                                                  */
'*   cwbSV_ErrHandle errorHandle - output                           */
'*     Any returned messages will be written to this object.  It    */
'*     is created with the cwbSV_CreateErrHandle API.  The          */
'*     messages may be retrieved through the cwbSV_GetErrText API.  */
'*     If the parameter is set to zero, no messages will be         */
'*     retrieved.                                                   */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_COMMUNICATIONS_ERROR - A communications error occurred.    */
'*   CWB_SERVER_PROGRAM_NOT_FOUND - AS/400 application not found.   */
'*   CWB_HOST_NOT_FOUND - AS/400 system inactive or does not exist. */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWB_SECURITY_ERROR - A security error has occurred.            */
'*   CWB_LICENSE_ERROR - A license error has occurred.              */
'*   CWB_CONFIG_ERROR - A configuration error has occurred.         */
'*   CWBDQ_INVALID_ATTRIBUTE_HANDLE - Invalid attributes handle.    */
'*   CWBDQ_BAD_QUEUE_NAME - Queue name is incorrect.                */
'*   CWBDQ_BAD_LIBRARY_NAME - Library name is incorrect.            */
'*   CWBDQ_BAD_SYSTEM_NAME - System name is incorrect.              */
'*   CWBDQ_REJECTED_USER_EXIT - Command rejected by user exit       */
'*                              program.                            */
'*   CWBDQ_USER_EXIT_ERROR - Error in user exit program.            */
'*   CWBDQ_LIBRARY_NOT_FOUND - Library not found on system.         */
'*   CWBDQ_NO_AUTHORITY - No authority to library.                  */
'*   CWBDQ_QUEUE_EXISTS - Queue already exists.                     */
'*   CWBDQ_QUEUE_SYNTAX - Queue syntax is incorrect.                */
'*   CWBDQ_LIBRARY_SYNTAX - Library syntax is incorrect.            */
'*                                                                  */
'* Usage Notes:                                                     */
'*   This function requires that you have previously issued:        */
'*     cwbDQ_CreateAttr()                                           */
'*     cwbDQ_SetMaxRecLen()                                         */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_Create Lib "CWBDQ" (ByVal queue As String, ByVal library As String, ByVal systemName As String, queueAttributes As Long, errorHandle As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_CreateAttr                                               */
'*                                                                  */
'* Purpose:                                                         */
'*   Create a data queue attribute object.  The handle returned     */
'*   by this API can be used to set the specific attributes         */
'*   you want for a data queue.  It is input for the API,           */
'*   cwbDQ_Create.                                                  */
'*                                                                  */
'* Parameters:                                                      */
'*   None.                                                          */
'*                                                                  */
'* Return Value:                                                    */
'*   cwbDQ_Attr - A handle to a cwbDQ_Attr object.                  */
'*                You can use this handle to                        */
'*                get and set attributes and in the  The other      */
'*                attributes will be set to the cwbDQ_Create        */
'*                API. following values:                            */
'*                    - Maximum Record Length - 0                   */
'*                    - Order - FIFO                                */
'*                    - Authority - LIBCRTAUT                       */
'*                    - Force to Storage - NO                       */
'*                    - Sender ID - NO                              */
'*                    - Key Length - 0                              */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_CreateAttr Lib "CWBDQ" () As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_CreateData                                               */
'*                                                                  */
'* Purpose:                                                         */
'*   Create the data object.  This data object can be used for      */
'*   both reading and writing data to a data queue.                 */
'*                                                                  */
'* Parameters:                                                      */
'*   None.                                                          */
'*                                                                  */
'* Return Value:                                                    */
'*   cwbDQ_Data - A handle to the data object.                      */
'*                After being opened the data parameters are set to:*/
'*                    - data - NULL and length 0                    */
'*                    - key - NULL and length 0                     */
'*                    - sender ID info - NULL                       */
'*                    - search order - NONE                         */
'*                    - convert - NO                                */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_CreateData Lib "CWBDQ" () As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_Delete                                                   */
'*                                                                  */
'* Purpose:                                                         */
'*   Remove all messages from an AS/400 data queue and delete       */
'*   the data queue object.                                         */
'*                                                                  */
'* Parameters:                                                      */
'*   char * queue - input                                           */
'*     Pointer to the data queue name contained in an ASCIIZ string.*/
'*                                                                  */
'*   char * library - input                                         */
'*     Pointer to the library name contained in an ASCIIZ string.   */
'*     If this pointer is NULL then the current library will be     */
'*     used (set library to "*CURLIB").                             */
'*                                                                  */
'*   char * systemName - input                                      */
'*     Pointer to the system name contained in an ASCIIZ string.    */
'*                                                                  */
'*   cwbSV_ErrHandle errorHandle - output                           */
'*     Any returned messages will be written to this object.  It    */
'*     is created with the cwbSV_CreateErrHandle API.  The          */
'*     messages may be retrieved through the cwbSV_GetErrText API.  */
'*     If the parameter is set to zero, no messages will be         */
'*     retrieved.                                                   */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_COMMUNICATIONS_ERROR - A communications error occurred.    */
'*   CWB_SERVER_PROGRAM_NOT_FOUND - AS/400 application not found.   */
'*   CWB_HOST_NOT_FOUND - AS/400 system inactive or does not exist. */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWB_SECURITY_ERROR - A security error has occurred.            */
'*   CWB_LICENSE_ERROR - A license error has occurred.              */
'*   CWB_CONFIG_ERROR - A configuration error has occurred.         */
'*   CWBDQ_QUEUE_NAME - Queue name is too long.                     */
'*   CWBDQ_LIBRARY_NAME - Library name is too long.                 */
'*   CWBDQ_SYSTEM_NAME - System name is too long.                   */
'*   CWBDQ_REJECTED_USER_EXIT - Command rejected by user exit       */
'*                              program.                            */
'*   CWBDQ_USER_EXIT_ERROR - Error in user exit program.            */
'*   CWBDQ_LIBRARY_NOT_FOUND - Library not found on system.         */
'*   CWBDQ_QUEUE_NOT_FOUND - Queue not found on system.             */
'*   CWBDQ_NO_AUTHORITY - No authority to queue.                    */
'*   CWBDQ_QUEUE_SYNTAX - Queue syntax is incorrect.                */
'*   CWBDQ_LIBRARY_SYNTAX - Library syntax is incorrect.            */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_Delete Lib "CWBDQ" (ByVal queue As String, ByVal library As String, ByVal systemName As String, errorHandle As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_DeleteAttr                                               */
'*                                                                  */
'* Purpose:                                                         */
'*   Delete the data queue attributes.                              */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Attr queueAttributes - input                             */
'*     Handle of the data queue attributes returned by a previous   */
'*     call to cwbDQ_CreateAttr.                                    */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_ATTRIBUTE_HANDLE - Invalid attributes handle.    */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_DeleteAttr Lib "CWBDQ" (queueAttributes As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_DeleteData                                               */
'*                                                                  */
'* Purpose:                                                         */
'*   Delete the data object.                                        */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Data data - input                                        */
'*     Handle of the data object that was returned by a previous    */
'*     call to cwbDQ_CreateData.                                    */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_DATA_HANDLE - Invalid data handle.               */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_DeleteData Lib "CWBDQ" (ByVal data As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetAuthority                                             */
'*                                                                  */
'* Purpose:                                                         */
'*   Get the attribute for the authority that other users will have */
'*   to the data queue.                                             */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Attr queueAttributes - input                             */
'*     Handle of the data queue attributes returned by a previous   */
'*     call to cwbDQ_CreateAttr.                                    */
'*                                                                  */
'*   unsigned short * authority - output                            */
'*           Pointer to an unsigned short where the authority will  */
'*           be written.                                            */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_ATTRIBUTE_HANDLE - Invalid attributes handle.    */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetAuthority Lib "CWBDQ" (queueAttributes As Long, authority As Long) As Integer

'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetConvert                                               */
'*                                                                  */
'* Purpose:                                                         */
'*   Get the convert flag.                                          */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Data data - input                                        */
'*     Handle of the data object that was returned by a previous    */
'*     call to cwbDQ_CreateData.                                    */
'*                                                                  */
'*   cwb_Boolean * convert - output                                 */
'*     Pointer to a Boolean where the convert flag will be written. */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_DATA_HANDLE - Invalid data handle.               */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetConvert Lib "CWBDQ" (data As Long, convert As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetData                                                  */
'*                                                                  */
'* Purpose:                                                         */
'*   Get the data attribute of the data object.                     */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Data data - input                                        */
'*     Handle of the data object that was returned by a previous    */
'*     call to cwbDQ_CreateData.                                    */
'*                                                                  */
'*   unsigned char * data - output                                  */
'*     Pointer to the data.  The data may contain embedded NULLs,   */
'*     so it is not an ASCIIZ string.                               */
'*                                                                  */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_DATA_HANDLE - Invalid data handle.               */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetData Lib "CWBDQ" (ByVal data As Long, ByVal DataBuffer As String) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetDataAddr                                              */
'*                                                                  */
'* Purpose:                                                         */
'*   Get the address of where the data buffer is.                   */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Data data - input                                        */
'*     Handle of the data object that was returned by a previous    */
'*     call to cwbDQ_CreateData.                                    */
'*                                                                  */
'*   unsigned char * * data - output                                */
'*     Pointer to where the buffer address will be written.         */
'*                                                                  */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_DATA_HANDLE - Invalid data handle.               */
'*   CWBDQ_ADDRESS_NOT_SET - Address not set with cwbDQ_SetDataAddr.*/
'*                                                                  */
'* Usage Notes:                                                     */
'*    Use this function to retrieve the address of where the data   */
'*    is stored.  The data addess must be set with the API          */
'*    cwbDQ_SetDataAddr(), otherwise the return code                */
'*    CWBDQ_ADDRESS_NOT_SET will be returned.                       */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetDataAddr Lib "CWBDQ" (data As Long, DataBuffer As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetDataLen                                               */
'*                                                                  */
'* Purpose:                                                         */
'*   Get the data length attribute of the data object.  This is the */
'*   total length of the data object.  To obtain the length of data */
'*   that was read use the cwbDQ_GetRetDataLen() API.               */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Data data - input                                        */
'*     Handle of the data object that was returned by a previous    */
'*     call to cwbDQ_CreateData.                                    */
'*                                                                  */
'*   unsigned long * dataLength - output                            */
'*     Pointer to an unsigned long where the length of the data     */
'*     will be written.                                             */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_DATA_HANDLE - Invalid data handle.               */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetDataLen Lib "CWBDQ" (data As Long, dataLength As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetDesc                                                  */
'*                                                                  */
'* Purpose:                                                         */
'*   Get the attribute for the description of the data queue.       */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Attr queueAttributes - input                             */
'*     Handle of the data queue attributes returned by a previous   */
'*     call to cwbDQ_CreateAttr.                                    */
'*                                                                  */
'*   char * description - output                                    */
'*     Pointer to a 51 character buffer where the description will  */
'*     be written.  The description is an ASCIIZ string.            */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_ATTRIBUTE_HANDLE - Invalid attributes handle.    */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetDesc Lib "CWBDQ" (queueAttributes As Long, ByVal description As String) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetForceToStorage                                        */
'*                                                                  */
'* Purpose:                                                         */
'*   Get the attribute for whether or not records will be forced to */
'*   auxiliary storage when they are enqueued.                      */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Attr queueAttributes - input                             */
'*     Handle of the data queue attributes returned by a previous   */
'*     call to cwbDQ_CreateAttr.                                    */
'*                                                                  */
'*   cwb_Boolean * forceToStorage - output                          */
'*     Pointer to a Boolean where the force to storage indicator    */
'*     will be written.                                             */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_ATTRIBUTE_HANDLE - Invalid attributes handle.    */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetForceToStorage Lib "CWBDQ" (queueAttributes As Long, forceToStorage As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetKey                                                   */
'*                                                                  */
'* Purpose:                                                         */
'*   Get the key attribute of the data object.  This is the key     */
'*   that is used for writing messages to a keyed data queue.       */
'*   Along with the search order, this key is used to read          */
'*   messages from a keyed data queue.  The key that is associated  */
'*   with the record retrieved can be obtained by calling the       */
'*   cwbDQ_GetRetKey API.                                           */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Data data - input                                        */
'*     Handle of the data object that was returned by a previous    */
'*     call to cwbDQ_CreateData.                                    */
'*                                                                  */
'*   unsigned char * key - output                                   */
'*     Pointer to the key.  The key may contain embedded NULLS,     */
'*     so it is not an ASCIIZ string.                               */
'*                                                                  */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_DATA_HANDLE - Invalid data handle.               */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetKey Lib "CWBDQ" (data As Long, key As Byte) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetKeyLen                                                */
'*                                                                  */
'* Purpose:                                                         */
'*   Get the key length attribute of the data object.               */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Data data - input                                        */
'*     Handle of the data object that was returned by a previous    */
'*     call to cwbDQ_CreateData.                                    */
'*                                                                  */
'*   unsigned short * keyLength - output                            */
'*     Pointer to an unsigned short where the length of the key     */
'*     will be written.                                             */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_DATA_HANDLE - Invalid data handle.               */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetKeyLen Lib "CWBDQ" (data As Long, keyLength As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetKeySize                                               */
'*                                                                  */
'* Purpose:                                                         */
'*   Get the attribute for the key size in bytes.                   */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Attr queueAttributes - input                             */
'*     Handle of the data queue attributes returned by a previous   */
'*     call to cwbDQ_CreateAttr.                                    */
'*                                                                  */
'*   unsigned short * keySize - output                              */
'*     Pointer to an unsigned short where the key size will         */
'*     written.                                                     */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_ATTRIBUTE_HANDLE - Invalid attributes handle.    */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetKeySize Lib "CWBDQ" (queueAttributes As Long, keySize As Byte) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetLibName                                               */
'*                                                                  */
'* Purpose:                                                         */
'*   Retrieve the library name used with the cwbDQ_Open API.        */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_QueueHandle queueHandle - input                          */
'*     Handle that was returned by a previous call to the           */
'*     cwbDQ_Open function.  This identifies the AS/400 data        */
'*     queue object.                                                */
'*                                                                  */
'*   char * libName - output                                        */
'*     Pointer to a buffer where the library name will be written.  */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_QUEUE_HANDLE - Invalid queue handle.             */
'*                                                                  */
'* Usage Notes:                                                     */
'*   This function requires that you have previously issued:        */
'*     cwbDQ_Open()                                                 */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetLibName Lib "CWBDQ" (queueHandle As Long, libName As Byte) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetMaxRecLen                                             */
'*                                                                  */
'* Purpose:                                                         */
'*   Get the maximum record length for the data queue.              */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Attr queueAttributes - input                             */
'*     Handle of the data queue attributes returned by a            */
'*     call to cwbDQ_CreateAttr.                                    */
'*                                                                  */
'*   unsigned long * maxRecordLength - output                       */
'*     Pointer to an unsigned long where the maximum record length  */
'*     will be written.                                             */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_ATTRIBUTE_HANDLE - Invalid attributes handle.    */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetMaxRecLen Lib "CWBDQ" (queueAttributes As Long, maxRecordLength As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetOrder                                                 */
'*                                                                  */
'* Purpose:                                                         */
'*   Get the attribute for the queue order.                         */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Attr queueAttributes - input                             */
'*     Handle of the data queue attributes returned by a previous   */
'*     call to cwbDQ_CreateAttr.                                    */
'*                                                                  */
'*   unsigned short * order - output                                */
'*     Pointer to an unsigned short where the order will be written.*/
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_ATTRIBUTE_HANDLE - Invalid attributes handle.    */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetOrder Lib "CWBDQ" (queueAttributes As Long, order As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetQueueAttr                                             */
'*                                                                  */
'* Purpose:                                                         */
'*   Retrieve the attributes of the AS/400 data queue object that   */
'*   is identified by the specified handle.  A handle to the data   */
'*   queue attributes will be returned.  The attributes can then    */
'*   be retrieved individually.                                     */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_QueueHandle queueHandle - input                          */
'*     Handle that was returned by a previous call to the           */
'*     cwbDQ_Open function.  This identifies the AS/400 data        */
'*     queue object.                                                */
'*                                                                  */
'*   cwbDQ_Attr queueAttributes - input/output                      */
'*     The attribute object.  This was the output from the          */
'*     cwbDQ_CreateAttr call.  The attributes will be filled in by  */
'*     this function, and you should call the cwbDQ_DeleteAttr      */
'*     function to delete this object when you have retrieved       */
'*     the attributes from it.                                      */
'*                                                                  */
'*   cwbSV_ErrHandle errorHandle - output                           */
'*     Any returned messages will be written to this object.  It    */
'*     is created with the cwbSV_CreateErrHandle API.  The          */
'*     messages may be retrieved through the cwbSV_GetErrText API.  */
'*     If the parameter is set to zero, no messages will be         */
'*     retrieved.                                                   */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_QUEUE_HANDLE - Invalid queue handle.             */
'*   CWBDQ_REJECTED_USER_EXIT - Command rejected by user exit       */
'*                              program.                            */
'*                                                                  */
'* Usage Notes:                                                     */
'*   This function requires that you have previously issued:        */
'*     cwbDQ_Open()                                                 */
'*     cwbDQ_CreateAttr()                                           */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetQueueAttr Lib "CWBDQ" (queueHandle As Long, queueAttributes As Long, errorHandle As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetQueueName                                             */
'*                                                                  */
'* Purpose:                                                         */
'*   Retrieve the queue name used with the cwbDQ_Open API.          */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_QueueHandle queueHandle - input                          */
'*     Handle that was returned by a previous call to the           */
'*     cwbDQ_Open function.  This identifies the AS/400 data        */
'*     queue object.                                                */
'*                                                                  */
'*   char * queueName - output                                      */
'*     Pointer to a buffer where the queue name will be written.    */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_QUEUE_HANDLE - Invalid queue handle.             */
'*                                                                  */
'* Usage Notes:                                                     */
'*   This function requires that you have previously issued         */
'*   cwbDQ_Open()                                                   */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetQueueName Lib "CWBDQ" (queueHandle As Long, queueName As Byte) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetRetDataLen                                            */
'*                                                                  */
'* Purpose:                                                         */
'*   Get the length of data that was returned.  The returned data   */
'*   length will be zero until a cwbDQ_Read() or cwbDQ_Peek() API   */
'*   is called, then it will have the length of the data that was   */
'*   actually returned.                                             */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Data data - input                                        */
'*     Handle of the data object that was returned by a previous    */
'*     call to cwbDQ_CreateData.                                    */
'*                                                                  */
'*   unsigned long * retDataLength - output                         */
'*     Pointer to an unsigned long where the length of the data     */
'*     returned will be written.                                    */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_DATA_HANDLE - Invalid data handle.               */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetRetDataLen Lib "CWBDQ" (data As Long, retDataLength As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetRetKey                                                */
'*                                                                  */
'* Purpose:                                                         */
'*   Get the returned key of the data object.  This is the key      */
'*   that is associated with the messages that is retrieved from    */
'*   a keyed data queue.  If the search order is a value other than */
'*   equal to, this key may be different than the key used to       */
'*   retrieve the message.                                          */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Data data - input                                        */
'*     Handle of the data object that was returned by a previous    */
'*     call to cwbDQ_CreateData.                                    */
'*                                                                  */
'*   unsigned char * retKey - output                                */
'*     Pointer to the returned key.  The key may contain embedded   */
'*     NULLs, so it is not an ASCIIZ string.                        */
'*                                                                  */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_DATA_HANDLE - Invalid data handle.               */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetRetKey Lib "CWBDQ" (data As Long, key As Byte) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetRetKeyLen                                             */
'*                                                                  */
'* Purpose:                                                         */
'*   Get the returned key length attribute of the data object.      */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Data data - input                                        */
'*     Handle of the data object that was returned by a previous    */
'*     call to cwbDQ_CreateData.                                    */
'*                                                                  */
'*   unsigned short * retKeyLength - output                         */
'*     Pointer to an unsigned short where the length of the key     */
'*     will be written.                                             */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_DATA_HANDLE - Invalid data handle.               */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetRetKeyLen Lib "CWBDQ" (data As Long, retKeyLength As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetSearchOrder                                           */
'*                                                                  */
'* Purpose:                                                         */
'*   Get the search order of the open attributes.                   */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Data data - input                                        */
'*     Handle of the data object that was returned by a previous    */
'*     call to cwbDQ_CreateData.                                    */
'*                                                                  */
'*   unsigned short * searchOrder - output                          */
'*     Pointer to an unsigned short where the order will be written.*/
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_DATA_HANDLE - Invalid data handle.               */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetSearchOrder Lib "CWBDQ" (data As Long, searchOrder As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetSenderID                                              */
'*                                                                  */
'* Purpose:                                                         */
'*   Get the attribute for whether or not information about the     */
'*   sender is kept with each record on the queue.                  */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Attr queueAttributes - input                             */
'*     Handle of the data queue attributes returned by a previous   */
'*     call to cwbDQ_CreateAttr.                                    */
'*                                                                  */
'*   cwb_Boolean * senderID - output                                */
'*     Pointer to a Boolean where the sender ID indicator will be   */
'*     written.                                                     */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_ATTRIBUTE_HANDLE - Invalid attributes handle.    */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetSenderID Lib "CWBDQ" (queueAttributes As Long, senderID As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetSenderInfo                                            */
'*                                                                  */
'* Purpose:                                                         */
'*   Get the Sender Information attribute of the open attributes.   */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Data data - input                                        */
'*     Handle of the data object that was returned by a previous    */
'*     call to cwbDQ_CreateData.                                    */
'*                                                                  */
'*   unsigned char * senderInfo - output                            */
'*     Pointer to a 36 character buffer where the sender            */
'*     information will be written.                                 */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_DATA_HANDLE - Invalid data handle.               */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetSenderInfo Lib "CWBDQ" (data As Long, senderInfo As Byte) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_GetSysName                                               */
'*                                                                  */
'* Purpose:                                                         */
'*   Retrieve the system name used with the cwbDQ_Open API.         */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_QueueHandle queueHandle -input                           */
'*     Handle that was returned by a previous call to the           */
'*     cwbDQ_Open function.  This identifies the AS/400 data        */
'*     queue object.                                                */
'*                                                                  */
'*   char * systemName - output                                     */
'*     Pointer to a buffer where the system name will be written.   */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_QUEUE_HANDLE - Invalid queue handle.             */
'*                                                                  */
'* Usage Notes:                                                     */
'*   This function requires that you have previously issued         */
'*     cwbDQ_Open()                                                 */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_GetSysName Lib "CWBDQ" (queueHandle As Long, systemName As String) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_Open                                                     */
'*                                                                  */
'* Purpose:                                                         */
'*   Start a connection to the specified data queue.  This will     */
'*   start a conversation with the AS/400 system.  If the           */
'*   connection is not successful, a non-zero handle will be        */
'*   returned.                                                      */
'*                                                                  */
'* Parameters:                                                      */
'*   char * queue - input                                           */
'*     Pointer to the data queue name contained in an ASCIIZ string.*/
'*                                                                  */
'*   char * library - input                                         */
'*     Pointer to the library name contained in an ASCIIZ string.   */
'*     If this pointer is NULL then the library list will be        */
'*     used (set library to "*LIBL").                               */
'*                                                                  */
'*   char * systemName - input                                      */
'*     Pointer to the system name contained in an ASCIIZ string.    */
'*                                                                  */
'*   cwbDQ_QueueHandle * queueHandle - output                       */
'*     Pointer to a cwbDQ_QueueHandle where the handle will be      */
'*     returned.  This handle should be used in all subsequent      */
'*     calls.                                                       */
'*                                                                  */
'*   cwbSV_ErrHandle errorHandle - output                           */
'*     Any returned messages will be written to this object.  It    */
'*     is created with the cwbSV_CreateErrHandle API.  The          */
'*     messages may be retrieved through the cwbSV_GetErrText API.  */
'*     If the parameter is set to zero, no messages will be         */
'*     retrieved.                                                   */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_COMMUNICATIONS_ERROR - A communications error occurred.    */
'*   CWB_SERVER_PROGRAM_NOT_FOUND - AS/400 application not found.   */
'*   CWB_HOST_NOT_FOUND - AS/400 system inactive or does not exist. */
'*   CWB_COMM_VERSION_ERROR - Data Queues will not run with this    */
'*                            version of communications.            */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWB_SECURITY_ERROR - a security error has occurred.            */
'*   CWB_LICENSE_ERROR - a license error has occurred.              */
'*   CWB_CONFIG_ERROR - a configuration error has occurred.         */
'*   CWBDQ_BAD_QUEUE_NAME - Queue name is too long.                 */
'*   CWBDQ_BAD_LIBRARY_NAME - library name is too long.             */
'*   CWBDQ_BAD_SYSTEM_NAME - system name is too long.               */
'*   CWBDQ_REJECTED_USER_EXIT - Command rejected by user exit       */
'*                              program.                            */
'*   CWBDQ_USER_EXIT_ERROR - error in user exit program.            */
'*   CWBDQ_LIBRARY_NOT_FOUND - library not found on system.         */
'*   CWBDQ_QUEUE_NOT_FOUND - Queue not found on system.             */
'*   CWBDQ_NO_AUTHORITY - No authority to queue or library.         */
'*   CWBDQ_DAMAGED_QUE - Queue is in unusable state.                */
'*   CWBDQ_CANNOT_CONVERT - Data cannot be converted for this queue.*/
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_Open Lib "CWBDQ" (ByVal queue As String, ByVal library As String, ByVal systemName As String, queueHandle As Long, ByVal errorHandle As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_Peek                                                     */
'*                                                                  */
'* Purpose:                                                         */
'*   Read a record from the AS/400 data queue object that is        */
'*   identified by the specified handle.  When a record is peeked   */
'*   from a data queue it remains in the data queue.  You may       */
'*   wait for a record if the data queue is empty by specifying a   */
'*   wait time from 0 to 99,999 or forever (-1).  A wait time of    */
'*   zero will return immediately if there is no data in the data   */
'*   queue.                                                         */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_QueueHandle queueHandle - input                          */
'*     Handle that was returned by a previous call to the           */
'*     cwbDQ_Open API.  This identifies the AS/400 data             */
'*     queue object.                                                */
'*                                                                  */
'*   cwbDQ_Data data - input                                        */
'*     The data object to be read from the AS/400 data queue.       */
'*                                                                  */
'*   signed long waitTime - input                                   */
'*     Length of time in seconds to wait for data, if the data      */
'*     queue is empty.  A wait time of -1 indicates to wait forever.*/
'*                                                                  */
'*   cwbSV_ErrHandle errorHandle - output                           */
'*     Any returned messages will be written to this object.  It    */
'*     is created with the cwbSV_CreateErrHandle API.  The          */
'*     messages may be retrieved through the cwbSV_GetErrText API.  */
'*     If the parameter is set to zero, no messages will be         */
'*     retrieved.                                                   */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_TIME - Invalid wait time.                        */
'*   CWBDQ_INVALID_QUEUE_HANDLE - Invalid queue handle.             */
'*   CWBDQ_INVALID_SEARCH - Invalid search order.                   */
'*   CWBDQ_DATA_TRUNCATED - Data truncated.                         */
'*   CWBDQ_TIMED_OUT - Wait time expired and no data returned.      */
'*   CWBDQ_REJECTED_USER_EXIT - Command rejected by user exit       */
'*                              program.                            */
'*   CWBDQ_QUEUE_DESTROYED - Queue was destroyed.                   */
'*   CWBDQ_CANNOT_CONVERT - Unable to convert data.                 */
'*                                                                  */
'* Usage Notes:                                                     */
'*   This function requires that you have previously issued:        */
'*     cwbDQ_Open()                                                 */
'*     cwbDQ_CreateData()                                           */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_Peek Lib "CWBDQ" (ByVal queueHandle As Long, ByVal data As Long, ByVal WaitTime As Long, ByVal errorHandle As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_Read                                                     */
'*                                                                  */
'* Purpose:                                                         */
'*   Read a record from the AS/400 data queue object that is        */
'*   identified by the specified handle.  When a record is read     */
'*   from a data queue it is removed from the data queue.  You may  */
'*   wait for a record if the data queue is empty by specifying a   */
'*   wait time from 0 to 99,999 or forever (-1).  A wait time of    */
'*   zero will return immediately if there is no data in the data   */
'*   queue.                                                         */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_QueueHandle queueHandle - input                          */
'*     Handle that was returned by a previous call to the           */
'*     cwbDQ_Open function.  This identifies the AS/400 data        */
'*     queue object.                                                */
'*                                                                  */
'*   cwbDQ_Data data - input                                        */
'*      The data object to be read from the AS/400 data queue.      */
'*                                                                  */
'*   long waitTime - input                                          */
'*     Length of time in seconds to wait for data, if the data      */
'*     queue is empty.  A wait time of -1 indicates to wait forever.*/
'*                                                                  */
'*   cwbSV_ErrHandle errorHandle - output                           */
'*     Any returned messages will be written to this object.  It    */
'*     is created with the cwbSV_CreateErrHandle API.  The          */
'*     messages may be retrieved through the cwbSV_GetErrText API.  */
'*     If the parameter is set to zero, no messages will be         */
'*     retrieved.                                                   */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_TIME - Invalid wait time.                        */
'*   CWBDQ_INVALID_QUEUE_HANDLE - Invalid queue handle.             */
'*   CWBDQ_INVALID_SEARCH - Invalid search order.                   */
'*   CWBDQ_DATA_TRUNCATED - Data truncated.                         */
'*   CWBDQ_TIMED_OUT - Wait time expired and no data returned.      */
'*   CWBDQ_REJECTED_USER_EXIT - Command rejected by user exit       */
'*                              program.                            */
'*   CWBDQ_QUEUE_DESTROYED - Queue was destroyed.                   */
'*   CWBDQ_CANNOT_CONVERT - Unable to convert data.                 */
'*                                                                  */
'* Usage Notes:                                                     */
'*   This function requires that you have previously issued         */
'*   cwbDQ_Open()                                                   */
'*     cwbDQ_CreateData()                                           */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_Read Lib "CWBDQ" (ByVal queueHandle As Long, ByVal data As Long, WaitTime As Long, ByVal errorHandle As Long) As Integer
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_SetAuthority                                             */
'*                                                                  */
'* Purpose:                                                         */
'*   Set the attribute for the authority that other users will have */
'*   to the data queue.                                             */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Attr queueAttributes - input                             */
'*     Handle of the data queue attributes returned by a previous   */
'*     call to cwbDQ_CreateAttr.                                    */
'*                                                                  */
'*   unsigned short authority - input                               */
'*     Authority that other users on the AS/400 system have to      */
'*     access the data queue.  Use one of the defined types for     */
'*     authority.                                                   */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_ATTRIBUTE_HANDLE - Invalid attributes handle.    */
'*   CWBDQ_INVALID_AUTHORITY - Invalid queue authority.             */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_SetAuthority Lib "CWBDQ" (queueAttributes As Long, ByVal authority As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_SetConvert                                               */
'*                                                                  */
'* Purpose:                                                         */
'*   Set the convert flag.  If the flag is set then all data being  */
'*   written will be converted from ASCII to EBCDIC and all data    */
'*   being read will be converted from EBCDIC to ASCII.  Default    */
'*   behavior is no conversion of data.                             */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Data data - input                                        */
'*     Handle of the data object that was returned by a previous    */
'*     call to cwbDQ_CreateData.                                    */
'*                                                                  */
'*   cwb_Boolean convert - input                                    */
'*     Flag indicating if data written to the queue will be         */
'*     converted from ASCII to EBCDIC, and data read from the       */
'*     queue will be converted from EBCDIC to ASCII.                */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_DATA_HANDLE - Invalid data handle.               */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_SetConvert Lib "CWBDQ" (ByVal data As Long, ByVal convert As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_SetData                                                  */
'*                                                                  */
'* Purpose:                                                         */
'*   Set the data and data length attributes of the data object.    */
'*   The default is to have no data with zero length.  This         */
'*   function will make a copy of the data.                         */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Data data - input                                        */
'*     Handle of the data object that was returned by a previous    */
'*     call to cwbDQ_CreateData.                                    */
'*                                                                  */
'*   unsigned char * dataBuffer - input                             */
'*     Pointer to the data.  The data may contain embedded  NULLS,  */
'*     so it is not an ASCIIZ string.                               */
'*                                                                  */
'*   unsigned long dataLength - input                               */
'*     Length of the data in bytes.                                 */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_DATA_HANDLE - Invalid data handle.               */
'*   CWBDQ_BAD_DATA_LENGTH - Length of data is not correct.         */
'*                                                                  */
'* Usage Notes:                                                     */
'*   Use this function if you want to write a small amount of data  */
'*   or you do not want to manage the memory for the data in your   */
'*   application.  Data will be copied and this may affect your     */
'*   application's performance.                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_SetData Lib "CWBDQ" (ByVal data As Long, ByVal DataBuffer As String, ByVal dataLength As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_SetDataAddr                                              */
'*                                                                  */
'* Purpose:                                                         */
'*   Set the data and data length attributes of the data object.    */
'*   The default is to have no data with zero length.  This         */
'*   function will not copy the data.                               */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Data data - input                                        */
'*     Handle of the data object that was returned by a previous    */
'*     call to cwbDQ_CreateData.                                    */
'*                                                                  */
'*   unsigned char * dataBuffer - input                             */
'*     Pointer to the data.  The data may contain embedded  NULLS,  */
'*     so it is not an ASCIIZ string.                               */
'*                                                                  */
'*   unsigned long dataLength - input                               */
'*     Length of the data in bytes.                                 */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_DATA_HANDLE - Invalid data handle.               */
'*   CWBDQ_BAD_DATA_LENGTH - Length of data is not correct.         */
'*                                                                  */
'* Usage Notes:                                                     */
'*   This function is better for large amounts of data, or if you   */
'*   want to manage memory in your application.  Data will not be   */
'*   copied so performance will be improved.                        */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_SetDataAddr Lib "CWBDQ" (ByVal data As Long, ByVal DataBuffer As String, ByVal dataLength As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_SetDesc                                                  */
'*                                                                  */
'* Purpose:                                                         */
'*   Set the attribute for the description of the data queue.       */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Attr queueAttributes - input                             */
'*     Handle of the data queue attributes returned by a  previous  */
'*     call to cwbDQ_CreateAttr.                                    */
'*                                                                  */
'*   char * description - input                                     */
'*     Pointer to an ASCIIZ string that contains the  description   */
'*     for the data queue.  The maximum length for the              */
'*     description is 50 characters.                                */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWB_INVALID_POINTER - Bad or null pointer.                     */
'*   CWBDQ_INVALID_ATTRIBUTE_HANDLE - Invalid attributes handle.    */
'*   CWBDQ_INVALID_QUEUE_TITLE - Queue title is too long.           */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_SetDesc Lib "CWBDQ" (queueAttributes As Long, description As String) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_SetForceToStorage                                        */
'*                                                                  */
'* Purpose:                                                         */
'*   Set the attribute for whether or not records will be forced to */
'*   auxiliary storage when they are enqueued.                      */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Attr queueAttributes - input                             */
'*     Handle of the data queue attributes returned by a previous   */
'*     call to cwbDQ_CreateAttr.                                    */
'*                                                                  */
'*   cwb_Boolean forceToStorage - input                             */
'*     Boolean indicator of whether or not each record is forced    */
'*     to auxiliary storage when it is enqueued.                    */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_ATTRIBUTE_HANDLE - Invalid attributes handle.    */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_SetForceToStorage Lib "CWBDQ" (queueAttributes As Long, forceToStorage As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_SetKey                                                   */
'*                                                                  */
'* Purpose:                                                         */
'*   Set the key and key length attributes of the data attributes.  */
'*   The default is to have no key with zero length.  This is the   */
'*   correct value for a non-keyed (LIFO or FIFO) data queue.       */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Data data - input                                        */
'*     Handle of the data object that was returned by a previous    */
'*     call to cwbDQ_CreateData.                                    */
'*                                                                  */
'*   unsigned char * key - input                                    */
'*     Pointer to the key.  The key may contain embedded NULLS,     */
'*     so it is not an ASCIIZ string.                               */
'*                                                                  */
'*   unsigned short keyLength - input                               */
'*     Length of the key in bytes.                                  */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_DATA_HANDLE - Invalid data handle.               */
'*   CWBDQ_BAD_KEY_LENGTH - Length of key is not correct.           */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_SetKey Lib "CWBDQ" (data As Long, key As Byte, ByVal keyLength As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_SetKeySize                                               */
'*                                                                  */
'* Purpose:                                                         */
'*   Set the attribute for the key size in bytes.                   */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Attr queueAttributes - input                             */
'*     Handle of the data queue attributes returned by a previous   */
'*     call to cwbDQ_CreateAttr.                                    */
'*                                                                  */
'*   unsigned short keySize - input                                 */
'*     Size in bytes of the key.  This value should be zero if      */
'*     the order is LIFO or FIFO.                                   */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_KEY_LENGTH - Invalid key length.                 */
'*   CWBDQ_INVALID_ATTRIBUTE_HANDLE - Invalid attributes handle.    */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_SetKeySize Lib "CWBDQ" (queueAttributes As Long, ByVal keySize As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_SetMaxRecLen                                             */
'*                                                                  */
'* Purpose:                                                         */
'*   Get the maximum record length for the data queue.              */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Attr queueAttributes - input                             */
'*     Handle of the data queue attributes returned by a  previous  */
'*     call to cwbDQ_CreateAttr.                                    */
'*                                                                  */
'*   unsigned long maxLength - input                                */
'*     Maximum length for a data queue record.                      */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_ATTRIBUTE_HANDLE - Invalid attributes handle.    */
'*   CWBDQ_INVALID_QUEUE_LENGTH - Invalid queue record length.      */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_SetMaxRecLen Lib "CWBDQ" (ByVal queueAttributes As Long, ByVal maxRecordLength As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_SetOrder                                                 */
'*                                                                  */
'* Purpose:                                                         */
'*   Set the attribute for the queue order.                         */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Attr queueAttributes - input                             */
'*     Handle of the data queue attributes returned by a previous   */
'*     call to cwbDQ_CreateAttr.                                    */
'*                                                                  */
'*   unsigned short order - input                                   */
'*     Order in which new entries will be enqueued.  Use one of     */
'*     the defined types for order.                                 */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_ATTRIBUTE_HANDLE - Invalid attributes handle.    */
'*   CWBDQ_INVALID_ORDER - Invalid queue order.                     */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_SetOrder Lib "CWBDQ" (ByVal queueAttributes As Long, ByVal order As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_SetSearchOrder                                           */
'*                                                                  */
'* Purpose:                                                         */
'*   Set the search order of the open attributes.  The default is   */
'*   no search order.  If the cwbDQ_SetKey API is called, the       */
'*   search order is changed to equal.  Use this API to set it to   */
'*   something else.                                                */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Data data - input                                        */
'*     Handle of the data object that was returned by a previous    */
'*     call to cwbDQ_CreateData.                                    */
'*                                                                  */
'*   unsigned short searchOrder - input                             */
'*     Order to use when reading from a keyed queue.                */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_DATA_HANDLE - Invalid data handle.               */
'*   CWBDQ_INVALID_SEARCH - Invalid search order.                   */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_SetSearchOrder Lib "CWBDQ" (data As Long, ByVal searchOrder As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_SetSenderID                                              */
'*                                                                  */
'* Purpose:                                                         */
'*   Set the attribute for whether or not information about the     */
'*   sender is kept with each record on the queue.                  */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_Attr queueAttributes - input                             */
'*     Handle of the data queue attributes returned by a previous   */
'*     call to cwbDQ_CreateAttr.                                    */
'*                                                                  */
'*   cwb_Boolean senderID - input                                   */
'*     Boolean indicator of whether or not information about the    */
'*     sender is kept with record on the queue.                     */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_INVALID_ATTRIBUTE_HANDLE - Invalid attributes handle.    */
'*                                                                  */
'* Usage Notes:                                                     */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_SetSenderID Lib "CWBDQ" (queueAttributes As Long, senderID As Long) As Integer
 
'********************************************************************/
'*                                                                  */
'* API:                                                             */
'*   cwbDQ_Write                                                    */
'*                                                                  */
'* Purpose:                                                         */
'*   Write a record to the AS/400 data queue object that is         */
'*   identified by the specified handle.  Writing with commit ON    */
'*   means that your application will not have control returned to  */
'*   it until after the message has been enqueued.                  */
'*                                                                  */
'* Parameters:                                                      */
'*   cwbDQ_QueueHandle queueHandle - input                          */
'*     Handle that was returned by a previous call to the           */
'*     cwbDQ_Open function.  This identifies the AS/400 data        */
'*     queue object.                                                */
'*                                                                  */
'*   cwbDQ_Data data - input                                        */
'*     The data object to be written to the AS/400 data queue.      */
'*                                                                  */
'*   cwb_Boolean commit - input                                     */
'*     Boolean flag indicating if the data should be committed on   */
'*     write or not.                                                */
'*                                                                  */
'*   cwbSV_ErrHandle errorHandle - output                           */
'*     Any returned messages will be written to this object.  It    */
'*     is created with the cwbSV_CreateErrHandle API.  The          */
'*     messages may be retrieved through the cwbSV_GetErrText API.  */
'*     If the parameter is set to zero, no messages will be         */
'*     retrieved.                                                   */
'*                                                                  */
'* Return Codes:                                                    */
'*   The following list shows common return values.                 */
'*                                                                  */
'*   CWB_OK - Successful completion.                                */
'*   CWBDQ_BAD_DATA_LENGTH - Length of data is not correct.         */
'*   CWBDQ_INVALID_MESSAGE_LENGTH - Invalid message length.         */
'*   CWBDQ_INVALID_QUEUE_HANDLE - Invalid queue handle.             */
'*   CWBDQ_REJECTED_USER_EXIT - Command rejected by user exit       */
'*                              program.                            */
'*   CWBDQ_CANNOT_CONVERT - Unable to convert data.                 */
'*                                                                  */
'* Usage Notes:                                                     */
'*   This function requires that you have previously issued:        */
'*     cwbDQ_Open()                                                 */
'*     cwbDQ_CreateData()                                           */
'*                                                                  */
'********************************************************************/
Declare Function cwbDQ_Write Lib "CWBDQ" (ByVal queueHandle As Long, ByVal data As Long, ByVal commit As Long, ByVal errorHandle As Long) As Integer
 

