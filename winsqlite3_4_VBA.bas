option explicit

public const SQLITE_OK      =    0  ' Successful result
public const SQLITE_ERROR   =    1  ' Generic Error
public const SQLITE_BUSY    =    5  ' The database file is locked
public const SQLITE_TOOBIG  =   18  ' String or BLOB exceeds size limit
public const SQLITE_MISUSE  =   21  ' Library used incorrectly
public const SQLITE_ROW     =  100  ' sqlite3_step() has another row ready
public const SQLITE_DONE    =  101  ' sqlite3_step() has finished executing

' Constants that identify SQLite data types ' {
'
' Returned by sqlite3_column_type()
'
public const SQLITE_INTEGER =  1
public const SQLITE_FLOAT   =  2
public const SQLITE_TEXT    =  3  ' TODO: Should it be named SQLITE3_TEXT?
public const SQLITE_BLOB    =  4
public const SQLITE_NULL    =  5
' }

' { WinAPI constants
private const CP_UTF8        = 65001
' }

' sqlite3_open {
declare ptrSafe function sqlite3_open        lib "winsqlite3.dll" (  _
     byVal    zFilename        as string , _
     byRef    ppDB             as longPtr  _
) as longPtr ' }

' sqlite3_close {
declare ptrSafe function sqlite3_close       lib "winsqlite3.dll" ( _
     byVal    db               as longPtr  _
) as longPtr ' }

' sqlite3_exec ' {
'
'   TODO: how can/should errmsg be handled in case of an error
'
declare ptrSafe function sqlite3_exec        lib "winsqlite3.dll" ( _
     byVal    db               as any    , _
     byVal    sql              as string , _
     byVal    callback         as longPtr, _
     byVal    argument_1       as longPtr, _
     byRef    errmsg           as string   _
) as longPtr ' }

' sqlite3_prepare_v2 {
'
'   TODO pzTail is actually a char**
'
declare ptrSafe function sqlite3_prepare_v2  lib "winsqlite3.dll" ( _
     byVal    db               as any    , _
     byVal    zSql             as string , _
     byVal    nByte            as longPtr, _
     byRef    ppStatement      as longPtr, _
     byRef    pzTail           as any      _
) as longPtr ' }

' sqlite3_finalize {
'
'
declare ptrSafe function sqlite3_finalize    lib "winsqlite3.dll" ( _
     byVal    stmt             as longPtr  _
) as longPtr ' }

' sqlite3_bind_* ' {

' sqlite3_bind_int {
declare ptrSafe function sqlite3_bind_int    lib "winsqlite3.dll" ( _
     byVal    stmt            as longPtr , _
     byVal    pos             as long    , _
     byVal    val             as long      _
) as long ' }

' sqlite3_bind_text {
'
'     TODO: what is the «whatIsThis» parameter used for?
'
declare ptrSafe function sqlite3_bind_text_  lib "winsqlite3.dll" alias "sqlite3_bind_text" ( _
     byVal    stmt            as longPtr , _
     byVal    pos             as long    , _
     byVal    val             as longPtr , _
     byVal    len_            as integer , _
     byVal    whatIsThis      as longPtr   _
) as long ' }

' sqlite3_bind_null {
declare ptrSafe function sqlite3_bind_null   lib "winsqlite3.dll" ( _
     byVal    stmt            as longPtr , _
     byVal    pos             as long      _
) as long
' }

' }

' sqlite3_step ' {
declare ptrSafe function sqlite3_step          lib "winsqlite3.dll" ( _
     byVal     stmt           as longPtr   _
) as long ' }


' sqlite3_reset ' {
declare ptrSafe function sqlite3_reset         lib "winsqlite3.dll" ( _
     byVal     stmt           as longPtr   _
) as long ' }


' sqlite3_column_* {

' sqlite3_column_double ' {
declare ptrSafe function sqlite3_column_double lib "winsqlite3.dll" ( _
     byVal     stmt           as longPtr , _
     byVal     iCol           as integer   _
) as double ' }

' sqlite3_column_int ' {
declare ptrSafe function sqlite3_column_int    lib "winsqlite3.dll" ( _
     byVal     stmt           as longPtr , _
     byVal     iCol           as integer   _
) as integer ' }

' sqlite3_column_text ' {
'
' The string returned from SQLite needs to be converted to
' a wide character string in VBA. The following declaration (whose
' name ends in an underscore) first gets the pointer to the (ASCII or UTF8) string.
' Further below, another function is declared, sqlite3_column_text, that
' takes the pointer and converts it into a wide character string, suitable
' for a VBA string.
'

declare ptrSafe function sqlite3_column_text_  lib "winsqlite3.dll" alias "sqlite3_column_text" ( _
     byVal     stmt           as longPtr , _
     byVal     iCol           as integer   _
) as longPtr ' }

' }

' sqlite3_column_type ' {
'
' Returns one of the five data types (SQLITE_INTEGER, …) of
' a selected column.
'
declare ptrSafe function sqlite3_column_type   lib "winsqlite3.dll" ( _
     byVal     stmt           as longPtr , _
     byVal     iCol           as integer   _
) as integer ' }

' MultiByteToWideChar {
private declare ptrSafe function MultiByteToWideChar lib "kernel32" ( _
   byVal CodePage       as long   , _
   byVal dwFlags        as long   , _
   byVal lpMultiByteStr as longPtr, _
   byVal cbMultiByte    as long   , _
   byVal lpWideCharStr  as longPtr, _
   byVal cchWideChar    as long     _
) as long ' }

' WideCharToMultiByte {
private declare ptrSafe function WideCharToMultiByte lib "kernel32" ( _
  byVal CodePage          as long   , _
  byVal dwFlags           as long   , _
  byVal lpWideCharStr     as longPtr, _
  byVal cchWideChar       as long   , _
  byVal lpMultiByteStr    as longPtr, _
  byVal cchMultiByte      as long   , _
  byVal lpDefaultChar     as longPtr, _
  byVal lpUsedDefaultChar as longPtr  _
) as long ' }

function utf8ptrToString(byVal pUtf8String as longPtr) as string ' {
'
' Found @ https://github.com/govert/SQLiteForExcel/blob/master/Source/SQLite3VBAModules/Sqlite3_64.bas
'
    dim buf     as string
    dim cSize   as long
    dim retVal  as long

  ' cSize includes the terminating null character
    cSize = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, 0, 0)

    if cSize <= 1 then ' {
        Utf8ptrToString = ""
        exit function
    end if ' }

    Utf8ptrToString = string(cSize - 1, "*") ' and a termintating null char.

    retVal = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, strPtr(Utf8ptrToString), cSize)
    if retVal = 0 then ' {
       err.raise 1000, "utf8ptrToString", "Utf8ptrToString error: " & err.lastDllError
       exit function
    end if ' }

end function ' }

function stringToUtf8bytes(byVal txt as string) as byte() ' {

    dim bSize  as long
    dim retVal as long
    dim buf()  as byte

    bSize = WideCharToMultiByte(CP_UTF8, 0, strPtr(txt), -1, 0, 0, 0, 0)

    if bSize = 0 then ' {
        exit function
    end if ' }

    ReDim buf(bSize)

    retVal = WideCharToMultiByte(CP_UTF8, 0, strPtr(txt), -1, varPtr(buf(0)), bSize, 0, 0)

    if retVal = 0 then
        err.raise 1000, "stringToUtf8bytes", "stringToUtf8bytes error: " & err.lastDllError
        exit function
    end if

    stringToUtf8bytes = buf

end function ' }

' { sqlite3_bind_text

function sqlite3_bind_text  ( _
     byVal    stmt            as longPtr , _
     byVal    pos             as long    , _
     byVal    val             as string  , _
     byVal    len_            as integer , _
     byVal    whatIsThis      as longPtr   _
) as long ' }

  dim arrayVariant as variant
  arrayVariant = stringToUtf8bytes(val)

' dim x() as byte
' x = stringToUtf8bytes(val)

  sqlite3_bind_text = sqlite3_bind_text_(stmt, pos, varPtr(arrayVariant          ), len_, whatIsThis)
' sqlite3_bind_text = sqlite3_bind_text_(stmt, pos, varPtr(stringToUtf8bytes(val)), len_, whatIsThis)
' sqlite3_bind_text = sqlite3_bind_text_(stmt, pos,        stringToUtf8bytes(val ), len_, whatIsThis)

end function ' }

' { sqlite3_column_text
function sqlite3_column_text (             _
     byVal     stmt           as longPtr , _
     byVal     iCol           as integer   _
) as string

    sqlite3_column_text = utf8ptrToString(sqlite3_column_text_(stmt, iCol))

end function ' }
