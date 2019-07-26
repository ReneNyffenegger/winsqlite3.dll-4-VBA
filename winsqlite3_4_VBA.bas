option explicit

public const SQLITE_OK      =    0  ' Successful result
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

' sqlite3_open {
declare ptrSafe function sqlite3_open        lib "winsqlite3.dll" (  _
     byVal    zFilename        as string , _
     byRef    ppDB             as any      _
) as longPtr ' }

' sqlite3_close {
declare ptrSafe function sqlite3_close       lib "winsqlite3.dll" ( _
     byVal    db               as any      _
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
     byVal    stmt             as any      _
) as longPtr ' }

' sqlite3_bind_* ' {

' sqlite3_bind_int {
declare ptrSafe function sqlite3_bind_int    lib "winsqlite3.dll" ( _
     byVal    stmt            as any     , _
     byVal    pos             as long    , _
     byVal    val             as long      _
) as long ' }

' sqlite3_bind_text {
'
'     TODO: whatIsThis
'
declare ptrSafe function sqlite3_bind_text   lib "winsqlite3.dll" ( _
     byVal    stmt            as any     , _
     byVal    pos             as long    , _
     byVal    val             as string  , _
     byVal    len_            as integer , _
     byVal    whatIsThis      as longPtr   _
) as long ' }

' sqlite3_bind_null {
declare ptrSafe function sqlite3_bind_null   lib "winsqlite3.dll" ( _
     byVal    stmt            as any     , _
     byVal    pos             as long      _
) as long
' }

' }

' sqlite3_step ' {
declare ptrSafe function sqlite3_step          lib "winsqlite3.dll" ( _
     byVal     stmt           as any       _
) as longPtr ' }

' sqlite3_column_* {
' sqlite3_column_double ' {
declare ptrSafe function sqlite3_column_double lib "winsqlite3.dll" ( _
     byVal     stmt           as any     , _
     byVal     iCol           as integer   _
) as double ' }

' sqlite3_column_int ' {
declare ptrSafe function sqlite3_column_int    lib "winsqlite3.dll" ( _
     byVal     stmt           as any     , _
     byVal     iCol           as integer   _
) as integer ' }

' sqlite3_column_text ' {
declare ptrSafe function sqlite3_column_text   lib "winsqlite3.dll" ( _
     byVal     stmt           as any     , _
     byVal     iCol           as integer   _
) as string ' }

' }

' sqlite3_column_type ' {
'
' Returns one of the five data types (SQLITE_INTEGER, …) of
' a selected column.
'
declare ptrSafe function sqlite3_column_type   lib "winsqlite3.dll" ( _
     byVal     stmt           as any     , _
     byVal     iCol           as integer   _
) as integer ' }
