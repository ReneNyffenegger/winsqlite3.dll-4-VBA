option explicit

Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long

sub main() ' {

    dim db as longPtr

    db = openDB("C:\Users\OMIS~1.REN\AppData\Local\Temp\test.db")

    execSQL db, "create table tab(foo, bar, baz)"

    execSQL db, "insert into tab values(1, 'one', null);"
    execSQL db, "insert into tab values(2,  2.2 ,'two');"

    dim stmt as longPtr
    stmt = prepareStmt(db, "insert into tab values(?, ?, ?)")

    sqlite3_bind_int  stmt, 1, 3 
    sqlite3_bind_text stmt, 2,"three", -1, 0 
    sqlite3_bind_int  stmt, 3, 333 
    sqlite3_step      stmt 

    sqlite3_bind_int  stmt, 1, 55
    sqlite3_bind_text stmt, 2,"four" , -1, 0 
    sqlite3_bind_null stmt, 3 
    sqlite3_step      stmt 

    sqlite3_finalize  stmt

    selectFromTab(db)

    closeDB(db)

end sub ' }


function utf8PtrToString(byVal pUtf8String as longPtr) as string ' {
'
' Found @ https://github.com/govert/SQLiteForExcel/blob/master/Source/SQLite3VBAModules/Sqlite3_64.bas
'

    dim buf     as string
    dim cSize   as long
    dim RetVal  as long
    dim CP_UTF8 as long

    CP_UTF8 = 65001

  ' cSize includes the terminating null character
    cSize = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, 0, 0)

    if cSize <= 1 then
        Utf8PtrToString = ""
        exit function
    end if

    Utf8PtrToString = string(cSize - 1, "*") ' and a termintating null char.

    retVal = MultiByteToWideChar(CP_UTF8, 0, pUtf8String, -1, strPtr(Utf8PtrToString), cSize)
    if retVal = 0 Then
       err.raise "Utf8PtrToString error: " & err.lastDllError
       exit function
    end if

end function ' }

function openDB(fileName as string) as longPtr ' {

    dim res as longPtr

    res = sqlite3_open(fileName, openDB)
    if res <> SQLITE_OK then
       err.raise("sqlite_open failed, res = " & res)
    end if

    debug.print("SQLite db opened, db = " & openDB)

end function ' }

sub closeDB(db as longPtr) ' {

    dim res as longPtr

    res = sqlite3_close(db)
    if res <> SQLITE_OK then
       err.raise("sqlite_open failed, res = " & res)
    end if

end sub ' }

sub execSQL(db as longPtr, sql as string) ' {

    dim res    as longPtr
    dim errmsg as string

    res = sqlite3_exec(db, sql, 0, 0, errmsg)
    if res <> SQLITE_OK then
       err.raise("sqlite3_exec failed, res = " & res)
    end if

end sub ' }

function prepareStmt(db as longPtr, sql as string) as longPtr ' {

    dim res    as longPtr

    res = sqlite3_prepare_v2(db, sql, -1, prepareStmt, 0)
    if res <> SQLITE_OK then
       err.raise("sqlite3_prepare failed, res = " & res)
    end if

    debug.print("stmt = " & prepareStmt)

end function ' }

sub selectFromTab(db as longPtr) ' {

    dim stmt as longPtr
    stmt = prepareStmt(db, "select * from tab where foo > ? order by foo")

    sqlite3_bind_int stmt, 1, 2

    dim rowNo as long

    while sqlite3_step(stmt) <> SQLITE_DONE ' {

      rowNo = rowNo + 1

      dim colNo as long
      colNo = 0
      while colNo <= 2 ' {

         if     sqlite3_column_type(stmt, colNo) = SQLITE_INTEGER then

                cells(rowNo, colNo + 1) = sqlite3_column_int(stmt, colNo)

         elseIf sqlite3_column_type(stmt, colNo) = SQLITE_FLOAT   then

                cells(rowNo, colNo + 1) = sqlite3_column_double(stmt, colNo)

         elseIf sqlite3_column_type(stmt, colNo) = SQLITE_TEXT    then

                cells(rowNo, colNo + 1) = Utf8PtrToString(sqlite3_column_text(stmt, colNo)) ' TODO: wrap call to Utf8PtrToString

         elseIf sqlite3_column_type(stmt, colNo) = SQLITE_NULL    then

                cells(rowNo, colNo + 1) ="n/a"

         else

                cells(rowNo, colNo + 1) ="?"

         end if

         colNo = colNo + 1

      wend ' }

    wend ' }

    sqlite3_finalize stmt

end sub ' }
