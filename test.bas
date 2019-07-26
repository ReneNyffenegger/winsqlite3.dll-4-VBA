option explicit

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

    sqlite3_bind_int  stmt, 1, 4 
    sqlite3_bind_text stmt, 2,"four" , -1, 0 
    sqlite3_bind_null stmt, 3 
    sqlite3_step      stmt 

    sqlite3_finalize  stmt

    closeDB(db)

end sub ' }

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
