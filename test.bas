option explicit


sub main() ' {

    dim db as longPtr

    db = openDB(environ("temp") & "\test.db")

    execSQL db, "create table tab(foo, bar, baz)"

    execSQL db, "insert into tab values(1, 'one', null);"
    execSQL db, "insert into tab values(2,  2.2 ,'two');"

    dim stmt as longPtr
    stmt = prepareStmt(db, "insert into tab values(?, ?, ?)")

    checkBindRetval(sqlite3_bind_int (stmt, 1, 3              ))
    checkBindRetval(sqlite3_bind_text(stmt, 2,"three", -1, 0  ))
    checkBindRetval(sqlite3_bind_int (stmt, 3, 333            ))
    checkStepRetval(sqlite3_step     (stmt))

  ' sqlite3_reset(stmt) still seems necesssary although the documentation says
  ' that in version after 3.6.something, it should not be necessary anymore...
  '
  ' TODO: or should sqlite3_clear_bindings be used?
  '
    sqlite3_reset(stmt) ' Or sqlite3_clear_bindings() ?

    checkBindRetval(sqlite3_bind_int (stmt, 1, 55             ))
    checkBindRetval(sqlite3_bind_text(stmt, 2,"four" , -1, 0  ))
    checkBindRetval(sqlite3_bind_null(stmt, 3                 ))
    checkStepRetval(sqlite3_step     (stmt))
    sqlite3_reset(stmt) ' Or sqlite3_clear_bindings() ?

    checkBindRetval(sqlite3_bind_int (stmt, 1, 42                 ))
    checkBindRetval(sqlite3_bind_text(stmt, 2,"Umlauts"   , -1, 0 ))
    checkBindRetval(sqlite3_bind_text(stmt, 3,"äöü ÄÖÜ éÉ", -1, 0 ))
    checkStepRetval(sqlite3_step     (stmt))
'   sqlite3_reset(stmt) ' Or sqlite3_clear_bindings() ?

    sqlite3_finalize  stmt

    selectFromTab(db)

    closeDB(db)

end sub ' }

function openDB(fileName as string) as longPtr ' {

    dim res as longPtr

    res = sqlite3_open(fileName, openDB)
    if res <> SQLITE_OK then
       err.raise 1000, "openDB", "sqlite_open failed, res = " & res
    end if

    debug.print("SQLite db opened, db = " & openDB)

end function ' }

sub closeDB(db as longPtr) ' {

    dim res as longPtr

    res = sqlite3_close(db)
    if res <> SQLITE_OK then
       err.raise 1000, "closeDB", "sqlite_open failed, res = " & res
    end if

end sub ' }

sub checkBindRetval(retVal as long) ' {

    if retVal = SQLITE_OK then
       exit sub
    end if

    if retVal = SQLITE_TOOBIG then
       err.raise 1000, "checkBindRetval", "bind failed: String or BLOB exceeds size limit"
    end if

    if retVal = SQLITE_MISUSE then
       err.raise 1000, "checkBindRetval", "bind failed: Library used incorrectly"
    end if

    err.raise 1000, "checkBindRetval", "bind failed, retVal = " & retVal

end sub ' }

sub checkStepRetval(retVal as long) ' {

    if retVal = SQLITE_DONE then
       exit sub
    end if

    err.raise 1000, "checkStepRetval", "step failed, retVal = " & retVal

end sub ' }

sub execSQL(db as longPtr, sql as string) ' {

    dim res    as longPtr
    dim errmsg as string

    res = sqlite3_exec(db, sql, 0, 0, errmsg)
    if res <> SQLITE_OK then
       err.raise 1000, "execSQL", "sqlite3_exec failed, res = " & res
    end if

end sub ' }

function prepareStmt(db as longPtr, sql as string) as longPtr ' {

    dim res    as longPtr

    res = sqlite3_prepare_v2(db, sql, -1, prepareStmt, 0)
    if res <> SQLITE_OK then
       err.raise 1000, "prepareStmt", "sqlite3_prepare failed, res = " & res
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

                cells(rowNo, colNo + 1) = sqlite3_column_text(stmt, colNo)

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
