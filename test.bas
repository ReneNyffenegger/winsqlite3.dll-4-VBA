#
#   Compare with https://renenyffenegger.ch/notes/development/databases/SQLite/VBA/index
#

$dbFileName = 'c:\users\rene\test.db'
remove-item $dbFileName -errorAction ignore

add-type -typeDefinition @"
using System;
using System.Runtime.InteropServices;
 
public static partial class sqlite {
//
// Version 0.01 

   [DllImport("winsqlite3.dll", CharSet=CharSet.Ansi)]
    public static extern IntPtr sqlite3_open(
           String zFilename,
       ref IntPtr ppDB       // db handle
    );

   [DllImport("winsqlite3.dll", CharSet=CharSet.Ansi)]
    public static extern IntPtr sqlite3_exec(
           IntPtr db      ,    /* An open database                                               */
           String sql     ,    /* SQL to be evaluated                                            */
           IntPtr callback,    /*  int (*callback)(void*,int,char**,char**) -- Callback function */
           IntPtr cb1stArg,    /* 1st argument to callback                                       */
       ref String errMsg       /* Error msg written here  ( char **errmsg)                       */
    );

   [DllImport("winsqlite3.dll", CharSet=CharSet.Ansi)]
    public static extern IntPtr sqlite3_prepare_v2(
           IntPtr db      ,     /* Database handle */
           String zSql    ,     /* SQL statement, UTF-8 encoded */
           IntPtr nByte   ,     /* Maximum length of zSql in bytes. */
      ref  IntPtr sqlite3_stmt, /* int **ppStmt -- OUT: Statement handle */
    //  ref  String pzTail        /*  const char **pzTail  --  OUT: Pointer to unused portion of zSql */
           IntPtr pzTail       /*  const char **pzTail  --  OUT: Pointer to unused portion of zSql */
    );

   [DllImport("winsqlite3.dll")]
    public static extern IntPtr sqlite3_bind_int(
           IntPtr    stmt,
           IntPtr /* int */ index,
           IntPtr /* int */ value);

   [DllImport("winsqlite3.dll", CharSet=CharSet.Ansi)]
    public static extern IntPtr sqlite3_bind_text(
           IntPtr    stmt,
           IntPtr    index,
           String    value , /* const char*  */
           IntPtr    x     , /* What does this parameter do? */
           IntPtr    y       /* void(*)(void*) */
     );
   [DllImport("winsqlite3.dll")]
    public static extern IntPtr sqlite3_bind_null(
           IntPtr    stmt,
           IntPtr    index
    );

   [DllImport("winsqlite3.dll")]
    public static extern IntPtr sqlite3_step(
           IntPtr    stmt
    );

   [DllImport("winsqlite3.dll")]
    public static extern IntPtr sqlite3_reset(
           IntPtr    stmt
    );

   [DllImport("winsqlite3.dll")]
    public static extern IntPtr sqlite3_clear_bindings(
           IntPtr    stmt
    );

}
"@

[IntPtr]$db = 0
$res = [sqlite]::sqlite3_open($dbFileName, [ref] $db)
echo "$res , $db"

[String]$errMsg = ''
$res = [sqlite]::sqlite3_exec($db, 'create table tab(foo, bar, baz', 0, 0, [ref] $errMsg)
echo "$res, $errMsg"

$res = [sqlite]::sqlite3_exec($db, 'create table tab(foo, bar, baz)', 0, 0, [ref] $errMsg)
echo "$res, $errMsg"

$res = [sqlite]::sqlite3_exec($db, 'create table tab(foo, bar, baz)', 0, 0, [ref] $errMsg)
echo "$res, $errMsg"

echo "preparing statement"

[IntPtr] $stmt = 0
[String] $pzTail = ''
$res = [sqlite]::sqlite3_prepare_v2($db, 'insert into tab values(?, ?, ?)', -1, [ref] $stmt, 0)
echo "$res"

echo "Binding values for first row"

$res = [sqlite]::sqlite3_bind_int($stmt, 1, 55)
echo "$res"

$res = [sqlite]::sqlite3_bind_text($stmt, 2, 'four', -1, 0)
echo "$res"

$res = [sqlite]::sqlite3_bind_int($stmt, 3, 333)
echo "$res"

echo "Inserting 1st row"
$res = [sqlite]::sqlite3_step($stmt)
echo "$res (101 is expected!)"

# $res = [sqlite]::sqlite3_clear_bindings($stmt)
$res = [sqlite]::sqlite3_reset($stmt)
echo "$res"

echo "Binding values for second row"

$res = [sqlite]::sqlite3_bind_int($stmt, 1, 42)
echo "$res"

$res = [sqlite]::sqlite3_bind_text($stmt, 2, 'forty-two', -1, 0)
echo "$res"

$res = [sqlite]::sqlite3_bind_null($stmt, 3)
echo "$res"

echo "Inserting 2nd row"
$res = [sqlite]::sqlite3_step($stmt)
echo "$res (101 is expected!)"

# $res = [sqlite]::sqlite3_clear_bindings($stmt)
# $res = [sqlite]::sqlite3_reset($stmt)
# echo "$res"
