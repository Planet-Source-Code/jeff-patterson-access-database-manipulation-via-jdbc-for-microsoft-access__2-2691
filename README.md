<div align="center">

## Access Database Manipulation via JDBC \(for Microsoft Access\)


</div>

### Description

This will teach you how to connect to a Microsoft Access database. It's also a great overview of JDBC. Once you are connected, you may run any SQL statement that is allowable on Access, such as SELECT, etc. You don't even have to have MS Access installed to run this tutorial - it shows you how to make a blank one without Access!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jeff Patterson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeff-patterson.md)
**Level**          |Beginner
**User Rating**    |4.8 (2833 globes from 591 users)
**Compatibility**  |Java \(JDK 1\.1\), Java \(JDK 1\.2\)
**Category**       |[Databases/ JDBC](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-jdbc__2-61.md)
**World**          |[Java](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/java.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jeff-patterson-access-database-manipulation-via-jdbc-for-microsoft-access__2-2691/archive/master.zip)





### Source Code

<BR><BR>
<font size="2" face="arial"><I>Sorry if the formatting is a little screwed up on this - PlanetSourceCode seems to modify my HTML just a little when I upload it...it should still all be readable enough...</I></font>
<BR><BR>
<P>
<a href="#VOTE">If you find this useful, please vote for me!</a>
<P>
<center>
<font size="5"><B>How to manipulate a Microsoft Access Database via JDBC</B></font><BR>
and it's also<BR><font size="4"><B>A Super Quick Overview of JDBC Basics</B></font>
</center>
<P>
This will teach you how to connect to a Microsoft Access database. Once you are connected, you may run any SQL statement that is allowable on Access, such as:
<font size="3">
<ul>
<li>a <code><B>SELECT</B></code> statement to retrieve data
<li>an <code><B>INSERT</B></code> statement to add data
<li>a <code><B>DELETE</B></code> statement to remove data
<li>an <code><B>CREATE TABLE</B></code> statement to build a new table
<li>a <code><B>DROP TABLE</B></code> statement to destroy a table
</ul>
</font>
This document goes at a pretty slow pace, so you may not need to cover every little detail here. If you are entirely new to JDBC, you shouldn't have too much trouble following along. So let's get going!
<P>
<a name="MY_TOP"><h2>Steps to take:</h2></a>
<P>
There are three things we need to do to manipulate a MS Access database:<BR> 1) Set up Java to undestand ODBC, <BR> 2) Get a connection to our MS Access Database, <BR> 3) Run a SQL statement.
<P>
<font size="3"><B>1) First we need to set up Java to understand how to communicate with an ODBC data source</B><BR></font>
<ul>
<font size="3"><li><a href="#SECTION0">Set up your DriverManager to understand ODBC data sources</a></font><BR>
</ul>
<font size="3"><a name="NEXT"><B>2) After we set up the DriverManager, we need to get a Connection</B></a><BR></font>
  There are two ways to get a connection from your Microsoft Access Database:
<ol>
<font size="3"><li><a href="#SECTION1">Get a connection by accessing the Database Directly</a></font><BR>
The simpler way, but may not work on all systems!
<font size="3"><li><a href="#SECTION2">Set the Access Database up as an ODBC DSN and get a connection through that</a></font><BR>
A little more complex, but will work on any system, and will work even if you don't already have a Microsoft Access Database!
</ol>
<P>
<font size="3"><a name="SQL"><B>3) Once you have gained access to the Database (been granted a connection), you are ready to try:</B></a><BR></font>
<ul>
<font size="3"><li><a href="#SECTION_SQL">Running a SQL Statement on your Access Database</a></font><BR>
This is the section that you will be most interested in - if you're impatient, you might want to start here...<I>but please come back and read it all!</I>
</ul>
<BR><BR>
In addition, please refer to the section at the end of this document:
<ul>
<font size="3"><li><a href="#SECTION_LAST">What I assume you already know</a></font><BR>
Plus a little additional reading.
</ul>
<hr size=1 noshade>
<P>
<table border=0 cellpadding=4 cellspacing=0 bgcolor=lightblue width=100%><tr><td>
<font size="4"><a name="SECTION0">Step 1) Set up your DriverManager to understand ODBC data sources</a></font>
</td><td align=right valign=top><font size="1"><a href="#MY_TOP">BACK TO TOP</a></font></td></tr></table>
The first thing we must do in order to manipulate data in the database is to be granted a connection to the database. This connection, referenced in the Java language as an Object of type <font size="+1"><code><B>java.sql.Connection</B></code></font>, is handed out by the <B>DriverManager</B>. We tell the DriverManager what type of driver to use to handle the connections to databases, and from there, ask it to give us a connection to a particular database of that type.<P>
For this tutorial, we are interested in accessing a Microsoft Access database. Microsoft has developed a data access method called <B>ODBC</B>, and MS Access databases understand this method. We cannot make a connection directly to an ODBC data source from Java, but Sun has provided a <B>bridge</B> from JDBC to ODBC. This bridge gives the DriverManager the understanding of how to communicate with an ODBC (ie a MS Access) data source.
<P>
So the first thing we'll do is set up our DriverManager and let it know that we want to communicate with ODBC data sources via the JDBC:ODBC bridge. We do this by calling the static <font size=+1><code>forName()</code></font> method of the Class class. Here is an entire program that accomplishes what we're after:
<table border=1 cellpadding=4 cellspacing=0 bgcolor=#DDDDDD>
<tr><td><pre><code>class Test
{
 public static void main(String[] args)
 {
 try {
  Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
 }
 catch (Exception e) {
  System.out.println("Error: " + e);
 }
 }
}</code></pre><font color=red>//save this code into a file called <B>Test.java</B> and compile it</font></td></tr></table>
Notice the TRY-CATCH block. The forName() method might throw a <I>ClassNotFoundException</I>. This really can't happen with the JDBC:ODBC bridge, since it's built in to the Java API, but we still have to catch it. If you compile and run this code, it's pretty boring. In fact, if it produces any output, then that means that you've encountered an error! But it shows how to get your DriverManager set.
<P>
We're now ready to try and <a href="#NEXT">get a connection</a> to our specific database so we can start to run SQL statements on it!
<BR><BR><BR>
<table border=0 cellpadding=4 cellspacing=0 bgcolor=lightblue width=100%><tr><td>
<font size="4"><a name="SECTION1">Step 2 method 1) Get a connection by direct access</a></font>
</td><td align=right valign=top><font size="1"><a href="#MY_TOP">BACK TO TOP</a></font></td></tr></table>
One way to get a connection is to go directly after the MS Access database file. This can be a quick and easy way to do things, but I have seen this not work on some windows machines. Don't ask me why - I just know that it works sometimes and it doesn't others...
<P>
Here is a complete sample program getting a connection to a MS Access database on my hard drive at <B>D:\java\mdbTEST.mdb</B>. This sample includes the lines required to set the DriverManager up for ODBC data sources:
<table border=1 cellpadding=4 cellspacing=0 bgcolor=#DDDDDD>
<tr><td><pre><code>import java.sql.*;
class Test
{
 public static void main(String[] args)
 {
 try {
  Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
  <font color=red>// set this to a MS Access DB you have on your machine</font>
  String filename = "d:/java/mdbTEST.mdb";
  String database = "jdbc:odbc:Driver={Microsoft Access Driver (*.mdb)};DBQ=";
  database+= filename.trim() + ";DriverID=22;READONLY=true}"; <font color=green>// add on to the end</font>
  <font color=green>// now we can get the connection from the DriverManager</font>
  Connection con = DriverManager.getConnection( database ,"","");
 }
 catch (Exception e) {
  System.out.println("Error: " + e);
 }
 }
}</code></pre><font color=red>//save this code into a file called <B>Test.java</B> and compile it</font></td></tr></table>
<P>
Notice that this time I imported the <B>java.sql</B> package - this gives us usage of the <font size=+1><code>java.sql.Connection</code></font> object.
<P>
The line that we are interested in here is the line<font size=+1><pre><code> Connection con = DriverManager.getConnection( database ,"","");</code></pre></font>
What we are trying to do is get a <B>Connection</B> object (named <I>con</I>) to be built for us by the DriverManager. The variable <I>database</I> is the URL to the ODBC data source, and the two sets of empty quotes ("","") indicate that we are not using a username or password.
<P>
In order to have this program run successfully, you have to have an MS Access database located at <I>filename</I> location. Edit this line of code and set it to a valid MS Access database on your machine. If you do not already have an MS Access database, please jump down to <a href="#SECTION2">Set the Access Database up as an ODBC DSN</a> section, which shows how to create an empty MS Access database.
<P>
If you do have a MS Access database, and this is working correctly, then you're ready to <a href="#SECTION_SQL">Run an SQL Statement</a>!
<P>
<table border=0 cellpadding=4 cellspacing=0 bgcolor=lightblue width=100%><tr><td>
<font size="4"><a name="SECTION2">Step 2 method 2) Set up a DSN and get a connection through that</a></font>
</td><td align=right valign=top><font size="1"><a href="#MY_TOP">BACK TO TOP</a></font></td></tr></table>
Microsoft has provided a method to build a quick Jet-Engine database on your computer without the need for any specific database software (it comes standard with Windows). Using this method, we can even create a blank Microsoft Access database without having MS Access installed!
<P>
As we learned earlier, MS Access data bases can be connected to via ODBC. Instead of accessing the database directly, we can access it via a Data Source Name (DSN). Here's how to set up a DSN on your system:
<P>
<ol>
<li>Open Windows' ODBC Data Source Administrator as follows:
 <ul>
 <li>In Windows 95, 98, or NT, choose Start > Settings > Control Panel, then double-click the ODBC Data Sources icon. Depending on your system, the icon could also be called ODBC or 32bit ODBC.
 <li>In Windows 2000, choose Start > Settings > Control Panel > Administrative Tools > Data Sources.
 </ul>
<li>In the ODBC Data Source Administrator dialog box, click the System DSN tab.
<li>Click Add to add a new DSN to the list.
<li>Scroll down and select the Microsoft Access (.MDB) driver
<li>Type in the name "mdbTEST" (no quotes, but leave the cases the same) for the Data Source Name
<li>Click CREATE and select a file to save the database to (I chose "d:\java\mdbTEST.mdb") - this creates a new blank MS Access database!
<li>Click "ok" all the way out
</ol>
Now our data source is done! Here's a complete program showing how to access your new DSN data source:
<table border=1 cellpadding=4 cellspacing=0 bgcolor=#DDDDDD>
<tr><td><pre><code>import java.sql.*;
public class Test
{
 public static void main(String[] args)
 {
 <font color=red>// change this to whatever your DSN is</font>
 String dataSourceName = "mdbTEST";
 String dbURL = "jdbc:odbc:" + dataSourceName;
 try {
  Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
  Connection con = DriverManager.getConnection(dbURL, "","");
 }
 catch (Exception err) {
  System.out.println( "Error: " + err );
 }
 }
}</code></pre><font color=red>//save this code into a file called <B>Test.java</B> and compile it</font></td></tr></table>
<P>
As stated in the code, modify the variable <i>dataSourceName</i> to whatever you named your DSN in step 5 from above.
<P>
If this complies and runs successfully, it should produce no output. If you get an error, something isn't set up right - give it another shot!
<P>
Once this is working correctly, then you're ready to <a href="#SECTION_SQL">Run an SQL Statement</a>!
<P>
<table border=0 cellpadding=4 cellspacing=0 bgcolor=lightblue width=100%><tr><td>
<font size="4"><a name="SECTION_SQL">Step 3) Running a SQL Statement on your Access Database</a></font>
</td><td align=right valign=top><font size="1"><a href="#MY_TOP">BACK TO TOP</a></font></td></tr></table>
Once you have your connection, you can manipulate data within the database. In order to run a SQL query, you need to do 2 things:
<ol>
<li>Create a <B>Statement</B> from the connection you have made
<li>Get a <B>ResultSet</B> by executing a query (your insert/delete/etc. statement) on that statement
</ol>
Now lets learn how to make a <B>statement</B>, execute a query and display a the <B>ResultSet</B> from that query.
<P>
Refer to the following complete program for an understanding of these concepts (details follow):
<P>
<font size=2><I>This code assumes that you have used the <a href="#SECTION2">DSN method (Step 2 method 2)</a> to create a DSN named <B>mdbTest</B>. If you have not, you'll need to modify this code to work for a direct connection as explained in <a href="#SECTION1">Step 2 method 1</a>.</I></font>
<table border=1 cellpadding=4 cellspacing=0 bgcolor=#DDDDDD>
<tr><td><pre><code>import java.sql.*;
public class Test
{
 public static void main(String[] args)
 {
 try {
  Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
  <font color=green>/* the next 3 lines are Step 2 method 2 from above - you could use the direct
  access method (Step 2 method 1) istead if you wanted */</font>
  String dataSourceName = "mdbTEST";
  String dbURL = "jdbc:odbc:" + dataSourceName;
  Connection con = DriverManager.getConnection(dbURL, "","");
  <font color=green>// try and create a java.sql.Statement so we can run queries</font>
  Statement s = con.createStatement();
  s.execute("create table TEST12345 ( column_name integer )"); <font color=green>// create a table</font>
  s.execute("insert into TEST12345 values(1)"); <font color=green>// insert some data into the table</font>
  s.execute("select column_name from TEST12345"); <font color=green>// select the data from the table</font>
  ResultSet rs = s.getResultSet(); <font color=green>// get any ResultSet that came from our query</font>
  if (rs != null) <font color=green>// if rs == null, then there is no ResultSet to view</font>
  while ( rs.next() ) <font color=green>// this will step through our data row-by-row</font>
  {
  <font color=green>/* the next line will get the first column in our current row's ResultSet
   as a String ( <I>getString( columnNumber)</I> ) and output it to the screen */</font>
  System.out.println("Data from column_name: " + rs.getString(1) );
  }
  s.execute("drop table TEST12345");
  s.close(); <font color=green>// close the Statement to let the database know we're done with it</font>
  con.close(); <font color=green>// close the Connection to let the database know we're done with it</font>
 }
 catch (Exception err) {
  System.out.println("ERROR: " + err);
 }
 }
}</code></pre><font color=red>//save this code into a file called <B>Test.java</B> and compile it</font></td></tr></table>
<P>
If this program compiles and runs successfully, you should see some pretty boring output:
<table border=0 cellpadding=0 cellspacing=0 bgcolor=black><tr><td>
<font size=+1 color=#DDDDDD><code><pre><P>
&nbsp;&nbsp; Data from column_name: 1 &nbsp;&nbsp;
</pre></code></font>
</td></tr></table>
<P>
While that may not seem like much, let's take a quick look at what we've accomplished in the code.
<ol>
<li>First, we set the DriverManager to understand ODBC data sources.
<code><pre>
 Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
</pre></code></li>
<li>Then, we got a connection via the DSN as per <a href="#SECTION2">Step 2 method 2</a>:
<code><pre>
 String dataSourceName = "mdbTEST";
 String dbURL = "jdbc:odbc:" + dataSourceName;
 Connection con = DriverManager.getConnection(dbURL, "","");
</pre></code>
We could have used the <a href="#SECTION1">direct method</a> instead to get our connection.</li><P>
<li>Next, we created a <code><font size=+1>java.sql.Statement</font></code> Object so we could run some queries:
<code><pre>
 Statement s = con.createStatement();
</pre></code></li><P>
<li>Then came the exciting stuff - we ran some queries and made some changes!
<code><pre>
 s.execute("create table TEST12345 ( column_name integer )"); // create a table
 s.execute("insert into TEST12345 values(1)"); // insert some data into the table
 s.execute("select column_name from TEST12345"); // select the data from the table
</pre></code></li><P>
<li>The next part might be a little strange - when we ran our <B>select</B> query (see above), it produced a <code><font size=+1>java.sql.ResultSet</font></code>. A ResultSet is a Java object that contains the resulting data from the query that was run - in this case, all the data from the column <B>column_name</B> in the table <B>TEST12345</B>.
<code><pre>
 ResultSet rs = s.getResultSet(); // get any ResultSet that came from our query
 if (rs != null) // if rs == null, then there is no ResultSet to view
 while ( rs.next() ) // this will step through our data row-by-row
 {
  /* the next line will get the first column in our current row's ResultSet
  as a String ( getString( columnNumber) ) and output it to the screen */
  System.out.println("Data from column_name: " + rs.getString(1) );
 }
</pre></code></li><P>
As you can see, if the ResultSet object <B>rs</B> equals null, then we just skip by the entire <B>while</B> loop. But since we should have some data in there, we do this <I>while ( rs.next() )</I> bit.
<P>
What that means is: <i><B>while there is still data to be had in this result set, loop through this block of code and do something with the current row in the result set, then move on to the next row.</B></i>
<P>
What we're doing is looping through the result set, and for every row grabbing the first column of data and printing it to the screen. We are using the method provided in the result set called <code><font size=+1>getString(int columnNumber)</font></code> to get the data from the first column in our result set as as <B>String</B> object, and then we're just printing it out via <i>System.out.println</i>.
<P>
We know that the data in our ResultSet is of type String, since we just built the table a couple of lines before. There are other <code><font size=+1>getXXX</font></code> methods provided by ResultSet, like getInt() and getFloat(), depending on what type of data you are trying to get out of the ResultSet. Please refer to the <a href="http://java.sun.com/j2se/1.3/docs/api/java/sql/ResultSet.html" target=_new>JSDK API</a> for a full description of the ResultSet methods.</li><P>
<li>After that we just cleaned up our database by dropping (completely removing) the newly created table:
<code><pre>
 s.execute("drop table TEST12345");
</pre></code></li><P>
<li>Lastly, we need to close the Statement and Connection objects. This tells the database that we are done using them and that the database can free those resources up for someone else to use. <B>It is very important to close your connections - failure to do so can over time crash your database!</B> While this isn't too important with a MS Access database, the same rules apply for any data base (like Oracle, MS SQL, etc.)
<code><pre>
 s.close(); // close the Statement to let the database know we're done with it
 con.close(); // close the Connection to let the database know we're done with it
</pre></code></li><P>
</ol>
<P>
<h3><B>That's it!!</B> Now you know the basics for connecting to a MS Access Database via JDBC!</h3>
<BR>
<a href="#VOTE">If you found this useful, please vote for me!</a>
<BR><BR><BR><BR><BR><BR><BR>
<P>
<table border=0 cellpadding=4 cellspacing=0 bgcolor=lightblue width=100%><tr><td>
<font size="4"><a name="SECTION_LAST">What I assume you already know</a></font>
</td><td align=right valign=top><font size="1"><a href="#MY_TOP">BACK TO TOP</a></font></td></tr></table>
This document assumes that you are working on a Windows machines since we'll be connecting to a Microsoft Access database.
<P>
I assume you are familiar with database concepts. If you don't know anything about what a database is or what it is for, please take 5 minutes and read <a href="http://www.webopedia.com/TERM/d/database.html" target=_blank>this description</a> from Webopedia.
<P>
I do assume that you understand Java syntax to a degree, and that you are comfortable compiling and executing Java code. If not, please point your browser to the <a href="http://java.sun.com/docs/books/tutorial/" target=_blank>Java Tutorials</a> provided by Sun Microsystems - they'll get you started.
<BR><BR><BR>
<a name="VOTE"><B><font size="5" color="green">If you found this useful, please vote for me!</font></B></a>

