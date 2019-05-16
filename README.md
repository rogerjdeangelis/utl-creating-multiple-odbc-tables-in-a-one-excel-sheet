# utl-creating-multiple-odbc-tables-in-a-one-excel-sheet
Creating multiple odbc tables in a one excel sheet
    Creating multiple odbc tables in a one excel sheet

    By ODBC tables a mean named-ranges

    github
    https://tinyurl.com/y3yref49
    https://github.com/rogerjdeangelis/utl-creating-multiple-odbc-tables-in-a-one-excel-sheet

    The code below creates two tables, named ranges, in one sheet.
    It works with an existing workbook or will dynamically create a workbook.

    The tables in excel can be directly accessed using passthru, proc sql or a datastep.
    You can even do a passtru join.

    ODS excel can stack reports and graphs but cannot create named ranges at arbitrary
    locations on the same sheet.

    You may need classic SAS for this not EG server?

    SAS FDorum
    https://tinyurl.com/y4p3eu7m
    https://communities.sas.com/t5/SAS-Enterprise-Guide/Using-the-ODS-Excel-Destination-how-to-use-multiple-tables-in-a/m-p/559354

    *_                   _
    (_)_ __  _ __  _   _| |_
    | | '_ \| '_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    ;

    options validvarname=upcase;
    libname sd1 "d:/sd1";
    data sd1.males sd1.females;
      set sashelp.class(obs=6);
       keep name sex age;
      if sex="M" then output sd1.males;
      else output sd1.females;
    run;quit;

    SD1.FEMALES total obs=3

    Obs     NAME      SEX    AGE

     1     Alice       F      13
     2     Barbara     F      13
     3     Carol       F      14

    SD1.MALES total obs=3

    Obs     NAME     SEX    AGE

     1     Alfred     M      14
     2     Henry      M      14
     3     James      M      12



     WANT excel sheet with males starting at A3 and females at G3
     ==============================================================

     d:/xls/gender.xlsx

         +---------------------+---------------------------------+-------+
         |  A  |  B    |  C    |  D  |  E  |  F    |  G    |  H  |  D    |
         +---------------------+---------------------------------+-------+
     1   |     |       |       |     |     |       |       |     |       |
         |-----+-------+-------|-----+-----+-------+-------+-----|-------|
     2   |     |       |       |     |     |       |       |     |       |
         |-----+-------+-------+-----+-----+-------+-------+-----+-------+
     3   |     |NAME   |SEX    |AGE  |     |       |NAME   |SEX  |AGE    |
         |-----+-------+-------|-----+-----+-------+-------+-----|-------|
     4   |     |Alfred |M      |13   |     |       |Alice  |F    |14     |
         |-----+-------+-------+-----+-----+-------+-------+-----+-------+
     5   |     |Alex   |M      |13   |     |       |Barbara|F    |14     |
         |-----+-------+-------+-----+-----+-------+-------+-----+-------+
     6   |     |JAMES  |M      |14   |     |       |Carol  |F    |12     |
         -----------------------------------------------------------------
     ...

     [GENDER]


     Formulas->Name Manger

     females  =gender!$B$3:$D$6
     males    =gender!$G$3:$I$6



    *          _       _   _
     ___  ___ | |_   _| |_(_) ___  _ __
    / __|/ _ \| | | | | __| |/ _ \| '_ \
    \__ \ (_) | | |_| | |_| | (_) | | | |
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|

    ;

    %utlfkil(d:/xls/gender.xlsx); * delete if exist - it works with an existing workbook;

    %utl_submit_r64('
    library(XLConnect);
    library(haven);

    females<-read_sas("d:/sd1/females.sas7bdat");
    wb <- loadWorkbook("d:/xls/gender.xlsx", create = TRUE);
    createSheet(wb, name = "gender");
    createName(wb, name = "females", formula = "gender!$B$3");
    writeNamedRegion(wb, females, name = "females");

    males<-read_sas("d:/sd1/males.sas7bdat");
    createName(wb, name = "males", formula = "gender!$G$3");
    writeNamedRegion(wb, males, name = "males");
    saveWorkbook(wb);
    ');


    *_
    | | ___   __ _
    | |/ _ \ / _` |
    | | (_) | (_| |
    |_|\___/ \__, |
             |___/
    ;

    Stderr output:
    Loading required package: XLConnectJars
    XLConnect 0.2-15 by Mirai Solutions GmbH [aut],
      Martin Studer [cre],
      The Apache Software Foundation [ctb, cph] (Apache POI),
      Graph Builder [ctb, cph] (Curvesapi Java library)
    http://www.mirai-solutions.com
    https://github.com/miraisolutions/xlconnect

