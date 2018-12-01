**FIFOStockSellProfitCalculator.py** calculates FIFO profits on a LibreOffice calc document with buys and sells in rows.    
    
PROFIT column will contain the "FIFO (First In First Out) profit" after a sell.   
Also the mean profit (see *TYPE_OF_CALCULATION*), or even FIFO and mean row blocks one after another can be calculated (see *changes_in_type_of_calculation* array).   
   
Each Row represents a single stock operation.   
The sell is identified by SELL string on TYPE column, using buys (BUY string on TYPE column) to acumulate previous buys.   
ASSET column contains the identification of the type of the asset: each different asset is treated apart.   
Other columns needed are PRICE, FEE, VOLUME and COST.   
PROFIT column will show FIFO profit (sell - FIFO buys) and on PROFIT\_DESC column the ASSET id will be repeated.   
Number of significant digits is DECIMAL\_NUMBER\_PRECISION, and asset volumes below ROUND\_IF\_BELOW are set to zero.   
See **Configuration** below.

See [this page for a step by step guide](https://circulosmeos.wordpress.com/2017/04/23/fifo-profits-stock-sell-calculation-with-libreoffice-calc).

Columns T, U and V and the three accumulated values have been written by the script on the sample file *sample.csv* (C1:J21):

![](https://circulosmeos.files.wordpress.com/2017/04/calc-stock-ops-example-after-script-exec.png?w=680)

**Tested at least under LibreOffice versions 5 and 6.**


Download and use with your own data
===================================

You can download [sample.withScript.ods](https://github.com/circulosmeos/FIFOStockSellProfitCalculator/blob/master/sample.withScript.ods) and paste your data there with as many rows as needed, respecting each column data type.   
Then just do `Tools / Macros / Run macro... `, unfold the file name from the new window, and select `FIFOStockSellProfitCalculator` there - then click `Run`.

If your regional configuration uses '**,**' as decimal point separator instead of '**.**', please, download instead [sample.withScriptCommaSep.ods](https://github.com/circulosmeos/FIFOStockSellProfitCalculator/blob/master/sample.withScriptCommaSep.ods). This script contains the same script changing the proper configuration line: `DECIMAL_POINT = ','`.   
Note that you will know that you're running into this *regional comma vs point separator* if an error like this appears : 
    
    com.sun.star.uno.RuntimeException: Error during invoking function FIFOStockSellProfitCalculator in module vnd.sun.star.tdoc:/3/Scripts/python/FIFOStockSellProfitCalculator.py (<class 'decimal.InvalidOperation'>: [<class 'decimal.ConversionSyntax'>]
      File "C:\LibreOffice 6\program\pythonscript.py", line 879, in invoke
    ret = self.func( *args )
      File "vnd.sun.star.tdoc:/3/Scripts/python/FIFOStockSellProfitCalculator.py", line 146, in FIFOStockSellProfitCalculator
    )


Manually embed the script into a Calc document
==============================================

The script `FIFOStockSellProfitCalculator.py` can be manually embedded into any LibreOffice Calc doc (.ods): [follow these instructions to do that](https://github.com/circulosmeos/LibreOfficeScriptInsert).


Installation
============

Copy the file **FIFOStockSellProfitCalculator.py** to your LibreOffice python scripts directory in order to have it available for any of your projects.   
In Windows this would tipically be:

    C:\Program Files (x86)\LibreOffice 6\share\Scripts\python\

Then, it can be run on the current Calc document from Tools / Macros... 


Running as commandline script
=============================

If the LOG script variable is changed to "LOG = 1", **useful internal operations are shown on stdout** which allows you to follow the operations as they have been processed.   
In order to run the script from command line, LibreOffice must be executed with special parameters:

    $ ./soffice.bin --calc \
    --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"

Or in Windows:

    C:\Program Files (x86)\LibreOffice 6\program> soffice.exe --calc --accept="socket,host=localhost,port=2002;urp;"

Now, open the desired Calc document with the stocks operations.
Then, the *python* executable from LibreOffice installation must be used, so change your pwd to that path and append to it the complete path to your `FIFOStockSellProfitCalculator.py` script:

    C:\Program Files (x86)\LibreOffice 6\program> python.exe C:\FIFOStockSellProfitCalculator.py


Configuration
=============

Modify these script's variable values as convenient:

    # make FIFO (1), or average (0) calculation
    TYPE_OF_CALCULATION = 1;

    # array of changes in TYPE_OF_CALCULATION depending on rows:
    # pairs of the type [ row (from this number of row on, inclusive), TYPE_OF_CALCULATION]
    # Please, note that rows must be in order of magnitude in the array.
    # Example:
    # from row 13 on, change TYPE_OF_CALCULATION to 0, and from row 50 to the end change TYPE_OF_CALCULATION to 1:
    # changes_in_type_of_calculation = deque( [ [13, 0], [50, 1] ] )
    # Default value is no change from initial TYPE_OF_CALCULATION value:  = deque( [] )
    changes_in_type_of_calculation = deque( [] )

    # decimal numbers precision:
    DECIMAL_NUMBER_PRECISION = 6; # 6 significant digits (both integer & decimal places, depending of number size...)

    # regional comma separator:
    # DECIMAL_POINT is substitued by PYTHON_DECIMAL_POINT, just in case a regional configuration
    # be different from the standard decimal point '.'
    DECIMAL_POINT = '.' # change this to your regional configuration
    PYTHON_DECIMAL_POINT = '.'

    # column and value strings definitions:
    BEGINNING_OF_DATA = 2   # first row with data
    ASSET       = 'C'       # Unique asset id (i.e. USD, EUR, STCKXXXX, etc)
    TYPE        = 'E'       # SELL | BUY | other
    SELL        = 'sell'    # sell string identifier on TYPE column
    BUY         = 'buy'     # buy  string identifier on TYPE column
    PRICE       = 'G'
    FEE         = 'I'
    VOLUME      = 'J'
    PROFIT      = 'T'       # column of results
    PROFIT_DESC = 'U'       # here the ASSET id will be repeated
    ASSETS_VOL  = 'V'       # volume of assets of type ASSET after each buy/sell op
    BUY_TO_FIAT = 'W'       # volume of fiat currency (or origin asset) invested in the buy(s) now sold
    BUYS_TO_FIAT_VOLUME = 'A'   # volume of fiat currency (or origin asset) accumulated
                                # in buys of assets sold again to fiat currency (or origin asset)
                                # This column will be used in a row after the end of rows of data,
                                # in which a resume of all origin/dest assets will presented.
                                # If fiat currency (or origin asset) is the same in all cases,
                                # the rows can be added in a global result, but this is not done
                                # as a preventive measure to avoid errors of origin fiat/asset judgements.

    # round to zero if this quantity of assets is left on some sell:
    # NOTE: do no set ROUND_IF_BELOW to values below (previously set) DECIMAL_NUMBER_PRECISION 
    ROUND_IF_BELOW = Decimal('5e-06')

    # print logs to stdout
    LOG = 0


Examples
========

* [sample.withScript.ods](https://github.com/circulosmeos/FIFOStockSellProfitCalculator/blob/master/sample.withScript.ods), LibreOffice file with embedded script

* [sample.withScriptCommaSep.ods](https://github.com/circulosmeos/FIFOStockSellProfitCalculator/blob/master/sample.withScriptCommaSep.ods), LibreOffice file with embedded script for regions that use '**,**' as decimal point separator

* [sample.csv](https://github.com/circulosmeos/FIFOStockSellProfitCalculator/blob/master/sample.csv), plain text data file with the example rows


License
=======

Licensed as [GPL v3](http://www.gnu.org/licenses/gpl-3.0.en.html) or higher.   
