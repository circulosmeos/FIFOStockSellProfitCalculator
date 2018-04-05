#
# Calculate FIFO profits on a LibreOffice calc with 
# buys and sells in rows. 
# See FIFOStockSellProfitCalculator() comments for columns' values.
#
# by circulosmeos, 2017-04, 2017-11, 2017-12, 2018-04
# https://circulosmeos.wordpress.com/2017/04/23/fifo-profits-stock-sell-calculation-with-libreoffice-calc
# https://github.com/circulosmeos/FIFOStockSellProfitCalculator
# licensed under GPLv3
#

import sys
import socket  # only needed on win32-OOo3.0.0
import uno
from collections import deque
from decimal import *

def FIFOStockSellProfitCalculator(*args):
    # Calculate in PROFIT column the "FIFO (First In First Out) profit" after a sell.
    # Each Row represent a single stock operation.
    # The sell is identified by SELL string on TYPE column,
    # using buys (BUY string on TYPE column) to acumulate previous buys.
    # ASSET column contains the identification of the type of the asset:
    # each different asset is treated apart.
    # Other columns needed are PRICE, FEE, VOLUME and COST.
    # PROFIT column will show FIFO profit (sell - FIFO buys) and on
    # PROFIT_DESC column the ASSET id will be repeated.
    # Number of significant digits is DECIMAL_NUMBER_PRECISION, 
    # and asset volumes below ROUND_IF_BELOW are set to zero.

    # make FIFO (1), or average (0) calculation
    TYPE_OF_CALCULATION = 1;

    # array of changes in TYPE_OF_CALCULATION depending on rows:
    # pairs of the type [ row (from this number of row on, inclusive), TYPE_OF_CALCULATION]
    # Please, note that rows must be in order of magnitude in the array.
    # ( deque implements popleft() )
    # Example:
    # from row 13 on, change TYPE_OF_CALCULATION to 0, and from row 50 to the end change TYPE_OF_CALCULATION to 1:
    # changes_in_type_of_calculation = deque( [ [13, 0], [50, 1] ] )
    # Default value is no change from initial TYPE_OF_CALCULATION value:  = deque( [] )
    changes_in_type_of_calculation = deque( [] )

    # decimal numbers precision:
    DECIMAL_NUMBER_PRECISION = 6; # 6 significant digits (both integer & decimal places, depending on number size...)

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

    # .................................................

    getcontext().prec = DECIMAL_NUMBER_PRECISION

    # .................................................
    # connect to the running office
    # first trying a command line call from a living $ soffice.bin --calc --accept="socket,host=localhost,port=2002;urp;"
    # and in case of error, trying a call from inside the opened .odt file
    if LOG==1: print("connecting to LibreOffice...")
    try:
        # get the uno component context from the PyUNO runtime
        localContext = uno.getComponentContext()
        # create the UnoUrlResolver
        resolver = localContext.ServiceManager.createInstanceWithContext(
                        "com.sun.star.bridge.UnoUrlResolver", localContext )
        ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
        smgr = ctx.ServiceManager
        # get the central desktop object
        desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
        # access the current writer document
        model = desktop.getCurrentComponent()
        if LOG==1: print("successfully connected to running LibreOffice")
    except Exception:
        # connect using XSCRIPTCONTEXT
        # get the doc from the scripting context which is made available to all scripts
        desktop = XSCRIPTCONTEXT.getDesktop()
        model = desktop.getCurrentComponent()
    # .................................................

    # check whether there's already an opened document. Otherwise, die
    if not hasattr(model, "Sheets"):
        sys.exit()
    # get the XText interface
    sheet = model.Sheets.getByIndex(0)

    # calculate length of data
    # taking into account that data starts at BEGINNING_OF_DATA
    END_OF_DATA = BEGINNING_OF_DATA
    while (sheet.getCellRangeByName(TYPE + str(END_OF_DATA))).String != '' :
        END_OF_DATA+=1

    # data array with buys is a dictionary of deque's (implements popleft())
    # Buy element = [ price, quantity ]
    fifos = {}
    # accumulator for the FIFO cost of each sell
    accumulator = 0
    # (for BUYS_TO_FIAT_VOLUME): dictionary for accumulate buys value
    # converted again in sells, for every asset
    buys_to_fiat_volume = {}

    # move along TYPE column
    # accumulating BUYs and decrementing them with SELLs using FIFO
    i = BEGINNING_OF_DATA
    asset = ''
    while (i < END_OF_DATA):
        asset = sheet.getCellRangeByName(ASSET + str(i)).String
        if ( fifos.get(asset) == None ):
            fifos[asset]=deque()
        fifo=fifos[asset]
        # buy:
        # buys are accumulated one after another until a sell arrives
        if ( sheet.getCellRangeByName(TYPE + str(i)).String == BUY ):
            if LOG==1: print("buy  " + str(i))
            # Buy = [ price, quantity ]
            fifo.append( 
                [ Decimal(sheet.getCellRangeByName(PRICE  + str(i)).String.replace(DECIMAL_POINT,PYTHON_DECIMAL_POINT)), 
                  Decimal(sheet.getCellRangeByName(VOLUME + str(i)).String.replace(DECIMAL_POINT,PYTHON_DECIMAL_POINT)) 
                ] )
            # also show the assets' volume after this buy
            sheet.getCellRangeByName(PROFIT_DESC + str(i)).String = asset
            assets_remaining = Decimal('0')
            for buys in fifo:
                assets_remaining+=buys[1]
            sheet.getCellRangeByName(ASSETS_VOL + str(i)).Value = float(assets_remaining)            
        # sell:
        elif ( sheet.getCellRangeByName(TYPE + str(i)).String == SELL ):
            # first, some logs
            if LOG==1: print("sell " + str(i) + "\t" +
                "[" + sheet.getCellRangeByName(PRICE + str(i)).String + ", " + 
                    sheet.getCellRangeByName(VOLUME  + str(i)).String + "] \t" + asset
                )
            if LOG==1: print(fifo)

            quantity    = Decimal(sheet.getCellRangeByName(VOLUME + str(i)).String.replace(DECIMAL_POINT,PYTHON_DECIMAL_POINT))
            accumulator = Decimal('0')

            if ( TYPE_OF_CALCULATION == 1 ):
                # FIFO:
                # add previous buy values from first on until sell's asset quantity is exhausted
                # in order to calculate buy value and with the sell value obtain the (FIFO) benefit of this sell
                if ( len(fifo) == 0 ):
                    # if there hasn't been buys previous to a sell, the sell value is directly computed
                    accumulator = Decimal('0')
                else:
                    while ( len(fifo) > 0 ):
                        if ( fifo[0][1] >= quantity ):
                            # first buy has enough funds (quantity) for this sell operation
                            fifo[0][1] -= quantity
                            price = fifo[0][0]
                            accumulator += quantity*price
                            # ROUND_IF_BELOW:
                            if ( fifo[0][1] <= ROUND_IF_BELOW ):
                                fifo.popleft()
                            break
                        else:
                            # first buy has not enough funds (quantity) for this operation
                            # so it will have to drain next buys later in the while loop
                            quantity -= fifo[0][1]
                            accumulator += fifo[0][1]*fifo[0][0]
                            fifo.popleft()
            elif ( TYPE_OF_CALCULATION == 0 ):
                # average:
                # add previous buy values from first to last to calculate an average to which single value fifo[][] will be set
                # in order to calculate buy value and with the sell value obtain the (average) benefit of this sell
                if ( len(fifo) == 0 ):
                    # if there hasn't been buys previous to a sell, the sell value is directly computed
                    accumulator = Decimal('0')
                else:
                    # sum up all previous buys
                    total_quantity = Decimal('0')
                    total_accumulator = Decimal('0')
                    while ( len(fifo) > 0 ):
                        total_quantity += fifo[0][1]
                        total_accumulator += fifo[0][1]*fifo[0][0]
                        fifo.popleft()
                    # substitute all previous buys with a single operation with all funds valued at medium price
                    # if there's a remaining quantity of assets (if not, assets' array is left empty)
                    if ( total_quantity - quantity > 0 ):
                        fifo.append(
                            [ Decimal(total_accumulator / total_quantity),
                              Decimal(total_quantity - quantity)
                            ] )
                    # now calculate "buy value" with medium costs:
                    accumulator = quantity * ( total_accumulator / total_quantity )

            # sell "buy value" has been calculated in accumulator variable
            # let's write the profit:
            sheet.getCellRangeByName(PROFIT + str(i)).Value = float ( 
                Decimal(sheet.getCellRangeByName(VOLUME + str(i)).String.replace(DECIMAL_POINT,PYTHON_DECIMAL_POINT)) * 
                Decimal(sheet.getCellRangeByName(PRICE  + str(i)).String.replace(DECIMAL_POINT,PYTHON_DECIMAL_POINT)) - 
                accumulator
                )
            sheet.getCellRangeByName(PROFIT_DESC + str(i)).String = asset

            # also show the remaining assets' volume
            assets_remaining = Decimal('0')
            for buys in fifo:
                assets_remaining+=buys[1]
            sheet.getCellRangeByName(ASSETS_VOL + str(i)).Value = float(assets_remaining)
            if LOG==1: print(fifo)
            if LOG==1: print(PROFIT + str(i) + "=\t" + sheet.getCellRangeByName(PROFIT + str(i)).String)

            # show the volume of fiat currency (or origin asset) invested in the buy(s) now sold
            sheet.getCellRangeByName(BUY_TO_FIAT + str(i)).Value = float(accumulator)
            if LOG==1: print(BUY_TO_FIAT + str(i) + "=\t" + sheet.getCellRangeByName(BUY_TO_FIAT + str(i)).String)

            # (for BUYS_TO_FIAT_VOLUME): Accumulate buy(s) value spent in this asset sell
            if ( buys_to_fiat_volume.get(asset) == None ):
                buys_to_fiat_volume[asset] = accumulator
            else:
                buys_to_fiat_volume[asset] += accumulator

        # next row
        i+=1

        # check changes_in_type_of_calculation array:
        if ( len(changes_in_type_of_calculation) > 0 ):
            if ( changes_in_type_of_calculation[0][0] == i ):
                TYPE_OF_CALCULATION = changes_in_type_of_calculation[0][1]
                changes_in_type_of_calculation.popleft()
                if LOG==1: print("\nTYPE_OF_CALCULATION = " + str(TYPE_OF_CALCULATION) + "\n")

    # write sum of gross profits
    sheet.getCellRangeByName(PROFIT + str(END_OF_DATA)).Formula = \
        '=SUM(' + str(PROFIT) + str(BEGINNING_OF_DATA) + ':' + str(PROFIT) + str(END_OF_DATA-1) + ')'
    sheet.getCellByPosition(ord(PROFIT)-ord('A')+1, END_OF_DATA-1).String = 'Gross profit'
    
    # write sum of fees
    sheet.getCellRangeByName(FEE    + str(END_OF_DATA)).Formula = \
        '=SUM(' + str(FEE) + str(BEGINNING_OF_DATA) + ':' + str(FEE) + str(END_OF_DATA-1) + ')'
    sheet.getCellByPosition(ord(FEE)-ord('A')+1, END_OF_DATA-1).String = 'Fees'

    # write sum of net profits
    sheet.getCellRangeByName(PROFIT + str(END_OF_DATA+1)).Formula = \
        '=' + str(PROFIT) + str(END_OF_DATA) + '-' + str(FEE) + str(END_OF_DATA)
    sheet.getCellByPosition(ord(PROFIT)-ord('A')+1, END_OF_DATA-1+1).String = 'Net profit'

    # (for BUYS_TO_FIAT_VOLUME):
    # write resume of the volume of fiat currency (or origin asset) accumulated
    # in buys of assets sold again to fiat currency (or origin asset)
    i+=4 # row with this results after every row with data
    sheet.getCellRangeByName(BUYS_TO_FIAT_VOLUME + str(i)).String = 'Accumulated value of buys converted to sells'
    for asset, accumulator in sorted( buys_to_fiat_volume.items() ):
        i+=1
        sheet.getCellByPosition(ord(BUYS_TO_FIAT_VOLUME)-ord('A'), i-1).String = asset
        sheet.getCellByPosition(ord(BUYS_TO_FIAT_VOLUME)-ord('A')+1, i-1).Value = float(accumulator)

    return None

# direct run as script, without function() entry point indication:
if __name__ == "__main__":
    FIFOStockSellProfitCalculator()
