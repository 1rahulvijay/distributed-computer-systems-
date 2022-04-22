from vb2py.vbfunctions import *
from vb2py.vbdebug import *



def insertTransCosts():
    wkb = Workbook()

    taaWorkbook = String()

    strSQL = String()

    oRS = ADODB.Recordset()

    oConn = ADODB.Connection()

    ccy = String()

    assetDef = String()

    profile = String()

    tmp = String()

    doInsert = Boolean()

    TAA = String()

    assetCat = String()

    dateOf = String()

    longname = String()
    taaWorkbook = Range('D6')
    Workbooks(taaWorkbook).Activate()
    if ( Len(taaWorkbook) == 0 ) :
        MsgBox('In order to proceed open the current TAA file ')
        sys.exit(0)
    oConn = ADODB.Connection()
    with_variable0 = oConn
    with_variable0.ConnectionString = '\\\\CH.AD.HEDANI.NET\\GROUPS_CH$\\SJMO14\\02_Projects\\BO Reporting Drafts\\Contribution Breakdown Report\\Transaction Costs\\pa_report.mdb'
    #*********************************
    #update 26.04.2019 Location update
    #*********************************
    #old    '"\\chca6024.eur.beluni.net\groups$\SJMO14\02_Projects\BO Reporting Drafts\Contribution Breakdown Report\Transaction Costs\pa_report.mdb"
    #.ConnectionString = "Data Source=\\Eur.beluni.net\cs-ch-groups$\APSA\zOzan\VVA\"
    #.Provider = "Microsoft Jet 4.0 OLE DB Provider"
    with_variable0.Provider = 'Microsoft.ACE.OLEDB.12.0'
    with_variable0.Open()
    oRS = ADODB.Recordset()
    dateOf = Application.InputBox(prompt= 'Please type in the \'dateof\' (mm/dd/yyyy)', Type= 2)
    currencies = Array('CHF', 'CHFI', 'CHF Focus', 'EUR', 'USD', 'USDE', 'GBP')
    HedgeCCYs = Array('CHF', 'EUR', 'USD', 'GBP', 'JPY', 'CAD')
    profiles = Array('F', 'I', 'B', 'G', 'E')
    #Schelife über Worksheets
    for j in vbForRange(0, 6):
        if ( j == 0 ) :
            Sheets('PMACS_CHF').Select()
            ccy = 'CHF'
        elif  ( j == 1 ) :
            Sheets('PMACS_CHFI').Select()
            ccy = 'CHFI'
        elif  ( j == 2 ) :
            Sheets('PMACS_CHF_Focus').Select()
            ccy = 'CHFF'
        elif  ( j == 3 ) :
            Sheets('PMACS_EUR').Select()
            ccy = 'EUR'
        elif  ( j == 4 ) :
            Sheets('PMACS_USD').Select()
            ccy = 'USD'
        elif  ( j == 5 ) :
            Sheets('PMACS_USDE').Select()
            ccy = 'USDE'
        elif  ( j == 6 ) :
            Sheets('PMACS_GBP').Select()
            ccy = 'GBP'
        #Schleife über Assets
        r = 0
        i = 0
        while not Cells(3 + i, 2) == '':
            taaRow = 1
            oRS = ADODB.Recordset()
            assetCat = ''
            assetDef = ''
            #+ var declaration + trim function added 24.10.2018
            excelID = Cells(3 + i, 2)
            excelID = Trim(excelID)
            ExcelAssetCat = Trim(ExcelAssetCat)
            Debug.Print(excelID)
            for k in vbForRange(0, 4):
                r = 1
                while not Cells(2, taaRow + r) == 'Current':
                    r = r + 1
                Debug.Print(excelID)
                Debug.Print(r)
                r = r + 1
                Debug.Print(excelID)
                taaRow = taaRow + r
                profile = profiles(k)
                Debug.Print(excelID)
                if Abs(Cells(3 + i, taaRow).Value) > 0:
                    Spread = 0
                    Stamp = 0
                    FXSpread = 0
                    strSQL = 'select Spread, Stamp, FXSpread from TransactionscostMapping where ExcelID like "' + excelID + '" and CCY like "' + ccy + '" ;'
                    Debug.Print(strSQL)
                    oRS.Open(strSQL, oConn, adOpenKeyset, adLockOptimistic)
                    Spread = oRS.Fields(0).Value
                    Stamp = oRS.Fields(1).Value
                    FXSpread = oRS.Fields(2).Value
                    oRS.Close()
                    TAAweight = Cells(3 + i, taaRow).Value
                    TAAweight = Abs(TAAweight)
                    strSQL = 'insert into TransactionsCosts (Profile, ccy, ' + ' dateOf, ExcelAssetCat,DeltaWeight, TransactionsCost) values ("' + profile + '", "' + ccy + '", #' + dateOf + '#' + ', "' + excelID + '", ' + TAAweight + ',' + TAAweight *  ( Spread + Stamp + FXSpread )  + ')'
                    Debug.Print(strSQL)
                    oRS = ADODB.Recordset()
                    oRS.Open(strSQL, oConn, adOpenKeyset, adLockOptimistic)
            i = i + 1
    Workbooks('uploadWeightsNeu.xls').Activate()
    MsgBox('Done')

def insertTAABM():
    wkb = Workbook()

    taaWorkbook = String()

    strSQL = String()

    oRS = ADODB.Recordset()

    oConn = ADODB.Connection()

    ccy = String()

    assetDef = String()

    profile = String()

    tmp = String()

    doInsert = Boolean()

    TAA = String()

    assetCat = String()

    dateOf = String()

    longname = String()
    taaWorkbook = Range('D6')
    Workbooks(taaWorkbook).Activate()
    if ( Len(taaWorkbook) == 0 ) :
        MsgBox('In order to proceed open the current TAA file ')
        sys.exit(0)
    oConn = ADODB.Connection()
    with_variable1 = oConn
    with_variable1.ConnectionString = 'Data Source=\\\\Eur.beluni.net\\cs-ch-groups$\\APSA\\zOzan\\VVA\\pa_report.mdb'
    with_variable1.Provider = 'Microsoft Jet 4.0 OLE DB Provider'
    with_variable1.Open()
    oRS = ADODB.Recordset()
    dateOf = Application.InputBox(prompt= 'Please type in the \'dateof\' (mm/dd/yy)', Type= 2)
    currencies = Array('CHF', 'CHFI', 'CHF Focus', 'EUR', 'USD', 'USDE', 'GBP')
    HedgeCCYs = Array('CHF', 'EUR', 'USD', 'GBP', 'JPY', 'CAD')
    profiles = Array('F', 'I', 'B', 'G', 'E')
    for j in vbForRange(0, 6):
        if ( j == 0 ) :
            Sheets('PMACS_CHF').Select()
            ccy = 'CHF'
        elif  ( j == 1 ) :
            Sheets('PMACS_CHFI').Select()
            ccy = 'CHFI'
        elif  ( j == 2 ) :
            Sheets('PMACS_CHF_Focus').Select()
            ccy = 'CHFF'
        elif  ( j == 3 ) :
            Sheets('PMACS_EUR').Select()
            ccy = 'EUR'
        elif  ( j == 4 ) :
            Sheets('PMACS_USD').Select()
            ccy = 'USD'
        elif  ( j == 5 ) :
            Sheets('PMACS_USDE').Select()
            ccy = 'USDE'
        elif  ( j == 6 ) :
            Sheets('PMACS_GBP').Select()
            ccy = 'GBP'
        r = 0
        i = 0
        while not Cells(3 + i, 2) == '':
            taaRow = 1
            oRS = ADODB.Recordset()
            indexID = ''
            assetCat = ''
            assetDef = ''
            #trim function added 24.10.2018
            excelID = Trim(excelID)
            excelID = Cells(3 + i, 2)
            strSQL = 'select IndexID from ExcelIDIndexIDMatch where ExcelID like "' + excelID + '" and Profile like "' + ccy + '" and startDate <= #' + dateOf + '# and endDate >= #' + dateOf + '#;'
            Debug.Print(strSQL)
            oRS.Open(strSQL, oConn, adOpenKeyset, adLockOptimistic)
            indexID = oRS.Fields(0).Value
            oRS.Close()
            if not indexID == '':
                strSQL = 'select AssetCat,AssetLocalCCY from IndexDefinitions where IndexID= ' + indexID + ';'
                Debug.Print(strSQL)
                oRS.Open(strSQL, oConn, adOpenKeyset, adLockOptimistic)
                assetCat = oRS.Fields(0).Value
                assetDef = oRS.Fields(1).Value
                oRS.Close()
            for k in vbForRange(0, 4):
                r = 1
                while not Cells(2, taaRow + r) == 'TAA':
                    r = r + 1
                taaRow = taaRow + r
                profile = profiles(k)
                if Abs(Cells(3 + i, taaRow).Value) > 0:
                    if indexID == '':
                        MsgBox('no indexid given for ' + excelID + ' in profile ' + profile)
                        sys.exit(0)
                    TAAweight = Cells(3 + i, taaRow).Value
                    TAAweight = TAAweight
                    strSQL = 'insert into WEIGHTS (PortfolioClass, Profile, AssetLocalCCY, ccy, ' + ' assetcat, weight,indexid, dateOf, isGrownWeight) values ("TAA", "' + profile + '", "' + assetDef + '", "' + ccy + '", "' + assetCat + '", ' + TAAweight + ',' + indexID + ', #' + dateOf + '#' + ', 0)'
                    Debug.Print(strSQL)
                    oRS = ADODB.Recordset()
                    oRS.Open(strSQL, oConn, adOpenKeyset, adLockOptimistic)
            i = i + 1
        #Hedges
        #HedgeCCYs = Array("CHF", "EUR", "USD", "GBP", "JPY", "CAD")
        h = 0
        i = 61
        if not Cells(i, 2).Value == 'CHF - Hedge':
            MsgBox('check Hedge Zeile - anscheinend nicht bei 61')
            sys.exit(0)
        while not Cells(i, 2) == 'AUD - Hedge':
            assetDef = ''
            if Cells(i, 2) == 'CHF - Hedge':
                assetDef = 'CHF'
            if Cells(i, 2) == 'EUR - Hedge':
                assetDef = 'EUR'
            if Cells(i, 2) == 'USD - Hedge':
                assetDef = 'USD'
            if Cells(i, 2) == 'GBP - Hedge':
                assetDef = 'GBP'
            if Cells(i, 2) == 'JPY - Hedge':
                assetDef = 'JPY'
            if Cells(i, 2) == 'CAD - Hedge':
                assetDef = 'CAD'
            if assetDef == '':
                Cells(i, 2).Select()
                MsgBox('Hedgewährung nicht bekannt')
                sys.exit(0)
            assetCat = 'HEDGE'
            taaRow = 1
            for k in vbForRange(0, 4):
                r = 1
                while not Cells(2, taaRow + r) == 'TAA':
                    r = r + 1
                taaRow = taaRow + r
                profile = profiles(k)
                if Abs(Cells(i, taaRow).Value) > 0:
                    TAAweight = Cells(i, taaRow).Value
                    TAAweight = TAAweight * 1
                    strSQL = 'insert into WEIGHTS (PortfolioClass,  profile, assetLocalccy, ccy, ' + ' assetcat, weight, indexid,dateOf,isGrownWeight) values ("TAA",  "' + profile + '", "' + assetDef + '", "' + ccy + '", "' + assetCat + '", ' + TAAweight + ',0, #' + dateOf + '#' + ',0)'
                    Debug.Print(strSQL)
                    oRS = ADODB.Recordset()
                    oRS.Open(strSQL, oConn, adOpenKeyset, adLockOptimistic)
            i = i + 3
            h = h + 1
    Workbooks('uploadWeightsNeu.xls').Activate()
    MsgBox('Done')

def UploadBM():
    wkb = Workbook()

    taaWorkbook = String()

    strSQL = String()

    oRS = ADODB.Recordset()

    oConn = ADODB.Connection()

    ccy = String()

    assetDef = String()

    profile = String()

    tmp = String()

    doInsert = Boolean()

    TAA = String()

    assetCat = String()

    dateOf = String()

    longname = String()
    taaWorkbook = Range('D6')
    Workbooks(taaWorkbook).Activate()
    if ( Len(taaWorkbook) == 0 ) :
        MsgBox('In order to proceed open the current TAA file ')
        sys.exit(0)
    oConn = ADODB.Connection()
    with_variable2 = oConn
    with_variable2.ConnectionString = 'Data Source=X:\\APSA\\zOzan\\VVA\\pa_report.mdb'
    with_variable2.Provider = 'Microsoft Jet 4.0 OLE DB Provider'
    with_variable2.Open()
    oRS = ADODB.Recordset()
    dateOf = Application.InputBox(prompt= 'Please type in the \'dateof\' (mm/dd/yy)', Type= 2)
    currencies = Array('CHF', 'CHFI', 'CHF Focus', 'EUR', 'USD', 'USDE', 'GBP')
    HedgeCCYs = Array('CHF', 'EUR', 'USD', 'GBP', 'JPY', 'CAD')
    profiles = Array('F', 'I', 'B', 'G', 'E')
    for j in vbForRange(0, 6):
        if ( j == 0 ) :
            Sheets('PMACS_CHF').Select()
            ccy = 'CHF'
        elif  ( j == 1 ) :
            Sheets('PMACS_CHFI').Select()
            ccy = 'CHFI'
        elif  ( j == 2 ) :
            Sheets('PMACS_CHF_Focus').Select()
            ccy = 'CHFF'
        elif  ( j == 3 ) :
            Sheets('PMACS_EUR').Select()
            ccy = 'EUR'
        elif  ( j == 4 ) :
            Sheets('PMACS_USD').Select()
            ccy = 'USD'
        elif  ( j == 5 ) :
            Sheets('PMACS_USDE').Select()
            ccy = 'USDE'
        elif  ( j == 6 ) :
            Sheets('PMACS_GBP').Select()
            ccy = 'GBP'
        r = 0
        i = 0
        while not Cells(3 + i, 1) == '':
            taaRow = 1
            oRS = ADODB.Recordset()
            indexID = ''
            longname = ''
            assetCat = ''
            assetDef = ''
            excelID = Cells(3 + i, 2)
            strSQL = 'select IndexID from ExcelIDIndexIDMatch where ExcelID like "' + excelID + '" and Profile like "' + ccy + '" and startDate <= #' + dateOf + '# and endDate >= #' + dateOf + '#;'
            Debug.Print(strSQL)
            oRS.Open(strSQL, oConn, adOpenKeyset, adLockOptimistic)
            indexID = oRS.Fields(0).Value
            oRS.Close()
            if not indexID == '':
                strSQL = 'select index_names,Asset_Cat,Asset_Def from tbl_IndexNames_Definitions where IndexID= ' + indexID + ' and dateof <= #' + dateOf + '# order by dateof desc;'
                Debug.Print(strSQL)
                oRS.Open(strSQL, oConn, adOpenKeyset, adLockOptimistic)
                longname = oRS.Fields(0).Value
                assetCat = oRS.Fields(1).Value
                assetDef = oRS.Fields(2).Value
                oRS.Close()
            for k in vbForRange(0, 4):
                r = 1
                while not Cells(2, taaRow + r) == 'BMK':
                    r = r + 1
                taaRow = taaRow + r
                profile = profiles(k)
                if Abs(Cells(3 + i, taaRow).Value) > 0:
                    if indexID == '':
                        MsgBox('no indexid given for ' + excelID + ' in profile ' + profile)
                        sys.exit(0)
                    TAAweight = Cells(3 + i, taaRow).Value
                    TAAweight = TAAweight * 100
                    strSQL = 'insert into WEIGHTS_TABLE (weight_type, contract_type,  profile, asset_def_ccy, ccy, ' + ' asset_cat, weight,index_id, dateOf, index_name) values ("BM", "CACF",  "' + profile + '", "' + assetDef + '", "' + ccy + '", "' + assetCat + '", ' + TAAweight + ',' + indexID + ', #' + dateOf + '#' + ', "' + longname + '")'
                    Debug.Print(strSQL)
                    oRS = ADODB.Recordset()
                    oRS.Open(strSQL, oConn, adOpenKeyset, adLockOptimistic)
            i = i + 1
    Workbooks('uploadWeightsNeu.xls').Activate()
    MsgBox('Done')

