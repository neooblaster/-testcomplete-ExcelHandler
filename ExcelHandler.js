const TableJs = require('TableJs');

/**
 * Version v0.3.3
 *
 * @author: Nicolas DUPRE (VISEO)
 *
 * Excel Handler for better usability & readability in scripts.
 *
 * @param {String}  $sPath        path to the Excel File.
 * @param {Boolean} $bHasHeaders  Indicates if the Excel file has header.
 *                                Implies :
 *                                  - Data start at line 2
 *                                  - Auto-generates Cols definition.
 *
 * @return {ExcelHandler} Current Instance
 *
 * @constructor
 */
function ExcelHandler (
    $sPath,
    $bHasHeaders = true
) {
    let self = this;

    self._sFilePath      = null;
    self._oExcel         = null;
    self._oExcelSheet    = null;
    self._sActiveSheet   = null;
    self._bHasHeader     = $bHasHeaders;
    self._nRowStartAt    = ($bHasHeaders) ? 2 : 1;
    self._oCols          = {};
    self._aColNames      = {};
    self._oSheetRead     = {};
    self._oTables        = {};
    self._oIndexes       = {};
    self._inDecimalSep   = '.';
    self._inThousandSep  = ',';
    self._outDecimalSep  = ',';
    self._outThousandSep = ' ';

    /**
     * Open a handler on the Excel file with provided path during instantiation.
     *
     * @return {ExcelHandler}
     */
    self.open = function () {
        if (!self._sFilePath) {
            Log.error('Excel data file not specified.');
        }
        self._oExcel = Excel.Open(self._sFilePath);

        return self;
    };

    /**
     * Set handling on the Excel Sheet using it title.
     *
     * @param {String} $sExcelSheet  Title of the Sheet.
     *
     * @return {ExcelHandler}
     */
    self.sheet = function ($sExcelSheet) {
        try {
            self._oExcelSheet = self._oExcel.SheetByTitle($sExcelSheet);
            self._sActiveSheet = $sExcelSheet;

            // If the Sheet not currently read. Analyze Columns
            if (!self._oSheetRead[$sExcelSheet]) {
                let nColNumber = self._oExcelSheet.ColumnCount;
                let oCols = {};

                for (let i = 0; i < nColNumber; i++) {
                    // Generates Methods with the column name
                    let sColumnName = self.getColName(i + 1);
                    oCols[sColumnName] = sColumnName;

                    if(!self._aColNames[$sExcelSheet]) self._aColNames[$sExcelSheet] = {};

                    // At least, Method is the name of the column
                    self._aColNames[$sExcelSheet][sColumnName] = sColumnName;

                    // If file has headers, first line stand for another column name
                    if (self._bHasHeader) {
                        let sHeaderName = self._oExcelSheet.Cell(sColumnName, 1).Value;

                        // Generates Methods with column header value
                        if (sHeaderName) {
                            if (/^[a-zA-Z]/.test(sHeaderName)) {
                                sHeaderName = sHeaderName.replace(/\s/g, '');
                                oCols[sHeaderName] = sColumnName;

                                // Overwrite ColumnName by HeaderName
                                self._aColNames[$sExcelSheet][sColumnName] = sHeaderName;
                            }
                        }
                    }
                }

                // If the developper already set Column Definition,
                // Do not overwrite it definition by our default generation
                if (self._oCols[$sExcelSheet]) {
                    self._oCols[$sExcelSheet] = Object.assign(
                        oCols,                                          // Generated
                        self._oCols[$sExcelSheet]     // Existing definition
                    );
                }
                // If does not exist, our generated definition can be use as default
                else {
                    self._oCols[$sExcelSheet] = oCols;
                }
            }

            // Makes Method using Sheet Column Definition.
            self.cols(self._oCols[$sExcelSheet]);
        } catch ($err) {
            Log.error(`Sheet ${$sExcelSheet} not found in Excel file ${self._sFilePath}`);
        }
        return self;
    };

    /**
     * Make the appropriate Excel Column Letter with column position.
     *
     * @param {number}  $nColumn  Column number (starting at 1).
     *
     * @return {string}
     */
    self.getColName = function ($nColumn) {
        var temp, letter = '';
        while ($nColumn > 0)
        {
            temp = ($nColumn - 1) % 26;
            letter = String.fromCharCode(temp + 65) + letter;
            $nColumn = ($nColumn - temp - 1) / 26;
        }
        return letter;
    };


    /**
     * Indicates the line where data start. This settings is use to retrieve
     * All data of a column when Dynamic methods are called without parameter.
     *
     * @param {Number} $nStartAt Line where data start (starting from 1 not 0).
     *
     * @return {ExcelHandler}
     */
    self.rowStartAt = function ($nStartAt) {
        $nStartAt = parseInt($nStartAt);

        self._nRowStartAt = (!isNaN($nStartAt)) ? $nStartAt : self._nRowStartAt;

        return self;
    };

    /**
     * Set Excel Sheet Column definition to get dynamically generated method.
     *
     * @param {object} $oColSettings Your Column name pointing to Column letter index
     * @private
     *
     * @return {ExcelHandler}
     */
    self.cols = function ($oColSettings) {
        if (self._oExcel) {
            if(!self._oCols[self._sActiveSheet]){
                self._oCols[self._sActiveSheet] = {};
            }

            self._oCols[self._sActiveSheet] = Object.assign(
                self._oCols[self._sActiveSheet],
                $oColSettings
            );
        }

        // Making Method to retrieve Line
        for (let colName in $oColSettings) {
            let column = $oColSettings[colName];

            self._aColNames[self._sActiveSheet][column] = colName;

            try {
                self[colName] = self.value.bind({
                    instance: self,
                    name: colName,
                    col: column
                });
            } catch (err) {
                Log.error(`ExcelHandlerError 1/3 :: Can not generate method for column with name '${colName}'.`);
                Log.error(`ExcelHandlerError 2/3 :: Rename your header in the Excel File.`);
                Log.error(`ExcelHandlerError 3/3 :: Or set new name for column ${column} with method cols().`);
            }
        }

        return self;
    };

    /**
     * Generic method for generated method according to defined cols
     * to return celle value / set new value.
     *
     * @TODO, Si table ouverte quand maj Excel Direct, maj Table pour pas écraser
     *        la donnée updated dans l'excel en direct
     *        -> Syncrhonisé (binder) Excel / TableJS
     *
     * @param {Number|Array} $nRows
     * @param {String|Array} $sNewValue
     *
     * @return {ExcelHandler}
     */
    self.value = function ($nRows, $sNewValue = null) {
        // Process as line
        if ($nRows) {
            // Individual Line
            if (typeof $nRows === 'number') {
                // When a value is specified -> implies cell updating (handle in synchronize)
                //
                // Updating Excel Content via ExcelHanlder need to replicates the new value
                // in the table (if instantiated)
                // Method synchronize performs the operation
                return self.synchronize.call(this, $nRows, $sNewValue);
            }
            // Mass update
            else {
                if ($nRows instanceof Array) {
                    // @TODO (For mass update)
                    // return self.synchronize.call(this, $nRows, $sNewValue);
                }
            }
        }

        // If no row specified, return the column value
        else {
            let nRowCount = self._oExcelSheet.RowCount;
            let aValues = [];

            for (let r = self._nRowStartAt; r <= nRowCount; r++) {
                let sValue = self[this.name](r);
                aValues.push(sValue);
            }

            return aValues;
        }
    };

    /**
     * Save indexes for data to keep link between Excel Line and TableJS Row Id
     */
    self.index = function () {
        return {
            indexing: function () {
                // Considering First Excel Line = First Line of the table
                // Rule is : Line Excel x = Table Id x - RowStart - 1 (Index starting from 0)
                //
                //  Also considering if table order change, order change in Excel.
                //  In the main case, the user which uses the table will not switch back to Excel
                //  But we have to considering the possibility
                //
                //  That why we have to sync Excel & Table.
                //  The table functionalities are more advanced than Excel in TC
                //  So user have to know what is currently doing when he change data order.
                //
                let nAssociatedExcelLine = self._nRowStartAt;

                // Even if the table is empty, instantiate "Index" Object for the
                // Current Sheet to prevent undefined error
                //if(!self._oIndexes[self._sActiveSheet]) self._oIndexes[self._sActiveSheet] = {
                self._oIndexes[self._sActiveSheet] = {
                    "ExcelLine": {},
                    "TableId": {},
                    "TableIdx": {},
                    "Max": {
                        "Excel": 0,
                        "TableId": 0,
                        "TableIdx": 0
                    }
                };

                self._oTables[self._sActiveSheet].forEach(function ($aRow, $nIdx) {
                    let nRowId = $aRow.id;

                    self._oIndexes[self._sActiveSheet].ExcelLine[nAssociatedExcelLine] = nRowId;
                    self._oIndexes[self._sActiveSheet].TableId[nRowId] = nAssociatedExcelLine;
                    self._oIndexes[self._sActiveSheet].TableIdx[nRowId] = $nIdx;
                    self._oIndexes[self._sActiveSheet].Max.TableIdx = $nIdx;

                    if (nRowId > self._oIndexes[self._sActiveSheet].Max.TableId) {
                        self._oIndexes[self._sActiveSheet].Max.TableId = nRowId;
                    }

                    nAssociatedExcelLine++;
                });
            },

            rowToId: function ($nRow) {
                return self._oIndexes[self._sActiveSheet].ExcelLine[$nRow];
            },

            rowToIndex: function ($nRow) {
                return self._oIndexes[self._sActiveSheet].TableIdx[self.index().rowToId($nRow)];
            },

            idToRow: function ($nId) {

            },

            indexToRow: function () {

            },

            isRowIndexed: function ($nRow) {
                return (self._oTables[self._sActiveSheet][self._oIndexes[self._sActiveSheet].ExcelLine[$nRow]]);
            },

            lastTableIndex: function () {
                return self._oIndexes[self._sActiveSheet].Max.TableId;
            }
        }
    };

    /**
     * Maintain data depending of life cycle of the ExcelHandler instance.
     * If table() method has been used, we have to work with TableJs instance
     * instead of Excel while keeping Excel methods.
     *
     * @param {Number}   $nRow       Expected Excel Line for data
     * @param {String}   $sNewValue  [Optional] New value to set for the cell
     *
     * @return {*}
     */
    self.synchronize = function ($nRow, $sNewValue) {
        let aTableSource = null;
        let oExcelSource = null;

        //  Handle Index before updating/returning
        // ------------------------------------------
        if (self._oTables[self._sActiveSheet]) {
            // Has we can not know about change made on the table
            // We have to reindex the table before synchronization
            self.index().indexing();

            // Depending of the Excel Row we are handling,
            // Maybe there is no associated TableRow
            if (self.index().isRowIndexed($nRow)) {
                aTableSource = self._oTables[self._sActiveSheet][self.index().rowToIndex($nRow)][this.name];
            } else {
                // Specific Case for lines before rowStart :
                //  - Data are not stored in the Table.
                //  - We have to work with Excel
                if ($nRow >= self._nRowStartAt) {
                    // Add missing lines to meet with the last Excel
                    let nExpectedIdx = $nRow - self._nRowStartAt;
                    let nTableMissinLine = nExpectedIdx - self.index().lastTableIndex();

                    for (let i = 0; i < nTableMissinLine; i++) {
                        self._oTables[self._sActiveSheet].push();
                    }

                    // Reindex next to additions
                    self.index().indexing();

                    // Update Table
                    aTableSource = self._oTables[self._sActiveSheet][self.index().rowToIndex($nRow)][this.name];
                }
            }
        }

        // In any case, update Excel
        oExcelSource = self._oExcelSheet.Cell(this.col, $nRow);

        //  Handling Update / Return
        // -----------------------------
        if ($sNewValue !== null) {
            if(aTableSource) aTableSource(self.number().in().format($sNewValue));
            oExcelSource.Value = self.number().in().format($sNewValue);
        } else {
            if (aTableSource) {
                return self.number().out().format(aTableSource());
            } else {
                return self.number().out().format(oExcelSource.Value);
            }
        }
    };

    self.number = function () {
        //  Handling Numbers
        // -----------------------------
        // Excel Number format :
        //  - Decimal separator : ,
        //  - Thousand separator (optional) : space
        //  - Decimal place : 2 to n where 2 = default value
        //
        // As ExcelHandler is the interface between Excel Data and
        // any application through JavaScript, it have to provided
        // the appropriate type in any case.
        //
        // Interface must identified numbers even provided as string.
        //
        // Test Complete Excel Object :
        //
        //  Rules for Output: Excel to JavaScript
        //  {General} Excel (1234,5)  ---> {Number} Javascript  1234.5
        //  {General} Excel (1234.5)  ---> {String} Javascript "1234.5"
        //  {Number}  Excel (1234,50) ---> {Number} Javascript  1234.5
        //  {Number}  Excel (1234.50) ---> {String} Javascript "1234.5"
        //  {String}  Excel (1234,5)  ---> {String} Javascript "1234.5"
        //  {String}  Excel (1234.5)  ---> {String} Javascript "1234,5"
        //
        //  Rules for Input: JavaScript to Excel
        //      Excel Type : CellType -> ResultType
        //
        //  {String}  JavaScript ("1234.5") ---> {String->String}   Excel (1234.5 )
        //  {String}  JavaScript ("1234.5") ---> {Number->String}   Excel (1234.5 )
        //  {String}  JavaScript ("1234.5") ---> {General->String}  Excel (1234.5 )
        //  {Number}  JavaScript ( 1234.5 ) ---> {String->String}   Excel (1234,5 ) /!\
        //  {Number}  JavaScript ( 1234.5 ) ---> {Number->Number}   Excel (1234,50) /!\
        //  {Number}  JavaScript ( 1234.5 ) ---> {General->Number}  Excel (1234,5 ) /!\
        //
        //
        return {
            in: function () {
                return {
                    // JS --> Excel (All as string)
                    format: function ($mInput) {
                        // Conversion must be done for string (which match with number format)
                        if (typeof $mInput === 'string' && self.number().isNumberAsString($mInput)) {
                            $mInput = $mInput.trim();

                            // Remove Thousand Separator if exist
                            if (self.number().hasSeparator($mInput)) {
                                // Try to identify separators
                                let oSeparator = self.number().separatorIdentifier($mInput);

                                // If there is separators but nothing found, use defaut
                                if (oSeparator.thousand === null && oSeparator.decimal === null) {
                                    oSeparator.thousand = self._inThousandSep;
                                    oSeparator.decimal = self._inDecimalSep;
                                }

                                if (oSeparator.thousand) {
                                    let oRegExp = new RegExp(`[${oSeparator.thousand}]`, 'g');
                                    $mInput = $mInput.replace(oRegExp, '');
                                }
                            }
                        }
                        return $mInput.toString();
                    },

                    setDecimalSeparator: function ($sSeparator = ',') {
                        self._inDecimalSep = $sSeparator.toString();

                        return self;
                    },

                    setThousandSeparator: function ($sSeparator = ',') {
                        self._inThousandSep = $sSeparator.toString();

                        return self;
                    }
                };
            },

            out: function () {
                return {
                    // Excel --> JS (String number as Number)
                    format: function ($mOutput) {
                        // Conversion must be done for string (which match with number format)
                        if (typeof $mOutput === 'string' && self.number().isNumberAsString($mOutput)) {
                            $mOutput = $mOutput.trim();

                            // Remove Thousand Separator if exist
                            if (self.number().hasSeparator($mOutput)) {
                                // Try to identify separators
                                let oSeparator = self.number().separatorIdentifier($mOutput);

                                // If there is separators but nothing found, use defaut
                                if (oSeparator.thousand === null && oSeparator.decimal === null) {
                                    oSeparator.thousand = self._inThousandSep;
                                    oSeparator.decimal = self._inDecimalSep;
                                }

                                if (oSeparator.thousand) {
                                    let oRegExp = new RegExp(`[${oSeparator.thousand}]`, 'g');
                                    $mOutput = $mOutput.replace(oRegExp, '');
                                }

                                if (oSeparator.decimal !== '.') {
                                    $mOutput = $mOutput.replace(oSeparator.decimal, '.');
                                }
                            }
                            $mOutput = parseFloat($mOutput)
                        }
                        return $mOutput;
                    },

                    setDecimalSeparator: function ($sSeparator = ',') {
                        self._outDecimalSep = $sSeparator.toString();

                        return self;
                    },

                    setThousandSeparator: function ($sSeparator = ',') {
                        self._outThousandSep = $sSeparator.toString();

                        return self;
                    }
                };
            },

            isNumberAsString: function ($sString) {
                let oRegExp = /^([0-9]{1}|[0-9,.\s]{2,})$/;

                if (oRegExp.test($sString)) {
                    // Now check if its not corresponding to a dat
                    let oDateRegExp = /^(?:([0-9]{2}[.][0-9]{2}[.][0-9]{4})|([0-9]{4}[.][0-9]{2}[.][0-9]{2}))$/;

                    if (oDateRegExp.test($sString)) {
                        return false
                    } else {
                        return true
                    }
                } else {
                    return false;
                }
            },

            hasSeparator: function ($sString) {
                let oSepRegExp = /([,.\s])/;

                return oSepRegExp.test($sString);
            },

            separatorIdentifier: function ($sString) {
                let separators = {
                    decimal: null,
                    thousand: null
                };

                // Rule for number as string where there is only on separator
                // 1,234 can be 1234 (, for thousand sep) or 1.234 (floating number)
                // If , or . is at position 3 or more, than can be only a floating number
                // Else, we can not guess expected. Here, we considering using default separators
                //
                // 1
                // 1.1
                // 1,1
                // 123
                // 123.3
                // 123,4
                // --> Can be only decimal char
                if ($sString.length <= 4) {
                    // Check for separator
                    let oSepRegExp = /([,.\s])/;
                    if (oSepRegExp.test($sString)) {
                        separators.decimal = $sString.match(oSepRegExp)[0];
                    }
                }

                // For other format
                // Checking if there is 2 kind of separators.
                // If there is only one sep, check if his role is for decimal or thousand
                //
                // 1234
                // 1234,4
                // 1234,40
                // 1234.400
                // 1,234
                // 1.234,1
                // 1,234.1
                // 1 234.1
                // 1,234.5
                // 1.234,5
                // 1 234.5
                // 1 234,5
                // 1,234,567
                // 1.234.567
                // 1,234,567.89
                // 1.234.567,89
                // 1 234 567,89
                // 1 234 567.89
                //
                else {
                    // Check for Multiple separator
                    let nSeparator = 0;
                    let aRegExp = [
                        {regexp: /([,])/, index: 0, match: ''},
                        {regexp: /([.])/, index: 0, match: ''},
                        {regexp: /([\s])/, index: 0, match: ''},
                    ];
                    // let oCommaSeparator = /[,]/;
                    // let oDotSeparator = /[.]/;
                    // let oSpaceSeparator = /[\s]/;

                    aRegExp.forEach(function ($oRegExp) {
                        if($oRegExp.regexp.test($sString)){
                            let match = $sString.match($oRegExp.regexp);
                            $oRegExp.match = match[0];
                            $oRegExp.index = match.index;
                            nSeparator++;
                        }
                    });

                    // If multiple Separator -> First = thousand, other = decimal
                    if (nSeparator > 1) {
                        let sThousandSep = null;
                        let sDecimalSep = null;
                        let nLastIndex = 0;

                        aRegExp.forEach(function ($oSepReg) {
                            if ($oSepReg.index > nLastIndex) {
                                sThousandSep = sDecimalSep;
                                sDecimalSep = $oSepReg.match;
                                nLastIndex = $oSepReg.index;
                            } else {
                                if(!sThousandSep) sThousandSep = $oSepReg.match;
                                if(!sDecimalSep) sDecimalSep = $oSepReg.match;
                            }
                        });

                        separators = {
                            thousand: sThousandSep,
                            decimal: sDecimalSep
                        };
                    }

                    // Else...
                    // We have to check for Position / Occurences
                    else {
                        // Search for Char
                        let oSeparator = null;
                        aRegExp.forEach(function ($oSepReg) {
                            if ($oSepReg.index) oSeparator = $oSepReg;
                        });

                        let aParts = $sString.split(oSeparator.match);

                        // 1,234     = 1 and 234           --> ??
                        // 1,234,567 = 1 and 234 and 567   --> Thousand
                        if (aParts.length >= 3) {
                            separators.thousand = oSeparator.match;
                        }
                        else {
                            if ($sString.length >= 5 && $sString.length <= 7) {
                                // Specific case : 0.123 or 0,123 -> Separator is for decimal
                                if ($sString.length === 5 && $sString[0] === '0') {
                                    separators.decimal = oSeparator.match;
                                } else {
                                    let nUnknownIndex = $sString.length - 4;
                                    if (oSeparator.index !== nUnknownIndex) {
                                        separators.decimal = oSeparator.match;
                                    }
                                }
                            } else {
                                separators.decimal = oSeparator.match;
                            }
                        }
                    }
                }

                return separators;
            }
        }
    };

    /**
     * Get the Excel File data as table (reference)
     *
     * @param {Array} $aKeys Fields which compose the key.
     *
     * @return {*}
     */
    self.table = function ($aKeys = []) {
        if (!self._oTables[self._sActiveSheet] && self._oExcelSheet) {
            let aTableData = [];
            let aFields = [];
            let nRowCount = self._oExcelSheet.RowCount;

            for (let colName in self._aColNames[self._sActiveSheet]) {
                if(!self._aColNames[self._sActiveSheet].hasOwnProperty(colName)) continue;
                aFields.push(self._aColNames[self._sActiveSheet][colName]);
            }

            for (let r = self._nRowStartAt; r <= nRowCount; r++) {
                let aRow = [];

                for (let colName in self._aColNames[self._sActiveSheet]) {
                    if(!self._aColNames[self._sActiveSheet].hasOwnProperty(colName)) contine;
                    let sField = self._aColNames[self._sActiveSheet][colName];
                    let sValue = self[sField](r);

                    // With Excel, if cell is empty, it returns 'undefined'
                    if (typeof sValue == 'undefined') {
                        sValue = '';
                    }

                    try {
                        sValue = sValue.trim();
                    } catch(err) { }

                    aRow.push(sValue);
                }

                aTableData.push(aRow);
            }

            // Do not serve empty Excel lines
            for (let i = aTableData.length - 1; i >= 0; i--) {
                let aLastRow = aTableData[i];
                let bFound = false;

                aLastRow.forEach(function ($sCellValue) {
                    if($sCellValue) bFound = true;
                });

                if (!bFound) {
                    aTableData.pop();
                } else {
                    break;
                }
            }

            self._oTables[self._sActiveSheet] = new TableJs(
                aFields,
                $aKeys,
                aTableData
            );

            // Perform Indexing to link Excel Line to Row ID
            self.index().indexing();
        } else {
            // Set / Update keys fields
            self._oTables[self._sActiveSheet].keys($aKeys);
        }

        return self._oTables[self._sActiveSheet];
    };

    /**
     * Close current handled Excel File.
     *
     * @param {boolean} $bWithSave Indicates if the Excel file must be save
     *                             before closing it.
     */
    self.close = function ($bWithSave = false) {
        if ($bWithSave) self.save();

        self._oExcel      = null;
        self._oExcelSheet = null;
    };

    /**
     * Commit changes made on Excel file on this disk.
     */
    self.save = function ($bWithClose = false) {
        let activeSheet = self._sActiveSheet;

        // Update Excels handled Sheets
        for (let sheet in self._oCols) {
            if(!self._oCols.hasOwnProperty(sheet)) continue;

            let oCols = self._aColNames[sheet];
            let aData = self._oTables[sheet];

            // Update Excel only if we get a table previously
            if (aData) {
                // Switch to the appropriate sheet
                self.sheet(sheet);

                // If file has header, check them (mainly to add header for new columns)
                if (self._bHasHeader) {
                    for (let colName in oCols) {
                        if (!oCols.hasOwnProperty(colName)) continue;
                        let sHeaderName = oCols[colName];

                        //@TODO : Must be rowStart - 1 instead of 1
                        if (!self[sHeaderName](1)) self[sHeaderName](1, sHeaderName);
                    }
                }

                // Update Data
                for (let r = self._nRowStartAt; r <= (aData.length + self._nRowStartAt - 1); r++) {
                    let aRow = aData[r - self._nRowStartAt];

                    for (let colName in oCols) {
                        if(!oCols.hasOwnProperty(colName)) continue;
                        let sField = oCols[colName];
                        self[sField](r, aRow[sField]());
                    }
                }
            }
        }

        // Revert back to active sheet
        self.sheet(activeSheet);

        // Save data in Excel
        self._oExcel.Save();

        if ($bWithClose) {
            self.close();
        }
    };

    // Store path
    self._sFilePath = $sPath;

    return self;
}

module.exports = ExcelHandler;