const TableJs = require('TableJs');

/**
 * Version v0.1.3
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

    self._sFilePath    = null;
    self._oExcel       = null;
    self._oExcelSheet  = null;
    self._sActiveSheet = null;
    self._bHasHeader   = $bHasHeaders;
    self._nRowStartAt  = ($bHasHeaders) ? 2 : 1;
    self._oCols        = {};
    self._aColNames    = {};
    self._oSheetRead   = {};
    self._oTables      = {};

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
     * @param {Number|Array} $nRows
     * @param {String|Array} $sNewValue
     *
     * @return {ExcelHandler}
     */
    self.value = function ($nRows, $sNewValue = null) {
        if ($nRows) {
            if (typeof $nRows === 'number') {
                if ($sNewValue !== null) {
                    self._oExcelSheet.Cell(this.col, $nRows).Value = $sNewValue;
                }
                return self._oExcelSheet.Cell(this.col, $nRows).Value;
            } else {
                if ($nRows instanceof Array) {
                    // @TODO
                }
            }
        } else {
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

            // let oCols = self._oCols[sheet];
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


