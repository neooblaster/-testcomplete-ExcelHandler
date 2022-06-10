clog = console.log;

// @TODO : sur ajout de champs, mettre à jour chaque entrée pour prendre en charge la nouvelle zone
// @T0D0 : creer une method "Update" pour tourner les fields functions en tant que
//          setters plutôt que getter : table.FIELD('xxx').update().FIELD('newValue');
// @TODO : Creer une méthod "Copy" pour detacher les liaison des tables (ou detach)

/**
 * Callback System (Inspired from SAP Exit Concept) :
 *  - Core.add() :
 *      - pre,  receive {Arguments},    must return {Arguments}.
 *      - push, receive {String|Array}, must return {String|Array}.
 *      - post, receive {Array},        must return {Array}.
 *
 */

/**
 * Instantiates an enhanced Array, which works as Array with extra features
 * to manipulates rows/cells.
 *
 * @param {Array} $fields
 * @param {Array} $keys
 * @param {Array} $array
 *
 * @return {TableJs}
 *
 * @constructor
 */
function TableJs($fields, $keys, $array) {
    let self = this;

    // Master Array
    self._data = [];

    // Side Data
    self._fields = [];
    self._keys = [];
    self._indexes = {
        byKeys: {},
        byFields: {}
    };

    /**
     * Enhance Master Array for chaining features while keeping manipulating Array
     */
    Object.defineProperties(self._data, {
        /**
         * Keeping TableJs reference in Array
         */
        instance: {
            enumerable: false,
            writable: false,
            value: self
        },

        /**
         * Rewrite native Push method to add table row instead of unique value.
         *
         * @return {Number} return the next index number (as original one).
         */
        push: {
            enumerable: false,
            writable: false,
            value: function () {
                let nextIndex = this.length;

                for (let a in arguments) {
                    if(!arguments.hasOwnProperty(a)) continue;
                    let argv = arguments[a];

                    if (!(argv instanceof Array)) {
                        argv = [argv];
                    }

                    let row =  this.data().consolidate(argv);

                    // Re-indexing next to push
                    this.core().indexing(row, nextIndex);

                    this[nextIndex] = row;
                    nextIndex++;
                }

                return nextIndex;
            }
        },

        /**
         * Make a full new copy of the table (no link between rows)
         *
         * @return {TableJs}
         */
        copy: {
            enumerable: false,
            writable: false,
            value: function () {
                let copy = [];

                for (let i in this) {
                    let row = this[i];
                    let newRow = [];

                    for (let f in row) {
                        newRow.push(row[f]);
                    }
                    copy.push(newRow)
                }

                return new TableJs(
                    self._fields,
                    self._keys,
                    copy
                );
            }
        },

        // /**
        //  * Rewrite native ForEach method to return row (instead of table)
        //  * in callback parameter
        //  *
        //  * @param {Function} $callback
        //  */
        // forEach: {
        //     enumerable: false,
        //     writable: false,
        //     value: function ($callback) {
        //         for (let i in this) {
        //             if(!this.hasOwnProperty(i)) continue;
        //             let entry = new TableJs(
        //                 self._fields,
        //                 self._keys,
        //                 this[i]
        //             );
        //             $callback.call(this, entry.data().getRow(0));
        //         }
        //     }
        // },

        /**
         * Fields, Keys & Data have the same working process.
         * - Pooling by using Core
         * - Specialisation using callbacks
         *
         * @return ...
         */
        core: {
            enumerable: false,
            writable: false,
            value: function () {
                // For better usage, each variant can be called directly
                // Call directly, return instance.
                // Call with empty parameter will return sub-methods
                if (arguments.length > 0) {
                    return self._data.core.apply(this).set.apply(this, arguments);
                }

                // Return basic method + those for call (fields, keys or data)
                let returning = Object.assign({
                    /**
                     * Flush & Set provided arguments values in bound variable name (this.data).
                     *
                     * @return {TableJs}
                     */
                    set: function () {
                        let data = self[`_${this.data}`];

                        // Refresh data
                        data = [];

                        // Add Data
                        return self._data.core.apply(this).add.apply(this, arguments);
                    },

                    /**
                     * Add to the end arguments values in bound variable name (this.data).
                     *
                     * @return {TableJs}
                     */
                    add: function () {
                        let data = self[`_${this.data}`];

                        // Pre processor for arguments
                        if (this.callbacks && this.callbacks.add && this.callbacks.add.pre) {
                            arguments = this.callbacks.add.pre.apply(this, arguments);
                        }

                        // Process arguments
                        for (let a = 0; a < arguments.length; a++) {
                            let argv = arguments[a];

                            // Argument value is an Array
                            if (argv instanceof Array) {
                                argv.forEach(function ($value) {
                                    if (data.lastIndexOf($value) < 0) {
                                        if (this.callbacks && this.callbacks.add && this.callbacks.add.push) {
                                            $value = this.callbacks.add.push.call(this, $value);
                                        }
                                        data.push($value);
                                    }
                                }.bind(this));
                            }
                            // Argument value is a String
                            if (typeof argv === 'string') {
                                if (data.lastIndexOf(argv) < 0) {
                                    if (this.callbacks && this.callbacks.add && this.callbacks.add.push) {
                                        argv = this.callbacks.add.push.call(this, argv);
                                    }
                                    data.push(argv);
                                }
                            } else {
                                // What can we do for other type ?
                            }
                        }

                        // Post Processing
                        if (this.callbacks && this.callbacks.add && this.callbacks.add.post) {
                            data = this.callbacks.add.post.call(this, data);
                        }

                        // Indexing Data ---> Made by Push method
                        // self._data.core.apply(this).indexing();

                        return self._data;
                    },

                    /**
                     * Return values stored in bounded variable name (this.data).
                     *
                     * @return {Array}
                     */
                    get: function () {
                        return self[`_${this.data}`];
                    },

                    /**
                     * Read data an creates index to easily return values.
                     */
                    indexing: function ($row, $index) {
                        let sourceData = null;

                        // When row is provided, only index the provided row
                        if ($row !== undefined) {
                            sourceData = [$row];
                        }
                        // Indexing game data
                        else {
                            // Flush Indexes
                            self._indexes = {
                                byKeys: {},
                                byFields: {}
                            };
                            sourceData = self._data;
                        }

                        // Reading Data
                        for (let r = 0; r < sourceData.length; r++) {
                            let row = sourceData[r];

                            // Reading for fields
                            for (let f = 0; f < self._fields.length; f++) {
                                let field = self._fields[f];
                                let value = row[f];

                                if (!self._indexes.byFields[field]) self._indexes.byFields[field] = {};
                                if (!self._indexes.byFields[field][value]) self._indexes.byFields[field][value] = [];
                                self._indexes.byFields[field][value].push(($index === undefined) ? r : $index);
                            }

                            // Reading Key
                            if (self._keys.length > 0) {
                                let rowKey = "";
                                let keyFields = [];

                                // Get fields with their index to have
                                // a unique key (according to fields definition)
                                // regardless the order of the key definition.
                                for (let k = 0; k < self._keys.length; k++) {
                                    let key = self._keys[k];
                                    let keyFieldIndex = self._fields.lastIndexOf(key);
                                    keyFields[keyFieldIndex] = key;
                                }

                                // Making the unique key
                                for (let i in keyFields) {
                                    if(!keyFields.hasOwnProperty(i)) continue;
                                    rowKey += String(row[i]);
                                }

                                // Check if the key is not already defined
                                if (self._indexes.byKeys.hasOwnProperty(rowKey)) {
                                    throw `Key '${rowKey}' already exist for row index ${self._indexes.byKeys[rowKey]}`;
                                } else {
                                    self._indexes.byKeys[rowKey] = ($index === undefined) ? r : $index;
                                }
                            }
                        }
                    },

                    /**
                     * Return rows where the cells respond to the requested values.
                     *
                     * @param {String}    arguments[0] Fields to control.
                     * @param {Arguments} arguments[1] List of value to retrieve.
                     *
                     * @return {Array}    Table with corresponding rows.
                     */
                    values: function () {
                        let field = arguments[0];
                        let requestedValues = arguments[1];
                        let values = [];
                        let data = [];

                        // Process All argument (except first which is field)
                        for (let a = 0; a < requestedValues.length; a++) {
                            let forValue = requestedValues[a];

                            // For common processing, transform string to array
                            if (!(forValue instanceof Array)) {
                                forValue = [forValue];
                            }

                            // Manage duplicated values
                            forValue.forEach(function ($value) {
                                if(values.lastIndexOf($value) < 0) values.push($value);
                            });
                        }

                        // Retrieve rows for provided values
                        if (values.length > 0) {
                            values.forEach(function ($value) {
                                if (self._indexes.byFields[field][$value]) {
                                    self._indexes.byFields[field][$value].forEach(function ($index) {
                                        if (self._data[$index]) data.push(self._data[$index]);
                                    });
                                }
                            });
                        } else {
                            let distinct = [];
                            for(let value in self._indexes.byFields[field]) {
                                distinct.push(value);
                            }
                            return distinct;
                        }


                        // Result is a part of table, so create a new TableJs
                        // to get all features
                        return new TableJs(
                            self._fields,
                            self._keys,
                            data
                        );
                    },

                    /**
                     * Set and/or Return the value of the corresponding field.
                     *
                     * @return {*}  Value of the field
                     */
                    value: function ($field, $row, $value) {
                        let fieldIndex = self._fields.lastIndexOf($field);

                        // If a value is set, that implies we want to set
                        // a new value
                        if ($value !== undefined) {
                            // Update all references

                            // Update Locally
                            $row[fieldIndex] = $value;
                        }

                        return $row[fieldIndex];
                    }

                }, this.returns);

                // Create methods to get rows from fields
                let coreContext = this;
                self._fields.forEach(function ($field) {
                    returning[$field] = function () {
                        return self._data.core.apply(this).values.call(this, $field, arguments);
                    }.bind(coreContext);
                });

                // To standardize & for extended functions,
                // Make function wrapper to bind "this".
                for (let fName in returning) {
                    if (!returning.hasOwnProperty(fName)) continue;

                    let fn = returning[fName];
                    returning[fName] = function () {
                        return fn.apply(this, arguments);
                    }.bind(this);
                }

                return returning;
            }
        },

        /**
         * Field Manager.
         *
         * @return ...
         */
        fields: {
            enumerable: false,
            writable: false,
            value: function () {
                let functions = {
                    createMethod: function ($field) {
                        Object.defineProperty(self._data, $field, {
                            enumerable: false,
                            writable: false,
                            value: function () {
                                return self._data.core.apply(this).values.call(this, $field, arguments);
                            }
                        });

                        return $field;
                    }
                };

                let extended = {
                    data: 'fields',
                    callbacks: {
                        add: {
                            push: functions.createMethod
                        }
                    },
                    returns: functions
                };

                return self._data.core.apply(extended, arguments);
            }
        },

        /**
         * Keys Manager.
         *
         * @return ...
         */
        keys: {
            enumerable: false,
            writable: false,
            value: function () {
                let functions = {

                };

                let extended = {
                    data: 'keys',
                    callbacks: {},
                    returns: functions
                };

                return self._data.core.apply(extended, arguments);
            }
        },

        /**
         * Data Manager.
         *
         * @return ...
         */
        data: {
            enumerable: false,
            writable: false,
            value: function () {
                let functions = {
                    /**
                     * Add to the end values in the table data.
                     * - String value stands for one line
                     * - Array of string stands for one line
                     * - Array of Array stands for data matrix
                     * All here before type can be mixted at once
                     *
                     * @return ...
                     */
                    append: function () {
                        return self._data.core.apply(extended).add.apply(extended, arguments);
                    },

                    /**
                     * Wraps array of string into an array to become data matrix.
                     *
                     * @return {IArguments}
                     */
                    preprocArgs: function () {
                        for (let i in arguments) {
                            if(!arguments.hasOwnProperty(i)) continue;
                            let argv = arguments[i];

                            if (argv instanceof Array) {
                                if (argv.length > 0) {
                                    if (typeof argv[0] === 'string') {
                                        arguments[i] = [argv];
                                    }
                                }
                            }
                        }

                        return arguments;
                    },

                    /**
                     * Complete row according to number of set fields.
                     *
                     * @param {Array} $row  Data matrix row to process
                     *
                     * @return {Array} Consolidated row
                     */
                    consolidate: function ($row) {
                        if (typeof $row === 'string') {
                            $row = [$row];
                        }

                        let fieldsNumber = self._fields.length;
                        let rowNumber = $row.length;
                        let missingCell = fieldsNumber - rowNumber;

                        if (missingCell > 0) {
                            for (let i = 0; i < missingCell; i++) {
                                $row.push('');
                            }
                        }

                        self._fields.forEach(function ($field) {
                            Object.defineProperty($row, $field, {
                                enumerable: false,
                                writable: false,
                                configurable: true,
                                value: function ($value) {
                                    return self._data.core.apply(this).value.call(this, $field, $row, $value);
                                }
                            })
                        });

                        return $row;
                    },

                    /**
                     * Return a table row using it index.
                     *
                     * @param {Number} $index  Row index starting from 0.
                     *
                     * @return {Array} Table row.
                     */
                    getRow: function ($index = 0) {
                        if (typeof $index !== 'number') $index = 0;
                        return self._data[$index];
                    }
                };

                let extended = {
                    data: 'data',
                    callbacks: {
                        add: {
                            pre: functions.preprocArgs,
                            push: functions.consolidate
                        }
                    },
                    returns: functions
                };

                return self._data.core.apply(extended, arguments);
            }
        }
    });



    /** ------------------------------------------------------------------
     *   Internal Processing
     * -------------------------------------------------------------------
     */
    // Setting Up Fields
    if ($fields !== undefined) {
        self._data.fields($fields);
    }

    // Setting Up Keys
    if ($keys !== undefined) {
        self._data.keys($keys);
    }

    // Setting Up Data
    if ($array !== undefined)  {
        self._data.data($array);
    }

    // Return enhanced Array
    return self._data;
}

try {
    module.exports = TableJs;
} catch ($err) {
    // Not in NodeJs or require.js not available
}