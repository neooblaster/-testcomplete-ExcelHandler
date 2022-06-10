/**
 * Instantiates an enhanced Array, which works as Array with extra features
 * to manipulates rows/cells.
 *
 * Callback System (Inspired from SAP Enhancement Concept) :
 *  - Core.add() :
 *      - pre,  receive {Arguments} func params,           must return {Arguments}.
 *      - push, receive {String|Array} one argument value, must return {String|Array}.
 *      - post, receive {Array} on game data,              must return {Array}.
 *
 *
 * @param {Array} $fields  Field list of the new table.
 * @param {Array} $keys    Field from field list which will compose the unique key.
 * @param {Array} $array   2D Table data.
 *
 * @return {Array}
 *
 * @constructor
 */
function TableJs($fields, $keys, $array) {

    // -----------------------------------------------------
    //    Master Data
    // -----------------------------------------------------
    /**
     * @type {TableJs} Current TableJS Instance.
     */
    let self = this;

    /**
     * @type {Array} Array which contains data. Master Array
     * @private
     */
    self._data = [];


    // -----------------------------------------------------
    //    Internal Data
    // -----------------------------------------------------
    /**
     * @type {null|TableJs} Next to a selection, a new TableJs is created and
     * parent TableJs is link in this property.
     * @private
     */
    self._parent = null;

    /**
     * @type {Array} List of defined fields.
     * @private
     */
    self._fields = [];

    /**
     * @type {Array} List of defined fields which compose the key.
     * @private
     */
    self._keys = [];

    /**
     * Indexes for performances
     *
     * @type {{
     *      nextId:     number,     Each row must have a unique ID
     *      deprecated: boolean,    Flag indicating indexes are deprecated
     *      byId:       {Object},   Return row  index using row ID
     *      byKeys:     {Object},   Return row  index using row key (composition of fields)
     *      byFields:   {Object}    Return rows indexes for a value for a field
     * }}
     * @private
     */
    self._indexes = {
        deprecated: true,
        nextId: 0,
        byId: {},
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

                    if (!(Object.prototype.toString.call(argv) === "[object Array]")) {
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

                // Note : do not attach parent (core().setParent()).
                // The copy must fully detached.
                return new TableJs(
                    self._fields,
                    self._keys,
                    copy
                );
            }
        },

        /**
         * Update field(s) of selection with provided values
         *
         * @param {Object}   $newFieldValues Each properties must be the field with value
         *
         * @return {Array}
         */
        update: {
            enumerable: false,
            writable: false,
            value: function ($newFieldValues) {
                for (let field in $newFieldValues) {
                    if(!$newFieldValues.hasOwnProperty(field)) continue;
                    // Check if function exist (to prevent issues for wrong property in update)
                    if (!this.hasOwnProperty(field)) continue;

                    // Make this update
                    this.forEach(function ($row) {
                        $row[field]($newFieldValues[field]);
                    });
                }

                return self._data;
            }
        },

        /**
         * Delete rows of the table. Purposed for sub TableJs.
         *
         * @return {Array}
         */
        delete: {
            enumerable: false,
            writable: false,
            value: function () {
                // The order is very important,
                // Because on splice, indexes are updated.
                // We have to process in reversed order.
                // We have to delete row at same time
                // Get RowId. Method deleteRows is in charge to delete them in
                // the appropriate order
                let rowsId = [];

                this.forEach(function ($row) {
                    rowsId.push($row.id);
                });

                this.core().deleteRows(rowsId);

                // Next to the deletion,
                // We have to update indexes
                self._data.core().indexing();

                // Parent are deprecated
                self._data.core().setDeprecated(true);

                return self._data;
            }
        },

        /**
         * WIP ?
         */
        insert: {
            enumerable: false,
            writable: false,
            value: function () {
                return self._data;
            }
        },

        /**
         * Create a new table row with appropriates fields and methods
         *
         * @param [Array|String]  [Optional] Data for the new row.
         *
         * @return [Array]		  The TableJs Row with all generated methods.
         */
        new: {
            enumerable: false,
            writable: false,
            value: function ($aRowData = []) {
                // Create a new empty row, then push it in the table.
                // Return the created row
                return this[this.push($aRowData) - 1];
            }
        },

        /**
         * Return indexes table. Not purpose for handling but to take acknowledge
         * about data.
         *
         * @return {object}
         */
        indexes: {
            enumerable: false,
            writable: false,
            value: function () {
                return self._indexes;
            }
        },

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
                            if (Object.prototype.toString.call(argv) === "[object Array]") {
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
                                //
                                // Note : TestComplete not able to see arguments
                                // passed directly with bracket as an instance of Array
                                // For unknown, try as object which is an table
                                try {
                                    argv.forEach(function ($value) {
                                        if (data.lastIndexOf($value) < 0) {
                                            if (this.callbacks && this.callbacks.add && this.callbacks.add.push) {
                                                $value = this.callbacks.add.push.call(this, $value);
                                            }
                                            data.push($value);
                                        }
                                    }.bind(this));
                                } catch ($err) {

                                }
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
                            self._indexes.byId     = {};
                            self._indexes.byKeys   = {};
                            self._indexes.byFields = {};

                            sourceData = self._data;

                            // Index in now refreshed
                            self._indexes.deprecated = false;
                        }

                        // Reading Data
                        for (let r = 0; r < sourceData.length; r++) {
                            let row   = sourceData[r];
                            let rowId = row.id;

                            // Read Row Id and store its index
                            self._indexes.byId[rowId] = ($index === undefined) ? r : $index;

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
                                    let message = `Error : The generated key '${rowKey}' already exist for the row index ${self._indexes.byKeys[rowKey]}\n`;
                                    message += `You should add some fields as key component or check your data.\n`;
                                    message += `Your selected fields for the unique keys are : \n`;
                                    self._keys.map(function($field){
                                        message += ` • ${$field}\n`;
                                    });

                                    throw message ;
                                } else {
                                    self._indexes.byKeys[rowKey] = ($index === undefined) ? r : $index;
                                }
                            }
                        }
                    },

                    /**
                     * Indicates the index table is deprecated.
                     * Reflection for all parents TableJs.
                     *
                     * @return {boolean}
                     */
                    setDeprecated: function ($parentOnly = false) {
                        // In some case, we only want to deprecated
                        // index of parent(s)
                        if (!$parentOnly) {
                            self._indexes.deprecated = true;
                        }

                        if (self._parent) {
                            self._parent._data.core().setDeprecated();
                        }

                        return true;
                    },

                    /**
                     * Set in current table the parent table for communication.
                     *
                     * @param {TableJs} $parent
                     *
                     * @return {[]}
                     */
                    setParent: function ($parent) {
                        self._parent = $parent;
                        self._indexes.nextId = $parent._indexes.nextId;

                        return self._data;
                    },

                    /**
                     * Increment for all TableJs the self._indexes.nextId
                     * when a new row is added.
                     */
                    incrementNextId: function () {
                        self._indexes.nextId++;

                        if (self._parent) {
                            self._parent._data.core().incrementNextId();
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
                            if (!(Object.prototype.toString.call(forValue) === "[object Array]")) {
                                forValue = [forValue];
                            }

                            // Manage duplicated values
                            forValue.forEach(function ($value) {
                                if(values.lastIndexOf($value) < 0) values.push($value);
                            });
                        }

                        // Retrieve rows for provided values
                        if (values.length > 0) {
                            // If index is flagged as deprecated,
                            // perform a full reindex to get right values
                            if (self._indexes.deprecated) {
                                this.core().indexing();
                            }

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
                        ).core().setParent(self);
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
                            // Update Locally
                            $row[fieldIndex] = $value;

                            // Set Index deprecated
                            self._data.core().setDeprecated();
                        }

                        // In any case, return the current value.
                        return $row[fieldIndex];
                    },

                    /**
                     * Delete definitely the provided row (reflected in all parents).
                     *
                     * @param {Array} $rowsId  List of Row Id to delete.
                     */
                    deleteRows: function ($rowsId) {
                        // Retrieve indexes for rows ID
                        let indexes = {};

                        $rowsId.forEach(function ($id) {
                            indexes[self._indexes.byId[$id]] = $id;
                        });

                        let reversed = [];

                        for (let idx in indexes) {
                            reversed.push(parseInt(idx));
                        }

                        reversed.reverse().forEach(function ($index) {
                            self._data.splice($index, 1);
                        });

                        // Delete in all parents
                        if (self._parent) {
                            self._parent._data.core().deleteRows($rowsId);
                        }
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
                    /**
                     * Define new property which is a function where the name
                     * is the provided field name.
                     *
                     * Purpose : Core().add()/push callback point.
                     *
                     * @param {String} $field
                     *
                     * @return {*}
                     */
                    createMethod: function ($field) {
                        Object.defineProperty(self._data, $field, {
                            enumerable: false,
                            writable: false,
                            value: function () {
                                return self._data.core.apply(this).values.call(this, $field, arguments);
                            }
                        });

                        return $field;
                    },

                    /**
                     * Process all existing entries to consolidate rows according
                     * to the current field list.
                     *
                     * Purpose : Core().add()/post callback point.
                     *
                     * @param {Array} $fields In its context, receive field list
                     *
                     */
                    consolidateAll: function ($fields) {
                        if (self._data.length > 0) {
                            self._data = self._data.map(function ($row) {
                                return self._data.data().consolidate($row);
                            });
                        }

                        return $fields;
                    }
                };

                let extended = {
                    data: 'fields',
                    callbacks: {
                        add: {
                            push: functions.createMethod,
                            post: functions.consolidateAll
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
                    /**
                     * Before pushing the requested field as key, check if the field
                     * is defined.
                     *
                     * Purpose : Core().add()/push callback point.
                     *
                     * @param {String} $field
                     *
                     * @return {String} $field
                     */
                    isFieldExist: function ($field) {
                        if (self._fields.lastIndexOf($field) < 0) {
                            let field = $field;
                            if(field === '') field = '<empty>';
                            if(field === null) field = '<null>';
                            if(field === 'undefined') field = '<undefined>';
                            let message = `The requested field '${field}' as key component does not exist`;
                            throw message;
                        }

                        return $field;
                    },


                    /**
                     * Once fields are added as key component, trigger a new
                     * indexing process to update byKey index.
                     *
                     * Purpose : Core().add()/post callback point.
                     *
                     * @param {Array}  $keys
                     *
                     * @return {Array} $keys
                     */
                    triggerIndexing: function ($keys) {
                        self._data.core().indexing();

                        return $keys;
                    }
                };

                let extended = {
                    data: 'keys',
                    callbacks: {
                        add: {
                            push: functions.isFieldExist,
                            post: functions.triggerIndexing
                        }
                    },
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
                     * Purpose : Core().add()/pre callback point.
                     *
                     * @return {IArguments}
                     */
                    preprocArgs: function () {
                        for (let i in arguments) {
                            if(!arguments.hasOwnProperty(i)) continue;
                            let argv = arguments[i];

                            if (Object.prototype.toString.call(argv) === "[object Array]") {
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

                        // Set a unique Row Id for internal use
                        let id = $row.id;
                        if (id === undefined) {
                            id = self._indexes.nextId;
                            self._data.core().incrementNextId();
                        }

                        // Set a unique Id for the row to identify them
                        // easily internally.
                        Object.defineProperty($row, 'id', {
                            enumerable: false,
                            writable: false,
                            configurable: true,
                            value: id
                        });

                        // Create 'field method' to return/set field value
                        // for the current row.
                        self._fields.forEach(function ($field) {
                            Object.defineProperty($row, $field, {
                                enumerable: false,
                                writable: false,
                                configurable: true,
                                value: function ($value) {
                                    return self._data.core.apply(this).value.call(this, $field, $row, $value);
                                }
                            });
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
                            // push: functions.consolidate //---> Best integrated in Array push rewrited method
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