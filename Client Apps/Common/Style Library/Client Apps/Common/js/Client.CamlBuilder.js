"use strict";
var SpData;
(function (SpData) {
    /// <summary>
    ///     CAML Operations
    /// </summary>
    var CamlOperator;
    (function (CamlOperator) {
        CamlOperator[CamlOperator["Contains"] = 0] = "Contains";
        CamlOperator[CamlOperator["BeginsWith"] = 1] = "BeginsWith";
        CamlOperator[CamlOperator["Eq"] = 2] = "Eq";
        CamlOperator[CamlOperator["Geq"] = 3] = "Geq";
        CamlOperator[CamlOperator["Leq"] = 4] = "Leq";
        CamlOperator[CamlOperator["Lt"] = 5] = "Lt";
        CamlOperator[CamlOperator["Gt"] = 6] = "Gt";
        CamlOperator[CamlOperator["Neq"] = 7] = "Neq";
        CamlOperator[CamlOperator["DateRangesOverlap"] = 8] = "DateRangesOverlap";
        CamlOperator[CamlOperator["IsNotNull"] = 9] = "IsNotNull";
        CamlOperator[CamlOperator["IsNull"] = 10] = "IsNull";
    })(CamlOperator = SpData.CamlOperator || (SpData.CamlOperator = {}));
    var EventRecurranceOverlap;
    (function (EventRecurranceOverlap) {
        EventRecurranceOverlap[EventRecurranceOverlap["Now"] = 0] = "Now";
        EventRecurranceOverlap[EventRecurranceOverlap["Today"] = 1] = "Today";
        EventRecurranceOverlap[EventRecurranceOverlap["Week"] = 2] = "Week";
        EventRecurranceOverlap[EventRecurranceOverlap["Month"] = 3] = "Month";
        EventRecurranceOverlap[EventRecurranceOverlap["Year"] = 4] = "Year";
    })(EventRecurranceOverlap = SpData.EventRecurranceOverlap || (SpData.EventRecurranceOverlap = {}));
    var AggregationType;
    (function (AggregationType) {
        AggregationType[AggregationType["SUM"] = 0] = "SUM";
        AggregationType[AggregationType["COUNT"] = 1] = "COUNT";
        AggregationType[AggregationType["AVG"] = 2] = "AVG";
        AggregationType[AggregationType["MAX"] = 3] = "MAX";
        AggregationType[AggregationType["MIN"] = 4] = "MIN";
        AggregationType[AggregationType["STDEV"] = 5] = "STDEV";
        AggregationType[AggregationType["VAR"] = 6] = "VAR";
    })(AggregationType = SpData.AggregationType || (SpData.AggregationType = {}));
    var CamlBuilder = (function () {
        function CamlBuilder() {
            this.camlClauses = [];
            this.orderByFields = [];
            this.requireAll = true;
            this.viewFieldsString = "";
            this.groupByFieldsString = "";
            this.aggregationFieldsString = "";
            this.begin(true);
        }
        Object.defineProperty(CamlBuilder.prototype, "query", {
            get: function () {
                return this.buildQuery();
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CamlBuilder.prototype, "viewFields", {
            get: function () {
                return this.viewFieldsString;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(CamlBuilder.prototype, "totalClauses", {
            get: function () {
                return this.camlClauses.length;
            },
            enumerable: true,
            configurable: true
        });
        CamlBuilder.prototype.buildQuery = function () {
            if (this.recurrence) {
                return "<Where>" + this.recurrence + "</Where>";
            }
            var query = "";
            var openOperators = 0;
            if (this.camlClauses.length > 0) {
                query += "<Where>";
                //When we have just one clause we can't use AND or OR.
                if (this.camlClauses.length > 1) {
                    var totalCamlPairs = this.camlClauses.length - 1;
                    while (totalCamlPairs > 0) {
                        query += (this.requireAll ? "<And>" : "<Or>");
                        totalCamlPairs--;
                        openOperators++;
                    }
                }
                var clausesAdded = 0;
                for (var i = 0; i < this.camlClauses.length; i++) {
                    var clause = this.camlClauses[i];
                    query += clause;
                    clausesAdded++;
                    if (clausesAdded > 1) {
                        query += this.requireAll ? "</And>" : "</Or>";
                        openOperators--;
                    }
                }
                query += "</Where>";
            }
            if (this.orderByFields.length > 0) {
                query += "<OrderBy>";
                for (var i = 0; i < this.orderByFields.length; i++) {
                    var item = this.orderByFields[i];
                    query += "<FieldRef Name=\"" + item.fieldRef + "\" Ascending=\"" + (item.ascending ? "TRUE" : "FALSE") + "\" />";
                }
                query += "</OrderBy>";
            }
            return query;
        };
        Object.defineProperty(CamlBuilder.prototype, "viewXml", {
            get: function () {
                var viewFields = "";
                if (this.viewFields) {
                    viewFields = "<ViewFields>" + this.viewFields + "</ViewFields>";
                }
                var groupBy = "";
                if (this.groupByFieldsString) {
                    groupBy = "<GroupBy Collapse=\"TRUE\" GroupLimit=\"1999\">" + this.groupByFieldsString + "</GroupBy>";
                }
                var aggregations = "";
                if (this.aggregationFieldsString) {
                    aggregations = "<Aggregations Value=\"On\">" + this.aggregationFieldsString + "</Aggregations>";
                }
                return "<View" + (this.viewScope ? " " + this.viewScope : "") + ">" + viewFields + "<Query>" + groupBy + (this.totalClauses === 0 ? "" : this.query) + "</Query>" + aggregations + (this.rowLimit ? "<RowLimit>" + this.rowLimit + "</RowLimit>" : "") + "</View>";
            },
            enumerable: true,
            configurable: true
        });
        CamlBuilder.prototype.begin = function (requireAll) {
            this.camlClauses = [];
            this.requireAll = requireAll;
            this.recurrence = null;
            this.viewFieldsString = "";
            this.orderByFields = [];
        };
        CamlBuilder.prototype.getNullClause = function (fieldRef) {
            var retVal = "";
            if (fieldRef) {
                retVal = "<IsNull><FieldRef Name=\"" + fieldRef + "\" /></IsNull>";
            }
            return retVal;
        };
        CamlBuilder.prototype.getClause = function (operation, fieldRef, value, valueType) {
            var retVal = "";
            if (value) {
                retVal = "<" + CamlOperator[operation] + "><FieldRef Name=\"" + fieldRef + "\" /><Value Type=\"" + valueType + "\">" + value + "</Value></" + CamlOperator[operation] + ">";
            }
            return retVal;
        };
        CamlBuilder.prototype.getDateTimeClause = function (operation, fieldRef, value, includeTime) {
            var retVal = "";
            if (value) {
                retVal = "<" + CamlOperator[operation] + "><FieldRef Name=\"" + fieldRef + "\" /><Value Type=\"DateTime\" " + (includeTime ? "IncludeTimeValue='TRUE'" : "") + ">" + value + "</Value></" + CamlOperator[operation] + ">";
            }
            return retVal;
        };
        CamlBuilder.prototype.addNullClause = function (fieldRef) {
            this.camlClauses.push(this.getNullClause(fieldRef));
        };
        CamlBuilder.prototype.addTextClause = function (operation, fieldRef, value) {
            this.camlClauses.push(this.getClause(operation, fieldRef, value, "Text"));
        };
        CamlBuilder.prototype.addBooleanClause = function (operation, fieldRef, value) {
            this.camlClauses.push(this.getClause(operation, fieldRef, value ? "1" : "0", "Integer"));
        };
        CamlBuilder.prototype.addNumberClause = function (operation, fieldRef, value) {
            this.camlClauses.push(this.getClause(operation, fieldRef, value.toString(), "Number")); //TODO: verify type
        };
        CamlBuilder.prototype.addDateTimeClause = function (operation, fieldRef, value, includeTime) {
            if (value) {
                this.camlClauses.push(this.getDateTimeClause(operation, fieldRef, value, includeTime));
            }
        };
        CamlBuilder.prototype.addLookupClause = function (operation, fieldRef, value, valueType) {
            if (value) {
                var clause = "<" + CamlOperator[operation] + "><FieldRef Name=\"" + fieldRef + "\" LookupId=\"True\" /><Value Type=\"" + valueType + "\">" + value + "</Value></" + CamlOperator[operation] + ">";
                this.camlClauses.push(clause);
            }
        };
        CamlBuilder.prototype.addAggregation = function (fieldRef, aggregationType) {
            this.aggregationFieldsString = this.aggregationFieldsString + ("<FieldRef Name=\"" + fieldRef.replace(" ", "_x0020_") + "\" Type=\"" + AggregationType[aggregationType] + "\" />");
        };
        CamlBuilder.prototype.recurrenceQuery = function (overlapType) {
            this.recurrence = "<DateRangesOverlap><FieldRef Name=\"EventDate\"/><FieldRef Name=\"EndDate\"/><FieldRef Name=\"RecurrenceID\"/><Value><" + EventRecurranceOverlap[overlapType] + "/></Value></DateRangesOverlap>";
        };
        CamlBuilder.prototype.addViewField = function (fieldRef) {
            this.viewFieldsString = this.viewFieldsString + ("<FieldRef Name=\"" + fieldRef.replace(" ", "_x0020_") + "\"/>");
        };
        CamlBuilder.prototype.addViewFields = function (fieldRefs) {
            for (var i = 0; i < fieldRefs.length; i++) {
                this.addViewField(fieldRefs[i]);
            }
        };
        CamlBuilder.prototype.addGroupByField = function (fieldRef) {
            this.groupByFieldsString = this.groupByFieldsString + ("<FieldRef Name=\"" + fieldRef.replace(" ", "_x0020_") + "\"/>");
        };
        CamlBuilder.prototype.addGroupByFields = function (fieldRefs) {
            for (var i = 0; i < fieldRefs.length; i++) {
                this.addGroupByField(fieldRefs[i]);
            }
        };
        CamlBuilder.prototype.addOrderBy = function (fieldRef, ascending) {
            var orderBy = new OrderBy();
            orderBy.fieldRef = fieldRef;
            orderBy.ascending = ascending;
            this.orderByFields.push(orderBy);
        };
        CamlBuilder.prototype.addDaysToDate = function (date, days) {
            date.setDate(date.getDate() + days);
            return date;
        };
        return CamlBuilder;
    }());
    SpData.CamlBuilder = CamlBuilder;
    var OrderBy = (function () {
        function OrderBy() {
        }
        return OrderBy;
    }());
    SpData.OrderBy = OrderBy;
})(SpData || (SpData = {}));
//# sourceMappingURL=Client.CamlBuilder.js.map