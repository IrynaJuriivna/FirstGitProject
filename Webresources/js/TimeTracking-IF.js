var Enavate = Enavate || {};
if (!Enavate.TimeTracking) Enavate.TimeTracking = {};

Enavate.TimeTracking.Utils = {
    //+
    setLookupFieldToCurrentUser: function (formContext, fieldName) {
        var globalContext = Xrm.Utility.getGlobalContext();

        let currentUser = [];
        currentUser[0] = {};
        currentUser[0].entityType = "systemuser";
        currentUser[0].name = "name";


        formContext.getAttribute(fieldName).setValue(currentUser);
    },

    //- unavailable CalculateRollupField FakeMessageExecutor
    calcRollupField: function (formContext, strTargetEntitySetName, strTargetRecordId, strTargetFieldName) {
        strTargetRecordId = strTargetRecordId.replace("{", "").replace("}", "");
        let req = new XMLHttpRequest();
        let requestUrl = `${Xrm.Page.context.getClientUrl()}/api/data/v9.0/CalculateRollupField(Target=@p1,FieldName=@p2)?@p1={'@odata.id':'${strTargetEntitySetName}(${strTargetRecordId})'}&@p2='${strTargetFieldName}'`;

        req.open("GET", requestUrl, false);

        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("OData-MaxVersion", "4.0");

        req.send();
    },

    //+
    formatDate: function (dateValue) {
        const offset = dateValue.getTimezoneOffset()
        dateValue = new Date(dateValue.getTime() - (offset * 60 * 1000))
        let result = dateValue.toISOString().slice(0, 10);
        return result;
    },

    //+
    setLookupField: function (formContext, fieldName, entityName, id, name) {
        let values = [];
        values[0] = {};
        values[0].entityType = entityName;
        values[0].id = id;
        values[0].name = name;
        formContext.getAttribute(fieldName).setValue(values);
    }
}

