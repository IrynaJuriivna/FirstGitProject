var Enavate = Enavate || {};
if (!Enavate.TimeTracking) Enavate.TimeTracking = {};
if (!Enavate.TimeTracking.Forms) Enavate.TimeTracking.Forms = {};

Enavate.TimeTracking.Utils = {
    //+
    setLookupFieldToCurrentUser: function (formContext, fieldName) {
        var globalContext = Xrm.Utility.getGlobalContext();

        let currentUser = [];
        currentUser[0] = {};
        currentUser[0].entityType = "client";
        currentUser[0].name = "Ira";


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

Enavate.TimeTracking.Forms.TimeTrackingRecordForms = {
    formOnLoaяd: function (executionContext) {
        var formContext = executionContext.getFormContext();
        let currentRecordId = formContext.data.entity.getId();
        if (!currentRecordId) {
            Enavate.TimeTracking.Utils.setLookupFieldToCurrentUser(formContext, "st08_employee");
        }
        formContext.getAttribute("st08_reportedhours").addOnChange(this.onReportHoursChangeHandler.bind(this));
        formContext.getAttribute("st08_employee").addOnChange(this.onEmployeeChangeHandler.bind(this));
        formContext.getAttribute("st08_date").addOnChange(this.onDateChangeHandler.bind(this));
        //      formContext.data.entity.addOnSave(this.onReportHoursSaveHandler.bind(this));
        formContext.data.entity.addOnSave(this.formOnSaveHandler.bind(this));
    },

    onReportHoursChangeHandler: function (executionContext) {
        var formContext = executionContext.getFormContext();
        let reportedhours = formContext.getAttribute("st08_reportedhours").getValue();

        if (reportedhours > 8) {
            formContext.ui.setFormNotification("Please look to 'Reported hours' field. It shouldn`t be more then 8 hours!", "WARNING", "HoursNotification")
        } else {
            formContext.ui.clearFormNotification("HoursNotification")
        }
    },

    onEmployeeChangeHandler: function (executionContext) {
        this.setDayRecord(executionContext);
    },

    //-
    onDateChangeHandler: function (executionContext) {
        this.setDayRecord(executionContext);

        var formContext = executionContext.getFormContext();
        let selectedDate = formContext.getAttribute("st08_date").getValue();

        //При зміні дати додаєм фільтр
        if (selectedDate) {
            formContext.getControl("st08_project").addPreSearch(this.setProjectLookupFilter);
            //     formContext.ui.setFormNotification("1", "WARNING", "HoursNotification");
        } else {
            formContext.getControl("st08_project").removePreSearch(this.setProjectLookupFilter);
        }
    },

    setDayRecord: function (executionContext) {
        var formContext = executionContext.getFormContext();
        var ttdate = formContext.getAttribute("st08_date").getValue();
        let employeeRef = formContext.getAttribute("st08_employee").getValue();
        let epmloyeeId = (employeeRef && employeeRef.length > 0) ? employeeRef[0].id : null;

        if (ttdate && epmloyeeId) {
            let dateStr = Enavate.TimeTracking.Utils.formatDate(ttdate);
            console.log(dateStr);

            Xrm.WebApi.retrieveMultipleRecords("st08_timetrackingday", `?$select=st08_name,st08_timetrackingdayid&$filter=st08_date eq ${dateStr} and _st08_employee_value eq ${epmloyeeId}&$top=1`).then(
                function success(rezult) {
                    if (rezult.entities.length > 0) {
                        let id = rezult.entities[0].st08_timetrackingdayid;
                        let name = rezult.entities[0].st08_name;
                        Enavate.TimeTracking.Utils.setLookupField(formContext, "st08_day", "st08_timetrackingday", id, name);
                    } else {
                        formContext.getAttribute("st08_day").setValue(null);
                    }
                    formContext.getAttribute("st08_day").fireOnChange();
                },
                function (error) {
                    console.log(error.message);
                }
            )
        } else {
            formContext.getAttribute("st08_day").setValue(null);
        }
    },

    //- (calcRollup)
    formOnSaveHandler: function (executionContext) {
        var eventArgs = executionContext.getEventArgs();
        if (eventArgs.getSaveMode() == 70) {
            eventArgs.preventDefault();
            return;
        }
        var formContext = executionContext.getFormContext();
        let id = formContext.data.entity.getId();
        var dayIdRef = formContext.getAttribute("st08_day").getValue();
        var thisObject = this;

        if (id && (dayIdRef && dayIdRef.length > 0)) {
            let dayId = dayIdRef[0].id;

            setTimeout(() => {
                thisObject.allowSaveUpdate = false;
                Enavate.TimeTracking.Utils.calcRollupField(formContext, "st08_timetrackingdaies", dayId, "st08_totalhours");
                //Xrm.Utility.openEntityForm(entity, id);
                formContext.data.refresh(false);
            }, 500);
        }
    },

    //-
    setProjectLookupFilter: function (executionContext) {
        var formContext = executionContext.getFormContext();
        var selectedDate = formContext.getAttribute("st08_date").getValue();

        if (selectedDate) {
            let dateStr = Enavate.TimeTracking.Utils.formatDate(selectedDate);
            let filter = `<filter><filter type="or" ><condition attribute="st08_startdate" operator="le" value="${dateStr}" /><condition attribute="st08_startdate" operator="null" /></filter><filter type="or" ><condition attribute="st08_enddate" operator="null" /><condition attribute="st08_enddate" operator="ge" value="${dateStr}" /></filter></filter>`
            formContext.getControl("st08_project").addCustomFilter(filter);
        }
    }
}
