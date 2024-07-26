sap.ui.define([
    "sap/ui/core/Fragment",
    "sap/m/MessageToast", "xlsx",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator",
    "sap/ui/model/json/JSONModel",
],
function (Fragment, MessageToast, XLSX, Filter, FilterOperator, JSONModel) {
    "use strict";
    return {

        excelSheetsData: [],
        pDialog: null,

        onInit: function () {
            var oDataModel = this.getOwnerComponent().getModel("mainModel");
            this.getView().setModel(oDataModel, "odataModel");
        },

        openExcelUploadDialog: function (oEvent) {
            console.log(XLSX.version);
            this.excelSheetsData = [];
            var oView = this.getView();
            if (!this.pDialog) {
                Fragment.load({
                    id: "excel_upload",
                    name: "project1.ext.fragment.ExcelUpload",
                    type: "XML",
                    controller: this
                }).then((oDialog) => {
                    var oFileUploader = Fragment.byId("excel_upload", "uploadSet");
                    oFileUploader.removeAllItems();
                    this.pDialog = oDialog;
                    this.pDialog.open();
                })
                    .catch(error => alert(error.message));
            } else {
                var oFileUploader = Fragment.byId("excel_upload", "uploadSet");
                oFileUploader.removeAllItems();
                this.pDialog.open();
            }
        },

        checkForExistingRecords: function (oEvent) {
            var oView = this.getView();
            var oDataResultModel = oView.getModel("ODataResultModel");
            var formattedDate = sap.ui.core.format.DateFormat.getDateInstance({ pattern: "dd.MM.yyyy" }).format(new Date());

            if (oDataResultModel) {
                var aResults = oDataResultModel.getData();
                if (aResults && aResults.length > 0) {
                    sap.m.MessageToast.show(formattedDate + " tarihinde zaten kayıt yapılmış. Tekrar kayıt yapamazsın.");
                    return;
                } else {
                    this.uploadExcelData(oEvent); 
                }
            }
        },

        oncondition: function () {
            var currentDate = new Date();
            var dateFormat = sap.ui.core.format.DateFormat.getDateInstance({ pattern: "dd.MM.yyyy" });
            var formattedDate = dateFormat.format(currentDate);

            var oModel = this.getView().getModel("odataModel");
            var oView = this.getView();

            oModel.read("/Building", {
                filters: [
                    new Filter({
                        path: "created_at",
                        operator: FilterOperator.EQ,
                        value1: formattedDate
                    })
                ],
                success: function (oData) {
                    console.log("OData isteği başarılı:", oData);

                    // Gelen veriyi bir JSONModel'de tutma
                    var oDataResultModel = new JSONModel();
                    oDataResultModel.setData(oData.results);
                    oView.setModel(oDataResultModel, "ODataResultModel");

                    // Koşul kontrol fonksiyonunu çağır
                    this.checkForExistingRecords();
                }.bind(this),
                error: function () {
                    console.log("Liste Alınmadı");
                }
            });

        },
        onUploadSet: function (oEvent) {

            this.oncondition();

        },

        uploadExcelData: function (oEvent) {

            if (!this.excelSheetsData.length) {
                MessageToast.show("Select file to Upload");
                return;
            }

            var that = this;
            var oSource = oEvent ? oEvent.getSource() : { getText: () => "Your custom text" }; 

            var fnAddMessage = function () {
                return new Promise((fnResolve, fnReject) => {
                    that.callOdata(fnResolve, fnReject);
                });
            };

            var mParameters = {
                sActionLabel: oSource.getText()
            };

            this.extensionAPI.securedExecution(fnAddMessage, mParameters);

            this.pDialog.close();
        },


        onTempDownload: function (oEvent) {
            var oModel = this.getOwnerComponent().getModel();
            var oBuilding = oModel.getServiceMetadata().dataServices.schema[0].entityType.find(x => x.name === 'BuildingType');
            var propertyList = ['building_name', 'n_rooms', 'address_line',
                'city', 'state', 'country'];

            var excelColumnList = [];
            var colList = {};

            propertyList.forEach((value, index) => {
                let property = oBuilding.property.find(x => x.name === value);
                colList[property.extensions.find(x => x.name === 'label').value] = '';
            });
            excelColumnList.push(colList);

            const ws = XLSX.utils.json_to_sheet(excelColumnList);
            const wb = XLSX.utils.book_new();

            XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
            XLSX.writeFile(wb, 'RAP - Building.xlsx');

            MessageToast.show("Template File Downloading...");
        },
        onCloseDialog: function (oEvent) {
            this.pDialog.close();
        },
        onBeforeUploadStart: function (oEvent) {

        },
        onUploadSetComplete: function (oEvent) {

            var oFileUploader = Fragment.byId("excel_upload", "uploadSet");
            var oFile = oFileUploader.getItems()[0].getFileObject();

            var reader = new FileReader();
            var that = this;

            reader.onload = (e) => {
                let xlsx_content = e.currentTarget.result;

                let workbook = XLSX.read(xlsx_content, { type: 'binary' });
                var excelData = XLSX.utils.sheet_to_row_object_array(workbook.Sheets["Sheet1"]);

                workbook.SheetNames.forEach(function (sheetName) {

                    that.excelSheetsData.push(XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]));
                });
                console.log("Excel Data", excelData);
                console.log("Excel Sheets Data", this.excelSheetsData);
            };
            reader.readAsBinaryString(oFile);

            MessageToast.show("Upload Successful");
        },
        onItemRemoved: function (oEvent) {
            this.excelSheetsData = [];
        },

        onBuildRead: function (oEvent) {
        },
        callOdata: function (fnResolve, fnReject) {

            var oModel = this.getView().getModel();
            var payload = {};

            this.excelSheetsData[0].forEach((value, index) => {
                payload = {
                    "building_name": value["Building Name"],
                    "n_rooms": value["No of Rooms"],
                    "address_line": value["Adress Line"],
                    "city": value["City"],
                    "state": value["State"],
                    "country": value["Country"]
                };
                oModel.create("/Building", payload, {
                    success: (result) => {
                        console.log(result);
                        var oMessageManager = sap.ui.getCore().getMessageManager();
                        var oMessage = new sap.ui.core.message.Message({
                            message: "İşlem başarıyla gerçekleşmiştir.",
                            persistent: true,
                            type: sap.ui.core.MessageType.Success
                        });
                        oMessageManager.addMessages(oMessage);
                        fnResolve();
                        location.reload();
                    },
                    error: fnReject
                })
            });
        }
    };
});
