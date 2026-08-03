"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var getPublic_1 = require("./getPublic");
var WorkbenchPage_1 = require("./WorkbenchPage");
var workbenchPage = new WorkbenchPage_1.WorkbenchPage();
// tslint:disable-next-line:export-name
exports.default = {
    '/workbench': workbenchPage.handleWorkbenchRequest,
    '*/*': getPublic_1.default
};

//# sourceMappingURL=index.js.map
