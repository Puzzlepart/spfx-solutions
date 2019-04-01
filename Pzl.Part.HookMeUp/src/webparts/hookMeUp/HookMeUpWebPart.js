"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_webpart_base_1 = require("@microsoft/sp-webpart-base");
var PropertyPaneLogo_1 = require("./PropertyPaneLogo");
var sp_core_library_2 = require("@microsoft/sp-core-library");
var strings = require("HookMeUpWebPartStrings");
var HookMeUpWebPart = (function (_super) {
    __extends(HookMeUpWebPart, _super);
    function HookMeUpWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    HookMeUpWebPart.prototype.render = function () {
        if (this.displayMode === sp_core_library_2.DisplayMode.Edit) {
            this.domElement.innerHTML = "\n                <div id=\"" + this.properties.anchorId + "\" style=\"margin-bottom:50px\">\n                    <b>Anchor:</b> " + this.properties.anchorId + " <i>(" + strings.NotVisible + ")</i>\n                </div>\n            ";
        }
        else {
            var element = this.domElement.parentElement;
            // check up to 5 levels up for padding
            for (var i = 0; i < 5; i++) {
                var style = window.getComputedStyle(element);
                var hasPadding = style.paddingTop !== '0px';
                if (hasPadding) {
                    element.style.paddingTop = '0';
                    element.style.paddingBottom = '0';
                    element.style.marginTop = '0';
                    element.style.marginBottom = '0';
                }
                if (element.className === 'ControlZone') {
                    break;
                }
                element = element.parentElement;
            }
            this.domElement.innerHTML = this.domElement.innerHTML = "<span id=\"" + this.properties.anchorId + "\"></span>";
            if (document.location.hash === '#' + this.properties.anchorId) {
                this.domElement.scrollIntoView();
            }
        }
    };
    Object.defineProperty(HookMeUpWebPart.prototype, "dataVersion", {
        get: function () {
            return sp_core_library_1.Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    HookMeUpWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    groups: [
                        {
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneTextField('anchorId', {
                                    label: strings.AnchorId,
                                    description: strings.Description,
                                    value: this.properties.anchorId
                                }),
                                new PropertyPaneLogo_1.default()
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return HookMeUpWebPart;
}(sp_webpart_base_1.BaseClientSideWebPart));
exports.default = HookMeUpWebPart;
//# sourceMappingURL=HookMeUpWebPart.js.map