"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var sp_webpart_base_1 = require("@microsoft/sp-webpart-base");
var PropertyPaneLogo = (function () {
    function PropertyPaneLogo() {
        this.type = sp_webpart_base_1.PropertyPaneFieldType.Custom;
        this.properties = {
            key: 'Logo',
            onRender: this.onRender.bind(this)
        };
    }
    PropertyPaneLogo.prototype.onRender = function (elem) {
        elem.innerHTML = "\n    <div style=\"margin-top: 30px\">\n      <div style=\"float:right\">Author: <a href=\"https://twitter.com/mikaelsvenson\" tabindex=\"-1\" target=\"_blank\">Mikael Svenson</a></div>\n      <div style=\"float:right\"><a href=\"https://www.puzzlepart.com/\" target=\"_blank\"><img src=\"//www.puzzlepart.com/wp-content/uploads/2017/08/Pzl-LogoType-200.png\" onerror=\"this.style.display = 'none'\";\"></a></div>\n    </div>";
    };
    return PropertyPaneLogo;
}());
exports.PropertyPaneLogo = PropertyPaneLogo;
exports.default = PropertyPaneLogo;
//# sourceMappingURL=PropertyPaneLogo.js.map