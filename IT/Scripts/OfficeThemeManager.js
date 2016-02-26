(function () {
    //Modify the native browser CSSStyleSheet object to add the events required
    var EventInfo = function (source) {
        var obj = {};
        obj.source = source;
        return obj;
    };

    var nativeInsertRule = CSSStyleSheet.prototype.insertRule;
    if (nativeInsertRule) {
        CSSStyleSheet.prototype.insertRule = function () {
            if (this.onRuleChange) {
                this.onRuleChange.call(this, new EventInfo("insertRule"));
            }
            nativeInsertRule.apply(this, arguments);
        };
    }

    var nativeAddRule = CSSStyleSheet.prototype.addRule;
    if (nativeAddRule) {
        CSSStyleSheet.prototype.addRule = function () {
            if (this.onRuleChange) {
                this.onRuleChange.call(this, new EventInfo("addRule"));
            }
            nativeAddRule.apply(this, arguments);
        };
    }


    var nativeDeleteRule = CSSStyleSheet.prototype.deleteRule;
    if (nativeDeleteRule) {
        CSSStyleSheet.prototype.deleteRule = function () {
            if (this.onRuleChange) {
                this.onRuleChange.call(this, new EventInfo("deleteRule"));
            }
            nativeDeleteRule.apply(this, arguments);
        };
    }
})();

var OfficeThemeManager = function OfficeThemeManager() {
    var self = this;
    var handlerCallback = null;
    var cssFileName = "officethemes.css";
    var officeThemeCss = null;
    var cssRuleChangeCount = 0;

    //Used to retrieve values out of OfficeThemes.css and convert them to a key/value map
    var themeCssMapping = [
                    { name: "primaryFontColor", cssSelector: ".office-docTheme-primary-fontColor", cssProperty: "color" },
                    { name: "primaryBackgroundColor", cssSelector: ".office-docTheme-primary-bgColor", cssProperty: "background-color" },
                    { name: "secondaryFontColor", cssSelector: ".office-docTheme-secondary-fontColor", cssProperty: "color" },
                    { name: "secondaryBackgroundColor", cssSelector: ".office-docTheme-secondary-bgColor", cssProperty: "background-color" },
                    { name: "accent1", cssSelector: ".office-contentAccent1-color", cssProperty: "color" },
                    { name: "accent2", cssSelector: ".office-contentAccent2-color", cssProperty: "color" },
                    { name: "accent3", cssSelector: ".office-contentAccent3-color", cssProperty: "color" },
                    { name: "accent4", cssSelector: ".office-contentAccent4-color", cssProperty: "color" },
                    { name: "accent5", cssSelector: ".office-contentAccent5-color", cssProperty: "color" },
                    { name: "accent6", cssSelector: ".office-contentAccent6-color", cssProperty: "color" },
                    { name: "hyperlink", cssSelector: ".office-a", cssProperty: "color" },
                    { name: "followedHyperlink", cssSelector: ".office-a:visited", cssProperty: "color" },
                    { name: "headerLatinFont", cssSelector: ".office-headerFont-latin", cssProperty: "font-family" },
                    { name: "headerEastAsianFont", cssSelector: ".office-headerFont-eastAsian", cssProperty: "font-family" },
                    { name: "headerScriptFont", cssSelector: ".office-headerFont-script", cssProperty: "font-family" },
                    { name: "headerLocalizedFont", cssSelector: ".office-headerFont-localized", cssProperty: "font-family" },
                    { name: "bodyLatinFont", cssSelector: ".office-bodyFont-latin", cssProperty: "font-family" },
                    { name: "bodyEastAsianFont", cssSelector: ".office-bodyFont-eastAsian", cssProperty: "font-family" },
                    { name: "bodyScriptFont", cssSelector: ".office-bodyFont-script", cssProperty: "font-family" },
                    { name: "bodyLocalizedFont", cssSelector: ".office-bodyFont-localized", cssProperty: "font-family" },
                    { name: "officePrimaryFontColor", cssSelector: ".office-officeTheme-primary-fontColor", cssProperty: "color" },
                    { name: "officePrimaryBackgroundColor", cssSelector: ".office-officeTheme-primary-bgColor", cssProperty: "background-color" },
                    { name: "officeSecondaryFontColor", cssSelector: ".office-officeTheme-secondary-fontColor", cssProperty: "color" },
                    { name: "officeSecondaryBackgroundColor", cssSelector: ".office-officeTheme-secondary-bgColor", cssProperty: "background-color" }
    ];

    var rgbToHex = function (rgb) {
        var results = rgb.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/i);
        if (results == null) {
            return rgb;
        }

        var red = parseInt(results[1]);
        var green = parseInt(results[2]);
        var blue = parseInt(results[3]);

        var rgbNumber = blue | (green << 8) | (red << 16);
        return "#" + (Number(rgbNumber) + 0x1000000).toString(16).slice(-6).toUpperCase();
    };

    var lookupCssStyle = function (officeCss, selector, style) {
        var styleMapping = {
            "font-family": "fontFamily",
            "border-color": "borderColor",
            "background-color": "backgroundColor",
            "color": "color"
        };
        var length = officeCss.cssRules ? officeCss.cssRules.length : officeCss.rules.length;
        for (var i = 0; i < length; i++) {
            var rule;
            if (officeCss.cssRules) {
                rule = officeCss.cssRules[i];
            } else {
                rule = officeCss.rules[i];
            }
            var ruleSelector = rule.selectorText;
            if (ruleSelector !== "" && ruleSelector.toLowerCase() == selector.toLowerCase()) {
                var styleValue = "";
                var property = styleMapping[style];
                if (officeCss.cssRules) {
                    styleValue = officeCss.cssRules[i].style[property];
                } else {
                    styleValue = officeCss.rules[i].style[property];
                }
                if (style === "color" || style === "border-color" || style === "background-color") {
                    styleValue = rgbToHex(styleValue);
                }
                return styleValue;
            }
        }

        return null;
    };

    var constructThemeTable = function (stylesheet) {

        var obj = {};
        for (var i = 0; i < themeCssMapping.length; i++) {
            var themeMapping = themeCssMapping[i];
            var themeFormattedValue = lookupCssStyle(stylesheet, themeMapping.cssSelector, themeMapping.cssProperty);
            if (themeFormattedValue != null) {
                obj[themeMapping.name] = themeFormattedValue;
            }
        }
        return obj;
    };

    var fireThemeChangeEvent = function () {
        //Fires the theme change callback to the client
        var obj = {
            theme: constructThemeTable(officeThemeCss)
        }
        handlerCallback(obj);
    };

    var stylesheetChangeHandler = function (event) {
        cssRuleChangeCount++;

        //Wait for stylesheet to stop modifying rules to fire off an event to the client
        //Office.js calls addRule many times during a single themChange event.  We need to wait for operations to stop
        if (cssRuleChangeCount === 1) {
            var stylesheetChangeEndDetection = function () {
                var currentCount = cssRuleChangeCount;
                //If for some reason an addRule call takes more than 5 ms, it is not an issue to the app.  It will receieve 2 event notifications
                //instead of a single notification in this case.  We are only waiting for performance reasons if this is removed and fireThemeChangeEvent() is called
                //the app will still work, but will receive many events.
                setTimeout(function () {
                    if (currentCount === cssRuleChangeCount) {
                        cssRuleChangeCount = 0;
                        fireThemeChangeEvent();
                    }
                    else {
                        stylesheetChangeEndDetection();
                    }
                }, 5);
            }
            stylesheetChangeEndDetection();
        }
    };

    for (var i = 0; i < document.styleSheets.length; i++) {
        var ss = document.styleSheets[i];
        if (ss.href && (cssFileName == (ss.href.substring(ss.href.length - cssFileName.length, ss.href.length)).toLowerCase())) {
            officeThemeCss = ss;
            //Register to the CSSStyleSheet onRuleChange handler.  This will occur any time addRule, insertRule, or deleteRule is called on the object.
            officeThemeCss.onRuleChange = stylesheetChangeHandler;
            break;
        }
    };


    /// <summary>
    /// Register to an event that gets fired whenever the theme changes of the host application.  Fires instantly after registration to account for the initial theme state.
    /// Object is returned as eventData.theme: {"primaryFontColor":"#000000","primaryBackgroundColor":"#2FA3EE", etc... }
    /// </summary>
    self.onThemeChange = function (handler) {
        handlerCallback = handler;
        fireThemeChangeEvent(handlerCallback);
    };
};