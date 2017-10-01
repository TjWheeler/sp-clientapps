var ClientAppConfig;
(function (ClientAppConfig) {
    /* The ClientAppConfig module is a reusable config component.
     * The app registers its custom properties.
     * The config component saves to the Client App Config list.
     * The config component outputs the html elements to manage saving the lists.
     * The showConfigForm should only be called when inDesignMode() returns true.
     * Dependencies: CamlBuilder.ts.
     */
    var ConfigListFields = (function () {
        function ConfigListFields() {
        }
        return ConfigListFields;
    }());
    ConfigListFields.PropertyName = "Title";
    ConfigListFields.AppName = "AppName";
    ConfigListFields.AppInstance = "AppInstance";
    ConfigListFields.PropertyValue = "PropertyValue";
    ConfigListFields.PropertyDescription = "PropertyDescription";
    var ConfigItem = (function () {
        function ConfigItem() {
        }
        return ConfigItem;
    }());
    ClientAppConfig.ConfigItem = ConfigItem;
    var ResourceLoader = (function () {
        function ResourceLoader() {
        }
        return ResourceLoader;
    }());
    ResourceLoader.loadCss = function (name, url) {
        var cssId = name;
        if (!document.getElementById(cssId)) {
            var path = url;
            var head = document.getElementsByTagName('head')[0];
            var link = document.createElement('link');
            link.id = cssId;
            link.rel = 'stylesheet';
            link.type = 'text/css';
            link.href = path;
            link.media = 'all';
            head.appendChild(link);
        }
    };
    ClientAppConfig.ResourceLoader = ResourceLoader;
    var Config = (function () {
        function Config() {
            this.executeQuery = function (clientContext) {
                var deferred = $.Deferred();
                clientContext.executeQueryAsync(function (sender, args) {
                    deferred.resolve(sender, args);
                }, function (sender, args) {
                    deferred.reject(sender, args);
                });
                return deferred.promise();
            };
            this.appProperties = new Array();
        }
        Config.prototype.registerAppProperty = function (appName, appInstance, propertyName, propertyValue, propertyDescription) {
            var item = new ConfigItem();
            item.id = 0;
            item.appName = appName;
            item.appInstance = appInstance;
            item.propertyName = propertyName;
            item.propertyValue = propertyValue;
            item.propertyDescription = propertyDescription;
            this.appProperties.push(item);
        };
        Config.prototype.load = function (appName, appInstance) {
            var _this = this;
            var deferred = $.Deferred();
            this.get(appName, appInstance).done(function (items) {
                for (var i = 0; i < items.length; i++) {
                    var existing = _this.getProperty(items[i].propertyName);
                    if (existing) {
                        existing.id = items[i].id;
                        existing.appInstance = items[i].appInstance;
                        existing.appName = items[i].appName;
                        existing.propertyValue = items[i].propertyValue;
                        existing.propertyName = items[i].propertyName;
                    }
                    else {
                        _this.appProperties.push(items[i]);
                    }
                }
                deferred.resolve();
            })
                .fail(function (sender, args) {
                deferred.reject(sender, args);
            });
            return deferred.promise();
        };
        Config.prototype.save = function (appName, appInstance) {
            var _this = this;
            var deferred = $.Deferred();
            var clientContext = SP.ClientContext.get_current();
            var list = clientContext.get_site().get_rootWeb().get_lists().getByTitle(Config.ConfigListName);
            for (var i = 0; i < this.appProperties.length; i++) {
                var prop = this.appProperties[i];
                var listItem = null;
                if (prop.id) {
                    listItem = list.getItemById(prop.id);
                }
                else {
                    var itemCreateInfo = new SP.ListItemCreationInformation();
                    listItem = list.addItem(itemCreateInfo);
                    listItem.set_item(ConfigListFields.PropertyDescription, prop.propertyDescription);
                }
                listItem.set_item(ConfigListFields.PropertyName, prop.propertyName);
                listItem.set_item(ConfigListFields.AppInstance, prop.appInstance);
                listItem.set_item(ConfigListFields.AppName, prop.appName);
                listItem.set_item(ConfigListFields.PropertyValue, prop.propertyValue);
                listItem.update();
            }
            this.executeQuery(clientContext)
                .done(function (sender, args) {
                _this.load(appName, appInstance).done(function () {
                    deferred.resolve(); //reload to ensure ID's are set
                }).fail(function (sender, args) {
                    deferred.reject(sender, args);
                });
            })
                .fail(function (sender, args) {
                deferred.reject(sender, args);
            });
            return deferred.promise();
        };
        Config.submitUpdates = function (event, appName, appInstance, configElementId) {
            //called my the click of the update button
            event.preventDefault();
            var getFeedbackTemplate = function (colour, message) {
                return "<div class=\"config-feedback\">\n                    <label style=\"color:" + colour + "\">" + message + "</label>\n                </div>";
            };
            $("#" + configElementId + " .config-feedback").remove();
            console.log("AppName:" + appName);
            console.log("AppInstance:" + appInstance);
            console.log("ConfigElementId:" + configElementId);
            var config = new Config();
            config.load(appName, appInstance).done(function () {
                $("#" + configElementId + " .config-value input").each(function (index, element) {
                    var propertyName = $(element).attr("name");
                    var propertyValue = $(element).val();
                    console.log(propertyName + ":" + propertyValue);
                    var property = config.getProperty(propertyName.substr(2));
                    if (!property) {
                        console.log("Warning: The property could not be retrieved.");
                    }
                    else {
                        property.propertyValue = propertyValue;
                    }
                });
                config.save(appName, appInstance).done(function () {
                    $("#" + configElementId).append(getFeedbackTemplate("green", "Properties updated successfully."));
                }).fail(function (sender, args) {
                    $("#" + configElementId).append(getFeedbackTemplate("red", "Error: Failed to update the Client Apps Config list - " + args.get_message()));
                });
            }).fail(function (sender, args) {
                $("#" + configElementId).append(getFeedbackTemplate("red", "Error: Failed to update the Client Apps Config list - " + args.get_message()));
            });
        };
        Config.prototype.showConfigForm = function (appName, appInstance, configElement, settingsIconUrl) {
            if (!($(configElement).attr("id"))) {
                $(configElement).attr("id", Math.floor(Math.random() * 26) + Date.now());
            }
            var index = 0;
            var getPropertyTemplate = function (property) {
                index++;
                return "\n                    <dt>" + property.propertyName + "</dt>\n                    <dd class=\"config-value\"><input name=\"id" + property.propertyName + "\" value=\"" + (property.propertyValue != null ? property.propertyValue : "") + "\">\n                    <span>" + property.propertyDescription + "</span>\n                    </dd>\n                ";
            };
            var getSubmitTemplate = function () {
                return "<div class=\"config-submit\">\n                    <input type=\"button\" value=\"Update\" onclick=\"ClientAppConfig.Config.submitUpdates(event, '" + appName + "','" + appInstance + "','" + $(configElement).attr('id') + "')\"}\">\n                </div>";
            };
            var nodes = new Array();
            var displayState = "display:none;";
            nodes.push("<img class='config-icon' src='" + settingsIconUrl + "' style='cursor:pointer' onclick='$(this).next().first().slideToggle()'><div class='config-wp-panel' style='" + displayState + "'>");
            nodes.push("<h3>" + appName + " - Settings</h3>"); //close the dl
            nodes.push("<dl>"); //open the dl
            for (var i = 0; i < this.appProperties.length; i++) {
                nodes.push(getPropertyTemplate(this.appProperties[i]));
            }
            nodes.push("</dl>"); //close the dl
            nodes.push(getSubmitTemplate()); //update button
            nodes.push("</div>"); //close the panel
            $(configElement).empty().append(nodes.join(' '));
        };
        ;
        Config.prototype.getProperty = function (propertyName) {
            if (this.appProperties) {
                for (var i = 0; i < this.appProperties.length; i++) {
                    if (this.appProperties[i].propertyName === propertyName) {
                        return this.appProperties[i];
                    }
                }
            }
            return null;
        };
        ;
        Config.prototype.ensureSaved = function (appName, appInstance) {
            var deferred = $.Deferred();
            var requiresUpdate = false;
            for (var i = 0; i < this.appProperties.length; i++) {
                if (!this.appProperties[i].id) {
                    requiresUpdate = true;
                    break;
                }
            }
            if (requiresUpdate) {
                this.save(appName, appInstance).done(deferred.resolve)
                    .fail(deferred.reject);
            }
            else {
                deferred.resolve();
            }
            return deferred.promise();
        };
        ;
        Config.prototype.isConfigured = function () {
            if (!this.appProperties || this.appProperties.length === 0) {
                return false;
            }
            for (var i = 0; i < this.appProperties.length; i++) {
                if (!this.appProperties[i].id) {
                    return false;
                }
            }
            return true;
        };
        Config.prototype.getPropertyValue = function (propertyName) {
            var property = this.getProperty(propertyName);
            if (property) {
                return property.propertyValue;
            }
            return null;
        };
        ;
        Config.prototype.getPropertyValueBoolean = function (propertyName) {
            var property = this.getProperty(propertyName);
            if (property && property.propertyValue) {
                if (property.propertyValue.toLowerCase() === "true") {
                    return true;
                }
                return false;
            }
            return null;
        };
        ;
        Config.prototype.getPropertyValueNumber = function (propertyName) {
            var property = this.getProperty(propertyName);
            if (property && !(isNaN(Number(property.propertyValue)))) {
                return Number(property.propertyValue);
            }
            return null;
        };
        ;
        Config.prototype.get = function (appName, appInstance) {
            var _this = this;
            var deferred = $.Deferred();
            var camlBuilder = new SpData.CamlBuilder();
            camlBuilder.begin(true); //Use an AND query
            camlBuilder.addViewFields([ConfigListFields.AppName, ConfigListFields.AppInstance, ConfigListFields.PropertyName, ConfigListFields.PropertyValue, ConfigListFields.PropertyDescription]);
            camlBuilder.addTextClause(SpData.CamlOperator.Eq, ConfigListFields.AppName, appName);
            camlBuilder.addTextClause(SpData.CamlOperator.Eq, ConfigListFields.AppInstance, appInstance);
            var clientContext = SP.ClientContext.get_current();
            var list = clientContext.get_site().get_rootWeb().get_lists().getByTitle(Config.ConfigListName);
            var query = new SP.CamlQuery();
            query.set_viewXml(camlBuilder.viewXml);
            var listItems = list.getItems(query);
            clientContext.load(listItems);
            this.executeQuery(clientContext)
                .done(function (sender, args) {
                var items = new Array();
                var iterator = listItems.getEnumerator();
                while (iterator.moveNext()) {
                    var listItem = iterator.get_current();
                    var item = new ConfigItem();
                    item.propertyName = listItem.get_item(ConfigListFields.PropertyName);
                    item.propertyValue = listItem.get_item(ConfigListFields.PropertyValue);
                    item.appName = listItem.get_item(ConfigListFields.AppName);
                    item.appInstance = listItem.get_item(ConfigListFields.AppInstance);
                    item.propertyDescription = listItem.get_item(ConfigListFields.PropertyDescription);
                    item.id = listItem.get_id();
                    items.push(item);
                    var registeredProp = _this.getProperty(item.propertyName);
                    if (registeredProp) {
                        _this.appProperties.splice(_this.appProperties.indexOf(registeredProp), 1, item);
                    }
                }
                deferred.resolve(items);
            })
                .fail(function (sender, args) {
                console.log("Error: " + args.get_message());
                deferred.reject(sender, args);
            });
            return deferred.promise();
        };
        return Config;
    }());
    Config.ConfigListName = "Client App Config";
    Config.inDesignMode = function () {
        var result = (window.MSOWebPartPageFormName != undefined) && ((document.forms[window.MSOWebPartPageFormName] && document.forms[window.MSOWebPartPageFormName].MSOLayout_InDesignMode && ("1" == document.forms[window.MSOWebPartPageFormName].MSOLayout_InDesignMode.value)) || (document.forms[window.MSOWebPartPageFormName] && document.forms[window.MSOWebPartPageFormName]._wikiPageMode && ("Edit" == document.forms[window.MSOWebPartPageFormName]._wikiPageMode.value)));
        return result || false;
    };
    ClientAppConfig.Config = Config;
})(ClientAppConfig || (ClientAppConfig = {}));
//# sourceMappingURL=ClientAppConfig.js.map