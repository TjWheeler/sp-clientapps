module ClientAppConfig {
    /* The ClientAppConfig module is a reusable config component.
     * The app registers its custom properties.  
     * The config component saves to the Client App Config list.
     * The config component outputs the html elements to manage saving the lists.
     * The showConfigForm should only be called when inDesignMode() returns true.
     * Dependencies: CamlBuilder.ts.
     */
    class ConfigListFields {
        public static readonly PropertyName = "Title";
        public static readonly AppName = "AppName";
        public static readonly AppInstance = "AppInstance";
        public static readonly PropertyValue = "PropertyValue";
        public static readonly PropertyDescription = "PropertyDescription";
    }
    export class ConfigItem {
        id: number;
        appName: string;
        appInstance: string;
        propertyName: string;
        propertyValue: string;
        propertyDescription: string;
    }
    export class ResourceLoader {
        public static loadCss = (name, url) => {
            var cssId = name;
            if (!document.getElementById(cssId)) {
                var path = url;
                var head = document.getElementsByTagName('head')[0];
                var link = document.createElement('link') as any;
                link.id = cssId;
                link.rel = 'stylesheet';
                link.type = 'text/css';
                link.href = path;
                link.media = 'all';
                head.appendChild(link);
            }
        };
    }
    export class Config {
        constructor() {
            this.appProperties = new Array();
        }
        public static readonly ConfigListName = "Client App Config";
        public appProperties: Array<ConfigItem>;
        public registerAppProperty(appName: string, appInstance: string, propertyName: string, propertyValue: string, propertyDescription: string) {
            var item = new ConfigItem();
            item.id = 0;
            item.appName = appName;
            item.appInstance = appInstance;
            item.propertyName = propertyName;
            item.propertyValue = propertyValue;
            item.propertyDescription = propertyDescription;
            this.appProperties.push(item);
        }
        private executeQuery = (clientContext: SP.ClientContext) => {
            var deferred = $.Deferred();
            clientContext.executeQueryAsync(
                (sender, args) => {
                    deferred.resolve(sender, args);
                },
                (sender, args) => {
                    deferred.reject(sender, args);
                }
            );
            return deferred.promise();
        };
        public load(appName: string, appInstance: string): JQueryPromise<any> {
            var deferred = $.Deferred();
            this.get(appName, appInstance).done((items) => {
                for (var i = 0; i < items.length; i++) {
                    var existing = this.getProperty(items[i].propertyName);
                    if (existing) {
                        existing.id = items[i].id;
                        existing.appInstance = items[i].appInstance;
                        existing.appName = items[i].appName;
                        existing.propertyValue = items[i].propertyValue;
                        existing.propertyName = items[i].propertyName;
                    }
                    else {
                        this.appProperties.push(items[i]);
                    }
                }
                deferred.resolve();
            })
                .fail((sender, args) => {
                    deferred.reject(sender, args);
                });
            return deferred.promise();
        }
        public save(appName: string, appInstance: string): JQueryPromise<any> {
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
                .done((sender: any, args: any) => {
                    this.load(appName, appInstance).done(() => {
                        deferred.resolve();//reload to ensure ID's are set
                    }).fail((sender: any, args: any) => {
                        deferred.reject(sender, args);
                    });
                })
                .fail((sender: any, args: any) => {
                    deferred.reject(sender, args);
                });
            return deferred.promise();
        }
        public static inDesignMode = () => {
            var result = ((<any>window).MSOWebPartPageFormName != undefined) && ((document.forms[(<any>window).MSOWebPartPageFormName] && (<any>document.forms[(<any>window).MSOWebPartPageFormName]).MSOLayout_InDesignMode && ("1" == (<any>document.forms[(<any>window).MSOWebPartPageFormName]).MSOLayout_InDesignMode.value)) || (document.forms[(<any>window).MSOWebPartPageFormName] && (<any>document.forms[(<any>window).MSOWebPartPageFormName])._wikiPageMode && ("Edit" == (<any>document.forms[(<any>window).MSOWebPartPageFormName])._wikiPageMode.value)));
            return result || false;
        };
        public static submitUpdates(event, appName: string, appInstance: string, configElementId: string) {
            //called my the click of the update button
            event.preventDefault();
            var getFeedbackTemplate = (colour: string, message: string) => {
                return `<div class="config-feedback">
                    <label style="color:${colour}">${message}</label>
                </div>`;
            };
            $("#" + configElementId + " .config-feedback").remove();
            console.log("AppName:" + appName);
            console.log("AppInstance:" + appInstance);
            console.log("ConfigElementId:" + configElementId);
            var config = new Config();
            config.load(appName, appInstance).done(() => {
                $("#" + configElementId + " .config-value input").each((index: number, element: Element) => {
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
                config.save(appName, appInstance).done(() => {
                    $("#" + configElementId).append(getFeedbackTemplate("green", "Properties updated successfully."));
                }).fail((sender, args) => {
                    $("#" + configElementId).append(getFeedbackTemplate("red", "Error: Failed to update the Client Apps Config list - " + args.get_message()));
                });
            }).fail((sender, args) => {
                $("#" + configElementId).append(getFeedbackTemplate("red", "Error: Failed to update the Client Apps Config list - " + args.get_message()));
            });
        }
        public showConfigForm(appName: string, appInstance: string, configElement: Element, settingsIconUrl: string) {
            if (!($(configElement).attr("id"))) {
                $(configElement).attr("id", Math.floor(Math.random() * 26) + Date.now());
            }
            var index = 0;
            var getPropertyTemplate = (property: ClientAppConfig.ConfigItem) => {
                index++;
                return `
                    <dt>${property.propertyName}</dt>
                    <dd class="config-value"><input name="id${property.propertyName}" value="${property.propertyValue != null ? property.propertyValue : ""}">
                    <span>${property.propertyDescription}</span>
                    </dd>
                `;
            };
            var getSubmitTemplate = () => {
                return `<div class="config-submit">
                    <input type="button" value="Update" onclick="ClientAppConfig.Config.submitUpdates(event, '${appName}','${appInstance}','${$(configElement).attr('id')}')"}">
                </div>`;
            };
            var nodes = new Array();
            var displayState = "display:none;";
            nodes.push("<img class='config-icon' src='" + settingsIconUrl + "' style='cursor:pointer' onclick='$(this).next().first().slideToggle()'><div class='config-wp-panel' style='" + displayState + "'>");
            nodes.push("<h3>" + appName + " - Settings</h3>");//close the dl
            nodes.push("<dl>");//open the dl
            for (var i = 0; i < this.appProperties.length; i++) {
                nodes.push(getPropertyTemplate(this.appProperties[i]));
            }
            nodes.push("</dl>");//close the dl
            nodes.push(getSubmitTemplate());//update button
            nodes.push("</div>");//close the panel
            $(configElement).empty().append(nodes.join(' '), );
        };
        public getProperty(propertyName: string) {
            if (this.appProperties) {
                for (var i = 0; i < this.appProperties.length; i++) {
                    if (this.appProperties[i].propertyName === propertyName) {
                        return this.appProperties[i];
                    }
                }
            }
            return null;
        };
        public ensureSaved(appName: string, appInstance: string): JQueryPromise<any> {
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
        public isConfigured(): boolean {
            if (!this.appProperties || this.appProperties.length === 0) {
                return false;
            }
            for (var i = 0; i < this.appProperties.length; i++) {
                if (!this.appProperties[i].id) {
                    return false;
                }
            }
            return true;
        }
        public getPropertyValue(propertyName: string): string {
            var property = this.getProperty(propertyName);
            if (property) {
                return property.propertyValue;
            }
            return null;
        };
        public getPropertyValueBoolean(propertyName: string): boolean {
            var property = this.getProperty(propertyName);
            if (property && property.propertyValue) {
                if (property.propertyValue.toLowerCase() === "true") {
                    return true;
                }
                return false;
            }
            return null;
        };
        public getPropertyValueNumber(propertyName: string): number {
            var property = this.getProperty(propertyName);
            if (property && !(isNaN(Number(property.propertyValue)))) {
                return Number(property.propertyValue);
            }
            return null;
        };

        private get(appName: string, appInstance: string): JQueryPromise<any> {
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
                .done((sender: any, args: any) => {
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
                        var registeredProp = this.getProperty(item.propertyName);
                        if (registeredProp) {
                            this.appProperties.splice(this.appProperties.indexOf(registeredProp), 1, item);
                        }
                    }
                    deferred.resolve(items);
                })
                .fail((sender: any, args: any) => {
                    console.log("Error: " + args.get_message());
                    deferred.reject(sender, args);
                });
            return deferred.promise();
        }
    }

}