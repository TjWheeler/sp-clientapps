var FacebookClientApp;
(function (FacebookClientApp) {
    var CustomProperties = (function () {
        function CustomProperties() {
        }
        return CustomProperties;
    }());
    //Reuse: Update these Property Names
    CustomProperties.AccessToken = "Access Token";
    CustomProperties.RefreshSeconds = "Refresh Seconds";
    CustomProperties.APIVersion = "API Version";
    CustomProperties.PageId = "Page Id";
    CustomProperties.PageSize = "Page Size";
    CustomProperties.MoreLinkText = "Read More Link Text";
    CustomProperties.MoreLinkUrl = "Read More Link Url";
    CustomProperties.RollupChars = "Rollup Characters";
    CustomProperties.DefaultIcon = "Default Icon";
    CustomProperties.Title = "Title";
    var Resources = (function () {
        function Resources() {
        }
        return Resources;
    }());
    Resources.load = function () {
        var loadCss = function (name, url) {
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
        var baseUrl = _spPageContextInfo.siteServerRelativeUrl;
        if (baseUrl.substring(baseUrl.length - 1) === "/") {
            baseUrl = baseUrl.substring(0, baseUrl.length - 1);
        }
        loadCss('facebookappcss', baseUrl + "/Style Library/Client Apps/Facebook/css/facebook.css");
        loadCss('clientappconfig', baseUrl + "/Style Library/Client Apps/Common/css/clientapp-config.css");
        loadCss('font-awesome.min.css', baseUrl + "/Style Library/Client Apps/Common/libs/font-awesome-4.7.0/css/font-awesome.min.css");
        loadCss('bootstrap-grid.css', baseUrl + "/Style Library/Client Apps/Common/libs/bootstrap-grid-3.3.7/bootstrap-grid.min.css");
    };
    FacebookClientApp.Resources = Resources;
    var AppProperties = (function () {
        function AppProperties() {
            this.AppName = "Facebook";
            this.AppInstance = _spPageContextInfo.webServerRelativeUrl + "/" + _spPageContextInfo.pageItemId;
            this.RollupChars = 100;
        }
        return AppProperties;
    }());
    var Facebook = (function () {
        function Facebook() {
            var _this = this;
            this.renderVideo = function (item) {
                return _this.renderLink(item); //compatible
            };
            this.renderPhoto = function (item) {
                return _this.renderLink(item); //compatible
            };
            this.renderIcon = function (item) {
                if (item.type && (item.type === "photo" || item.type === "video" || item.type === "story" || item.type === "link")) {
                    var imageUrl = (_spPageContextInfo.siteServerRelativeUrl.length > 1 ? _spPageContextInfo.siteServerRelativeUrl : "") + "/Style Library/Client Apps/Facebook/img/" + item.type + ".png";
                    return "<img class=\"facebook-item-icon\" src=\"" + imageUrl + "\" alt=\"" + item.type + "\">";
                }
                return "";
            };
            this.renderDescription = function (item) {
                if (!item.description && item.message) {
                    item.description = item.message;
                }
                if (!item.description && item.story) {
                    item.description = item.story;
                }
                var rollup = "";
                var description = "";
                if (item.description && item.description.length > _this.appProperties.RollupChars) {
                    var lastSpace = item.description.substr(0, _this.appProperties.RollupChars).lastIndexOf(" ");
                    rollup = item.description.substr(0, lastSpace) + "...";
                    if (item.description) {
                        description = "<div class=\"facebook-item-description\">\n                                    <span class=\"facebook-item-description-rollup\">" + rollup + "\n                                        <span class=\"clickable read-more\" onclick=\"$(this).parent().toggle();$(this).parent().next().toggle();\"> read more</span>\n                                    </span>\n                                    <span class=\"facebook-item-description\" style=\"display:none\">" + FacebookParsers.parseHashtag(item.description) + "\n                                        <span class=\"clickable read-less\" onclick=\"$(this).parent().toggle();$(this).parent().prev().toggle();\"> (read less)</span>\n                                    </span>\n                               </div>";
                    }
                }
                else {
                    description = "<div class=\"facebook-item-description\">\n                                    <span class=\"facebook-item-description\">" + item.description + "</span>\n                               </div>";
                }
                return description;
            };
            this.renderLikes = function (item) {
                var likeDetails = "";
                if (item.likes.data.length) {
                    var likes = [];
                    for (var i = 0; i < item.likes.data.length; i++) {
                        likes.push(item.likes.data[i].name);
                    }
                    likeDetails = "Liked by " + likes.join(", ");
                    if (item.likes.data.length < item.likes.summary.total_count) {
                        likeDetails += " and " + (item.likes.summary.total_count - likes.length) + " others";
                    }
                }
                var likeTemplate = "";
                if (item.likes && item.likes.summary.total_count) {
                    likeTemplate = " <div class=\"facebook-item-likes\">\n                                    <span title=\"" + likeDetails + "\">\n                                        <i class=\"fa fa-thumbs-up\" aria-hidden=\"true\"></i> " + item.likes.summary.total_count + " \n                                    </span>\n                                </div>";
                }
                return likeTemplate;
            };
            this.renderComments = function (item) {
                var comments = "";
                if (item.comments && item.comments.summary.total_count) {
                    comments = " <div class=\"facebook-item-comments\" title=\"" + item.comments.summary.total_count + " comments\">\n                       <i class=\"fa fa-comments\" aria-hidden=\"true\"></i> " + item.comments.summary.total_count + "\n                    </div>";
                }
                return comments;
            };
            this.renderImage = function (item) {
                var postImage = "";
                if (item.picture && item.type !== "link") {
                    postImage = " <a href=\"" + item.link + "\" target=\"_blank\"><img class=\"facebook-item-image-thumbnail\" src=\"" + item.picture + "&type=thumbnail\" alt=\"Post Image\"></a>";
                }
                else if (_this.appProperties.DefaultIcon) {
                    postImage =
                        " <a href=\"" + item.link + "\" target=\"_blank\"><img class=\"facebook-item-image-thumbnail\" src=\"" + _this
                            .appProperties.DefaultIcon + "\" alt=\"Default Image\"></a>";
                }
                else {
                    postImage = " <a href=\"" + item.link + "\" target=\"_blank\"><img class=\"facebook-item-image-thumbnail\" src=\"" + item.picture + "\" alt=\"Post Image\"></a>";
                }
                return postImage;
            };
            this.renderLink = function (item) {
                if (!item.name && item.type) {
                    item.name = item.type.substr(0, 1).toUpperCase();
                    item.name = item.name + item.type.substr(1);
                }
                var node = "<div class=\"facebook-item row facebook-" + item.type + "\">\n                    <div class=\"icon col-xs-3 text-center\">\n                        " + _this.renderIcon(item) + "\n                        " + _this.renderImage(item) + "\n                    </div>\n                    <div class=\"details col-xs-7 no-padding-left\">\n                        <a href=\"" + item.link + "\" target=\"_blank\">\n                            <h4 class=\"facebook-item-title\">\n                                " + item.name + "\n                            </h4>\n                        </a>\n                        " + _this.renderDescription(item) + "\n                        <div class=\"social\">\n                            <div class=\"pull-left like\">\n                                " + _this.renderLikes(item) + "\n                                " + _this.renderComments(item) + "\n                            </div>\n                            <div class=\"facebook-item-date pull-right\" title=\"" + moment(item.created_time).format("dddd, MMMM Do YYYY, h:mm:ss a") + "\">\n                                <i class=\"fa fa-clock-o\" aria-hidden=\"true\"></i> " + _this.formatDate(item.created_time) + "\n                            </div>\n                        </div>\n                    </div>\n             </div>";
                return node;
            };
            this.renderMore = function () {
                return "<div class=\"facebook-wp-more\"><a href=\"" + _this.appProperties.MoreLinkUrl + "\" target=\"_blank\"><i class=\"fa fa-external-link\" aria-hidden=\"true\"></i> " + _this.appProperties.MoreLinkText + "</div>";
            };
            this.renderTitle = function () {
                return "<h1 class=\"facebookwp-heading\"><i class=\"fa fa-facebook-square\" aria-hidden=\"true\"></i> " + _this.appProperties.Title + "</h1> ";
            };
            this.lastDataSet = null;
            this.render = function (items, paging) {
                var nodes = new Array();
                nodes.push(_this.renderTitle());
                try {
                    for (var i = 0; i < items.length; i++) {
                        switch (items[i].type) {
                            case "photo":
                                nodes.push(_this.renderPhoto(items[i]));
                                break;
                            case "link":
                                nodes.push(_this.renderLink(items[i]));
                                break;
                            case "video":
                                nodes.push(_this.renderVideo(items[i]));
                                break;
                            default:
                                console.log("No renderer associated with type " + items[i].type + ", will fallback to link renderer");
                                nodes.push(_this.renderLink(items[i]));
                                break;
                        }
                    }
                    nodes.push(_this.renderMore());
                    $(".facebook-wp-panel").empty().append(nodes).hide().fadeIn();
                }
                catch (e) {
                    console.error("Error: " + e);
                }
            };
            this.appProperties = new AppProperties();
            this.appConfig = new ClientAppConfig.Config();
            Resources.load();
        }
        Facebook.prototype.loadConfig = function () {
            var _this = this;
            var deferred = $.Deferred();
            //Reuse: Register all custom properties and include default values
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.Title, "Facebook", "Web Part Title (Optional)");
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.AccessToken, null, "This property is generated by facebook.  You must register an app then use it to generate the access token.");
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.APIVersion, "v2.9", "The api version, generally should be left to the default value");
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.RefreshSeconds, "0", "To make the web part automatically refresh set to a number greater than 0");
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.PageId, null, "The page id of the facebook page. The Page Id is a series of numbers.");
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.PageSize, "5", "Total number of items to show in a single page");
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.MoreLinkText, "Read more", "Changes the text on the link to the facebook page");
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.MoreLinkUrl, "https://www.facebook.com/", "Changes the url to the facebook page, eg; https://www.facebook.com/MyCompanyName");
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.RollupChars, "100", "Maximum number of characters to show of the description field before using the read more function.");
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.DefaultIcon, "", "If no picture is available, default to this image.");
            this.appConfig.load(this.appProperties.AppName, this.appProperties.AppInstance)
                .done(function (properties) {
                //Reuse: Update the config - this is what the app will use when needing values
                _this.appProperties.AcccessToken = _this.appConfig.getPropertyValue(CustomProperties.AccessToken);
                _this.appProperties.RefreshSeconds = _this.appConfig.getPropertyValueNumber(CustomProperties.RefreshSeconds);
                _this.appProperties.APIVersion = _this.appConfig.getPropertyValue(CustomProperties.APIVersion);
                _this.appProperties.PageId = _this.appConfig.getPropertyValue(CustomProperties.PageId);
                _this.appProperties.PageSize = _this.appConfig.getPropertyValueNumber(CustomProperties.PageSize);
                _this.appProperties.MoreLinkText = _this.appConfig.getPropertyValue(CustomProperties.MoreLinkText);
                _this.appProperties.MoreLinkUrl = _this.appConfig.getPropertyValue(CustomProperties.MoreLinkUrl);
                _this.appProperties.RollupChars = _this.appConfig.getPropertyValueNumber(CustomProperties.RollupChars);
                _this.appProperties.DefaultIcon = _this.appConfig.getPropertyValue(CustomProperties.DefaultIcon);
                _this.appProperties.Title = _this.appConfig.getPropertyValue(CustomProperties.Title);
                deferred.resolve();
            })
                .fail(function (sender, args) {
                deferred.reject(sender, args);
            });
            return deferred.promise();
        };
        Facebook.start = function () {
            var app = new Facebook();
            app.loadConfig().done(function () {
                if (!app.appProperties.AcccessToken || !app.appProperties.PageId) {
                    $(".facebook-wp-panel").empty().append("Web part is not configured. Edit this page to configure");
                }
                else {
                    app.fetch();
                    if (app.appProperties.RefreshSeconds > 0) {
                        setInterval(function () { app.fetch(); }, app.appProperties.RefreshSeconds * 1000);
                    }
                }
            })
                .fail(function (sender, args) {
                console.log("Error loading config: " + args.get_message());
            });
        };
        Facebook.prototype.formatDate = function (dateValue) {
            var date = moment(dateValue);
            var dateString = "";
            var duration = moment.duration(new moment().diff(date));
            var hours = Math.round(duration.asHours());
            var days = Math.round(duration.asDays());
            if (hours < 24) {
                dateString = date.fromNow();
            }
            else if (hours <= 120) {
                dateString = Math.max(1, days) + " d";
            }
            else {
                dateString = date.format("MMM Do");
            }
            return dateString;
        };
        Facebook.prototype.fetch = function () {
            var _this = this;
            //fetch data from facebook.
            var url = "https://graph.facebook.com/" + this.appProperties.APIVersion + "/" + this.appProperties.PageId + "/posts?access_token=" + this.appProperties.AcccessToken + "&pretty=1&fields=type%2Cstory%2Cdescription%2Cmessage%2Cstatus_type%2Cfrom%2Cpicture%2Clink%2Cname%2Clikes.limit(3).summary(true)%2Ccomments.limit(0).summary(true)%2Ccreated_time&limit=" + this.appProperties.PageSize;
            $.getJSON(url, null, function (response, textStatus, xhr) {
                if (response && response.data) {
                    if (JSON.stringify(response.data) === _this.lastDataSet) {
                        //No changes
                        return;
                    }
                }
                _this.lastDataSet = JSON.stringify(response.data);
                _this.render(response.data, response.paging);
            }).fail(function (result) {
                console.error("Error - StatusText: " + result.statusText + " ResponseText: " + result.responseText + " Status Code:" + result.status);
                _this.clearDisplay();
            });
        };
        Facebook.prototype.clearDisplay = function () {
            $(".facebook-wp-panel").empty();
        };
        return Facebook;
    }());
    Facebook.configure = function (configElement, settingsIconUrl) {
        //configElement is a dom node to output the text controls for updating properties.
        //Load the configuration (which maybe the default) and then show the config controls.
        //The settingsIconUrl is a url to an image that is displayed in Deisgn Mode only.  Should be a gear icon.
        var app = new FacebookClientApp.Facebook();
        app.loadConfig().done(function () {
            app.appConfig.ensureSaved(app.appProperties.AppName, app.appProperties.AppInstance).done(function () {
                app.appConfig.showConfigForm(app.appProperties.AppName, app.appProperties.AppInstance, configElement, settingsIconUrl);
            })
                .fail(function (sender, args) {
                console.error("Failed to update initial web part state.  Please check your permissions to the root web Client App Config list.");
            });
        }).fail(function (sender, args) {
            console.error("Failed to load configuration: " + args.get_message());
        });
    };
    FacebookClientApp.Facebook = Facebook;
    var FacebookParsers = (function () {
        function FacebookParsers() {
        }
        return FacebookParsers;
    }());
    FacebookParsers.parseHashtag = function (hashTag) {
        return hashTag.replace(/[#]+[A-Za-z0-9-_]+/g, function (t) {
            var tag = t.replace("#", "");
            var url = "https://www.facebook.com/hashtag/" + tag;
            var result = "<a href=\"" + url + "\" target=\"_blank\">#" + tag + "</a>";
            return result;
        });
    };
})(FacebookClientApp || (FacebookClientApp = {}));
//# sourceMappingURL=Facebook.js.map