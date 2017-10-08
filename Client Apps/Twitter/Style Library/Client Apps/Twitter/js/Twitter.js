var TwitterClientApp;
(function (TwitterClientApp) {
    var CustomProperties = (function () {
        function CustomProperties() {
        }
        return CustomProperties;
    }());
    //Reuse: Update these Property Names
    CustomProperties.RefreshSeconds = "Refresh Seconds";
    CustomProperties.PageSize = "Page Size";
    CustomProperties.MoreLinkText = "Read More Link Text";
    CustomProperties.MoreLinkUrl = "Read More Link Url";
    CustomProperties.Title = "Title";
    var CacheFields = (function () {
        function CacheFields() {
        }
        return CacheFields;
    }());
    CacheFields.TwitterCache = "Twitter_x0020_Cache";
    CacheFields.Title = "Title";
    CacheFields.Modified = "Modified";
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
        loadCss('twitterappcss', _spPageContextInfo.siteAbsoluteUrl + "/Style Library/Client Apps/Twitter/css/twitter.css");
        loadCss('clientappconfig', _spPageContextInfo.siteAbsoluteUrl + "/Style Library/Client Apps/Common/css/clientapp-config.css");
        loadCss('font-awesome.min.css', baseUrl + "/Style Library/Client Apps/Common/libs/font-awesome-4.7.0/css/font-awesome.min.css");
        loadCss('bootstrap-grid.css', baseUrl + "/Style Library/Client Apps/Common/libs/bootstrap-grid-3.3.7/bootstrap-grid.min.css");
    };
    TwitterClientApp.Resources = Resources;
    var AppProperties = (function () {
        function AppProperties() {
            this.AppName = "Twitter";
            this.AppInstance = _spPageContextInfo.webServerRelativeUrl + "/" + _spPageContextInfo.pageItemId;
        }
        return AppProperties;
    }());
    var Twitter = (function () {
        function Twitter() {
            var _this = this;
            this.renderIcon = function (item) {
                if (item.user && item.user.profile_background_image_url_https) {
                    return "<img class=\"twitter-item-icon\" src=\"" + item.user.profile_image_url_https.replace("_normal", "") + "\" alt=\"" + item.user.screen_name + "\">";
                }
                return "";
            };
            this.renderDescription = function (item) {
                if (!item.text) {
                    return "";
                }
                var tweet = TwitterParsers.parseURL(item.text);
                tweet = TwitterParsers.parseUsername(tweet);
                tweet = TwitterParsers.parseHashtag(tweet);
                var description = "<div class=\"twitter-item-description\">\n                                <span class=\"twitter-item-description\">" + tweet + "</span>\n                            </div>";
                return description;
            };
            this.renderTitle = function (item) {
                if (!item.user) {
                    return "";
                }
                var link = TwitterParsers.parsePostUrl(item);
                var title = "<div class=\"twitter-item-header\">\n                                <span class=\"twitter-header-screenname\">" + item.user.name + "</span> <a href=\"" + link + "\" target=\"_blank\" title=\"View " + item.user.name + "\"><span class=\"twitter-header-displayname\">@" + item.user.screen_name + "</span></a>\n                            </div>";
                return title;
            };
            this.renderWebPartTitle = function () {
                return "<h1 class=\"twitterwp-heading\"><i class=\"fa fa-twitter-square\" aria-hidden=\"true\"></i> " + _this.appProperties.Title + "</h1> ";
            };
            this.renderDate = function (item) {
                var dateString = _this.formatDate(item.created_at);
                return "<i class=\"fa fa-clock-o\" aria-hidden=\"true\"></i> " + dateString;
            };
            this.renderStatus = function (item) {
                var retweet = "";
                var fav = "";
                if (item.favorite_count || item.retweet_count) {
                    retweet = "<span class=\"retweeted\" title=\"Retweeted " + item.retweet_count + " times\">\n                                        <i class=\"fa fa-retweet fa-2x\" aria-hidden=\"true\"></i> " + item.retweet_count + " \n                                    </span>";
                    fav = "<span class=\"favourited\" title=\"Liked " + item.favorite_count + " times\">\n                                 <i class=\"fa fa-heart fa-2x\" aria-hidden=\"true\"></i> " + item.favorite_count + " \n                          </span>";
                    var node = " <div class=\"twitter-item-retweet\">\n                                        " + retweet + "\n                                        " + fav + "\n                             </div>";
                    return node;
                }
                return "";
            };
            this.renderMore = function () {
                return "<div class=\"twitter-wp-more pull-right\"><a href=\"" + _this.appProperties.MoreLinkUrl + "\" target=\"_blank\"><i class=\"fa fa-external-link\" aria-hidden=\"true\"></i> " + _this.appProperties.MoreLinkText + "</div>";
            };
            this.renderLink = function (item) {
                var link = TwitterParsers.parsePostUrl(item);
                //  var date = TwitterParsers.parseDate(item.created_at);
                var node = " <div class=\"twitter-item row\">\n                    <div class=\"twitter-item-icon col-xs-3 text-center\">\n                        <a href=\"" + link + "\" target=\"_blank\">\n                            " + _this.renderIcon(item) + "\n                        </a>\n                    </div>\n                    <div class=\"twitter-item-details col-xs-7 no-padding-left\">\n                        <div class=\"twitter-item-title\">" + _this.renderTitle(item) + "</div>\n                        <div class=\"twitter-item-post\">" + _this.renderDescription(item) + "</div>\n                        <div class=\"social\">\n                            <div class=\"pull-left status\">\n                                " + _this.renderStatus(item) + "\n                            </div>\n                            <span class=\"twitter-item-date pull-right\" title=\"" + new moment(TwitterParsers.parseDate(item.created_at)).format("dddd, MMMM Do YYYY, h:mm:ss a") + "\">\n                                " + _this.renderDate(item) + "\n                            </span>\n                        </div>\n                    </div>\n                 </div>";
                return node;
            };
            this.executeQuery = function (clientContext) {
                var deferred = $.Deferred();
                clientContext.executeQueryAsync(function (sender, args) {
                    deferred.resolve(sender, args);
                }, function (sender, args) {
                    deferred.reject(sender, args);
                });
                return deferred.promise();
            };
            this.lastDataSet = null;
            this.render = function (items) {
                if (_this.lastDataSet === JSON.stringify(items)) {
                    return; //only update the dom if the data has changed.
                }
                var nodes = new Array();
                nodes.push(_this.renderWebPartTitle());
                try {
                    for (var i = 0; i < items.length; i++) {
                        nodes.push(_this.renderLink(items[i]));
                    }
                    nodes.push(_this.renderMore());
                    $(".twitter-wp-panel").empty().append(nodes).hide().fadeIn();
                    _this.lastDataSet = JSON.stringify(items);
                }
                catch (e) {
                    console.error("Error: " + e);
                }
            };
            this.appProperties = new AppProperties();
            this.appConfig = new ClientAppConfig.Config();
        }
        Twitter.prototype.loadConfig = function () {
            var _this = this;
            var deferred = $.Deferred();
            //Reuse: Register all custom properties and include default values
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.Title, "Twitter", "Web Part Title (Optional)");
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.RefreshSeconds, "0", "To make the web part automatically refresh set to a number greater than 0");
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.PageSize, "5", "Total number of items to show in a single page");
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.MoreLinkText, "Read more", "Changes the text on the link to the twitter page");
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.MoreLinkUrl, "https://twitter.com/", "Changes the url to the facebook page, eg; https://twitter.com/MyCompanyName");
            this.appConfig.load(this.appProperties.AppName, this.appProperties.AppInstance)
                .done(function (properties) {
                //Reuse: Update the config - this is what the app will use when needing values
                _this.appProperties.RefreshSeconds = _this.appConfig.getPropertyValueNumber(CustomProperties.RefreshSeconds);
                _this.appProperties.PageSize = _this.appConfig.getPropertyValueNumber(CustomProperties.PageSize);
                _this.appProperties.MoreLinkText = _this.appConfig.getPropertyValue(CustomProperties.MoreLinkText);
                _this.appProperties.MoreLinkUrl = _this.appConfig.getPropertyValue(CustomProperties.MoreLinkUrl);
                _this.appProperties.Title = _this.appConfig.getPropertyValue(CustomProperties.Title);
                deferred.resolve();
            })
                .fail(function (sender, args) {
                deferred.reject(sender, args);
            });
            return deferred.promise();
        };
        Twitter.start = function () {
            var app = new TwitterClientApp.Twitter();
            app.loadConfig().done(function () {
                if (!app.appProperties.PageSize) {
                    $(".twitter-wp-panel").empty().append("Web part is not configured. Edit this page to configure");
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
        Twitter.prototype.formatDate = function (dateValue) {
            var date = moment(TwitterParsers.parseDate(dateValue));
            var dateString = "";
            var duration = moment.duration(moment().diff(date));
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
        Twitter.prototype.dayDiff = function (first, second) {
            return Math.round((second - first) / (1000 * 60 * 60 * 24));
        };
        ;
        Twitter.prototype.loadFromCache = function () {
            var _this = this;
            var deferred = $.Deferred();
            var camlBuilder = new SpData.CamlBuilder();
            camlBuilder.begin(true); //Use an AND query
            camlBuilder.addViewFields([CacheFields.TwitterCache, CacheFields.Title, CacheFields.Modified]);
            var clientContext = SP.ClientContext.get_current();
            var list = clientContext.get_site().get_rootWeb().get_lists().getByTitle(Twitter.cacheListName);
            var query = new SP.CamlQuery();
            query.set_viewXml(camlBuilder.viewXml);
            var listItems = list.getItems(query);
            clientContext.load(listItems);
            this.executeQuery(clientContext)
                .done(function (sender, args) {
                var items = new Array();
                if (listItems.get_count()) {
                    var cachedItem = listItems.itemAt(0);
                    var lastRefresh = cachedItem.get_item(CacheFields.Modified); //TODO: do anything with this?
                    var data = cachedItem.get_item(CacheFields.TwitterCache);
                    if (data) {
                        items = JSON.parse(data);
                    }
                    if (_this.appProperties.PageSize && items.length > _this.appProperties.PageSize) {
                        items.splice(_this.appProperties.PageSize);
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
        Twitter.prototype.fetch = function () {
            var _this = this;
            this.loadFromCache()
                .done(this.render)
                .fail(function (result) {
                console.error("Error - StatusText: " + result.statusText + " ResponseText: " + result.responseText + " Status Code:" + result.status);
                _this.clearDisplay();
            });
        };
        Twitter.prototype.clearDisplay = function () {
            $(".twitter-wp-panel").empty();
        };
        return Twitter;
    }());
    Twitter.cacheListName = "Twitter Cache";
    Twitter.configure = function (configElement, settingsIconUrl) {
        //configElement is a dom node to output the text controls for updating properties.
        //Load the configuration (which maybe the default) and then show the config controls.
        //The settingsIconUrl is a url to an image that is displayed in Deisgn Mode only.  Should be a gear icon.
        var app = new TwitterClientApp.Twitter();
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
    TwitterClientApp.Twitter = Twitter;
    var TwitterParsers = (function () {
        function TwitterParsers() {
        }
        TwitterParsers.parsePostUrl = function (item) {
            var userName = this.parseUsername(item.user.name);
            return "https://twitter.com/" + userName + "/status/" + item.id_str;
        };
        TwitterParsers.parseURL = function (urlString) {
            return urlString.replace(/[A-Za-z]+:\/\/[A-Za-z0-9-_]+\.[A-Za-z0-9-_:%&~\?\/.=]+/g, function (url) {
                if (url) {
                    return "<a href=\"" + url + "\" target=\"_blank\">" + url + "</a>";
                }
                return "";
            });
        };
        ;
        return TwitterParsers;
    }());
    TwitterParsers.parseUsername = function (userName) {
        return userName.replace(/[@]+[A-Za-z0-9-_]+/g, function (u) {
            var username = u.replace("@", "");
            var url = "http://twitter.com/" + username;
            return "<a href=\"" + url + "\" target=\"_blank\">@" + username + "</a>";
        });
    };
    TwitterParsers.parseHashtag = function (hashTag) {
        return hashTag.replace(/[#]+[A-Za-z0-9-_]+/g, function (t) {
            var tag = t.replace("#", "");
            var url = "https://twitter.com/hashtag/" + tag + "?src=hash";
            var result = "<a href=\"" + url + "\" target=\"_blank\">#" + tag + "</a>";
            return result;
        });
    };
    TwitterParsers.parseDate = function (str) {
        var v = str.split(' ');
        return new Date(Date.parse(v[1] + " " + v[2] + ", " + v[5] + " " + v[3] + " UTC"));
    };
})(TwitterClientApp || (TwitterClientApp = {}));
TwitterClientApp.Resources.load();
//# sourceMappingURL=Twitter.js.map