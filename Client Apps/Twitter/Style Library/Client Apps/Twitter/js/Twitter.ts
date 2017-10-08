/*  Title: Twitter Web Part
    Author: Tim Wheeler http://blog.timwheeler.io/author/tim/
    Purpose: Shows a twitter page feed, and displays in a custom layout.
    Can be styled to suit the organisation.
    Has an auto refresh option and a page size option.
    This is app uses a reusable config module to allow users to update the configuration in design mode.
    The config saves to the Client App Config list.
    The app does not access twitter directly, but loads from a Twitter Cache list.  There is a script that will automate populating the list.
 */
declare var moment;
declare module TwitterSchema {

    export interface Hashtag {
        text: string;
        indices: number[];
    }

    export interface Url {
        url: string;
        expanded_url: string;
        display_url: string;
        indices: number[];
    }

    export interface Entities {
        hashtags: Hashtag[];
        symbols: any[];
        user_mentions: any[];
        urls: Url[];
    }

    export interface Url3 {
        url: string;
        expanded_url: string;
        display_url: string;
        indices: number[];
    }

    export interface Url2 {
        urls: Url3[];
    }

    export interface Description {
        urls: any[];
    }

    export interface Entities2 {
        url: Url2;
        description: Description;
    }

    export interface User {
        id: number;
        id_str: string;
        name: string;
        screen_name: string;
        location: string;
        description: string;
        url: string;
        entities: Entities2;
        protected: boolean;
        followers_count: number;
        friends_count: number;
        listed_count: number;
        created_at: string;
        favourites_count: number;
        utc_offset: number;
        time_zone: string;
        geo_enabled: boolean;
        verified: boolean;
        statuses_count: number;
        lang: string;
        contributors_enabled: boolean;
        is_translator: boolean;
        is_translation_enabled: boolean;
        profile_background_color: string;
        profile_background_image_url: string;
        profile_background_image_url_https: string;
        profile_background_tile: boolean;
        profile_image_url: string;
        profile_image_url_https: string;
        profile_banner_url: string;
        profile_link_color: string;
        profile_sidebar_border_color: string;
        profile_sidebar_fill_color: string;
        profile_text_color: string;
        profile_use_background_image: boolean;
        has_extended_profile: boolean;
        default_profile: boolean;
        default_profile_image: boolean;
        following?: any;
        follow_request_sent?: any;
        notifications?: any;
        translator_type: string;
    }

    export interface Post {
        created_at: string;
        id: number;
        id_str: string;
        text: string;
        truncated: boolean;
        entities: Entities;
        source: string;
        in_reply_to_status_id?: any;
        in_reply_to_status_id_str?: any;
        in_reply_to_user_id?: any;
        in_reply_to_user_id_str?: any;
        in_reply_to_screen_name?: any;
        user: User;
        geo?: any;
        coordinates?: any;
        place?: any;
        contributors?: any;
        is_quote_status: boolean;
        retweet_count: number;
        favorite_count: number;
        favorited: boolean;
        retweeted: boolean;
        possibly_sensitive: boolean;
        lang: string;
    }
    
}

module TwitterClientApp {
    class CustomProperties {
        //Reuse: Update these Property Names
        public static readonly RefreshSeconds = "Refresh Seconds";
        public static readonly PageSize = "Page Size";
        public static readonly MoreLinkText = "Read More Link Text";
        public static readonly MoreLinkUrl = "Read More Link Url";
        public static readonly Title = "Title";
    }
    class CacheFields {
        public static readonly TwitterCache = "Twitter_x0020_Cache";
        public static readonly Title = "Title";
        public static readonly Modified = "Modified";
    }
    export class Resources {
        public static load = () => {
            var loadCss = (name, url) => {
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
            var baseUrl = _spPageContextInfo.siteServerRelativeUrl;
            if (baseUrl.substring(baseUrl.length - 1) === "/") {
                baseUrl = baseUrl.substring(0, baseUrl.length - 1);
            }
            loadCss('twitterappcss', _spPageContextInfo.siteAbsoluteUrl + "/Style Library/Client Apps/Twitter/css/twitter.css");
            loadCss('clientappconfig', _spPageContextInfo.siteAbsoluteUrl + "/Style Library/Client Apps/Common/css/clientapp-config.css");
            loadCss('font-awesome.min.css', baseUrl + "/Style Library/Client Apps/Common/libs/font-awesome-4.7.0/css/font-awesome.min.css");
            loadCss('bootstrap-grid.css', baseUrl + "/Style Library/Client Apps/Common/libs/bootstrap-grid-3.3.7/bootstrap-grid.min.css");
        }
    }

    class AppProperties {
        public readonly AppName = "Twitter";
        public AppInstance = `${_spPageContextInfo.webServerRelativeUrl}/${_spPageContextInfo.pageItemId}`;
        public RefreshSeconds: number;
        public PageSize: number;
        public MoreLinkText: string;
        public MoreLinkUrl: string;
        public Title: string;
    }

    export class Twitter {
        constructor() {
            this.appProperties = new AppProperties();
            this.appConfig = new ClientAppConfig.Config();
        }
        public static readonly cacheListName: string = "Twitter Cache";
        appProperties: AppProperties;
        appConfig: ClientAppConfig.Config;
        public static configure = (configElement: Element, settingsIconUrl: string) => {
            //configElement is a dom node to output the text controls for updating properties.
            //Load the configuration (which maybe the default) and then show the config controls.
            //The settingsIconUrl is a url to an image that is displayed in Deisgn Mode only.  Should be a gear icon.
            var app = new TwitterClientApp.Twitter();
            app.loadConfig().done(() => {
                app.appConfig.ensureSaved(app.appProperties.AppName, app.appProperties.AppInstance).done(() => {
                    app.appConfig.showConfigForm(app.appProperties.AppName, app.appProperties.AppInstance, configElement, settingsIconUrl);
                })
                    .fail((sender, args) => {
                        console.error("Failed to update initial web part state.  Please check your permissions to the root web Client App Config list.");
                    });

            }).fail((sender, args) => {
                console.error("Failed to load configuration: " + args.get_message());
            });
        }
        loadConfig(): JQueryPromise<Array<AppProperties>> {
            var deferred = $.Deferred();
            //Reuse: Register all custom properties and include default values
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.Title, "Twitter", "Web Part Title (Optional)");
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.RefreshSeconds, "0", "To make the web part automatically refresh set to a number greater than 0");
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.PageSize, "5", "Total number of items to show in a single page");
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.MoreLinkText, "Read more", "Changes the text on the link to the twitter page");
            this.appConfig.registerAppProperty(this.appProperties.AppName, this.appProperties.AppInstance, CustomProperties.MoreLinkUrl, "https://twitter.com/", "Changes the url to the facebook page, eg; https://twitter.com/MyCompanyName");
            this.appConfig.load(this.appProperties.AppName, this.appProperties.AppInstance)
                .done((properties) => {
                    //Reuse: Update the config - this is what the app will use when needing values
                    this.appProperties.RefreshSeconds = this.appConfig.getPropertyValueNumber(CustomProperties.RefreshSeconds);
                    this.appProperties.PageSize = this.appConfig.getPropertyValueNumber(CustomProperties.PageSize);
                    this.appProperties.MoreLinkText = this.appConfig.getPropertyValue(CustomProperties.MoreLinkText);
                    this.appProperties.MoreLinkUrl = this.appConfig.getPropertyValue(CustomProperties.MoreLinkUrl);
                    this.appProperties.Title = this.appConfig.getPropertyValue(CustomProperties.Title);
                    deferred.resolve();
                })
                .fail((sender, args) => {
                    deferred.reject(sender, args);
                });
            return deferred.promise();
        }
        public static start() {
            var app = new TwitterClientApp.Twitter();
                app.loadConfig().done(() => {
                    if (!app.appProperties.PageSize) {
                        $(".twitter-wp-panel").empty().append("Web part is not configured. Edit this page to configure");
                    }
                    else {
                        app.fetch();
                        if (app.appProperties.RefreshSeconds > 0) {
                            setInterval(() => { app.fetch() }, app.appProperties.RefreshSeconds * 1000);
                        }
                    }
                })
                .fail((sender, args) => {
                    console.log("Error loading config: " + args.get_message());
                });
        }
        private formatDate(dateValue) {
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
        }
        renderIcon = (item: TwitterSchema.Post) => {
            if (item.user && item.user.profile_background_image_url_https) {
                return `<img class="twitter-item-icon" src="${item.user.profile_image_url_https.replace("_normal","")}" alt="${item.user.screen_name}">`;
            }
            return "";
        };
        renderDescription = (item: TwitterSchema.Post) => {
            if (!item.text) {
                return "";
            }
            var tweet = TwitterParsers.parseURL(item.text);
            tweet = TwitterParsers.parseUsername(tweet);
            tweet = TwitterParsers.parseHashtag(tweet);
            var description = `<div class="twitter-item-description">
                                <span class="twitter-item-description">${tweet}</span>
                            </div>`;
            return description;
        };
        dayDiff(first, second) {
            return Math.round((second - first) / (1000 * 60 * 60 * 24));
        };
        renderTitle = (item: TwitterSchema.Post) => {
            if (!item.user) {
                return "";
            }
            
            var link = TwitterParsers.parsePostUrl(item);

            var title = `<div class="twitter-item-header">
                                <span class="twitter-header-screenname">${item.user.name}</span> <a href="${link}" target="_blank" title="View ${item.user.name}"><span class="twitter-header-displayname">@${item.user.screen_name}</span></a>
                            </div>`;
            return title;
        };
        renderWebPartTitle = () => {
            return `<h1 class="twitterwp-heading"><i class="fa fa-twitter-square" aria-hidden="true"></i> ${this.appProperties.Title}</h1> `;
        };
        renderDate = (item: TwitterSchema.Post) => {
            var dateString = this.formatDate(item.created_at);
            return `<i class="fa fa-clock-o" aria-hidden="true"></i> ${dateString}`;
        };
        renderStatus = (item: TwitterSchema.Post) => {
            var retweet = "";
            var fav = "";
            if (item.favorite_count || item.retweet_count) {
                retweet = `<span class="retweeted" title="Retweeted ${item.retweet_count} times">
                                        <i class="fa fa-retweet fa-2x" aria-hidden="true"></i> ${item.retweet_count} 
                                    </span>`;

                fav = `<span class="favourited" title="Liked ${item.favorite_count} times">
                                 <i class="fa fa-heart fa-2x" aria-hidden="true"></i> ${item.favorite_count} 
                          </span>`;
                var node = ` <div class="twitter-item-retweet">
                                        ${retweet}
                                        ${fav}
                             </div>`;
                return node;
            }
            return "";
        };
        renderMore = () => {
            return `<div class="twitter-wp-more pull-right"><a href="${this.appProperties.MoreLinkUrl}" target="_blank"><i class="fa fa-external-link" aria-hidden="true"></i> ${this.appProperties.MoreLinkText}</div>`;
        };
        
        renderLink = (item: TwitterSchema.Post) => {
            var link = TwitterParsers.parsePostUrl(item);
          //  var date = TwitterParsers.parseDate(item.created_at);
            var node =
                ` <div class="twitter-item row">
                    <div class="twitter-item-icon col-xs-3 text-center">
                        <a href="${link}" target="_blank">
                            ${this.renderIcon(item)}
                        </a>
                    </div>
                    <div class="twitter-item-details col-xs-7 no-padding-left">
                        <div class="twitter-item-title">${this.renderTitle(item)}</div>
                        <div class="twitter-item-post">${this.renderDescription(item)}</div>
                        <div class="social">
                            <div class="pull-left status">
                                ${this.renderStatus(item)}
                            </div>
                            <span class="twitter-item-date pull-right" title="${new moment(TwitterParsers.parseDate(item.created_at)).format("dddd, MMMM Do YYYY, h:mm:ss a") }">
                                ${this.renderDate(item)}
                            </span>
                        </div>
                    </div>
                 </div>`;
            return node;
        };
       
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
        private loadFromCache(): JQueryPromise<Array<TwitterSchema.Post>> {
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
                .done((sender: any, args: any) => {
                    var items = new Array<TwitterSchema.Post>();
                    if (listItems.get_count()) {
                        var cachedItem = listItems.itemAt(0);
                        var lastRefresh = cachedItem.get_item(CacheFields.Modified); //TODO: do anything with this?
                        var data = cachedItem.get_item(CacheFields.TwitterCache);
                        if (data) {
                            items = JSON.parse(data);
                        }
                        if (this.appProperties.PageSize && items.length > this.appProperties.PageSize) {
                            items.splice(this.appProperties.PageSize);
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
        private lastDataSet = null as string;
       
        fetch() {
            this.loadFromCache()
                .done(this.render)
                .fail((result) => {
                    console.error("Error - StatusText: " + result.statusText + " ResponseText: " + result.responseText + " Status Code:" + result.status);
                    this.clearDisplay();
                });
        }
        private clearDisplay() {
            $(".twitter-wp-panel").empty();
        }
        
        render = (items: Array<TwitterSchema.Post>) => {
            if (this.lastDataSet === JSON.stringify(items)) {
                return; //only update the dom if the data has changed.
            }
            var nodes = new Array();
            nodes.push(this.renderWebPartTitle());
            try {
                for (var i = 0; i < items.length; i++) {
                    nodes.push(this.renderLink(items[i] as TwitterSchema.Post));
                }
                nodes.push(this.renderMore());
                $(".twitter-wp-panel").empty().append(nodes).hide().fadeIn();
                this.lastDataSet = JSON.stringify(items);
            }
            catch (e) {
                console.error("Error: " + e);
            }
        };
    }
    class TwitterParsers {
        public static parsePostUrl(item: TwitterSchema.Post) {
            var userName = this.parseUsername(item.user.name)
            return `https://twitter.com/${userName}/status/${item.id_str}`;
        }
        public static parseURL(urlString): any {
            return urlString.replace(/[A-Za-z]+:\/\/[A-Za-z0-9-_]+\.[A-Za-z0-9-_:%&~\?\/.=]+/g, function (url) {
                if (url) {
                    return `<a href="${url}" target="_blank">${url}</a>`;
                }
                return "";
            });
        };
        public static parseUsername = (userName) => {
            return userName.replace(/[@]+[A-Za-z0-9-_]+/g, function (u) {
                var username = u.replace("@", "");
                var url = "http://twitter.com/" + username;
                return `<a href="${url}" target="_blank">@${username}</a>`;
            });
        };
        public static parseHashtag = (hashTag) => {
            return hashTag.replace(/[#]+[A-Za-z0-9-_]+/g, function (t) {
                var tag = t.replace("#", "");
                var url = `https://twitter.com/hashtag/${tag}?src=hash`;
                var result = `<a href="${url}" target="_blank">#${tag}</a>`;
                return result;
            });
        };
        public static parseDate = (str) => {
            var v = str.split(' ');
            return new Date(Date.parse(v[1] + " " + v[2] + ", " + v[5] + " " + v[3] + " UTC"));
        }
    }
}
TwitterClientApp.Resources.load();




