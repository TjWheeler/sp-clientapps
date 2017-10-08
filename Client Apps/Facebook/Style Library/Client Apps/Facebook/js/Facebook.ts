/*  Title: Facebook Web Part
    Author: Tim Wheeler http://blog.timwheeler.io/author/tim/
    Purpose: Queries a facebook page feed, and displays in a custom layout.
    Has an auto refresh option and a page size option.
    This is app uses a reusable config module to allow users to update the configuration in design mode.
    The config saves to the Client App Config list.
    The app requires an access token.  This is obtained from facebook.  To get one you have to register an app, then call a url to return the access token.
    This app does not yet support paging, although the paging url is tracked.
 */
declare var moment;
declare module FacebookSchema {

    export interface From {
        name: string;
        id: string;
    }

    export interface Datum {
        id: string;
        name: string;
    }

    export interface Cursors {
        before: string;
        after: string;
    }

    export interface Paging {
        cursors: Cursors;
        next: string;
    }

    export interface Summary {
        total_count: number;
        can_like: boolean;
        has_liked: boolean;
    }

    export interface Likes {
        data: Datum[];
        paging: Paging;
        summary: Summary;
    }

    export interface Summary2 {
        order: string;
        total_count: number;
        can_comment: boolean;
    }

    export interface Comments {
        data: any[];
        summary: Summary2;
    }

    export interface Post {
        type: string;
        description: string;
        message: string;
        status_type: string;
        from: From;
        picture: string;
        link: string;
        name: string;
        likes: Likes;
        comments: Comments;
        created_time: Date;
        id: string;
        story: string;
    }



}

module FacebookClientApp {
    class CustomProperties {
        //Reuse: Update these Property Names
        public static readonly AccessToken = "Access Token";
        public static readonly RefreshSeconds = "Refresh Seconds";
        public static readonly APIVersion = "API Version";
        public static readonly PageId = "Page Id";
        public static readonly PageSize = "Page Size";
        public static readonly MoreLinkText = "Read More Link Text";
        public static readonly MoreLinkUrl = "Read More Link Url";
        public static readonly RollupChars = "Rollup Characters";
        public static readonly DefaultIcon = "Default Icon";
        public static readonly Title = "Title";
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
            loadCss('facebookappcss', baseUrl + "/Style Library/Client Apps/Facebook/css/facebook.css");
            loadCss('clientappconfig', baseUrl + "/Style Library/Client Apps/Common/css/clientapp-config.css");
            loadCss('font-awesome.min.css', baseUrl + "/Style Library/Client Apps/Common/libs/font-awesome-4.7.0/css/font-awesome.min.css");
            loadCss('bootstrap-grid.css', baseUrl + "/Style Library/Client Apps/Common/libs/bootstrap-grid-3.3.7/bootstrap-grid.min.css");
        }
    }

    class AppProperties {
        public readonly AppName = "Facebook";
        public AppInstance = `${_spPageContextInfo.webServerRelativeUrl}/${_spPageContextInfo.pageItemId}`; 
        public AcccessToken: string;
        public RefreshSeconds: number;
        public APIVersion: string;
        public PageId: string;
        public PageSize: number;
        public MoreLinkText: string;
        public MoreLinkUrl: string;
        public RollupChars: number = 100;
        public DefaultIcon: string;
        public Title: string;
    }
    export class Facebook {
        constructor() {
            this.appProperties = new AppProperties();
            this.appConfig = new ClientAppConfig.Config();
            Resources.load();
        }
        appProperties: AppProperties;
        appConfig: ClientAppConfig.Config;
        public static configure = (configElement: Element, settingsIconUrl: string) => {
            //configElement is a dom node to output the text controls for updating properties.
            //Load the configuration (which maybe the default) and then show the config controls.
            //The settingsIconUrl is a url to an image that is displayed in Deisgn Mode only.  Should be a gear icon.
            var app = new FacebookClientApp.Facebook();
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
                .done((properties) => {
                    //Reuse: Update the config - this is what the app will use when needing values
                    this.appProperties.AcccessToken = this.appConfig.getPropertyValue(CustomProperties.AccessToken);
                    this.appProperties.RefreshSeconds = this.appConfig.getPropertyValueNumber(CustomProperties.RefreshSeconds);
                    this.appProperties.APIVersion = this.appConfig.getPropertyValue(CustomProperties.APIVersion);
                    this.appProperties.PageId = this.appConfig.getPropertyValue(CustomProperties.PageId);
                    this.appProperties.PageSize = this.appConfig.getPropertyValueNumber(CustomProperties.PageSize);
                    this.appProperties.MoreLinkText = this.appConfig.getPropertyValue(CustomProperties.MoreLinkText);
                    this.appProperties.MoreLinkUrl = this.appConfig.getPropertyValue(CustomProperties.MoreLinkUrl);
                    this.appProperties.RollupChars = this.appConfig.getPropertyValueNumber(CustomProperties.RollupChars);
                    this.appProperties.DefaultIcon = this.appConfig.getPropertyValue(CustomProperties.DefaultIcon);
                    this.appProperties.Title = this.appConfig.getPropertyValue(CustomProperties.Title);
                    deferred.resolve();
                })
                .fail((sender, args) => {
                    deferred.reject(sender, args);
                });
            return deferred.promise();
        }
        public static start() {
            var app = new Facebook();
                app.loadConfig().done(() => {
                    if (!app.appProperties.AcccessToken || !app.appProperties.PageId) {
                        $(".facebook-wp-panel").empty().append("Web part is not configured. Edit this page to configure");
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
        }
        renderVideo = (item: FacebookSchema.Post) => {
            return this.renderLink(item); //compatible
        };
        renderPhoto = (item: FacebookSchema.Post) => {
            return this.renderLink(item); //compatible
        };
        renderIcon = (item: FacebookSchema.Post) => {
            if (item.type && (item.type === "photo" || item.type === "video" || item.type === "story" || item.type === "link")) {
                var imageUrl = (_spPageContextInfo.siteServerRelativeUrl.length > 1 ? _spPageContextInfo.siteServerRelativeUrl : "") + "/Style Library/Client Apps/Facebook/img/" + item.type + ".png";
                return `<img class="facebook-item-icon" src="${imageUrl}" alt="${item.type}">`;
            }
            return "";
        };
        
        renderDescription = (item: FacebookSchema.Post) => {
            if (!item.description && item.message) {
                item.description = item.message;
            }
            if (!item.description && item.story) {
                item.description = item.story;
            }
            
            var rollup = "";
            var description = "";
            if (item.description && item.description.length > this.appProperties.RollupChars) {
                var lastSpace = item.description.substr(0, this.appProperties.RollupChars).lastIndexOf(" ");
                rollup = item.description.substr(0, lastSpace) + "...";
                if (item.description) {
                    description = `<div class="facebook-item-description">
                                    <span class="facebook-item-description-rollup">${rollup}
                                        <span class="clickable read-more" onclick="$(this).parent().toggle();$(this).parent().next().toggle();"> read more</span>
                                    </span>
                                    <span class="facebook-item-description" style="display:none">${FacebookParsers.parseHashtag(item.description)}
                                        <span class="clickable read-less" onclick="$(this).parent().toggle();$(this).parent().prev().toggle();"> (read less)</span>
                                    </span>
                               </div>`;
                }
            }
            else {
                description = `<div class="facebook-item-description">
                                    <span class="facebook-item-description">${item.description}</span>
                               </div>`;
            }
            return description;
        };
        renderLikes = (item: FacebookSchema.Post) => {
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
                likeTemplate = ` <div class="facebook-item-likes">
                                    <span title="${likeDetails}">
                                        <i class="fa fa-thumbs-up" aria-hidden="true"></i> ${item.likes.summary.total_count} 
                                    </span>
                                </div>`;
            }
            return likeTemplate;
        };
        renderComments = (item: FacebookSchema.Post) => {
            var comments = "";
            if (item.comments && item.comments.summary.total_count) {
                comments = ` <div class="facebook-item-comments" title="${item.comments.summary.total_count} comments">
                       <i class="fa fa-comments" aria-hidden="true"></i> ${item.comments.summary.total_count}
                    </div>`;
            }
            return comments;
        };
        renderImage = (item: FacebookSchema.Post) => {
            var postImage = "";
            if (item.picture && item.type !== "link") {
                postImage = ` <a href="${item.link}" target="_blank"><img class="facebook-item-image-thumbnail" src="${item.picture}&type=thumbnail" alt="Post Image"></a>`;
            }
            else if (this.appProperties.DefaultIcon) {
                postImage =
                    ` <a href="${item.link}" target="_blank"><img class="facebook-item-image-thumbnail" src="${this
                    .appProperties.DefaultIcon}" alt="Default Image"></a>`;
            } else {
                postImage = ` <a href="${item.link}" target="_blank"><img class="facebook-item-image-thumbnail" src="${item.picture}" alt="Post Image"></a>`;
            }
            return postImage;
        };
        renderLink = (item: FacebookSchema.Post) => {
            if (!item.name && item.type) {
                item.name = item.type.substr(0,1).toUpperCase();
                item.name = item.name + item.type.substr(1);
            }
            var node =
                `<div class="facebook-item row facebook-${item.type}">
                    <div class="icon col-xs-3 text-center">
                        ${this.renderIcon(item)}
                        ${this.renderImage(item)}
                    </div>
                    <div class="details col-xs-7 no-padding-left">
                        <a href="${item.link}" target="_blank">
                            <h4 class="facebook-item-title">
                                ${item.name}
                            </h4>
                        </a>
                        ${this.renderDescription(item)}
                        <div class="social">
                            <div class="pull-left like">
                                ${this.renderLikes(item)}
                                ${this.renderComments(item)}
                            </div>
                            <div class="facebook-item-date pull-right" title="${moment(item.created_time).format("dddd, MMMM Do YYYY, h:mm:ss a") }">
                                <i class="fa fa-clock-o" aria-hidden="true"></i> ${this.formatDate(item.created_time)}
                            </div>
                        </div>
                    </div>
             </div>`;
            return node;
        };
        renderMore = () => {
            return `<div class="facebook-wp-more"><a href="${this.appProperties.MoreLinkUrl}" target="_blank"><i class="fa fa-external-link" aria-hidden="true"></i> ${this.appProperties.MoreLinkText}</div>`;
        };
        renderTitle = () => {
            return `<h1 class="facebookwp-heading"><i class="fa fa-facebook-square" aria-hidden="true"></i> ${this.appProperties.Title}</h1> `; 
        };
        private lastDataSet = null as string;
        fetch() {
            //fetch data from facebook.
            var url = `https://graph.facebook.com/${this.appProperties.APIVersion}/${this.appProperties.PageId}/posts?access_token=${this.appProperties.AcccessToken}&pretty=1&fields=type%2Cstory%2Cdescription%2Cmessage%2Cstatus_type%2Cfrom%2Cpicture%2Clink%2Cname%2Clikes.limit(3).summary(true)%2Ccomments.limit(0).summary(true)%2Ccreated_time&limit=${this.appProperties.PageSize}`;
            $.getJSON(url, null, (response, textStatus, xhr) => {
                if (response && response.data) {
                    if (JSON.stringify(response.data) === this.lastDataSet) {
                        //No changes
                        return;
                    }
                }
                this.lastDataSet = JSON.stringify(response.data);
                this.render(response.data, response.paging);
            }).fail((result) => {
                console.error("Error - StatusText: " + result.statusText + " ResponseText: " + result.responseText + " Status Code:" + result.status);
                this.clearDisplay();
            });
        }
        private clearDisplay() {
            $(".facebook-wp-panel").empty();
        }

        render = (items: Array<any>, paging: FacebookSchema.Paging) => {
            var nodes = new Array();
            nodes.push(this.renderTitle());
            try {
                for (var i = 0; i < items.length; i++) {
                    switch (items[i].type) {
                        case "photo":
                            nodes.push(this.renderPhoto(items[i] as FacebookSchema.Post));
                            break;
                        case "link":
                            nodes.push(this.renderLink(items[i] as FacebookSchema.Post));
                            break;
                        case "video":
                            nodes.push(this.renderVideo(items[i] as FacebookSchema.Post));
                            break;
                        default:
                            console.log("No renderer associated with type " + items[i].type + ", will fallback to link renderer");
                            nodes.push(this.renderLink(items[i] as FacebookSchema.Post));
                            break;
                    }
                }
                nodes.push(this.renderMore());
                $(".facebook-wp-panel").empty().append(nodes).hide().fadeIn();
            }
            catch (e) {
                console.error("Error: " + e);
            }
        };
    }
    class FacebookParsers {
        public static parseHashtag = (hashTag) => {
            return hashTag.replace(/[#]+[A-Za-z0-9-_]+/g, function (t) {
                var tag = t.replace("#", "");
                var url = `https://www.facebook.com/hashtag/${tag}`;
                var result = `<a href="${url}" target="_blank">#${tag}</a>`;
                return result;
            });
        };
        
    }
}
