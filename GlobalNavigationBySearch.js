//Models and Namespaces
var wvstrienCustomGlobalNav = wvstrienCustomGlobalNav || {};
wvstrienCustomGlobalNav.Models = wvstrienCustomGlobalNav.Models || {}
wvstrienCustomGlobalNav.Models.NavigationNode = function () {
    this.Url = ko.observable("");
    this.Title = ko.observable("");
    this.Parent = ko.observable("");
};

wvstrienCustomGlobalNav.ViewModel = function () {

    function isIEorEDGE(){
        if (navigator.appName == 'Microsoft Internet Explorer') {
            // IE
            return true; 
        } else if(window.navigator.userAgent.indexOf("Edge") > -1){
            // EDGE
            return true;
        } else if (!navigator.userAgent.match(/Trident\/7\./)){
            return true;
        }

        return false;
    }


    var baseRequest = {
        url: "",
        type: ""
    };

    //Parses a local object from JSON search result.
    function getNavigationFromDto(dto) {
        var item = null;
        if (dto != undefined) {
            var webTemplate = getSearchResultsValue(dto.Cells.results, 'WebTemplate');

            if (webTemplate != "APP") {
                item = new wvstrienCustomGlobalNav.Models.NavigationNode();
                item.Title(getSearchResultsValue(dto.Cells.results, 'Title')); //Key = Title
                item.Url(getSearchResultsValue(dto.Cells.results, 'Path').toLowerCase().replace(/ /g, "%20")); //Key = Path
                item.Parent(getSearchResultsValue(dto.Cells.results, 'ParentLink')); //Key = ParentLink
            }

        }
        return item;
    }

    function getSearchResultsValue(results, key) {
        for (i = 0; i < results.length; i++) {
            if (results[i].Key == key) {
                return results[i].Value;
            }
        }
        return null;
    }

    //Parse a local object from the serialized cache.
    function getNavigationFromCache(dto) {
        var item = new wvstrienCustomGlobalNav.Models.NavigationNode();

        if (dto != undefined) {
            item.Title(dto.Title);
            item.Url(dto.Url);
            item.Parent(dto.Parent);
        }

        return item;
    }

    /* create a new OData request for JSON response */
    function getRequest(endpoint) {
        var request = baseRequest;
        request.type = "GET";
        request.url = endpoint;
        request.headers = { ACCEPT: "application/json;odata=verbose" };
        return request;
    };


    /* Navigation Module*/
    function NavigationViewModel() {
        "use strict";
        var self = this;

        var isIEOrEdge = isIEorEDGE();
        var root, userAndSiteKeyId;

        if (_spPageContextInfo) {
            root = _spPageContextInfo.siteAbsoluteUrl;
            userAndSiteKeyId = _spPageContextInfo.siteId + "_" + _spPageContextInfo.userLoginName + "_";
        } else {
           root = document.location.href.match(/https\:\/\/wvstrien.sharepoint.com\/.[^\/]+\/.[^\/]+\//)[0];
           root = root.substring(0, root.length - 1);
           userAndSiteKeyId = root + "_";
        }
        var nodeCacheKey = userAndSiteKeyId + "nodesCache";
        var nodeCachedAtKey = userAndSiteKeyId + "nodesCachedAt";
        var baseUrl = root + "/_api/search/query?querytext=";
        var query = baseUrl + "'contentClass=\"STS_Web\"-WebTemplate:APP+path:" + root + "'&trimduplicates=false&rowlimit=300";

        self.nodes = ko.observableArray([]);
        self.hierarchy = ko.observableArray([]);;
        self.loadNavigatioNodes = function () {
            if (document.location.href.match(/refreshnavigation/) === null) {
                //Check local storage for cached navigation datasource.
                var fromStorage = localStorage[nodeCacheKey];
                if (fromStorage != null) {
                    var cachedNodes = JSON.parse(fromStorage);
                    var timeStamp = JSON.parse(localStorage[nodeCachedAtKey]);
                    if (cachedNodes && timeStamp) {
                        //Check for cache expiration. Currently set to 3 hrs.
                        var now = new Date();
                        var diff = now.getTime() - timeStamp;
                        if (Math.round(diff / (1000 * 60 * 60)) < 3) {
                            //return from cache.
                            var cacheResults = [];
                            $.each(cachedNodes, function (i, item) {
                                var nodeitem = getNavigationFromCache(item, true);
                                cacheResults.push(nodeitem);
                            });

                            self.buildHierarchy(cacheResults);
                            self.toggleView();
                            addEventsToElements();
                            return;
                        }
                    }
                }
            }
            //No cache hit, REST call required.
            self.queryRemoteInterface();
        };

        //Executes a REST call and builds the navigation hierarchy.
        self.queryRemoteInterface = function () {
            var oDataRequest = getRequest(query);
            $.ajax(oDataRequest).done(function (data) {
                var results = [];
                $.each(data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results, function (i, item) {

                    if (i == 0) {
                        //Add root element.
                        var rootItem = new wvstrienCustomGlobalNav.Models.NavigationNode();

                        rootItem.Title("Root");
                        rootItem.Url(root.toLowerCase().replace(/ /g, "%20"));
                        rootItem.Parent(null);
                        results.push(rootItem);
                    }
                    var navItem = getNavigationFromDto(item);
                    if (navItem != null) results.push(navItem);
                });

                self.additionalGlobalNavigationLinks(results);
            }).fail(function () {
                //Handle error here!!
                $("#loading").hide();
                $("#error").show();
            });
        };

        self.additionalGlobalNavigationLinks = function (results) {
            var listAdditionalLinks = "AdditionalGlobalNavigationLinks";
            var checkListExists = root + "/_api/Web/Lists?$filter=title eq '" + listAdditionalLinks  + "'";
            var checkRequest = getRequest(checkListExists);
            $.ajax(checkRequest).done(function (data) {
                if (data.d.results.length > 0) {
                    var getAddionalLinksRequest = root + "/_api/web/lists/getbytitle('" + listAdditionalLinks + "')/items"
                    var getNavLinksRequest = getRequest(getAddionalLinksRequest);
                    $.ajax(getNavLinksRequest).done(function (data) {
                        $(data.d.results).each(function(i, item) {
                            if (item.Url != null) {
                                var navItem = new wvstrienCustomGlobalNav.Models.NavigationNode();
                                navItem.Title(item.Title);
                                navItem.Url(item.Url.toLowerCase().replace(/ /g, "%20"));
                                navItem.Parent(root);
                                results.push(navItem);
                            }
                        });

                        self.cacheAndDisplayQueriedResults(results);

                    }).fail(function () {
                        self.cacheAndDisplayQueriedResults(results);
                    });

                } else {
                    self.cacheAndDisplayQueriedResults(results);
                }
            });
        };

        self.cacheAndDisplayQueriedResults = function (results) {
            //Add to local cache
            localStorage[nodeCacheKey] = ko.toJSON(results);
            localStorage[nodeCachedAtKey] = new Date().getTime();

            self.nodes(results);
            if (self.nodes().length > 0) {
                var unsortedArray = self.nodes();
                var sortedArray = unsortedArray.sort(self.sortObjectsInArray);

                self.buildHierarchy(sortedArray);
                self.toggleView();
                addEventsToElements();

                ExecuteOrDelayUntilScriptLoaded(self.loadTitleOfRoot, "sp.js");
            }
        };

        self.loadTitleOfRoot = function () {
            var clientcontext = new SP.ClientContext(root);
            var currentWeb = clientcontext.get_web();
            clientcontext.load(currentWeb, 'Title');
            clientcontext.executeQueryAsync(
            function () {
                    $.each(self.nodes(), function (i, item) {
                        if (item.Title() == "Root") {
                            item.Title(currentWeb.get_title());
                            localStorage[nodeCacheKey] = ko.toJSON(self.nodes());
                        }
                    });
                }, null
            );
        };
    
        self.toggleView = function () {
            var navContainer = document.getElementById("navContainer");
            ko.applyBindings(self, navContainer);
            $("#loading").hide();
            $("#navContainer").show();

            var siteUrl = document.location.href.toLowerCase();
            var menuNode = $("#navContainer").find("a[href='" + siteUrl + "']");
            if ($(menuNode).length > 0) {
                $(menuNode).addClass("ms-core-listMenu-selected");
            } else {
                var pageIndex = siteUrl.indexOf("/sitepages");
                if (pageIndex === -1) pageIndex = siteUrl.indexOf("/pages");
                if (pageIndex === -1) pageIndex = siteUrl.indexOf("/lists");
                if (pageIndex === -1) pageIndex = siteUrl.indexOf("/_layouts/");
                if (pageIndex === -1) pageIndex = siteUrl.lastIndexOf("/");
        
                if (pageIndex != -1) {
                    siteUrl = siteUrl.substring(0, pageIndex);
                    var menuNode = $("#navContainer").find("a[href='" + siteUrl + "']");
                    while ($(menuNode).length > 0) {
                        $(menuNode).addClass("ms-core-listMenu-selected");
                        var ulNode = $(menuNode).closest("ul");
                        menuNode = $(ulNode).siblings("a.menu-item[href!='" + root.toLowerCase() + "']");
                    }
                }
            }
        };

        //Uses linq.js to build the navigation tree.
        self.buildHierarchy = function (enumerable) {
            self.hierarchy(Enumerable.from(enumerable).ByHierarchy(function (d) {
                return d.Parent() == null;
            }, function (parent, child) {
                if (parent.Url() == null || child.Parent() == null)
                    return false;
                return parent.Url().toUpperCase() == child.Parent().toUpperCase();
            }).toArray());

            self.sortChildren(self.hierarchy()[0]);
        };

        self.sortChildren = function (parent) {

            // skip processing if no children
            if (!parent || !parent.children || parent.children.length === 0) {
                return;
            }

            parent.children = parent.children.sort(self.sortObjectsInArray2);

            for (var i = 0; i < parent.children.length; i++) {
                var elem = parent.children[i];

                if (elem.children && elem.children.length > 0) {
                    self.sortChildren(elem);
                }
            }
        };

        self.sortObjectsInArray = function (a, b) {
            //sorting stuff
            if (isIEOrEdge) {
                if (a.Title() > b.Title()) return -1;
                if (a.Title() < b.Title()) return 1;
                return 0;
            } else {
                if (a.Title() > b.Title()) return 1;
                if (a.Title() < b.Title()) return -1;
                return 0;
            }
        };

        // ByHierarchy method breaks the sorting
        // we need to resort as ascending
        self.sortObjectsInArray2 = function (a, b) {
            return (self.sortObjectsInArray(a.item, b.item) * -1);
        };
    }

    //Loads the navigation on load and binds the event handlers for mouse interaction.
    function ContinueOnceJqueryAndKoLoaded() {
        if (window.jQuery && window.ko) {
            var viewModel = new NavigationViewModel();
            viewModel.loadNavigatioNodes();
        } else {
            window.setTimeout(ContinueOnceJqueryAndKoLoaded, 100);
        }
    }

    function InitCustomGlobalNavigation() {
        if (!window.jQuery) {
            var script = document.createElement("script");
            script.src = "https://code.jquery.com/jquery-1.11.2.min.js";
            document.head.appendChild(script);
        }
        if (!window.ko) {
            var script = document.createElement("script");
            script.src = "//ajax.aspnetcdn.com/ajax/knockout/knockout-2.2.0.js";
            document.head.appendChild(script);
        }

        ContinueOnceJqueryAndKoLoaded();
    }

    function addEventsToElements() {
        //events.
        $("li.dynamic-children")
            .mouseover(function () {
                var position = $(this).position();
                if ($(this).hasClass("dynamic")) {
                    $(this).children("ul").css({ "min-Width": 125, left: (position.left + $(this).width()) - 10, top: position.top + 10 });
                } else {
                    $(this).children("ul").css({ "min-Width": 125, left: position.left - 10, top: 38 });
                }
            })
            .mouseout(function () {
                $(this).children("ul").css({ left: -99999, top: 0 });
            });
    }


    // Public interface
    return {
        InitCustomGlobalNavigation: InitCustomGlobalNavigation
    }

} ();

(function() {
    _spBodyOnLoadFunctionNames.push("wvstrienCustomGlobalNav.ViewModel.InitCustomGlobalNavigation");
}());