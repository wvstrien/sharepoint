//Models and Namespaces
var wvstrienCustomDisplayGlobalNav = wvstrienCustomDisplayGlobalNav || {};
wvstrienCustomDisplayGlobalNav.Models = wvstrienCustomDisplayGlobalNav.Models || {}
wvstrienCustomDisplayGlobalNav.Models.NavigationNode = function () {
    this.Url = ko.observable("");
    this.Title = ko.observable("");
};

wvstrienCustomDisplayGlobalNav.ViewModel = function () {

    /* create a new OData request for JSON response */
    function getRequest(endpoint) {
		var request = {
			type: "GET",
			url: endpoint,
			headers: { ACCEPT: "application/json;odata=verbose" }
		}
        return request;
    };

    /* Navigation Module*/
    function NavigationViewModel() {
        "use strict";
        var self = this;

        var root;
        if (_spPageContextInfo) {
            root = _spPageContextInfo.siteAbsoluteUrl;
        } else {
           root = document.location.href.match(/https\:\/\/.[^\/]+\/.[^\/]+\/.[^\/]+\//)[0];
           root = root.substring(0, root.length - 1);
        }
		
		var tenantUrl = root.match(/https\:\/\/.[^\/]+\//)[0].toLowerCase();
        tenantUrl = tenantUrl.substr(0, tenantUrl.length - 1);

		var cachedGlobalNavigation;

        self.hierarchy = ko.observableArray([]);;
        self.loadNavigatioNodes = function () {
            if (document.location.href.match(/refreshnavigation/) === null) {
				ExecuteOrDelayUntilScriptLoaded(self.retrieveFromCacheViaWebProperty, "sp.js");
			} else {
				//No cache hit, REST call required.
				self.queryRemoteInterface();
			}
        };
		
		self.retrieveFromCacheViaWebProperty = function() {
			clientContext = new SP.ClientContext.get_current();
			rootWeb = clientContext.get_site().get_rootWeb();
			clientContext.load(rootWeb);
			rootWebProperties = rootWeb.get_allProperties();
			clientContext.load(rootWebProperties);
			clientContext.executeQueryAsync(self.onWebPropertyQuerySucceeded, self.queryRemoteInterface);
		};
		
		self.onWebPropertyQuerySucceeded = function() {
			var allRootWebProperties = rootWebProperties.get_fieldValues();
			self.cachedGlobalNavigation = allRootWebProperties["CachedGlobalNavigation"];
			
			if (
				self.cachedGlobalNavigation !== null &&
				!document.location.pathname.toLowerCase().endsWith("/_layouts/15/settings.aspx")
			) {
				self.renderNavigationNodes();
			} else {
				// Either not cached in webproperty, or
				// potential get back here after Navigation-Settings. Check whether need to 
				// update the cached value of navigation for non-owner visitors of this site.
				self.queryRemoteInterface();
			}
		};
			
		self.renderNavigationNodes = function() {
			if (self.cachedGlobalNavigation !== null) { 
				var cachedNodes = JSON.parse(self.cachedGlobalNavigation);
				if (cachedNodes) {			
					self.hierarchy(cachedNodes);
		
					self.toggleView();
					addEventsToElements();
				}
			}
		};
		
		self.queryRemoteInterface = function () {
			self.getGlobalNavigationLinksFromProvider().then(
				function(results) {
					var queriedResults = ko.toJSON(results);			
					if (self.cachedGlobalNavigation != queriedResults) {				
					    self.cachedGlobalNavigation = queriedResults;
						
						if (document.location.pathname.toLowerCase().endsWith("/_layouts/15/settings.aspx")) {
							clientContext = new SP.ClientContext.get_current();
							rootWeb = clientContext.get_site().get_rootWeb();
							clientContext.load(rootWeb);
							rootWebProperties = rootWeb.get_allProperties();
							rootWebProperties.set_item("CachedGlobalNavigation", self.cachedGlobalNavigation);
							rootWeb.update();
							clientContext.load(rootWeb);
							clientContext.executeQueryAsync(null, null);
						}
					}
					self.renderNavigationNodes();
				}
			);
		}
			
		self.parseGlobalNavigationLinksFromProviderMenuStateNodes = function(nodes, results) {
			$.each(nodes, function (i, item) {
				if (!item.IsHidden) {
					var navItem = new wvstrienCustomDisplayGlobalNav.Models.NavigationNode();
					navItem.Title(item.Title);
					 								
					var itemUrl = item.SimpleUrl.toLowerCase().replace(/ /g, "%20");
					if (!itemUrl.startsWith("http")) {
						itemUrl = tenantUrl + itemUrl;
					}
					navItem.Url(itemUrl);
					
					var navItemItem = {
						item: navItem,
						children: []
					};
					results.push(navItemItem);
					if (item.Nodes && item.Nodes.results && item.Nodes.results.length > 0) {
  					    self.parseGlobalNavigationLinksFromProviderMenuStateNodes(item.Nodes.results, navItemItem.children);
					}
				}
			});			
		}
		
        self.getGlobalNavigationLinksFromProvider = function () {
          return new Promise(function (resolve, reject) {
            var globalMapProviderUrl = root + "/_api/navigation/menustate?mapprovidername=\'GlobalNavigationSwitchableProvider\'";
            var globalMapProviderRequest = getRequest(globalMapProviderUrl);
            $.ajax(globalMapProviderRequest).done(function (data) {
				var results = [];
				var rootItem = new wvstrienCustomDisplayGlobalNav.Models.NavigationNode();
				if (data.d ) {				 
					rootItem.Title(data.d.MenuState.StartingNodeTitle);
					rootItem.Url(tenantUrl + data.d.MenuState.SimpleUrl.toLowerCase().replace(/ /g, "%20"));
					 
					var rootItemItem = {
				 	    item: rootItem,
						children: []
					};
					results.push(rootItemItem);
					if (data.d.MenuState && data.d.MenuState.Nodes && data.d.MenuState.Nodes.results && data.d.MenuState.Nodes.results.length > 0) {
					    self.parseGlobalNavigationLinksFromProviderMenuStateNodes(data.d.MenuState.Nodes.results, rootItemItem.children);
					}
				}
 				 
                resolve(results);
             }).fail(function () {
                 resolve([]);
             });
          });
        };		
		
        self.toggleView = function () {
            var navContainer = document.getElementById("navContainer");
			ko.cleanNode(navContainer);
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
    }

    //Loads the navigation on load and binds the event handlers for mouse interaction.
    function ContinueOnceAllDependentLibsLoaded() {
        if (window.jQuery && window.ko && window.Promise) {
            var viewModel = new NavigationViewModel();
            viewModel.loadNavigatioNodes();
        } else {
            window.setTimeout(ContinueOnceAllDependentLibsLoaded, 100);
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
        if (!window.Promise) {
            var script = document.createElement("script");
            script.src = "//cdnjs.cloudflare.com/ajax/libs/bluebird/3.3.4/bluebird.min.js";
            document.head.appendChild(script);
        }
        ContinueOnceAllDependentLibsLoaded();
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
    _spBodyOnLoadFunctionNames.push("wvstrienCustomDisplayGlobalNav.ViewModel.InitCustomGlobalNavigation");
}());