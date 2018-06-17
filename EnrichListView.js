
// Overload inplview.js functions, to restore dynamic filtering after navigating to another view (paging, sorting, ...)

function OverloadReRenderListView() {
    if ($.prototype.GridView_base__ReRenderListView === undefined) {
        $.prototype.GridView_base__ReRenderListView = ReRenderListView;
        ReRenderListView = function (a, k, o, f) {
           $.prototype.GridView_base__ReRenderListView(a, k, o, f);
           EnrichListView.ViewModel.AddDynamicFilteringToListView();
       }
    }
}

//Models and Namespaces
var EnrichListView = EnrichListView || {};

EnrichListView.ViewModel = function () {

    function AddDynamicFilteringToListView() {
        viewMode = $(".ms-listviewgrid").length === 1 ? "gridview" : "listview";

        var $table = $(IdentifyingClassListView());

        // On page-load, the AddDynamicFilteringToListView() is already invoked. Don't redo in case later the rendering is still done.
        if (!$table.children("tbody").hasClass("list")) {
            // Inject filter-row within header.
            // Based on the header columns: add class to every data-cell; and add filter-cell
            var $header = $table.find(".ms-viewheadertr");

            var valueNames = [];
            $header.children("th").each(function() {
                var header = $(this).text().replace(/ /g, '');
                if (header != '') {
                    var ownerIndex = $(this).index();   
                    var filterClass =  header + "_" + ownerIndex;

                    $(IdentifyingClassListView() + ' td:nth-child(' + (ownerIndex+1) + ')').addClass(filterClass);  

                    // Do not re-add in header in case still present (on re-rendering; tbody is renewed but thead remains
                    if ($(this).find(".search").length === 0) {
                        $(this).append($('<div style="clear:both;height=2px;"></div>'));
                        $filterInput = $('<input type="text" id="' + filterClass + '" class="search" placeholder="Filter" style="width:95%;"></input>');
                        $(this).append($filterInput.wrap('<div style="position:relative;left:0px;"></div>'));  
 
                        // Need to capture the event, to overload inplview handling
                        $filterInput.click(function(e) {
                            $(this).focus();
                            e.stopPropagation();
                        });
                    }

   		    valueNames.push(filterClass);
                }
            });

            if ($table.find("tbody[groupstring]").length > 0) {
                // In case of grouped listview; the table contains multiple tbody elements. Each must individual by connected to a List element
                var tbodySeq = 0;
 	        $table.children("tbody[isLoaded]").each(function() {
                    if ($(this).find("td").length > 0) {
                        var uniqueClass = "list_" + tbodySeq;
                        $(this).addClass(uniqueClass );
                        var filterList = new List($table.attr("id"), { valueNames: valueNames, listClass: uniqueClass, page: 2000 }); 
                    }
                    tbodySeq = tbodySeq + 1;          
                });
            } else {
                if (viewMode.toLowerCase() === "gridview") {
                    $('<thead></thead>').insertBefore($table.children("tbody"));
                    $header.detach();
                    $header.appendTo($table.children("thead"));
                }
            
                // Declare the body of table as the 'island' to be subject for filtering
                $table.find("tbody").addClass("list");

                var filterList = new List($table.attr("id"), { valueNames: valueNames, page: 2000 });

                // Standard 'List' behavior hides all rows not matching the filter. In case of gridview, this is undesired for the last row, which is used to
                // potential add new entries.
                if (viewMode.toLowerCase() === "gridview") {
                    filterList.on("searchComplete", function() {
                        var $lastTr = $(IdentifyingClassListView()).find(".list").children("tr").last();
                        $lastTr.show();                                    
                    });
                }
            }
        }
    }

    //Loads the navigation on load and binds the event handlers for mouse interaction.
    function ContinueOnceAllDataLoaded() {
        viewMode = $(".ms-listviewgrid").length === 1 ? "gridview" : "listview";

        var $table = $(IdentifyingClassListView());
        if ($table.length === 1) {
            // Race-condition: can be that gridview is already rendered on first page-load; therefore add explicit.
            // For re-render scenarios; handled by adding to the inplview.js handling
            AddDynamicFilteringToListView();

            // If no grouping, ReRenderListView called. Otherwise, the rendering is done per tbody element via clienttemplates rendering.
            if ($table.find("tbody[groupstring]").length == 0) {   
                ExecuteOrDelayUntilScriptLoaded(OverloadReRenderListView, "inplview.js");
            } else {
                var viewId = "{" + $table.attr("view") + "}";
                var viewCounter = window["ctx" + g_ViewIdToViewCounterMap[viewId]];
                SPClientRenderer.AddPostRenderCallback(viewCounter, function() { 
                    EnrichListView.ViewModel.AddDynamicFilteringToListView();
                    var $table = $(EnrichListView.ViewModel.IdentifyingClassListView());
                    var $nonEmptyFilters = $table.find(".search").filter(function () { return this.value.length > 0 });
                    if ($nonEmptyFilters.length > 0) {
                        var evt;
                        try {
                            evt = new KeyboardEvent("keyup");
                        } catch (e) {
                            evt = document.createEvent('KeyboardEvent');
                            evt.initEvent('keyup', true, false);
                        }
                        $nonEmptyFilters[0].dispatchEvent(evt);
                    }
                });
            }
        } else {
            window.setTimeout(ContinueOnceAllDataLoaded, 100);
        }
    }

    //Loads the navigation on load and binds the event handlers for mouse interaction.
    function ContinueOnceAllDependentLibsLoaded() {
        if (window.jQuery && typeof(List) != 'undefined') {
            ContinueOnceAllDataLoaded();
        } else {
            window.setTimeout(ContinueOnceAllDependentLibsLoaded, 100);
        }
    }

    function InitEnrichment() {
        if (!window.jQuery) {
            var script = document.createElement("script");
            script.src = "https://code.jquery.com/jquery-1.11.2.min.js";
            document.head.appendChild(script);
        }
        if (typeof(List) == 'undefined') {
            var script = document.createElement("script");
            script.src = "//<personal-CDN>/CustomizeListView/List.min.js";
            document.head.appendChild(script);
        }

        ContinueOnceAllDependentLibsLoaded();
    }

    // Default ListView
    var viewMode = "ListView";

    function IdentifyingClassListView() {
        if (viewMode.toLowerCase() === "gridview") {
            return ".ms-listviewgrid";
        }

        // default
        return ".ms-listviewtable";
    }

    // Public interface
    return {
        InitEnrichment: InitEnrichment,
        AddDynamicFilteringToListView: AddDynamicFilteringToListView,
        IdentifyingClassListView: IdentifyingClassListView
    }
} ();

(function() {
    _spBodyOnLoadFunctionNames.push("EnrichListView.ViewModel.InitEnrichment");
}());