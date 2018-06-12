
// Overload inplview.js functions, to restore dynamic filtering after navigating to another view (paging, sorting, ...)

function OverloadPostRenderAfterJSGridRender() {
    if ($.prototype.GridView_base__PostRenderAfterJSGridRender === undefined) {
        $.prototype.GridView_base__PostRenderAfterJSGridRender = PostRenderAfterJSGridRender;
        PostRenderAfterJSGridRender = function (a) {
           $.prototype.GridView_base__PostRenderAfterJSGridRender(a);
           EnrichListView.ViewModel.AddDynamicFilteringToListView();
       }
    }
}

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

            // Declare the body of table as the 'island' to be subject for filtering
            $table.find("tbody").addClass("list");

            if (viewMode.toLowerCase() === "gridview") {
                $('<thead></thead>').insertBefore($table.children("tbody"));
                $header.detach();
                $header.appendTo($table.children("thead"));
            }

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

    //Loads the navigation on load and binds the event handlers for mouse interaction.
    function ContinueOnceAllDataLoaded() {
        var $table = $(IdentifyingClassListView());
        if ($table.length === 1) {
            // Race-condition: can be that gridview is already rendered on first page-load; therefore add explicit.
            // For re-render scenarios; handled by adding to the inplview.js handling
            AddDynamicFilteringToListView();

            if (viewMode.toLowerCase() === "gridview") {
                // Overload PostRender function to re-apply adding dynamic filtering in case of user-initiated modifications of grid via menu 
                // (sort, filter on a value)
                ExecuteOrDelayUntilScriptLoaded(OverloadPostRenderAfterJSGridRender, "inplview.js");
            } else {
                ExecuteOrDelayUntilScriptLoaded(OverloadReRenderListView, "inplview.js");
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
            script.src = "//<personal-CDN>/CustomizeListView/List.js";
            document.head.appendChild(script);
        }

        ContinueOnceAllDependentLibsLoaded();
    }

    // Default ListView
    var viewMode = "ListView";
    function SetViewMode(mode) {
        viewMode = mode;
    }

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
        SetViewMode: SetViewMode,
        IdentifyingClassListView: IdentifyingClassListView
    }
} ();

(function() {
    _spBodyOnLoadFunctionNames.push("EnrichListView.ViewModel.InitEnrichment");
}());