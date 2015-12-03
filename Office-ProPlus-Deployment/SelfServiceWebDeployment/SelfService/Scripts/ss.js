var productToInstall = "";
var versionToInstall = "";
var languages = "";
var languageDictionary = {
    "English": "en-us",
    "Arabic": "ar-sa",
    "Bulgarian": "bg-bg",
    "Chinese (Simplified)": "zh-cn",
    "Chinese": "zh-tw",
    "Croatian": "hr-hr",
    "Czech": "cs-cz",
    "Croatian": "hr-hr",
    "Danish": "da-dk",
    "Estonian": "et-ee",
    "Finnish": "fi-fi",
    "French": "fr-fr",
    "German": "de-de",
    "Greek": "el-gr",
    "Hebrew": "he-il",
    "Hindi": "hi-in",
    "Hungarian": "hu-hu",
    "Indonesian": "id-id",
    "Italian": "it-it",
    "Japanese": "ja-jp",
    "Kazakh": "kk-kh",
    "Korean": "ko-kr",
    "Latvian": "lv-lv",
    "Lithuanian": "lt-lt",
    "Malay": "ms-my",
    "Norwegian (Bokm�l)": "nb-no",
    "Polish": "pl-pl",
    "Portuguese (Brazil)": "pt-br",
    "Portuguese (Portugal)": "pt-pt",
    "Romanian": "ro-ro",
    "Russian": "ru-ru",
    "Serbian (Latin)": "sr-latn-rs",
    "Slovak": "sk-sk",
    "Slovenian": "sl-si",
    "Spanish": "es-es",
    "Swedish": "sv-se",
    "Thai": "th-th",
    "Turkish": "tr-tr",
    "Ukrainian": "uk-ua",

    "en-us": "English",
    "ar-sa": "Arabic",
    "bg-bg": "Bulgarian",
    "zh-cn": "Chinese (Simplified)",
    "zh-tw": "Chinese",
    "hr-hr": "Croatian",
    "cs-cz": "Czech",
    "hr-hr": "Croatian",
    "da-dk": "Danish",
    "et-ee": "Estonian",
    "fi-fi": "Finnish",
    "fr-fr": "French",
    "de-de": "German",
    "el-gr": "Greek",
    "he-il": "Hebrew",
    "hi-in": "Hindi",
    "hu-hu": "Hungarian",
    "id-id": "Indonesian",
    "it-it": "Italian",
    "ja-jp": "Japanese",
    "kk-kh": "Kazakh",
    "ko-kr": "Korean",
    "lv-lv": "Latvian",
    "lt-lt": "Lithuanian",
    "ms-my": "Malay",
    "nb-no": "Norwegian (Bokm�l)",
    "pl-pl": "Polish",
    "pt-br": "Portuguese (Brazil)",
    "pt-pt": "Portuguese (Portugal)",
    "ro-ro": "Romanian",
    "ru-ru": "Russian",
    "sr-latn-rs": "Serbian (Latin)",
    "sk-sk": "Slovak",
    "sl-si": "Slovenian",
    "es-es": "Spanish",
    "sv-se": "Swedish",
    "th-th": "Thai",
    "tr-tr": "Turkish",
    "uk-ua": "Ukrainian"
};

var availableFilters = [];
var searchBoxTaggle;
var currentLocation;
var currentFilter;
var listView = 0;

var appliedFilters = [];
var previousSearch = "";

function setProduct(product, build) {
    buildID = build;
    getLanguages();
    productToInstall = product;
    $('#productSpan').text(product);
    showModal('languageModal');
}

function setVersion(version) {
    versionToInstall = version;
    $('#versionSpan').text(version);
}

function setLanguage() {
    var checkboxes = null;
    languages = null;
    checkboxes = $(".languageCheckBox:checked");
    $('#languageSpan').text(languageDictionary[checkboxes[0].id]);
    languages = [checkboxes[0].id];
    if (checkboxes.length > 1) {
        for (var i = 1; i < checkboxes.length; i++) {
            languages[i] = checkboxes[i].id;
            $('#languageSpan')[0].innerText += ", "+languageDictionary[checkboxes[i].id];
        }
    }
    showModal('confirmationModal');
}

function startInstall() {

}

function showModal(modalId) {
    $(".custom-Dialog").removeClass("hidden").addClass("hidden");
    $("#" + modalId).removeClass("hidden");
    if (modalId === "downloadModal") {
        $('#directDL').text(versionToInstall);
    }

    if (modalId === 'productModal')
    {
        resetFilters();
    }
}

function resetFilters() {
    searchBoxTaggle.removeAll();
    appliedFilters = [];
    $(searchBoxTaggle.getInput()).val('');
    searchBoxFilter();
    $('#ul-Location li:first').click();
}

function verifyLanguageInput() {
    sl = $('.languageCheckBox:checked');
    if (sl.length > 0) {
        $('#languageButton').prop('disabled', false);
    } else {
        $('#languageButton').prop('disabled', true);
    }
}

function getLanguages() {

    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {
                $('#languagesGrid li').remove();
                $xml = $(xml);
                var languages = $xml.find("[ID='" + buildID + "']").attr('Languages').split(",");
                $.each(languages, function (index, value) {
                    var label = value;
                    var id = value.split(" ").pop().replace(")",'').replace("(",'');
                    $('#languagesGrid > ul').append("<li class='languageli'><input type='checkbox' id='" + id + "' class='languageCheckBox' onclick='verifyLanguageInput()'/> \
                                    <label> <span class='ms-font-m checkboxLabel'>" + label + "</span></label></li>");
                });  
            }
    });
}

function getBuild() {

    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {

                if (listView === 1)
                {
                    $("#buildsTable").append("<div class='ms-Table-row'>\
                            <span class='ms-Table-cell custom-cell' style='padding-left:4%'>Name</span>\
                            <span class='ms-Table-cell custom-cell' style=''>Location</span>\
                            <span class='ms-Table-cell custom-cell'>Tags</span>\
                            <span class='ms-Table-cell custom-cell'></span>\
                            </div>");
                }

                $(xml).find('Build').each(function () {
                    var buildType = $(this).attr('Type');
                    var filters = $(this).attr('Filters').split(',');
                    var classString = "";
                    var textString = "";

                    if (listView === 1)
                    {
                       

                        if (Array.isArray(filters)) {
                            filters.forEach(function (element) {
                                classString += element.toLocaleLowerCase() + "-filter ";
                                textString +=   " "+element + ",";
                            });
                        } else {
                            if (filters) {
                                classString += filters + "-filter ";
                                textString += " " + element + ",";
                            }
                        }

                       
                        $("#buildsTable").append("<div class='ms-Table-row custom-table-row shown " + $(this).attr('Location').toLocaleLowerCase() + "-filter " + classString + "'>\
                            <span class='ms-Table-cell ms-font-l custom-first-cell custom-cell filter-field'><i class='ms-Icon ms-Icon--people package-people-table'></i>"+ buildType + "</span>\
                            <span class='ms-Table-cell custom-cell'>"+ $(this).attr('Location') + "</span>\
                            <span class='ms-Table-cell custom-cell'><i class='ms-Icon ms-Icon--tag custom-table-tag'></i>"+ textString + "</span>\
                            <span class='ms-Table-cell custom-cell custom-last-cell' onclick='setProduct(\"2016\",\""+ $(this).attr('ID') + "\")'><i class='ms-Icon ms-Icon--download custom-table-tag' ></i><a class='ms-link'>Install</a></span>\
                        </div>");

                    }
                    else
                    {
                        if (Array.isArray(filters)) {
                            filters.forEach(function (element) {
                                classString += element.toLocaleLowerCase() + "-filter ";
                                textString += "<li class='ms-font-m " + classString + "'>" + element + "</li>";
                            });
                        } else {
                            if (filters) {
                                classString += filters + "-filter ";
                                textString += "<li class='ms-font-m " + classString + "'>" + filters + "</li>";
                            }
                        }

                        $("#buildsGrid").append("<div class='ms-Grid-col package-group shown " + $(this).attr('Location').toLocaleLowerCase() + "-filter " + classString + "'>\
                                                <div id='custom-callout' class='ms-Callout ms-Callout--OOBE ms-Callout--arrowLeft hidden'>\
                                                    <div class='ms-Callout-main'>\
                                                        <div class='ms-Callout-header custom-callout-header'>\
                                                            <div class='ms-Callout-title ms-font-xxl ms-fontWeight-regular' >Tags</div>\
                                                            <i class='ms-Icon ms-Icon--x custom-x' onclick='closeCallout(event)'></i>\
                                                        </div>\
                                                        <div class='ms-Callout-inner custom-callout-inner'>\
                                                            <div class='ms-Callout-content'>\
                                                                <ul id='tags-list' class='tags-list'>"
                                                                + textString + "\
                                                                </ul>\
                                                            </div>\
                                                        </div>\
                                                    </div>\
                                                </div>\
                                                <div class='package package-main'>\
                                                     <div class='package-inner'>\
                                                        <span>\
                                                            <i class='ms-Icon ms-Icon--people package-people'></i>\
                                                            <i class='ms-Icon ms-Icon--tag package-tag' onclick='toggleCallout(event)'></i>\
                                                        </span>\
                                                        <p class='ms-font-xl package-label filter-field'>"+ buildType + "</b></p><br /><br />\
                                                        <p class='ms-font-s-plus package-label package-label-two' >"+ $(this).attr('Location') + "</p>\
                                                    </div>\
                                                    <span class='package-bottom ms-font-m' onclick='setProduct(\"2016\",\""+ $(this).attr('ID') + "\")'>\
                                                        <i class=' ms-Icon ms-Icon--download package-download'></i>\
                                                        <a class='ms-link'>Install</a>\
                                                    </span>\
                                                </div>\
                                                </div>");
                    }
                });
            }
    });
}

function toggleCallout(event) {
    $(event.target).parents().eq(3).find('#custom-callout').toggleClass('hidden');
}

function closeCallout(event)
{
    $(event.target).parents().eq(3).find('#custom-callout').addClass('hidden');
}

function getLocations(callback) {

    var locations = [];
    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {
                $("#ddl-Location").siblings('span.ms-Dropdown-title').text("Location Filter");
                $("#ddl-Location").siblings('ul').append("<li class='ms-Dropdown-item'>Location Filter</li>");

                $(xml).find('Build').each(function () {
                    var location = $(this).attr('Location');
                    if($.inArray(location,locations) === -1)
                    {
                        availableFilters.push(location);
                        locations.push(location);
                        $("#ddl-Location").siblings('ul').attr('id','ul-Location');
                        $("#ddl-Location").siblings('ul').append("<li class='ms-Dropdown-item'>" + location + "</li>");
                    }
                });
                updateAutocomplete();
                callback();
            }
    });
}

function getFilters() {

    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {
                $(xml).find('Build').each(function () {

                    var filter = $(this).attr('Filters').split(',');
                    if (Array.isArray(filter)) {
                        filter.forEach(function (element) {
                            if (availableFilters.indexOf(element.toLocaleLowerCase()) < 0) {
                                availableFilters.push(element.toLocaleLowerCase());
                            }
                        });
                    } else {
                        if (filter) {
                            if (availableFilters.indexOf(filter.toLocaleLowerCase()) < 0) {
                                availableFilters.push(filter.toLocaleLowerCase());
                            }
                        }
                    }
                });
                updateAutocomplete();
            }
    });
}

function getHelp() {

    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {
                $(xml).find('Help').find('Item').each(function () {

                    var qtext = $(this).find('Question').text();
                    var atext = $(this).find('Answer').text();
                    $('#helpContent').append(
                    "<div class='questionDiv'>\
                        <h4 class='ms-font-xl'>"+ qtext + "</h4>\
                        <p class='ms-font-m questionAnswer'>"
                            + atext +
                        "</p>\
                    </div>");
                });
            }
    });
}

function getCompanyInfo() {

    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {
                $(xml).find('Company').each(function () {
                    $('.companyName').text($(this).attr('Name'));
                    if ($(this).attr('LogoSrc')) {
                        $('.companyLogo').src($(this).attr('LogoSrc'));
                    } else {
                        $('.companyLogo').addClass('hidden');
                    }
                    
                });
            }
    });
}

function searchBoxFilter() {
    var searchTerm = searchBoxTaggle.getInput().value;
    searchTerm = searchTerm.toLocaleLowerCase();
    if (listView === 0) {
        $(".package-group").removeClass('search-filter');
        removeFilter("search");
        if (searchTerm) {

            $(".package-main p").each(function () {
                if ($(this).text().toLocaleLowerCase().indexOf(searchTerm) >= 0) {
                    $(this).parent().parent().parent().addClass('search-filter');
                }
            });


            $(".tags-list li").each(function () {

                if ($(this).text().toLocaleLowerCase().indexOf(searchTerm) >= 0) {
                    $(this).parent().parent().parent().parent().parent().parent().addClass('search-filter');
                }
            });


            addFilter("search");
        }
    }
    else {

        $(".custom-table-row").removeClass('search-filter');
        removeFilter("search");
        if (searchTerm) {

            $(".custom-table-row span").each(function () {
                if ($(this).text().toLocaleLowerCase().indexOf(searchTerm) >= 0) {
                    console.log($(this).text().toLocaleLowerCase());
                    $(this).parent().addClass('search-filter');
                }
            });

            addFilter("search");
        }

    
    }
   
    applyFilters();
}

function setTaggleFilters() {
    taggles = searchBoxTaggle.getTagValues();
    taggles.forEach(function (element) {
        addFilter(element);
    });
    applyFilters();
}

function locationFilter(location) {
    if (location) {
        removeFilter(currentLocation);
        if (location.toLocaleLowerCase().indexOf('filter') < 0) {
            addFilter(location);
        }
        applyFilters();
    }
    currentLocation = location;
}

function addFilter(filter) {
    if (appliedFilters.indexOf(filter) == -1) {
        appliedFilters.push(filter);
    }
}

function applyFilters() {

    if (listView === 0)
    {
        var filterString = ".package-group";
        appliedFilters.forEach(function (element) {
            filterString += "." + element + "-filter";
        });


        $(".package-group").addClass("hidden");
        $(filterString).removeClass("hidden").addClass('shown');
    }
    else
    {

        var filterString = ".custom-table-row";
        appliedFilters.forEach(function (element) {
            filterString += "." + element + "-filter";
        });


        $(".custom-table-row").addClass("hidden");
        $(filterString).removeClass("hidden").addClass('shown');
    }
    
}

function removeFilter(filter) {
    if (appliedFilters.indexOf(filter) >= 0) {
        appliedFilters.splice(appliedFilters.indexOf(filter), 1);
    }
}

function addLocationClick() {
    $('#ul-Location li').each(function () {
        $(this).attr('onclick', "locationFilter('"+$(this).text().toLocaleLowerCase()+"')");
    });
            }

function prepTags() {
    searchBoxTaggle = new Taggle('outerSearchBox',
        {
            saveOnBlur: true,
            placeholder: "Search",
            onTagAdd: function (event, tag) {
                $(searchBoxTaggle.getInput()).val('');
                addFilter(tag);
                applyFilters();
            },
            onTagRemove: function (event, tag) {
                removeFilter(tag);
                applyFilters();
            }
        });
    $('.taggle_placeholder').prepend('<i class="ms-SearchBox-icon ms-Icon ms-Icon--search"></i>');
}

function updateAutocomplete() {
    var container = searchBoxTaggle.getContainer();
    var input = searchBoxTaggle.getInput();
    searchBoxTaggle.settings.allowedTags = availableFilters;
    $(input).autocomplete({
        source: availableFilters,
        appendTo: container,
        position: { at: "left bottom", my: "left top" },
        select: function (event, data) {
            event.preventDefault();
            //Add the tag if user clicks
            if (event.which === 1) {
                searchBoxTaggle.add(data.item.value);
            }
        }
    });
}

function isListView() {
    listView = 1;
    $('#buildsTable').empty();
    $('#buildsGrid').empty();
    resetFilters();
    $('#tileViewToggle').attr('background-color', '#EFF6FC');
    $('#listViewToggle').attr('background-color', '#C7E0F4');
    getBuild(); 

}

function isTileView() {
    listView = 0;
    $('#buildsTable').empty();
    $('#buildsGrid').empty();
    $('#tileViewToggle').attr('background-color', '#C7E0F4');
    $('#listViewToggle').attr('background-color', '#EFF6FC');
    resetFilters();
    getBuild();
}

function focusDialog() {
    $('html,body').animate({

        scrollTop: $('.custom-mini-banner').offset().top + 500 
    }, 500);
    
}

function toggleBanner() {
    $('#banner').toggleClass('hidden');
    $('#mini-banner').toggleClass('hidden');
}

$(document).ready(function () {

    setVersion('2016');
    getCompanyInfo();
    getLocations(addLocationClick);
    getFilters();
    getBuild();
    getHelp();

    //searchbox filter
    $("#outerSearchBox").keyup(function (e) {
        searchBoxFilter(e);
    });

    //filter reset --REMOVE THIS--
    $('#btn-Reset').click(function () {
        resetFilters();
    });

 

    prepTags();
});