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
}
var availableLanguages;
function setProduct(product) {
    productToInstall = product;
    $('#productSpan')[0].innerText = product;
    showModal('languageModal');
}

function setVersion(version) {
    versionToInstall = version;
    $('#versionSpan')[0].innerText = version;
    showModal('productModal');
}

function setLanguage() {
    var checkboxes = null;
    languages = null;
    checkboxes = $(".languageCheckBox:checked");
    $('#languageSpan')[0].innerText = languageDictionary[checkboxes[0].id];
    languages = [checkboxes[0].id];
    if (checkboxes.length > 1) {
        for (var i = 1; i < checkboxes.length; i++) {
            languages[i] = checkboxes[i].id
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
    if (modalId == "downloadModal") {
        $('#directDL').text(versionToInstall);
    }
}

function getLanguages() {

    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {
                $(xml).find('Language').each(function () {
                    $('#languagesGrid > .ms-Grid-row').append("<div class='ms-Grid-col ms-u-lg4 ms-u-xl4 ms-u-md4'><label> <input type='checkbox' id='" + $(this).attr('ID') + "' class='languageCheckBox' /> \
                                    <span class='ms-Label checkboxLabel'>" + $(this).attr('Label') + "</span></label></div>");
                });
            }
    });
}

function getVersions() {

    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {
                var optional = $(xml).find('Versions').attr('Optional');
                console.log(optional);

                if (optional === "true") {
                    $(xml).find('Version').each(function () {
                        var version = $(this).attr('ID');
                        if (version === "2013") {
                            $('#versions').prepend("<li class='squareButton'>\
                                        <div class='ms-Dialog-action ms-Button ms-Button--primary ms-bgColor-orangeLight' onclick='setVersion(\"2013\")' style='width:225px;height:250px;padding:50px'>\
                                        <img src='Content/imgs/office-icon-white.png' style='height:100px'/>\
                                        <p class='ms-font-xl ms-fontColor-white' style='display:block'>2013</p>\
                                        </div>\
                                        </li>");
                        }
                        if (version === "2016") {
                            $('#versions').prepend("<li class='squareButton'>\
                                    <div class='ms-Dialog-action ms-Button ms-Button--primary' onclick='setVersion(\"2016\")' style='width:225px;height:250px;padding:50px'>\
                                    <img src='Content/imgs/office-icon-white.png' style='height:100px'/>\
                                    <p class='ms-font-xl ms-fontColor-white' style='display:block'>2016</p>\
                                    </div>\
                                    </li>")
                        }
                    });
                }
                else
                {
                    $('#versions').prepend("<li class='squareButton'>\
                                    <div class='ms-Dialog-action ms-Button ms-Button--primary' onclick='setVersion(\"Office 365 ProPlus\")' style='text-align:center;width:225px;height:250px;max-height:250px;padding-top:50px;'>\
                                    <img src='Content/imgs/office-icon-white.png' style='height:100px;'/>\
                                    <p class='ms-font-xl ms-fontColor-white' style='display:block;r'>Office 365 ProPlus</p>\
                                    </div>\
                                    </li>")
                }
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
                $(xml).find('Build').each(function () {
                    var buildType = $(this).attr('Type');
                    $("#buildsGrid").append("<li class='squareButton-build'>\
                                    <button class='ms-Dialog-action ms-Button' onclick='setProduct(versionToInstall)' style='width:225px;height:250px;'>\
                                    <i class='ms-Icon ms-Icon--people' style='font-size:125px'></i>\
                                    <p class='ms-font-xl' style='display:block'>" + $(this).attr('Type') + "</p>\
                                    <div id='tag' class='ms-font-md' style='display:block:padding-bottom:2px;'>Tags: " + $(this).attr('Location') +","+$(this).attr('FilterOne') + "</div>\
                                    </button>\
                                    </li>");
                });
            }
    });
}


function getLocations() {

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
                        locations.push(location);
                        $("#ddl-Location").siblings('ul').append("<li class='ms-Dropdown-item'>" + location + "</li>");
                    }
                });
            }
    });
}

function getFilterOne() {

    var filters = [];
    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {

                var filterOneLabel = $(xml).find('FilterOne').attr('Filter');
                $("#ddl-FilterOne").siblings('span.ms-Dropdown-title').text(filterOneLabel + " Filter");
                $("#ddl-FilterOne").siblings('ul').append("<li class='ms-Dropdown-item'>" + filterOneLabel + " Filter</li>");
                $(xml).find('Build').each(function () {

                    var filter= $(this).attr('FilterOne');
                    if ($.inArray(filter, filters) === -1) {
                        filters.push(filter);
                        $("#ddl-FilterOne").siblings('ul').append("<li class='ms-Dropdown-item'>" + filter + "</li>");
                    }
                });
            }
    });
}
$(document).ready(function () {
    getVersions();
    getBuild();
    getLanguages();
    getLocations();
    getFilterOne();
});