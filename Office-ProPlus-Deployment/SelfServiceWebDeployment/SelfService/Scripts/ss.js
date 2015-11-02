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
    showLanguageModal();
}

function setVersion(version) {
    versionToInstall = version;
    $('#versionSpan')[0].innerText = version;
    showProductModal();
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
    showConfirmationModal();
}

function startInstall() {

}

function showProductModal() {
    $("#productModal")[0].style.display = "block";
    $("#versionModal")[0].style.display = "none";
    $("#languageModal")[0].style.display = "none";
    $('#confirmationModal')[0].style.display = "none";
    $('#downloadModal')[0].style.display = "none";
}

function showLanguageModal() {
    $("#productModal")[0].style.display = "none";
    $("#versionModal")[0].style.display = "none";
    $("#languageModal")[0].style.display = "block";
    $('#confirmationModal')[0].style.display = "none";
    $('#downloadModal')[0].style.display = "none";
}

function showVersionModal() {
    $("#helpModal")[0].style.display = "none";
    $("#productModal")[0].style.display = "none";
    $("#versionModal")[0].style.display = "block";
    $("#languageModal")[0].style.display = "none";
    $('#confirmationModal')[0].style.display = "none";
    $('#downloadModal')[0].style.display = "none";
}

function showConfirmationModal() {
    $("#productModal")[0].style.display = "none";
    $("#versionModal")[0].style.display = "none";
    $("#languageModal")[0].style.display = "none";
    $('#confirmationModal')[0].style.display = "block";
    $('#downloadModal')[0].style.display = "none";

}

function showDownloadModal() {
    $("#productModal")[0].style.display = "none";
    $("#versionModal")[0].style.display = "none";
    $("#languageModal")[0].style.display = "none";
    $('#confirmationModal')[0].style.display = "none";
    $('#downloadModal')[0].style.display = "block";

    $('#directDL').text(versionToInstall);
}

function getLanguages() {

    $.ajax({
        type: "GET",
        url: "SelfServiceConfig.xml",
        datatype: "xml",
        success:
            function (xml) {
                $(xml).find('Language').each(function () {
                    $('#languagesGrid > .ms-Grid-row').append("<div class='ms-Grid-col ms-u-lg4 ms-u-xl4 ms-u-md4'> <input type='checkbox' id='" + $(this).attr('ID') + "' class='languageCheckBox' /> \
                                    <span class='ms-Label checkboxLabel'>" + $(this).attr('Label') + "</span></div>");
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
                $(xml).find('Version').each(function () {
                    var version = $(this).attr('ID');
                    if (version === "2013") {
                        $('#versions').prepend("<li class='squareButton'>\
                                        <button class='ms-Dialog-action ms-Button ms-Button--primary ms-bgColor-orangeLight' onclick='setVersion(\"2013\")' style='width:225px;height:250px;'>\
                                        <img src='Content/imgs/office-icon-white.png' style='height:100px'/>\
                                        <p class='ms-font-xl ms-fontColor-white' style='display:block'>2013</p>\
                                        </button>\
                                        </li>");
                    }
                    if (version === "2016") {
                        $('#versions').prepend("<li class='squareButton'>\
                                    <button class='ms-Dialog-action ms-Button ms-Button--primary' onclick='setVersion(\"2016\")' style='width:225px;height:250px;'>\
                                    <img src='Content/imgs/office-icon-white.png' style='height:100px'/>\
                                    <p class='ms-font-xl ms-fontColor-white' style='display:block'>2016</p>\
                                    </button>\
                                    </li>")
                    }
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
                $(xml).find('Build').each(function () {
                    var buildType = $(this).attr('Type');
                    console.log("Type: " + buildType);
                    $("#buildsGrid").append("<li class='squareButton'>\
                                    <button class='ms-Dialog-action ms-Button' onclick='setProduct(versionToInstall)' style='width:225px;height:250px;'>\
                                    <i class='ms-Icon ms-Icon--people' style='font-size:125px'></i>\
                                    <p class='ms-font-xl' style='display:block'>" + $(this).attr('Type') + "</p>\
                                    </button>\
                                    </li>");
                });
            }
    });
}

$(document).ready(function () {
    getVersions();
    getBuild();
    getLanguages();
});