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

$(document).ready(function () {

    $.ajax({
        type: "GET",
        url: "/SelfService/Languages",
        success: function onSuccess(result) {
            console.log("success");
            availableLanguages = result.Split(";");
        }
    });
    $.ajax({
        type: "GET",
        url: "/SelfService/Versions",
        success: function onSuccess(result) {
            console.log("success");
            availableLanguages = result.Split(";");
        }
    });
    $.ajax({
        type: "GET",
        url: "/SelfService/Products",
        success: function onSuccess(result) {
            console.log("success");
            availableLanguages = result.Split(";");
        }
    });

});