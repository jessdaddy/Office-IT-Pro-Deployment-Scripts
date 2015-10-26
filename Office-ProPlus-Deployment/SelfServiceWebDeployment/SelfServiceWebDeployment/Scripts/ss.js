var productToInstall = "";
var versionToInstall = "";
var languages = "";

function setProduct(product) {
    productToInstall = product;
    showLanguageModal();
}

function setVersion(version) {
    versionToInstall = version;
    showProductModal();
}

function setLanguage() {
    checkboxes = $(".languageCheckBox:checked");
    languages = [checkboxes[0].id];
    if (checkboxes.length > 1) {
        for (var i = 1; i < checkboxes.length; i++) {
            languages[i] = checkboxes[i].id
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
}

function showLanguageModal() {
    $("#productModal")[0].style.display = "none";
    $("#versionModal")[0].style.display = "none";
    $("#languageModal")[0].style.display = "block";
    $('#confirmationModal')[0].style.display = "none";
}

function showVersionModal() {
    $("#productModal")[0].style.display = "none";
    $("#versionModal")[0].style.display = "block";
    $("#languageModal")[0].style.display = "none";
    $('#confirmationModal')[0].style.display = "none";
}

function showConfirmationModal() {
    $("#productModal")[0].style.display = "none";
    $("#versionModal")[0].style.display = "none";
    $("#languageModal")[0].style.display = "none";
    $('#confirmationModal')[0].style.display = "block";
}