var productToInstall = "";
var versionToInstall = "";
var languages = "";

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
    checkboxes = $(".languageCheckBox:checked");
    $('#productSpan')[0].innerText = checkboxes[0].id;
    languages = [checkboxes[0].id];
    if (checkboxes.length > 1) {
        for (var i = 1; i < checkboxes.length; i++) {
            languages[i] = checkboxes[i].id
            $('#languageSpan')[0].innerText += ", "+checkboxes[i].id;
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