/**
 * Check if url belongs to SharePoint site.
 * @param {} url 
 */
function isSharePointSite(url) {
    return url.indexOf(".sharepoint.com/") !== -1;
}

browser.tabs.onUpdated.addListener((tabId, changeInfo, tabInfo) => {
    if (isSharePointSite(tabInfo.url)) {
        browser.browserAction.setIcon({
            tabId: tabId,
            path: "icons/LogoTest1_48.png"
        });
    }
});