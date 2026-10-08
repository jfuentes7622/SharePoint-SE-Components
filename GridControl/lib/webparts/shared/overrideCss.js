"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var sp_webpart_base_1 = require("@microsoft/sp-webpart-base");
var pendingLoads = {};
var pageLoadVersion = String(new Date().getTime());
function applyOverrideCss(value, instanceId) {
    var elementId = 'spse-override-css-' + instanceId;
    if (pendingLoads[elementId]) {
        window.clearTimeout(pendingLoads[elementId]);
    }
    pendingLoads[elementId] = window.setTimeout(function () {
        var existing = document.getElementById(elementId);
        if (existing && existing.parentNode) {
            existing.parentNode.removeChild(existing);
        }
        var cssUrl = String(value || '').trim();
        if (!cssUrl) {
            delete pendingLoads[elementId];
            return;
        }
        var link = document.createElement('link');
        link.id = elementId;
        link.rel = 'stylesheet';
        link.type = 'text/css';
        var fragmentIndex = cssUrl.indexOf('#');
        var fragment = fragmentIndex >= 0 ? cssUrl.substring(fragmentIndex) : '';
        var urlWithoutFragment = fragmentIndex >= 0 ? cssUrl.substring(0, fragmentIndex) : cssUrl;
        link.href = urlWithoutFragment + (urlWithoutFragment.indexOf('?') >= 0 ? '&' : '?')
            + 'spseCssVersion=' + pageLoadVersion + fragment;
        link.onerror = function () {
            console.error('[OverrideCss] Unable to load stylesheet: ' + cssUrl);
        };
        (document.head || document.documentElement).appendChild(link);
        delete pendingLoads[elementId];
    }, 0);
}
exports.applyOverrideCss = applyOverrideCss;
function PropertyPaneOverrideCss(targetProperty, currentValue, context, onChange) {
    var pickerOverlay;
    var messageHandler;
    var closePicker = function () {
        if (messageHandler) {
            window.removeEventListener('message', messageHandler, false);
            messageHandler = undefined;
        }
        if (pickerOverlay && pickerOverlay.parentNode) {
            pickerOverlay.parentNode.removeChild(pickerOverlay);
        }
        pickerOverlay = undefined;
    };
    return {
        type: sp_webpart_base_1.PropertyPaneFieldType.Custom,
        targetProperty: targetProperty,
        properties: {
            key: 'overrideCssPicker-' + targetProperty,
            onRender: function (element) {
                element.innerHTML = '';
                var label = document.createElement('label');
                label.textContent = 'Override CSS';
                label.style.display = 'block';
                label.style.marginBottom = '4px';
                element.appendChild(label);
                var urlInput = document.createElement('input');
                urlInput.type = 'text';
                urlInput.value = currentValue || '';
                urlInput.placeholder = 'CSS URL or select a SharePoint .css file';
                urlInput.className = 'ms-TextField-field';
                urlInput.style.width = '100%';
                urlInput.onchange = function () {
                    onChange(urlInput.value);
                };
                element.appendChild(urlInput);
                var button = document.createElement('button');
                button.type = 'button';
                button.className = 'ms-Button ms-Button--default';
                button.textContent = 'Browse SharePoint';
                button.style.marginTop = '8px';
                button.onclick = function () {
                    closePicker();
                    pickerOverlay = document.createElement('div');
                    pickerOverlay.style.position = 'fixed';
                    pickerOverlay.style.left = '0';
                    pickerOverlay.style.top = '0';
                    pickerOverlay.style.right = '0';
                    pickerOverlay.style.bottom = '0';
                    pickerOverlay.style.zIndex = '1000000';
                    pickerOverlay.style.backgroundColor = 'rgba(0, 0, 0, 0.45)';
                    pickerOverlay.style.padding = '5vh 5vw';
                    pickerOverlay.style.boxSizing = 'border-box';
                    var pickerContainer = document.createElement('div');
                    pickerContainer.style.width = '100%';
                    pickerContainer.style.height = '100%';
                    pickerContainer.style.backgroundColor = '#ffffff';
                    pickerContainer.style.display = 'flex';
                    pickerContainer.style.flexDirection = 'column';
                    pickerOverlay.appendChild(pickerContainer);
                    var closeButton = document.createElement('button');
                    closeButton.type = 'button';
                    closeButton.className = 'ms-Button ms-Button--default';
                    closeButton.textContent = 'Close';
                    closeButton.style.alignSelf = 'flex-end';
                    closeButton.style.margin = '8px';
                    closeButton.onclick = closePicker;
                    pickerContainer.appendChild(closeButton);
                    var webUrl = context.pageContext.web.absoluteUrl;
                    var tenantUrl = webUrl.replace(context.pageContext.web.serverRelativeUrl, '');
                    var pickerUrl = webUrl + '/_layouts/15/onedrive.aspx?picker=';
                    pickerUrl += '%7B%22sn%22%3Afalse%2C%22v%22%3A%22files%22%2C%22id%22%3A%221%22%2C%22o%22%3A%22';
                    pickerUrl += encodeURI(tenantUrl);
                    pickerUrl += '%22%7D&id=' + encodeURI(context.pageContext.web.serverRelativeUrl);
                    pickerUrl += '&view=2&typeFilters=' + encodeURI('folder,.css') + '&p=2';
                    var iframe = document.createElement('iframe');
                    iframe.src = pickerUrl;
                    iframe.title = 'Select a CSS file from SharePoint';
                    iframe.style.border = '0';
                    iframe.style.width = '100%';
                    iframe.style.flex = '1 1 auto';
                    pickerContainer.appendChild(iframe);
                    messageHandler = function (event) {
                        if (event.origin !== window.location.origin || typeof event.data !== 'string') {
                            return;
                        }
                        var prefix = '[OneDrive-FromPicker]';
                        var prefixIndex = event.data.indexOf(prefix);
                        if (prefixIndex < 0) {
                            return;
                        }
                        try {
                            var result = JSON.parse(event.data.substring(prefixIndex + prefix.length));
                            if (result.type === 'cancel') {
                                closePicker();
                                return;
                            }
                            if (result.type === 'success' && result.items && result.items.length > 0
                                && result.items[0].sharePoint && result.items[0].sharePoint.url) {
                                var selectedUrl = String(result.items[0].sharePoint.url);
                                var normalizedUrl = selectedUrl.split('?')[0].split('#')[0].toLowerCase();
                                if (normalizedUrl.substr(normalizedUrl.length - 4) !== '.css') {
                                    console.error('[OverrideCss] The selected SharePoint file is not a CSS file: ' + selectedUrl);
                                    return;
                                }
                                urlInput.value = selectedUrl;
                                onChange(selectedUrl);
                                closePicker();
                            }
                        }
                        catch (error) {
                            console.error('[OverrideCss] Unable to process the SharePoint file picker response: ' + String(error));
                        }
                    };
                    window.addEventListener('message', messageHandler, false);
                    document.body.appendChild(pickerOverlay);
                };
                element.appendChild(button);
                var description = document.createElement('div');
                description.textContent = 'Saves the central SharePoint file URL and reloads its current contents on each page load.';
                description.style.fontSize = '12px';
                description.style.marginTop = '6px';
                element.appendChild(description);
            },
            onDispose: function () {
                closePicker();
            }
        }
    };
}
exports.PropertyPaneOverrideCss = PropertyPaneOverrideCss;

//# sourceMappingURL=overrideCss.js.map
