import {
  IPropertyPaneCustomFieldProps,
  IPropertyPaneField,
  PropertyPaneFieldType,
  WebPartContext
} from '@microsoft/sp-webpart-base';

const pendingLoads: { [id: string]: number } = {};
const pageLoadVersion = String(new Date().getTime());

export function applyOverrideCss(value: string, instanceId: string): void {
  const elementId = 'spse-override-css-' + instanceId;
  if (pendingLoads[elementId]) {
    window.clearTimeout(pendingLoads[elementId]);
  }

  pendingLoads[elementId] = window.setTimeout((): void => {
    const existing = document.getElementById(elementId);
    if (existing && existing.parentNode) {
      existing.parentNode.removeChild(existing);
    }

    const cssUrl = String(value || '').trim();
    if (!cssUrl) {
      delete pendingLoads[elementId];
      return;
    }

    const link = document.createElement('link');
    link.id = elementId;
    link.rel = 'stylesheet';
    link.type = 'text/css';
    const fragmentIndex = cssUrl.indexOf('#');
    const fragment = fragmentIndex >= 0 ? cssUrl.substring(fragmentIndex) : '';
    const urlWithoutFragment = fragmentIndex >= 0 ? cssUrl.substring(0, fragmentIndex) : cssUrl;
    link.href = urlWithoutFragment + (urlWithoutFragment.indexOf('?') >= 0 ? '&' : '?')
      + 'spseCssVersion=' + pageLoadVersion + fragment;
    link.onerror = (): void => {
      console.error('[OverrideCss] Unable to load stylesheet: ' + cssUrl);
    };
    (document.head || document.documentElement).appendChild(link);
    delete pendingLoads[elementId];
  }, 0);
}

export function PropertyPaneOverrideCss(
  targetProperty: string,
  currentValue: string,
  context: WebPartContext,
  onChange: (newValue: string) => void
): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  let pickerOverlay: HTMLElement | undefined;
  let messageHandler: ((event: MessageEvent) => void) | undefined;

  const closePicker = (): void => {
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
    type: PropertyPaneFieldType.Custom,
    targetProperty: targetProperty,
    properties: {
      key: 'overrideCssPicker-' + targetProperty,
      onRender: (element: HTMLElement): void => {
        element.innerHTML = '';

        const label = document.createElement('label');
        label.textContent = 'Override CSS';
        label.style.display = 'block';
        label.style.marginBottom = '4px';
        element.appendChild(label);

        const urlInput = document.createElement('input');
        urlInput.type = 'text';
        urlInput.value = currentValue || '';
        urlInput.placeholder = 'CSS URL or select a SharePoint .css file';
        urlInput.className = 'ms-TextField-field';
        urlInput.style.width = '100%';
        urlInput.onchange = (): void => {
          onChange(urlInput.value);
        };
        element.appendChild(urlInput);

        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'ms-Button ms-Button--default';
        button.textContent = 'Browse SharePoint';
        button.style.marginTop = '8px';
        button.onclick = (): void => {
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

          const pickerContainer = document.createElement('div');
          pickerContainer.style.width = '100%';
          pickerContainer.style.height = '100%';
          pickerContainer.style.backgroundColor = '#ffffff';
          pickerContainer.style.display = 'flex';
          pickerContainer.style.flexDirection = 'column';
          pickerOverlay.appendChild(pickerContainer);

          const closeButton = document.createElement('button');
          closeButton.type = 'button';
          closeButton.className = 'ms-Button ms-Button--default';
          closeButton.textContent = 'Close';
          closeButton.style.alignSelf = 'flex-end';
          closeButton.style.margin = '8px';
          closeButton.onclick = closePicker;
          pickerContainer.appendChild(closeButton);

          const webUrl = context.pageContext.web.absoluteUrl;
          const tenantUrl = webUrl.replace(context.pageContext.web.serverRelativeUrl, '');
          let pickerUrl = webUrl + '/_layouts/15/onedrive.aspx?picker=';
          pickerUrl += '%7B%22sn%22%3Afalse%2C%22v%22%3A%22files%22%2C%22id%22%3A%221%22%2C%22o%22%3A%22';
          pickerUrl += encodeURI(tenantUrl);
          pickerUrl += '%22%7D&id=' + encodeURI(context.pageContext.web.serverRelativeUrl);
          pickerUrl += '&view=2&typeFilters=' + encodeURI('folder,.css') + '&p=2';

          const iframe = document.createElement('iframe');
          iframe.src = pickerUrl;
          iframe.title = 'Select a CSS file from SharePoint';
          iframe.style.border = '0';
          iframe.style.width = '100%';
          iframe.style.flex = '1 1 auto';
          pickerContainer.appendChild(iframe);

          messageHandler = (event: MessageEvent): void => {
            if (event.origin !== window.location.origin || typeof event.data !== 'string') {
              return;
            }
            const prefix = '[OneDrive-FromPicker]';
            const prefixIndex = event.data.indexOf(prefix);
            if (prefixIndex < 0) {
              return;
            }

            try {
              const result = JSON.parse(event.data.substring(prefixIndex + prefix.length));
              if (result.type === 'cancel') {
                closePicker();
                return;
              }
              if (result.type === 'success' && result.items && result.items.length > 0
                && result.items[0].sharePoint && result.items[0].sharePoint.url) {
                const selectedUrl = String(result.items[0].sharePoint.url);
                const normalizedUrl = selectedUrl.split('?')[0].split('#')[0].toLowerCase();
                if (normalizedUrl.substr(normalizedUrl.length - 4) !== '.css') {
                  console.error('[OverrideCss] The selected SharePoint file is not a CSS file: ' + selectedUrl);
                  return;
                }
                urlInput.value = selectedUrl;
                onChange(selectedUrl);
                closePicker();
              }
            } catch (error) {
              console.error('[OverrideCss] Unable to process the SharePoint file picker response: ' + String(error));
            }
          };
          window.addEventListener('message', messageHandler, false);
          document.body.appendChild(pickerOverlay);
        };
        element.appendChild(button);

        const description = document.createElement('div');
        description.textContent = 'Saves the central SharePoint file URL and reloads its current contents on each page load.';
        description.style.fontSize = '12px';
        description.style.marginTop = '6px';
        element.appendChild(description);
      },
      onDispose: (): void => {
        closePicker();
      }
    }
  };
}
