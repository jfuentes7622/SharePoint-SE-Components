import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from "@microsoft/sp-webpart-base";
import { IScriptEditorProps } from './components/IScriptEditorProps';
import { IScriptEditorWebPartProps } from './components/IScriptEditorWebpartProps';
import PropertyPaneLogo from './PropertyPaneLogo';

export default class ScriptEditorWebPart extends BaseClientSideWebPart<IScriptEditorWebPartProps> {
  private _unqiueId: string;

  public render(): void {
    this._unqiueId = this.context.instanceId;

    if (this.displayMode === DisplayMode.Read) {
      if (this.properties.removePadding) {
        let element = this.domElement.parentElement;
        for (let i = 0; i < 5 && element; i++) {
          const style = window.getComputedStyle(element);
          if (style.paddingTop !== "0px") {
            element.style.paddingTop = "0px";
            element.style.paddingBottom = "0px";
            element.style.marginTop = "0px";
            element.style.marginBottom = "0px";
          }
          element = element.parentElement;
        }
      }

      ReactDom.unmountComponentAtNode(this.domElement);
      this.domElement.innerHTML = this.properties.script || "";
      this.executeScript(this.domElement);
    } else {
      this.renderEditor();
    }
  }

private async renderEditor(): Promise<void> {
    const editorPopUp:any = await require(/* webpackChunkName: 'scripteditor' */ './components/ScriptEditorWebpart');
    const element: React.ReactElement<IScriptEditorProps> = React.createElement(editorPopUp.default, {
      script: this.properties.script,
      title: this.properties.title,
      propPaneHandle: this.context.propertyPane,
      key: "editor" + new Date().getTime()
    });
    ReactDom.render(element, this.domElement);
  }


  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField("title", {
                  label: "Title to show in edit mode"
                }),
                PropertyPaneToggle("removePadding", {
                  label: "Remove top/bottom padding of web part container",
                  onText: "Remove padding",
                  offText: "Keep padding"
                }),
                PropertyPaneToggle("spPageContextInfo", {
                  label: "Enable classic _spPageContextInfo",
                  onText: "Enabled",
                  offText: "Disabled"
                }),
                PropertyPaneTextField("script", {
                  label: "HTML code",
                  multiline: true,
                  resizable: true,
                  rows: 12,
                  description: "Paste HTML markup to render in the web part."
                }),
                new PropertyPaneLogo()
              ]
            }
          ]
        }
      ]
    };
  }

  private evalScript(elem: HTMLScriptElement): void {
    const data: string = elem.text || elem.textContent || elem.innerHTML || "";
    const headTag: HTMLElement = document.getElementsByTagName("head")[0] || document.documentElement;
    const scriptTag: HTMLScriptElement = document.createElement("script");

    for (let i = 0; i < elem.attributes.length; i++) {
      const attr = elem.attributes[i];
      if (attr.name.toLowerCase() === "onload") continue;
      scriptTag.setAttribute(attr.name, attr.value);
    }

    scriptTag.type = (scriptTag.src && scriptTag.src.length) > 0 ? "pnp" : "text/javascript";
    scriptTag.setAttribute("pnpname", this._unqiueId);

    try {
      scriptTag.appendChild(document.createTextNode(data));
    } catch (e) {
      scriptTag.text = data;
    }

    headTag.insertBefore(scriptTag, headTag.firstChild);
  }

  private async executeScript(element: HTMLElement): Promise<void> {
    const headTag: HTMLElement = document.getElementsByTagName("head")[0] || document.documentElement;

    // Remove old scripts
    const oldScripts = headTag.getElementsByTagName("script");
    for (let i = 0; i < oldScripts.length; i++) {
      const scriptTag = oldScripts[i];
      if (scriptTag.hasAttribute("pnpname") && scriptTag.getAttribute("pnpname") === this._unqiueId) {
        headTag.removeChild(scriptTag);
      }
    }

    if (this.properties.spPageContextInfo && !window["_spPageContextInfo"]) {
      window["_spPageContextInfo"] = this.context.pageContext.legacyPageContext;
    }

    (window as any).ScriptGlobal = {};

    const scripts: HTMLScriptElement[] = [];
    const children_nodes = element.getElementsByTagName("script");

    for (let i = 0; children_nodes[i]; i++) {
      const child = children_nodes[i];
      if (!child.type || child.type.toLowerCase() === "text/javascript") {
        scripts.push(child);
      }
    }

    const urls: string[] = [];
    const onLoads: Array<() => void> = [];

    for (let i = 0; scripts[i]; i++) {
      const scriptTag = scripts[i];
      if (scriptTag.src && scriptTag.src.length > 0) {
        urls.push(scriptTag.src);
      }
      if (scriptTag.onload && typeof scriptTag.onload === "function") {
        onLoads.push(scriptTag.onload as () => void);
      }
    }

    let oldamd: any = null;
    if ((window as any)["define"] && (window as any)["define"].amd) {
      oldamd = (window as any)["define"].amd;
      (window as any)["define"].amd = null;
    }

    for (let i = 0; i < urls.length; i++) {
      try {
        let scriptUrl = urls[i];
        const prefix = scriptUrl.indexOf("?") === -1 ? "?" : "&";
        scriptUrl += prefix + "pnp=" + new Date().getTime();
        await SPComponentLoader.loadScript(scriptUrl, { globalExportsName: "ScriptGlobal" });
      } catch (error) {
        console.error(error);
      }
    }

    if (oldamd) {
      (window as any)["define"].amd = oldamd;
    }

    for (let i = 0; scripts[i]; i++) {
      const scriptTag = scripts[i];
      if (scriptTag.parentNode) {
        scriptTag.parentNode.removeChild(scriptTag);
      }
      this.evalScript(scriptTag);
    }

    for (let i = 0; onLoads[i]; i++) {
      onLoads[i]();
    }
  }
}
