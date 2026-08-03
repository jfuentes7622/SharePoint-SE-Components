import * as React from 'react';
import * as ReactDOM from 'react-dom';

import WorldClock from '../../components/WorldClock';
import { IWorldClockProps } from '../../Interfaces/IWorldClock';
import { ILoaderProps } from '../../Interfaces/ILoader';
import { ISPEventObserver } from '@microsoft/sp-core-library';
import { ISettings } from '../../Interfaces/ISettings';

import { loadStyles } from '@microsoft/load-themed-styles';


import { Log } from '@microsoft/sp-core-library';

//import { spfi, SPFx } from "@pnp/sp/presets/all";

// import {  sp } from "@pnp/sp";
// import "@pnp/sp/webs";
// import "@pnp/sp/files";
// import "@pnp/sp/folders";

import pnp from 'sp-pnp-js';

//import { textChangeRangeIsUnchanged } from 'typescript';
//import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";

class FileItem {
  public Name: string;
  public Type: string;

  constructor(Name: string, Type: string) {
    this.Name = Name;
    this.Type = Type;
  }
}

const FooterFileName: string = "FooterHtml.txt";
const HeaderFileName: string = "HeaderHtml.txt";
const SettingsFileName: string = "settings.txt";
const LoadFileName: string = "loadfiles.txt";
//const CentralFileName: string = "central.txt";
//const LogSource: string = "BrandingApplicationCustomizer";
const AssetsLib: string = "SiteAssets/";

export default class LoaderProvider implements ISPEventObserver {

  private settings: ISettings;
  private rootUrl: string;
  private props: ILoaderProps;

  constructor(props: ILoaderProps) {
    this.props = props;
  }

  public instanceId: string = 'loader-provider';
  public componentId: string = 'loader-provider-component';
  public isDisposed: boolean = false;

  public dispose(): void {
    this.isDisposed = true;
    // Clean up logic here
    const clockelement: HTMLElement = document.getElementById('worldclock') || document.documentElement;
    ReactDOM.unmountComponentAtNode(clockelement);
  }

  public async LoadFiles(onlyJs: boolean): Promise<void> {
    const head: HTMLElement = document.getElementsByTagName('head')[0] || document.documentElement;
    let appTop: HTMLElement;
    let appBottom: HTMLElement;
    let rootIframe: HTMLElement;
    let isCentral: boolean = false;

    //site is the sitecolleciton context.site
    //web is the current web (subsite as well) context.web

    try {

      if (this.props.rootUrl !== undefined) { //tenant wide vars set
        const normalizeUrl = (url: string) => url.replace(/\/+$/, '').toLowerCase();
        this.rootUrl = this.props.rootUrl;

        const normalizedRootUrl = normalizeUrl(this.props.rootUrl);
        const normalizedSiteUrl = normalizeUrl(this.props.context.pageContext.site.absoluteUrl);

        if (normalizedRootUrl !== normalizedSiteUrl) {
          isCentral = true;
        }
      } else { 
        if (this.props.context.pageContext.site.serverRelativeUrl === "/")
        {//is root site
          this.rootUrl = "/";
        }else{ //is not root must use full URL
          this.rootUrl = this.props.context.pageContext.site.absoluteUrl.replace(this.props.context.pageContext.site.serverRelativeUrl, "") + "/";
        }
        isCentral = true;
        // let settingsFilePath: string = this.cleanUrl(this.props.context.pageContext.site.absoluteUrl + AssetsLib + SettingsFileName);
        // this.rootUrl = new URL(this.props.context.pageContext.site.absoluteUrl);
        // if (this.props.context.pageContext.site.absoluteUrl !== this.props.context.pageContext.web.absoluteUrl) {
        //   const settingsfileFound: string = await this.readFile(settingsFilePath, true);
        //   if (settingsfileFound !== "") {
        //     isCentral = true;
        //   }
        // }

        // //Check the root site
        // settingsFilePath = this.cleanUrl(this.props.context.pageContext.site.absoluteUrl.replace(this.props.context.pageContext.site.serverRelativeUrl, "") + AssetsLib + SettingsFileName);
        // this.rootUrl = new URL(this.props.context.pageContext.site.absoluteUrl.replace(this.props.context.pageContext.site.serverRelativeUrl, ""));
        // const fileFound: string = await this.readFile(settingsFilePath, true);
        // if (fileFound !== "") {
        //   isCentral = true;
        // } else {
        //   //nothing worked, not at current site collection nor at root nor set in properties
        //   this.rootUrl = new URL('');
        // }
      }

      //This is to solve the problem if a subsite from the root.
      //With an iframe to the root or primary site collection, then auth is done on login
      //allowing access to files when we call to get the content of the files.
      if (onlyJs === false) {
        if (isCentral && this.rootUrl !== "/") {
          if (document.getElementById('bannersource') === null) {//Check to see if already added.
            const iframe: string = "<iframe id='bannersource' src='" + this.props.context.pageContext.site.absoluteUrl.replace(this.props.context.pageContext.site.serverRelativeUrl, "") + "' style='height: 1px; visibility: hidden;'></iframe>";
            rootIframe = document.getElementsByTagName('body')[0] || document.documentElement;
            rootIframe.insertAdjacentHTML("afterend", iframe);
          }
        }
      }

      if (this.rootUrl !== "") {
        const settingsJson: string = await this.readFile(this.cleanUrl(this.rootUrl + AssetsLib + SettingsFileName), isCentral);
        const footerHtml: string = await this.readFile(this.cleanUrl(this.rootUrl + AssetsLib + FooterFileName), isCentral);
        const headerHtml: string = await this.readFile(this.cleanUrl(this.rootUrl + AssetsLib + HeaderFileName), isCentral);
        const loadFilesJson: string = await this.readFile(this.cleanUrl(this.rootUrl + AssetsLib + LoadFileName), isCentral);

        const objArray = JSON.parse(loadFilesJson);
        const fileItemArray = objArray.map((item: { Name: string; Type: string; }) => new FileItem(item.Name, item.Type));
        this.settings = JSON.parse(settingsJson) as ISettings;

        if (onlyJs === false) {
          if (document.getElementById(this.settings.ModernBannerComId) !== null) //communications site
          {
            appTop = document.getElementById(this.settings.ModernBannerComId) || document.documentElement;
            appBottom = document.getElementById(this.settings.ModernFooterComId) || document.documentElement;
          } else if (document.getElementById(this.settings.ModernBannerTeamId) !== null) //teams site
          {
            appTop = document.getElementById(this.settings.ModernBannerTeamId) || document.documentElement;
            appBottom = document.getElementById(this.settings.ModernFooterTeamId) || document.documentElement;
          } else {
            appTop = document.getElementsByTagName('body')[0] || document.documentElement;
            appBottom = document.getElementsByTagName('body')[0] || document.documentElement;
          }

          if (document.getElementById('bannerheader') === null) {//Check to see if already added.
            if (headerHtml) {
              if (headerHtml.length > 0) {
                appTop.insertAdjacentHTML("afterbegin", headerHtml);
              }
            }
          }

          if (document.getElementById('footerbottom') === null) {//Check to see if already added.
            if (footerHtml) {
              if (footerHtml.length > 0) {
                appBottom.insertAdjacentHTML("beforeend", footerHtml);
              }
            }
          }
        }

        fileItemArray.forEach((Item: FileItem) => {
          switch (Item.Type.toLowerCase()) {
            case 'css': {
              //Process css
              /*   const LoadCSSItem: HTMLLinkElement = document.createElement('link');
                LoadCSSItem.rel = 'stylesheet';
                LoadCSSItem.type = 'text/css';
                LoadCSSItem.href = Item.Name;
                head.insertAdjacentElement('beforeend', LoadCSSItem); */
              if (onlyJs === false) {
                loadStyles(`@import url('${Item.Name}');`);
              }
              break;
            }
            case 'js': {
              //process js
              const LoadJSItem: HTMLScriptElement = document.createElement('script');
              LoadJSItem.src = Item.Name;
              LoadJSItem.type = 'text/javascript';
              head.insertAdjacentElement('beforeend', LoadJSItem);
              break;
            }
          }
        });
      }
    } catch (error) {
      console.log(error);
    }
  }


  public async render(): Promise<void> {
    try {
      await this.LoadFiles(false);

      //render webpart
      const clockelement: HTMLElement = document.getElementById('worldclock') || document.documentElement;
      if (clockelement && this.settings !== null) {
        const ListUrl: URL = new URL(this.props.context.pageContext.site.absoluteUrl.replace(this.props.context.pageContext.site.serverRelativeUrl, ""));
        const elementClock: React.ReactElement<IWorldClockProps> = React.createElement(
          WorldClock,
          {
            digitalOnly: this.settings.digitalOnly,
            hour12: this.settings.hour12,
            displayDay: this.settings.displayDay,
            spList: this.settings.spList,
            errorHandler: (errorMessage: string) => {
              //Log.info(LOG_SOURCE, errorMessage);
              console.error(errorMessage);
            },
            context: this.props.context,
            absoluteUrl: ListUrl.href,
            spHttpClient:this.props.context.spHttpClient            
            
          }
        );
        ReactDOM.render(elementClock, clockelement);
      }
    } catch (error) {
      console.log(error);
    }
  }

  private cleanUrl(url: string): string {
    return url.replace(/([^:]\/)\/+/g, "$1");
  }


  // private getPathUrl(fileUrl: string): string {
  //   const url = new URL(fileUrl);
  //   const segments = url.pathname.split("/");

  //   // Remove the last segment (file name)
  //   segments.pop();

  //   // Rebuild the URL
  //   url.pathname = segments.join("/");

  //   return url.toString();
  // }

  public async findExistingFile(fileUrl: string): Promise<string> {
    const url = new URL(fileUrl);
    const segments = url.pathname.split("/").filter(Boolean);

    // Extract the file name (last segment)
    const fileName = segments.pop();
    if (!fileName) { return ""; }

    // Traverse up the path and check for the file at each level
    for (let i = segments.length; i >= 0; i--) {
      const basePath = "/" + segments.slice(0, i).join("/");
      const testUrl = `${url.origin}${basePath}/${fileName}`;

      const exists = await this.readFile(testUrl, true);
      if (exists !== "") {
        return testUrl;
      }
    }

    return ""; // File not found at any level
  }



  public async findFileInSiteAssets(fileName: string, currentUrl: string): Promise<string> {
    const url = new URL(currentUrl);
    const segments = url.pathname.split("/").filter(Boolean);

    // Traverse up the site path, including the root
    for (let i = segments.length; i >= 0; i--) {
      const sitePath = i > 0 ? "/" + segments.slice(0, i).join("/") : "";
      const fileUrl = `${url.origin}${sitePath}/SiteAssets/${fileName}`;

      const exists = await this.readFile(fileUrl, true);
      if (exists !== "") {
        return fileUrl;
      }
    }

    return ""; // File not found in any SiteAssets folder
  }




  public async readFile(filename: string, isCentral: boolean): Promise<string> {
    Log.info("BannerClock", filename);

    if (isCentral) {
      try {
        const response = await fetch(`${filename}`, {
          method: 'GET',
          headers: {
            'Accept': 'application/json;odata=verbose'
            // Authorization header removed — AAD token provider not supported in SPFx 1.5.1
          }
        });

        if (response.ok) {
          return await response.text();
        } else {
          console.error('Error fetching file content:', response.statusText);
        }
      } catch (error) {
        console.error('Error fetching file content:', error);
      }
    } else {
      let text: string = "";
      if (await this.fileExists(filename)) {
        text = await pnp.sp.web.getFileByServerRelativeUrl(filename).getText();
        text = await pnp.sp.web.getFileByServerRelativePath(filename).getText();
      }
      return text;
    }

    return "";
  }





  public async fileExists(fileUrl: string): Promise<boolean> {
    try {
      await pnp.sp.web.getFileByServerRelativeUrl(fileUrl).getText();
      return true;
    } catch (error) {
      console.warn('File not found or inaccessible:', error);
      return false;
    }
  }



}