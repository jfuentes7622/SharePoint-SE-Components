import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
// import * as strings from 'BrandingApplicationCustomizerStrings';

// import * as React from "react";
import LoaderProvider from "./LoaderService";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IBrandingApplicationCustomizerProperties {
  rootUrl:string;
}

const LogSource:string = "BrandingApplicationCustomizer";


/** A Custom Action which can be run during execution of a Client Side Application */
export default class BrandingApplicationCustomizer
  extends BaseApplicationCustomizer<IBrandingApplicationCustomizerProperties> {
  
  private LoaderProviderService: LoaderProvider;

  public async onInit(): Promise<void> {
    
    Log.info(LogSource,"Starting int");
    this.LoaderProviderService = new LoaderProvider({context: this.context, rootUrl: this.properties.rootUrl});
    await this.LoaderProviderService.LoadFiles(true);
    this.context.placeholderProvider.changedEvent.add(this.LoaderProviderService, this.LoaderProviderService.render);

    return Promise.resolve();
  }
}
