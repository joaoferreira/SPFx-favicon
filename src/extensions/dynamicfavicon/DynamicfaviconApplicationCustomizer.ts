import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'DynamicfaviconApplicationCustomizerStrings';

const LOG_SOURCE: string = 'DynamicfaviconApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IDynamicfaviconApplicationCustomizerProperties {
  // This is an example; replace with your own property
  faviconpath: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class DynamicfaviconApplicationCustomizer
  extends BaseApplicationCustomizer<IDynamicfaviconApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {  

    let url: string = this.properties.faviconpath;
    if (!url) {
      Log.info(LOG_SOURCE, strings.Info);
    }else{

        //Add support for internet explorer and edge
        //The favicon must be removed before adding the new one 

        var links=document.getElementsByTagName('link');
        var head=document.getElementsByTagName('head')[0];
        for(var i=0; i<links.length; i++)
        {
          if(links[i].getAttribute('rel')==='shortcut icon')
          {
            head.removeChild(links[i]);
          }         
        } 
        
        var link = document.querySelector("link[rel*='icon']") as HTMLElement || document.createElement('link') as HTMLElement;
        link.setAttribute('type', 'image/x-icon');
        link.setAttribute('rel', 'shortcut icon');
        link.setAttribute('href', url);
        document.getElementsByTagName('head')[0].appendChild(link);
    }

    return Promise.resolve();
  }
}
