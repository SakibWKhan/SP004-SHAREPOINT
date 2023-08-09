import { override } from '@microsoft/decorators';

import { Log } from '@microsoft/sp-core-library';

import {

  BaseApplicationCustomizer,

  PlaceholderContent,

 

  PlaceholderName,

} from '@microsoft/sp-application-base';

import { Dialog } from '@microsoft/sp-dialog';

 

import * as strings from 'SalesPortalApplicationCustomizerStrings';

 

const LOG_SOURCE: string = 'SalesPortalApplicationCustomizer';

 

/**

 * If your command set uses the ClientSideComponentProperties JSON input,

 * it will be deserialized into the BaseExtension.properties object.

 * You can define an interface to describe it.

 */

export interface ISalesPortalApplicationCustomizerProperties {

  // This is an example; replace with your own property

  testMessage: string;

}

 

/** A Custom Action which can be run during execution of a Client Side Application */

export default class SalesPortalApplicationCustomizer

  extends BaseApplicationCustomizer<ISalesPortalApplicationCustomizerProperties> {

    private headerPlaceholder: PlaceholderContent;

 

    private footerPlaceholder: PlaceholderContent;

  @override

  public onInit(): Promise<void> {

    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

 

    // Register the header and footer placeholders

 

    this.context.placeholderProvider.changedEvent.add(this, this.renderPlaceholders);

 

    return Promise.resolve();

  }

  private renderPlaceholders(): void {

 

    // Check if the header and footer placeholders exist

 

    if (!this.headerPlaceholder) {

 

      this.headerPlaceholder = this.context.placeholderProvider.tryCreateContent(

 

        PlaceholderName.Top,

 

        { onDispose: this.onDispose },

 

      );

 

    }

 

 

 

 

    if (!this.footerPlaceholder) {

 

      this.footerPlaceholder = this.context.placeholderProvider.tryCreateContent(

 

        PlaceholderName.Bottom,

 

        { onDispose: this.onDispose },

 

      );

 

    }

 

 

 

 

    // Render your custom header and footer content with enhanced styles

 

    if (this.headerPlaceholder && this.footerPlaceholder) {

 

      this.headerPlaceholder.domElement.innerHTML = `

 

<div class="custom-header">

 

<h1>! Welcome to the Sale Portal !</h1>

 

</div>

 

        `;

 

 

 

      // Add custom CSS styles to the header

 

      const headerStyle = document.createElement('style');

 

      headerStyle.innerHTML = `

 

        .custom-header {

 

          background-color: #f2f2f2;

 

          padding: 20px;

 

          text-align: center;

 

        }

 

 

 

 

        .custom-header h1 {

 

          color: #787C85;

 

          font-size: 24px;

 

          margin: 0;

 

        }

 

      `;

 

      this.headerPlaceholder.domElement.appendChild(headerStyle);

 

 

 

 

      this.footerPlaceholder.domElement.innerHTML = `

 

<div class="custom-footer">

 

<p>Contact us: info@saleportal.com</p>

 

</div>

 

      `;

 

 

 

 

      // Add custom CSS styles to the footer

 

      const footerStyle = document.createElement('style');

 

      footerStyle.innerHTML = `

 

        .custom-footer {

 

          background-color: #333;

 

          color: #fff;

 

          padding: 10px;

 

          text-align: center;

 

        }

 

 

 

 

        .custom-footer p {

 

          margin: 0;

 

        }

 

      `;

 

      this.footerPlaceholder.domElement.appendChild(footerStyle);

 

    }

 

  }

 

 

 

 

 

 

 

 

 

 

  protected onDispose(): void {

 

    // Clean up resources

 

    Log.info(LOG_SOURCE, `Disposed ${strings.Title}`);

 

    this.headerPlaceholder = null;

 

    this.footerPlaceholder = null;

 

  }

}

