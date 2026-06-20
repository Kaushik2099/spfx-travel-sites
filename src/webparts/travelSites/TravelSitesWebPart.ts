import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { TravelSites, ITravelSitesProps } from './components/TravelSites';

export interface ITravelSitesWebPartProps {
  description: string;
}

export default class TravelSitesWebPart extends BaseClientSideWebPart<ITravelSitesWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ITravelSitesProps> = React.createElement(
      TravelSites,
      {
        // We pass the context here, but our component won't use it for REST calls anymore
        context: this.context 
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}