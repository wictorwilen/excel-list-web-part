import * as ko from 'knockout';
import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField
} from '@microsoft/sp-client-preview';

import {
  EnvironmentType
} from '@microsoft/sp-client-base';


import * as strings from 'excelListWebPartStrings';
import ExcelListWebPartViewModel, { IExcelListWebPartBindingContext } from './ExcelListWebPartViewModel';
import { IExcelListWebPartWebPartProps } from './IExcelListWebPartWebPartProps';

import * as ListService from './ListService';


let _instance: number = 0;

export default class ExcelListWebPartWebPart extends BaseClientSideWebPart<IExcelListWebPartWebPartProps> {
  private _id: number;
  private _koDescription: KnockoutObservable<string> = ko.observable('');
  private _shouter: KnockoutSubscribable<{}> = new ko.subscribable();

  private _dataService: ListService.IListsService;

  public constructor(context: IWebPartContext) {
    super(context);
    this._id = _instance++;
    const isDebug: boolean =
      DEBUG && (this.context.environment.type === EnvironmentType.Test || this.context.environment.type === EnvironmentType.Local);

    this._dataService = isDebug
      ? new ListService.MockListsService()
      : new ListService.ListsService(this.context);
  }

  public render(): void {
    if (!this.renderedOnce) {
      this.doInitStuff();
    }

    this._koDescription(this.properties.description);
  }

  private _createComponentElement(tagName: string): HTMLElement {
    const componentElement: HTMLElement = document.createElement('div');
    componentElement.setAttribute('data-bind', `component: { name: "${tagName}", params: $data }`);
    return componentElement;
  }

  private doInitStuff(): void {
    const tagName: string = `ComponentElement-${this._id}`;
    const componentElement: HTMLElement = this._createComponentElement(tagName);
    this._registerComponent(tagName);
    this.domElement.appendChild(componentElement);

    const bindings: IExcelListWebPartBindingContext = {
      description: this.properties.description,
      shouter: this._shouter,
      dataService: this._dataService
    };

    ko.applyBindings(bindings, this.domElement);

    this._koDescription.subscribe((newValue: string) => {
      this._shouter.notifySubscribers(newValue, 'description');
    });
  }

  private _registerComponent(tagName: string): void {
    ko.components.register(
      tagName,
      {
        viewModel: ExcelListWebPartViewModel,
        template: require('./ExcelListWebPart.template.html'),
        synchronous: false
      }
    );
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
