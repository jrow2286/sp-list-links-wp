import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneDropdownOption } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ListLinksWebPart.module.scss';
import * as strings from 'ListLinksWebPartStrings';
import "@pnp/polyfill-ie11";
import { sp } from '@pnp/sp';
import * as _ from 'lodash';

export interface IListLinksWebPartProps {
  headingHtml: string;
  imageUrl: string;
  listName: string;
  hyperlinkField: string;
  categoryField: string;
  categoryValue: string;
}

interface IListItem {
  Title?: string;
  Id: number;
  Link: HyperLinkField;
}

interface HyperLinkField {
  Url: string;
  Description?: string;
}

export default class ListLinksWebPart extends BaseClientSideWebPart<IListLinksWebPartProps> {
  private lists: IPropertyPaneDropdownOption[];
  private listsDropdownDisabled: boolean = true;
  
  private hyperlinkFields: IPropertyPaneDropdownOption[];
  private hyperlinkFieldDropdownDisabled: boolean = true;
  
  private categoryFields: IPropertyPaneDropdownOption[];
  private categoryFieldDropdownDisabled: boolean = true;

  private categoryValues: IPropertyPaneDropdownOption[];
  private categoryValueDropdownDisabled: boolean = true;

  protected onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error?: any) => void): void => {
      sp.setup({
        spfxContext: this.context,
        sp: {
          headers: {
            "Accept": "application/json; odata=nometadata"
          }
        }
      });
      resolve();
    });
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.listLinks }">
        <div class="${ styles.container }">
          ${(() => {
            var html = ``;
            if (this.properties.imageUrl) {
              html += `
                <div class="${ styles.center }">
                  <img src="${escape(this.properties.imageUrl)}" />
                </div>`;
            }
            if (this.properties.headingHtml) {
              html += `<div class="${ styles.title }">${this.properties.headingHtml}</div>`;
            }
            return html;
          })()}
          <span class="status ${ styles.status }"></span>
          <ul class="link-items"></ul>
        </div>
      </div>`;
    
    this.readItems();
  }

  private loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) => {
      sp.web.lists.filter("BaseTemplate eq 100").select("Title").get()
      .then((data): void => {
        resolve(
          data.map(item => {
            return {
              key: item.Title,
              text: item.Title
            };
          })
        );
      });
    });
  }

  private loadFieldsByType(fieldType: string): Promise<IPropertyPaneDropdownOption[]> {
    if (!this.properties.listName) {
      return Promise.resolve();
    }

    return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) => {
      sp.web.lists.getByTitle(this.properties.listName).fields
      .filter("ReadOnlyField eq false and TypeAsString eq '" + fieldType + "'").select("InternalName,Title").get()
      .then((data): void => {
        resolve(
          data.map(item => {
            return {
              key: item.InternalName,
              text: item.Title
            };
          })
        );
      });
    });

  }

  private loadUniqueCategoryValues(): Promise<IPropertyPaneDropdownOption[]> {
    if (!this.properties.categoryField) {
      return Promise.resolve();
    }

    return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) => {
      sp.web.lists.getByTitle(this.properties.listName).items
      .select(this.properties.categoryField).get()
      .then((data): void => {
        data = _.uniqBy(data, this.properties.categoryField);
        resolve(
          data.map(item => {
            return {
              key: item[this.properties.categoryField],
              text: item[this.properties.categoryField]
            };
          })
        );
      });
    });
  }

  private readItems(): void {
    if (this.properties.listName && this.properties.hyperlinkField) {
      let filter = "ID ne 0";
      if(this.properties.categoryField && this.properties.categoryValue){
        filter = this.properties.categoryField + " eq '" + this.properties.categoryValue + "'";
      }
      sp.web.lists.getByTitle(this.properties.listName)
      .items.filter(filter).select(this.properties.hyperlinkField).get()
      .then((items: IListItem[]): void => {
        this.updateLinkItemsHtml(items);
      }, (error: any): void => {
        error.response.json().then((json) => {
          this.updateStatus('Error retrieving list items: ' + json["odata.error"].message.value);
        });
      });
    } else {
      if (!this.properties.listName) {
        this.updateStatus('Please select the "Links List" in the webpart properties.');
      } else {
        this.updateStatus('Please select the "Hyperlink Field" in the webpart properties.');
      }
    }
  }

  private updateLinkItemsHtml(items: IListItem[]): void {
    this.domElement.querySelector('.link-items').innerHTML = 
      items.map(item => 
        `<li>
          <a href='${item[this.properties.hyperlinkField].Url}'>
            ${escape(item[this.properties.hyperlinkField].Description || item[this.properties.hyperlinkField].Url)}
          </a>
        </li>`)
      .join("");
  }

  private updateStatus(status: string): void {
    this.domElement.querySelector('.status').innerHTML = status;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {
    this.listsDropdownDisabled = !this.lists;
    this.hyperlinkFieldDropdownDisabled = !this.properties.listName || !this.hyperlinkFields;

    if (this.lists) {
      return;
    }

    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');

    this.loadLists()
    .then((listOptions: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
      this.lists = listOptions;
      this.listsDropdownDisabled = false;
      this.context.propertyPane.refresh();
      this.render();
      return this.loadFieldsByType("URL");
    })
    .then((hyperlinkFieldOptions: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
      this.hyperlinkFields = hyperlinkFieldOptions;
      this.hyperlinkFieldDropdownDisabled = !this.properties.listName;
      this.context.propertyPane.refresh();
      this.render();
      return this.loadFieldsByType("Text");
    })
    .then((categoryFieldOptions: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
      this.categoryFields = categoryFieldOptions;
      this.categoryFieldDropdownDisabled = !this.properties.listName;
      this.context.propertyPane.refresh();
      this.render();
      return this.loadUniqueCategoryValues();
    }).then((categoryValueOptions: IPropertyPaneDropdownOption[]): void => {
      this.categoryValues = categoryValueOptions;
      this.categoryValueDropdownDisabled = !this.properties.categoryField;
      this.context.propertyPane.refresh();
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      this.render();
    });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === 'listName' && newValue){
      const previousHyperlinkField: string = this.properties.hyperlinkField;
      this.properties.hyperlinkField = undefined;
      this.onPropertyPaneFieldChanged('hyperlinkField', previousHyperlinkField, this.properties.hyperlinkField);
      this.hyperlinkFieldDropdownDisabled = true;

      const previousCategoryField: string = this.properties.categoryField;
      this.properties.categoryField = undefined;
      this.onPropertyPaneFieldChanged('categoryField', previousCategoryField, this.properties.categoryField);
      this.categoryFieldDropdownDisabled = true;
      
      const previousCategoryValue: string = this.properties.categoryValue;
      this.properties.categoryValue = undefined;
      this.onPropertyPaneFieldChanged('categoryValue', previousCategoryValue, this.properties.categoryValue);
      this.categoryValueDropdownDisabled = true;
      
      this.context.propertyPane.refresh();
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'field data');
      this.loadFieldsByType("URL")
      .then((hyperLinkFieldOptions: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
        this.hyperlinkFields = hyperLinkFieldOptions;
        this.hyperlinkFieldDropdownDisabled = false;
        this.render();
        this.context.propertyPane.refresh();
        return this.loadFieldsByType("Text");
      })
      .then((categoryFieldOptions: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
        this.categoryFields = categoryFieldOptions;
        this.categoryFieldDropdownDisabled = false;
        this.render();
        this.context.propertyPane.refresh();
        return this.loadUniqueCategoryValues();
      }).then((categoryValueOptions: IPropertyPaneDropdownOption[]): void => {
        this.categoryValues = categoryValueOptions;
        this.categoryValueDropdownDisabled = false;
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
        this.context.propertyPane.refresh();
      });
    } else if (propertyPath === 'categoryField' && newValue){
      const previousCategoryValue: string = this.properties.categoryValue;
      this.properties.categoryValue = undefined;
      this.onPropertyPaneFieldChanged('categoryValue', previousCategoryValue, this.properties.categoryValue);
      this.categoryValueDropdownDisabled = true;

      this.context.propertyPane.refresh();
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'field values');
      this.loadUniqueCategoryValues()
      .then((categoryValueOptions: IPropertyPaneDropdownOption[]): void => {
        this.categoryValues = categoryValueOptions;
        this.categoryValueDropdownDisabled = false;
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
        this.context.propertyPane.refresh();
      });
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.Group1Name,
              groupFields: [
                PropertyPaneTextField('headingHtml', {
                  label: strings.HeadingHtmlFieldLabel
                }),
                PropertyPaneTextField('imageUrl', {
                  label: strings.ImageUrlFieldLabel
                })
              ]
            },
            {
              groupName: strings.Group2Name,
              groupFields: [
                PropertyPaneDropdown('listName', {
                  label: strings.ListNameFieldLabel,
                  options: this.lists,
                  disabled: this.listsDropdownDisabled
                }),
                PropertyPaneDropdown('hyperlinkField', {
                  label: strings.HyperlinkFieldFieldLabel,
                  options: this.hyperlinkFields,
                  disabled: this.hyperlinkFieldDropdownDisabled
                }),
                PropertyPaneDropdown('categoryField', {
                  label: strings.CategoryFieldFieldLabel,
                  options: this.categoryFields,
                  disabled: this.categoryFieldDropdownDisabled
                }),
                PropertyPaneDropdown('categoryValue', {
                  label: strings.CategoryValueFieldLabel,
                  options: this.categoryValues,
                  disabled: this.categoryValueDropdownDisabled
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
