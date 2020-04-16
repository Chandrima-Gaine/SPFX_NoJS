import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ChandStudentWebPart.module.scss';
import * as strings from 'ChandStudentWebPartStrings';

import {
  SPHttpClient,
  SPHttpClientResponse   
} from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

export interface IChandStudentWebPartProps {
  description: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  StudentName : string;
  Marks :number;
}

export default class ChandStudentWebPart extends BaseClientSideWebPart<IChandStudentWebPartProps> {

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('StudentList')/Items",SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
        return response.json();
        });
    }
    private _renderListAsync(): void {
    
      if (Environment.type == EnvironmentType.SharePoint || 
               Environment.type == EnvironmentType.ClassicSharePoint) {
       this._getListData()
         .then((response) => {
           this._renderList(response.value);
         });
     }
   }
    private _renderList(items: ISPList[]): void {
      let html: string = '<table border=1 width=100% style="border-collapse: collapse; border-color: white">';
      html += '<th>Subject</th> <th>Topper Student Name</th><th>Marks</th>';
      items.forEach((item: ISPList) => {
        html += `
        <tr>            
            <td>${item.Title}</td>
            <td>${item.StudentName}</td>
            <td>${item.Marks}</td>
            
            </tr>
            `;
      });
      html += '</table>';
    
      const listContainer: Element = this.domElement.querySelector('#spListContainer');
      listContainer.innerHTML = html;
    }

    public render(): void {
      this.domElement.innerHTML = `
        <div class="${ styles.chandStudent }">
          <div class="${ styles.container }">
            <div class="${ styles.row }">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
            <span class="${ styles.title }">Welcome to Chandrima's School</span>
          </div>
        </div> 
            <div class="${styles.row}">
            <div>2020 Topper Student Details</div>
            <br>
             <div id="spListContainer" />
          </div>
        </div>`;
        this._renderListAsync();
    }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
