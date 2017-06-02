import {
    IWebPartData,
    IWebPartContext,
    IClientSideWebPart,
    BaseClientSideWebPart,
    DisplayMode,
    IPropertyPanePage,
    IPropertyPaneFieldType,
    SPRequest,
    HostType
} from '@ms/sp-client-platform';

import './SharePointChart.css';
import Strings from './SharePointChart.strings';
import MockSPRequest from './test/MockSPRequest';

declare var Chart: any;

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export interface ISPFields {
  value: ISPField[];
}
export interface ISPField {
  Id: string;
  Title: string;
  InternalName: string;
  TypeAsString: string;
  ListId: string;
}
export interface ISPItems {
  value: ISPItem[];
}

export interface ISPItem {
  Id: number;
  Label: string;
  Data: number;
}

export interface ISharePointChartWebPartProps {
    description: string,
    charttype: string,
    datasource: string,
    labelcolumn: string,
    datacolumn: string,
    color: string
}

export class SharePointChartWebPart extends BaseClientSideWebPart<ISharePointChartWebPartProps> {

  datasources: Array<Object> = new Array();
  labelcolumns: Array<Object> = new Array();
  datacolumns: Array<Object> = new Array();
  alllistfields: Array<Object> = new Array();
  allValues: Array<Object> = new Array();
  labelcolumn: string = '';
  datacolumn: string = '';
  
  labels: string[] = [];
  chartdataset: number[] = [];
  
  colors: string[] = [];
  //selectedList: string = '';
  //selectedChartType: string = '';
  //selectedLabelColumn: string = '';
  //selectedDataColumn: string = '';

  public constructor(context: IWebPartContext, loadedModules: { [moduleName: string]: any }){
      super(context, loadedModules);
  }

  public render(mode: DisplayMode, data?: IWebPartData) {    
    this.domElement.innerHTML = `<canvas id="SPChart"></canvas>`;
    var canvas = <HTMLCanvasElement> document.getElementById("SPChart");
    var ctx = canvas.getContext("2d");
    
    var spchart = new Chart(ctx, {
            type: this.properties['charttype'],
            data: {
                labels: this.labels,
                datasets: [{
                    label: this.datacolumn,
                    data: this.chartdataset,
                    backgroundColor: this.colors,
                }]
            },
            options: {
                scales: {
                    yAxes: [{
                        ticks: {
                            beginAtZero:true
                        }
                    }]
                }
            }
        });
    this._renderListAsync();   
  }

  protected get propertyPaneSettings(): IPropertyPanePage[] {
        let properties: IPropertyPanePage[] = 
        [{
            header: {
                title: Strings.PropertyPaneHeader
            },
            groups: [
                {
                    groupName: Strings.BasicGroupName,
                    groupFields: [
                        {
                            type: IPropertyPaneFieldType.TextBox,
                            targetProperty: 'description',
                            properties: {
                                label: Strings.DescriptionFieldLabel
                            }
                        }]
                },
                {
                    groupName: Strings.ChartConfiguration,
                    groupFields: [
                        {
                            type: IPropertyPaneFieldType.DropDown,
                            targetProperty: 'charttype',
                            properties: {
                                label: Strings.ChartType,
                                options: [
                                     { key: 'line', text: 'Line Chart' },
                                     { key: 'bar', text: 'Bar Chart' },
                                     { key: 'radar', text: 'Radar Chart' },
                                     { key: 'polarArea', text: 'Polar Area Chart' },
                                     { key: 'pie', text: 'Pie & Doughnut Chart' }
                                ],
                                selectionChanged: function(index: number) { 
                                    console.log(index); 
                                }
                            }
                        },
                        {
                            type: IPropertyPaneFieldType.DropDown,
                            targetProperty: 'datasource',
                            properties: {
                                label: Strings.DataSource,
                                options: this.datasources,
                                selectionChanged: function(index: number) { 
                                    console.log(index); 
                                }
                            }
                        },
                        {
                            type: IPropertyPaneFieldType.DropDown,
                            targetProperty: 'labelcolumn',
                            properties: {
                                label: Strings.LabelColumn,
                                options: this.labelcolumns,
                                selectionChanged: function(index: number) { 
                                    console.log(index); 
                                }
                            }
                        },
                        {
                            type: IPropertyPaneFieldType.DropDown,
                            targetProperty: 'datacolumn',
                            properties: {
                                label: Strings.DataColumn,
                                options: this.datacolumns,
                                selectionChanged: function(index: number) { 
                                    console.log(index); 
                                }
                            }
                        },
                        {
                            type: IPropertyPaneFieldType.DropDown,
                            targetProperty: 'color',
                            properties: {
                                label: Strings.Color,
                                options:  [
                                     { key: '#FF6666', text: 'Red' },
                                     { key: '#66FF66', text: 'Green' },
                                     { key: '#6666FF', text: 'Blue' }
                                ],
                                selectionChanged: function(index: number) { 
                                    console.log(index); 
                                }
                            }
                        }
                    ]
                }]
        }];        
        return properties;
    }
    
   protected onPropertyChange(propertyName: string, newValue: any) {
     //this.selectedList = newValue;
     if(propertyName == 'datasource'){
        this.labelcolumns = new Array();
        this.datacolumns = new Array();
        this.alllistfields.forEach((field: ISPField) => {
          if(field.ListId == newValue){
              if(field.TypeAsString == 'Text'){
                this.labelcolumns.push({ key: field.Id, text: field.Title });
              }
              else if(field.TypeAsString == 'Number'){
                this.datacolumns.push({ key: field.Id, text: field.Title });
              }
          }
        });
        
     }
     //else if(propertyName == 'charttype'){
     //  this.selectedChartType = newValue;
     //}  
     //else if(propertyName == 'labelcolumn'){
     //  this.selectedLabelColumn = newValue;
     //}  
     //else if(propertyName == 'datacolumn'){
     //  this.selectedDataColumn = newValue;
     //}  
      this.properties[propertyName] = newValue;
      
    if(
        this.properties['charttype'] != undefined && this.properties['charttype'] != '' &&
        this.properties['datasource'] != undefined && this.properties['datasource'] != '' &&
        this.properties['labelcolumn'] != undefined && this.properties['labelcolumn'] != '' &&
        this.properties['datacolumn'] != undefined && this.properties['datacolumn'] != ''){
            
      this.alllistfields.forEach((field: ISPField) => {
          if(field.Id == this.properties['labelcolumn']){
            this.labelcolumn = field.InternalName;
          }
          else if(field.Id == this.properties['datacolumn']){
            this.datacolumn = field.InternalName;
          }
        });
        this._renderItemsAsync(
            this.properties['charttype'], 
            this.properties['datasource']);
        }   
      this.render(this.displayMode);
    }

    // =========================================
    // Loading the lists into WP Properties
    // =========================================
    private _getListData(): Promise<ISPLists> {
        return SPRequest.get(this.host.pageContext.webAbsoluteUrl + `/_api/web/lists?$select=Id,Title,Hidden&$filter=Hidden eq false`)
          .then((response: Response) => {
            return response.json();
          });
      }
     
    private _getMockListData(): Promise<ISPLists> {
          return MockSPRequest.get(this.host.pageContext.webAbsoluteUrl).then(() => {
              const listData: ISPLists = { value: [{ Title: 'Mock List 1', Id: '1' }, { Title: 'Mock List 2', Id: '2' }, { Title: 'Mock List 3', Id: '3' }] }
              return listData;
          }) as Promise<ISPLists>;
      }
      
    private _renderListAsync(): void {

      let items: ISPList[];

      // Test environment
      if (this.host.hostType === HostType.TestPage) {
          this._getMockListData().then((response) => {
              this._renderList(response.value);
          });

          // SharePoint environment
      } else if (this.host.hostType === HostType.ModernPage) {
          this._getListData()
              .then((response) => {
                  this._renderList(response.value);
              });
      }
    }

    private _renderList(lists: ISPList[]) {
        lists.forEach((list: ISPList) => {
            this.datasources.push({ key: list.Id, text: list.Title });
            this._renderFieldsAsync(list.Id);
        });
    }
    
    
    // =========================================
    // Loading fields from lists
    // =========================================
    private _getFields(listId: string): Promise<ISPFields> {
        return SPRequest.get(this.host.pageContext.webAbsoluteUrl + `/_api/web/lists(guid'` + listId + `')/Fields?$select=Id,Title,InternalName,TypeAsString`)
          .then((response: Response) => {
            return response.json();
        });
    }
      
    private _getMockFields(listId: string): Promise<ISPFields> {
          return MockSPRequest.get(this.host.pageContext.webAbsoluteUrl).then(() => {
              const fields: ISPFields =
              { value: [
                { Id: 'fa564e0f-0c70-4ab9-b863-0177e6ddd247', Title: 'Title', InternalName: 'InternalName', TypeAsString: 'Text', ListId: '5E3CD8EF-673E-4713-8C49-2F3925314202' },
                { Id: '8de5e87a-7cbd-42e4-bfcf-067ff2d57432', Title: 'DataSet1', InternalName: 'InternalName', TypeAsString: 'Number', ListId: '5E3CD8EF-673E-4713-8C49-2F3925314202' }] }
              return fields;
          }) as Promise<ISPFields>;
    }
    
    private _renderFieldsAsync(listId: string): void {

      let items: ISPField[];

      // Test environment
      if (this.host.hostType === HostType.TestPage) {
          this._getMockFields(listId).then((response) => {
              this._renderFields(listId, response.value);
          });

          // SharePoint environment
      } else if (this.host.hostType === HostType.ModernPage) {
          this._getFields(listId)
              .then((response) => {
                  this._renderFields(listId, response.value);
              });
      }
    }
    
    private _renderFields(listId: string, fields: ISPField[]) {
      fields.forEach(field => this.alllistfields.push({ Id:field.Id, ListId:listId, Title:field.Title, InternalName:field.InternalName,TypeAsString:field.TypeAsString}));
    }
    
    
    // =========================================
    // Loading items from lists
    // =========================================
    private _getItems(listId: string): Promise<ISPItems> {      
        return SPRequest.get(this.host.pageContext.webAbsoluteUrl + `/_api/web/lists(guid'` + listId + `')/items?$select=ID,` + this.labelcolumn + `,` + this.datacolumn + `&$top=5000`)
          .then((response: Response) => {
            return response.json();
        });
    }
    
    private _getMockItems(listId: string): Promise<ISPItems> {
          return MockSPRequest.get(this.host.pageContext.webAbsoluteUrl).then(() => {
              const items: ISPItems = 
                { value: [
                  { Id: 1, Label: 'A', Data: 1 },
                  { Id: 2, Label: 'A', Data: 1 },
                  { Id: 3, Label: 'A', Data: 1 },
                  { Id: 4, Label: 'B', Data: 2 },
                  { Id: 5, Label: 'B', Data: 2 },
                  { Id: 6, Label: 'B', Data: 2 },
                  { Id: 7, Label: 'C', Data: 3 },
                  { Id: 8, Label: 'C', Data: 3 },
                  { Id: 9, Label: 'C', Data: 3 }
                ] }
              return items;
          }) as Promise<ISPItems>;
    }
    
    private _renderItemsAsync(chartType: string, listId: string): void {

      let items: ISPItems[];
      
      // Test environment
      if (this.host.hostType === HostType.TestPage) {
          this._getMockItems(listId).then((response) => {
              this._renderItems(chartType, response.value);
          });

          // SharePoint environment
      } else if (this.host.hostType === HostType.ModernPage) {
          this._getItems(listId)
              .then((response) => {
                  this._renderItems(chartType, response.value);
              });
      }
    }
    
    private _renderItems(chartType: string, items: ISPItem[]) {
        
        this.labels = [];
        this.chartdataset = [];    
        this.colors = [];                 
        items.forEach((item: ISPItem) => {
            this.labels.push(item[this.labelcolumn]);
            this.chartdataset.push(item[this.datacolumn]);
            this.colors.push(this.properties.color);
        });    
                
      this.render(this.displayMode);
        
    }
}
