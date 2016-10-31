import * as ko from 'knockout';
import styles from './ExcelListWebPart.module.scss';
import { IExcelListWebPartWebPartProps } from './IExcelListWebPartWebPartProps';
import * as ListService from './ListService';


export interface IExcelListWebPartBindingContext extends IExcelListWebPartWebPartProps {
  shouter: KnockoutSubscribable<{}>;
  dataService: ListService.IListsService;
}

export default class ExcelListWebPartViewModel {
  public description: KnockoutObservable<string> = ko.observable('');
  public cssClass: KnockoutObservable<string> = ko.observable('');
  public containerClass: KnockoutObservable<string> = ko.observable('');
  public rowClass: KnockoutObservable<string> = ko.observable('');
  public buttonClass: KnockoutObservable<string> = ko.observable('');

  public lists: KnockoutObservableArray<string> = ko.observableArray([]);
  public selectedList: KnockoutObservable<string> = ko.observable('');
  public columns: KnockoutObservableArray<string> = ko.observableArray([]);
  public selectedColumns: KnockoutObservableArray<string> = ko.observableArray([]);
  public dataHref: KnockoutObservable<string> = ko.observable('');

  dataService: ListService.IListsService;

  constructor(bindings: IExcelListWebPartBindingContext) {
    this.description(bindings.description);
    bindings.shouter.subscribe((value: string) => {
      this.description(value);
    }, this, 'description');

    this.cssClass(styles.excelListWebPart);
    this.containerClass(styles.container);
    this.rowClass(`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`);
    this.buttonClass(`ms-Button ${styles.button}`);

    this.dataService = bindings.dataService;

    this.selectedList.subscribe((selectedList) => {
      if (selectedList == null || selectedList.length == 0) {
        return;
      }
      this.dataService.getListColumns(selectedList).then((columns: string[]) => {
        this.columns.removeAll();
        columns.forEach(column => {
          this.columns.push(column);
        });
      });
    });

    this.dataService.getListNames().then((names: string[]) => {
      this.lists.removeAll();
      names.forEach(name => {
        this.lists.push(name);
      });
    });
  }

  public generateExcel() {
    this.dataService.getListItems(this.selectedList(), this.selectedColumns()).then((data: string[][]) => {
      var csv: string = '';
      csv += this.selectedColumns().join(';');
      csv += '\n';
      data.forEach(row => {
        csv+=row.join(';');
        csv += '\n';
      });
      this.dataHref('data:attachment/csv;charset=utf-8,' + encodeURI(csv));
    });
  };
}
