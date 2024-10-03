import { Component } from '@angular/core';
import { ConfirmationService } from 'primeng/api';
import { MessageService } from 'primeng/api';
import { ExecutionOrderService } from './execution-order.service';
import { ServiceMaster } from '../models/service-master.model';
import { UnitOfMeasure } from '../models/unitOfMeasure.model';
import { ApiService } from '../shared/ApiService.service';
import * as FileSaver from 'file-saver';


import { MainItem } from './execution-order.model';
import { MaterialGroup } from '../models/materialGroup.model';
import { ServiceType } from '../models/serviceType.model';
import { LineType } from '../models/lineType.model';
import { NavigationExtras, Router } from '@angular/router';

@Component({
  selector: 'app-execution-order',
  templateUrl: './execution-order.component.html',
  styleUrls: ['./execution-order.component.scss'],
  providers: [MessageService, ExecutionOrderService, ConfirmationService]
})
export class ExecutionOrderComponent {

  // Pagination:
  loading: boolean = true;
  ///
  searchKey: string = ""
  currency: any
  totalValue: number = 0.0

  selectedExecutionOrder:MainItem[]=[]
  //fields for dropdown lists
  recordsServiceNumber!: ServiceMaster[];
  selectedServiceNumberRecord?: ServiceMaster
  selectedServiceNumber!: number;
  updateSelectedServiceNumber!: number
  updateSelectedServiceNumberRecord?: ServiceMaster
  shortText: string = '';
  updateShortText: string = '';
  shortTextChangeAllowed: boolean = false;
  updateShortTextChangeAllowed: boolean = false;

  recordsUnitOfMeasure: UnitOfMeasure[] = [];
  selectedUnitOfMeasure!: string;

  recordsMaterialGroup: MaterialGroup[] = [];
  selectedMaterialGroup!: string;

  recordsServiceType: ServiceType[] = [];
  selectedServiceType!: string;


  recordsLineType: LineType[] = [];
  selectedLineType: string = "Standard line";

  recordsCurrency!: any[];
  selectedCurrency: string="";
  //
  public rowIndex = 0;

  mainItemsRecords: MainItem[] = [];

  constructor(private router: Router, private _ApiService: ApiService, private _ExecutionOrderService: ExecutionOrderService, private messageService: MessageService, private confirmationService: ConfirmationService) { }

  
  ngOnInit() {
    this._ApiService.get<ServiceMaster[]>('servicenumbers').subscribe(response => {
      this.recordsServiceNumber = response
      //.filter(record => record.deletionIndicator === false);
    });
    this._ApiService.get<MaterialGroup[]>('materialgroups').subscribe(response => {
      this.recordsMaterialGroup = response
    });
    this._ApiService.get<ServiceType[]>('servicetypes').subscribe(response => {
      this.recordsServiceType = response
    });
    this._ApiService.get<LineType[]>('linetypes').subscribe(response => {
      this.recordsLineType = response
    });
    this._ApiService.get<any[]>('currencies').subscribe(response => {
      this.recordsCurrency = response;
    });
    this._ApiService.get<MainItem[]>('executionordermain').subscribe(response => {
      this.mainItemsRecords = response.sort((a, b) => a.executionOrderMainCode - b.executionOrderMainCode);
      console.log(this.mainItemsRecords);
      this.loading = false;

      const filteredRecords = this.mainItemsRecords.filter(record => record.lineTypeCode != "Contingency line");

      this.totalValue = filteredRecords.reduce((sum, record) => sum + record.total, 0);
      console.log('Total Value:', this.totalValue);
    });
  }
  // Helper Functions:
  removePropertiesFrom(obj: any, propertiesToRemove: string[]): any {
    const newObj: any = {};

    for (let key in obj) {
      if (obj.hasOwnProperty(key)) {
        if (Array.isArray(obj[key])) {
          // If the property is an array, recursively remove properties from each element
          newObj[key] = obj[key].map((item: any) => this.removeProperties(item, propertiesToRemove));
        } else if (typeof obj[key] === 'object' && obj[key] !== null) {
          // If the property is an object, recursively remove properties from the object
          newObj[key] = this.removeProperties(obj[key], propertiesToRemove);
        } else if (!propertiesToRemove.includes(key)) {
          // Otherwise, copy the property if it's not in the list to remove
          newObj[key] = obj[key];
        }
      }
    }

    return newObj;
  }
  removeProperties(obj: any, propertiesToRemove: string[]): any {
    const newObj: any = {};
    Object.keys(obj).forEach(key => {
      if (!propertiesToRemove.includes(key)) {
        newObj[key] = obj[key];
      }
    });
    return newObj;
  }
  // to handel checkbox selection:
  selectedMainItem: MainItem | undefined;
  selectedMainItems: MainItem[] = [];

  onMainItemSelection(event: any, mainItem: MainItem) {
    console.log('Event:', event);
    console.log('MainItem before change:', mainItem);

    console.log('Event checked state:', event.checked);
    console.log('MainItem selected before update:', mainItem.selected);

    mainItem.selected = event.checked;
    console.log('MainItem selected after update:', mainItem.selected);

    if (event.checked.length) {
      this.selectedMainItem = mainItem;
      console.log('Selected MainItem:', this.selectedMainItem);

      // Check if mainItem is already in the list to avoid duplication
      if (!this.selectedMainItems.includes(mainItem)) {
        this.selectedMainItems.push(mainItem);
        console.log('Selected Main Items after addition:', this.selectedMainItems);
      }

    } else {
      console.log('Entering else block');
      this.selectedMainItem = undefined;
      console.log('Selected MainItem after deselection:', this.selectedMainItem);

      const index = this.selectedMainItems.indexOf(mainItem);
      console.log('Index of mainItem in selected list:', index);

      if (index !== -1) {
        this.selectedMainItems.splice(index, 1);
        console.log('Selected Main Items after removal:', this.selectedMainItems);
      }
    }
  }

  // to handle All Records Selection / Deselection 
  selectedAllRecords: MainItem[] = [];
  onSelectAllRecords(event: any): void {
    if (Array.isArray(event.checked) && event.checked.length > 0) {
      this.selectedAllRecords = [...this.mainItemsRecords];
      console.log(this.selectedAllRecords);
    } else {
      this.selectedAllRecords = [];
    }
  }

  //In Creation to handle shortTextChangeAlowlled Flag 
  onServiceNumberChange(event: any) {
    const selectedRecord = this.recordsServiceNumber.find(record => record.serviceNumberCode === this.selectedServiceNumber);
    if (selectedRecord) {
      this.selectedServiceNumberRecord = selectedRecord
      this.shortTextChangeAllowed = this.selectedServiceNumberRecord?.shortTextChangeAllowed || false;
      this.shortText = ""
    }
    else {
      console.log("no service number");
      //this.dontSelectServiceNumber = false
      this.selectedServiceNumberRecord = undefined;
    }
  }
  //In Update to handle shortTextChangeAlowlled Flag 
  onServiceNumberUpdateChange(event: any) {
    const updateSelectedRecord = this.recordsServiceNumber.find(record => record.serviceNumberCode === event.value);
    if (updateSelectedRecord) {
      this.updateSelectedServiceNumberRecord = updateSelectedRecord
      this.updateShortTextChangeAllowed = this.updateSelectedServiceNumberRecord?.shortTextChangeAllowed || false;
      this.updateShortText = ""
    }
    else {
      this.updateSelectedServiceNumberRecord = undefined;
    }
  }

  // For Edit  MainItem
  clonedMainItem: { [s: number]: MainItem } = {};
  onMainItemEditInit(record: MainItem) {
    console.log(record);
    
    this.clonedMainItem[record.executionOrderMainCode] = { ...record };
  }
  onMainItemEditSave(index: number, record: MainItem) {
    console.log(record);
    const updatedMainItem = this.removePropertiesFrom(record, ['executionOrderMainCode','total']);
    console.log(updatedMainItem);
    console.log(this.updateSelectedServiceNumber);

    if (this.updateSelectedServiceNumberRecord) {
      const newRecord: MainItem = {
      // ...record, // Copy all properties from the original record
      ...updatedMainItem,
        unitOfMeasurementCode: this.updateSelectedServiceNumberRecord.baseUnitOfMeasurement,
        description: this.updateSelectedServiceNumberRecord.description,
        materialGroupCode: this.updateSelectedServiceNumberRecord.materialGroupCode,
        serviceTypeCode: this.updateSelectedServiceNumberRecord.serviceTypeCode,
      };
      console.log(newRecord);
      console.log(newRecord);
         // Remove properties with empty or default values
         const filteredRecord = Object.fromEntries(
          Object.entries(newRecord).filter(([_, value]) => {
            return value !== '' && value !== 0 && value !== undefined && value !== null;
          })
        );
        console.log(filteredRecord);
      
      this._ApiService.patch<MainItem>('executionordermain', record.executionOrderMainCode, filteredRecord).subscribe({
        next: (res) => {
          console.log('executionordermain  updated:', res);
          this.totalValue = 0;
          this.ngOnInit()
        }, error: (err) => {
          console.log(err);
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Invalid Data' });
        },
        complete: () => {
          this.updateSelectedServiceNumberRecord=undefined;
          this.messageService.add({ severity: 'success', summary: 'Success', detail: 'Record updated successfully ' });
          // this.ngOnInit()
        }
       
      });
    }

    if (!this.updateSelectedServiceNumberRecord) {
      console.log(updatedMainItem);
       // Remove properties with empty or default values
       const filteredRecord = Object.fromEntries(
        Object.entries(updatedMainItem).filter(([_, value]) => {
          return value !== '' && value !== 0 && value !== undefined && value !== null;
        })
      );
      console.log(filteredRecord);
      this._ApiService.patch<MainItem>('executionordermain', record.executionOrderMainCode, filteredRecord).subscribe({
        next: (res) => {
          console.log('executionordermain  updated:', res);
          this.totalValue = 0;
          this.ngOnInit()
        }, error: (err) => {
          console.log(err);
          this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Invalid Data' });
        },
        complete: () => {
          this.messageService.add({ severity: 'success', summary: 'Success', detail: 'Record updated successfully ' });
        }
      });
    }
  }
  onMainItemEditCancel(row: MainItem, index: number) {
    this.mainItemsRecords[index] = this.clonedMainItem[row.executionOrderMainCode]
    delete this.clonedMainItem[row.executionOrderMainCode]
  }

  // Delete MainItem 
  deleteRecord() {
    console.log("delete");
    console.log(this.selectedExecutionOrder);
    
    if (this.selectedExecutionOrder.length) {
      this.confirmationService.confirm({
        message: 'Are you sure you want to delete the selected record?',
        header: 'Confirm',
        icon: 'pi pi-exclamation-triangle',
        accept: () => {
          for (const record of this.selectedExecutionOrder) {
            console.log(record);
            this._ApiService.delete<MainItem>('executionordermain', record.executionOrderMainCode).subscribe(response => {
              console.log('executionordermain deleted :', response);
              this.totalValue = 0;
              this.ngOnInit();
            });
          }
          this.messageService.add({ severity: 'success', summary: 'Successfully', detail: 'Deleted', life: 3000 });
          this.selectedMainItems = []; // Clear the selectedRecords array after deleting all records
        }
      });
    }
  }

  // For Add new  Main Item
  newMainItem: MainItem = {
    Type:'',
    executionOrderMainCode: 0,
    serviceNumberCode: 0,
    description: "",
    unitOfMeasurementCode: "",
    currencyCode: "",
    materialGroupCode: "",
    serviceTypeCode: "",
    personnelNumberCode: "",
    lineTypeCode: "",

    totalQuantity: 0,
    amountPerUnit: 0,
    total: 0,
    actualQuantity: 0,
    actualPercentage: 0,
    overFulfillmentPercentage: 0,
    unlimitedOverFulfillment: false,
    manualPriceEntryAllowed: false,
    externalServiceNumber: "",
    serviceText: "",
    lineText: "",
    lineNumber: "",

    biddersLine: false,
    supplementaryLine: false,
    lotCostOne: false,
    doNotPrint: false,


  };

  addMainItem() {
    if (!this.selectedServiceNumberRecord) { // if user didn't select serviceNumber 

      const newRecord = {
        unitOfMeasurementCode: this.selectedUnitOfMeasure,
        currencyCode: this.selectedCurrency,
        description: this.newMainItem.description,

        materialGroupCode: this.selectedMaterialGroup,
        serviceTypeCode: this.selectedServiceType,
        // lesa .....
        personnelNumberCode: this.newMainItem.personnelNumberCode,
        lineTypeCode: this.selectedLineType,

        totalQuantity: this.newMainItem.totalQuantity,
        amountPerUnit: this.newMainItem.amountPerUnit,
        total: this.newMainItem.total,
        actualQuantity: this.newMainItem.actualQuantity,
        actualPercentage: this.newMainItem.actualPercentage,

        overFulfillmentPercentage: this.newMainItem.overFulfillmentPercentage,
        unlimitedOverFulfillment: this.newMainItem.unlimitedOverFulfillment,
        manualPriceEntryAllowed: this.newMainItem.manualPriceEntryAllowed,
        externalServiceNumber: this.newMainItem.externalServiceNumber,
        serviceText: this.newMainItem.serviceText,
        lineText: this.newMainItem.lineText,
        lineNumber: this.newMainItem.lineNumber,

        biddersLine: this.newMainItem.biddersLine,
        supplementaryLine: this.newMainItem.supplementaryLine,
        lotCostOne: this.newMainItem.lotCostOne,
        doNotPrint: this.newMainItem.doNotPrint,
      }
      if (this.newMainItem.totalQuantity === 0 || this.newMainItem.description === "" || this.selectedCurrency === "") {
        // || this.newMainItem.unitOfMeasurementCode === ""  // till retrieved from cloud correctly
        this.messageService.add({
          severity: 'error',
          summary: 'Error',
          detail: 'Description & Quantity & Currency and UnitOfMeasurement are required',
          life: 3000
        });
      }
      else {
        console.log(newRecord);
        // Remove properties with empty or default values
        const filteredRecord = Object.fromEntries(
          Object.entries(newRecord).filter(([_, value]) => {
            return value !== '' && value !== 0 && value !== undefined && value !== null;
          })
        );
        console.log(filteredRecord);
        this._ApiService.post<MainItem>('executionordermain', filteredRecord).subscribe({
          next: (res) => {
            console.log('executionordermain created:', res);
            this.totalValue = 0;
            this.ngOnInit()
          }, error: (err) => {
            console.log(err);
          },
          complete: () => {
            this.resetNewMainItem();
            this.selectedUnitOfMeasure = "";
            this.selectedCurrency = "";
            this.selectedMaterialGroup = "";
            this.selectedServiceType = "";
            this.selectedLineType = "";

            this.messageService.add({ severity: 'success', summary: 'Success', detail: 'Record added successfully ' });
            // this.ngOnInit()
          }
        });
        // this._ApiService.post<MainItem>('executionordermain', filteredRecord).subscribe((response: MainItem) => {
        //   console.log('executionordermain created:', response);
        //   if (response) {
        //     this.resetNewMainItem();
        //   }
        //   console.log(response);
        //   this.totalValue = 0;
        //   this.ngOnInit()
        // });
      }
    }

    else if (this.selectedServiceNumberRecord) { // if user select serviceNumber 
      const newRecord = {
        serviceNumberCode: this.selectedServiceNumber,
        unitOfMeasurementCode: this.selectedServiceNumberRecord?.baseUnitOfMeasurement,
        currencyCode: this.selectedCurrency,

        description: this.selectedServiceNumberRecord?.description,
        // lesa .....
        materialGroupCode: this.selectedServiceNumberRecord?.materialGroupCode,
        serviceTypeCode: this.selectedServiceNumberRecord?.serviceTypeCode,
        personnelNumberCode: this.newMainItem.personnelNumberCode,
        lineTypeCode: this.selectedLineType,

        totalQuantity: this.newMainItem.totalQuantity,
        amountPerUnit: this.newMainItem.amountPerUnit,
        total: this.newMainItem.total,
        actualQuantity: this.newMainItem.actualQuantity,
        actualPercentage: this.newMainItem.actualPercentage,

        overFulfillmentPercentage: this.newMainItem.overFulfillmentPercentage,
        unlimitedOverFulfillment: this.newMainItem.unlimitedOverFulfillment,
        manualPriceEntryAllowed: this.newMainItem.manualPriceEntryAllowed,
        externalServiceNumber: this.newMainItem.externalServiceNumber,
        serviceText: this.newMainItem.serviceText,
        lineText: this.newMainItem.lineText,
        lineNumber: this.newMainItem.lineNumber,

        biddersLine: this.newMainItem.biddersLine,
        supplementaryLine: this.newMainItem.supplementaryLine,
        lotCostOne: this.newMainItem.lotCostOne,
        doNotPrint: this.newMainItem.doNotPrint,
      }
      if (this.newMainItem.totalQuantity === 0 || this.selectedServiceNumberRecord.description === "" || this.selectedCurrency === "") {
        // || this.newMainItem.unitOfMeasurementCode === ""  // till retrieved from cloud correctly
        this.messageService.add({
          severity: 'error',
          summary: 'Error',
          detail: 'Description & Quantity & Currency and UnitOfMeasurement are required',
          life: 3000
        });
      }
      else {
        console.log(newRecord);
        // Remove properties with empty or default values
        const filteredRecord = Object.fromEntries(
          Object.entries(newRecord).filter(([_, value]) => {
            return value !== '' && value !== 0 && value !== undefined && value !== null;
          })
        );
        console.log(filteredRecord);
        this._ApiService.post<MainItem>('executionordermain', filteredRecord).subscribe({
          next: (res) => {
            console.log('executionordermain created:', res);
            this.totalValue = 0;
            this.ngOnInit()
          }, error: (err) => {
            console.log(err);
          },
          complete: () => {
            this.resetNewMainItem();
            this.selectedServiceNumberRecord = undefined;
            //this.selectedServiceNumber = 0;
            this.selectedLineType = "";
            this.selectedCurrency = ""
            this.messageService.add({ severity: 'success', summary: 'Success', detail: 'Record added successfully ' });
            //this.ngOnInit()
          }
        });

        // this._ApiService.post<MainItem>('executionordermain', filteredRecord).subscribe((response: MainItem) => {
        //   console.log('executionordermain created:', response);
        //   if (response) {
        //     this.resetNewMainItem();
        //     this.selectedServiceNumberRecord = undefined;
        //     this.selectedServiceNumber = 0;
        //     this.selectedLineType = "";
        //     this.selectedCurrency = ""
        //   }
        //   console.log(response);
        //   this.totalValue = 0;
        //   this.ngOnInit()
        // });
      }
    }
  }

  resetNewMainItem() {
    this.newMainItem = {
      Type:'',
      executionOrderMainCode: 0,
      serviceNumberCode: 0,
      description: "",
      unitOfMeasurementCode: "",
      currencyCode: "",
      materialGroupCode: "",
      serviceTypeCode: "",
      personnelNumberCode: "",
      lineTypeCode: "",

      totalQuantity: 0,
      amountPerUnit: 0,
      total: 0,
      actualQuantity: 0,
      actualPercentage: 0,
      overFulfillmentPercentage: 0,
      unlimitedOverFulfillment: false,
      manualPriceEntryAllowed: false,
      externalServiceNumber: "",
      serviceText: "",
      lineText: "",
      lineNumber: "",

      biddersLine: false,
      supplementaryLine: false,
      lotCostOne: false,
      doNotPrint: false,
    },
      this.selectedUnitOfMeasure = '';
    // this.selectedServiceNumber=0
  }
  
  // Export to excel sheet:
  transformData(data: MainItem[]) {
    const transformed: MainItem[] = []

    data.forEach((mainItem) => {
      transformed.push({
        Type: 'Main Item',
        serviceNumberCode: mainItem.serviceNumberCode,
        description: mainItem.description,
        totalQuantity: mainItem.totalQuantity,
        unitOfMeasurementCode: mainItem.unitOfMeasurementCode,
        amountPerUnit: mainItem.amountPerUnit,
        currencyCode: mainItem.currencyCode,
        total: mainItem.total,
        actualQuantity: mainItem.actualQuantity,
        actualPercentage: mainItem.actualPercentage,
        overFulfillmentPercentage: mainItem.overFulfillmentPercentage,
        unlimitedOverFulfillment: mainItem.unlimitedOverFulfillment,
        manualPriceEntryAllowed: mainItem.manualPriceEntryAllowed,
        materialGroupCode: mainItem.materialGroupCode,
        serviceTypeCode: mainItem.serviceTypeCode,
        externalServiceNumber: mainItem.externalServiceNumber,
        serviceText: mainItem.serviceText,
        lineText: mainItem.lineText,
        personnelNumberCode: mainItem.personnelNumberCode,
        lineTypeCode: mainItem.lineTypeCode,
        lineNumber: mainItem.lineNumber,
        biddersLine: mainItem.biddersLine,
        supplementaryLine: mainItem.supplementaryLine,
        lotCostOne: mainItem.lotCostOne,

        doNotPrint: mainItem.doNotPrint,
        executionOrderMainCode: mainItem.executionOrderMainCode
      });

    });

    return transformed;
  }
  exportExcel() {
    import('xlsx').then((xlsx) => {
      const transformedData = this.transformData(this.mainItemsRecords);
      const worksheet = xlsx.utils.json_to_sheet(transformedData);
      const workbook = { Sheets: { data: worksheet }, SheetNames: ['data'] };
      const ws = workbook.Sheets.data;
      if (!ws['!ref']) {
        ws['!ref'] = 'A1:Z1000';
      }
      const range = xlsx.utils.decode_range(ws['!ref']);
      let rowStart = 1;

      transformedData.forEach((row, index) => {
        if (row.Type === 'Main Item') {

          if (index + 1 < transformedData.length && transformedData[index + 1].Type === 'Sub Item') {
            ws['!rows'] = ws['!rows'] || [];
            ws['!rows'][index] = { hidden: false };
            ws['!rows'][index + 1] = { hidden: false };
          } else {
            ws['!rows'] = ws['!rows'] || [];
            ws['!rows'][index] = { hidden: false };
          }
        } else {
          ws['!rows'] = ws['!rows'] || [];
          ws['!rows'][index] = { hidden: false };
        }
      });

      const excelBuffer: any = xlsx.write(workbook, { bookType: 'xlsx', type: 'array' });
      this.saveAsExcelFile(excelBuffer, 'Execution Order');
    });
  }
  saveAsExcelFile(buffer: any, fileName: string): void {
    let EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
    let EXCEL_EXTENSION = '.xlsx';
    const data: Blob = new Blob([buffer], {
      type: EXCEL_TYPE
    });
    FileSaver.saveAs(data, fileName + '_export_' + new Date().getTime() + EXCEL_EXTENSION);
  }

}
