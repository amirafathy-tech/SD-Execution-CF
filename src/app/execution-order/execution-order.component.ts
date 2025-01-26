import { ChangeDetectionStrategy, ChangeDetectorRef, Component } from '@angular/core';
import { ConfirmationService } from 'primeng/api';
import { MessageService } from 'primeng/api';
import { ExecutionOrderService } from './execution-order.service';
import { ServiceMaster } from '../models/service-master.model';
import { UnitOfMeasure } from '../models/unitOfMeasure.model';
import { ApiService } from '../shared/ApiService.service';
import * as FileSaver from 'file-saver';
import * as XLSX from 'xlsx';

import { MainItem } from './execution-order.model';
import { MaterialGroup } from '../models/materialGroup.model';
import { ServiceType } from '../models/serviceType.model';
import { LineType } from '../models/lineType.model';
import { Router } from '@angular/router';
import { MainItemSalesQuotation } from '../models/sales-quotation.model';
import { ModelEntity, ModelSpecDetails } from '../models/model-specs.model';
import { forkJoin } from 'rxjs';
@Component({
  selector: 'app-execution-order',
  templateUrl: './execution-order.component.html',
  styleUrls: ['./execution-order.component.scss'],
  providers: [MessageService, ExecutionOrderService, ConfirmationService],
  changeDetection: ChangeDetectionStrategy.Default
})
export class ExecutionOrderComponent {

  //cloud data:
  documentNumber!: number;
  itemNumber!: number;
  customerId!: number;
  referenceSDDocument!: number;
  itemText: string = "";
  cloudCurrency!:string;

  displayImportsDialog = false;
  displayTenderingDocumentDialog = false;
  displayModelSpecsDialog = false;
  displayModelSpecsDetailsDialog = false;
  displayExcelDialog = false;

  selectedSalesQuotations: MainItemSalesQuotation[] = [];
  salesQuotations: MainItemSalesQuotation[] = [];

  selectedModelSpecsDetails: ModelSpecDetails[] = [];
  models: ModelEntity[] = [];
  modelSpecsDetails: ModelSpecDetails[] = [];

  savedInMemory: boolean = false;

  // Pagination:
  loading: boolean = true;
  ///
  searchKey: string = ""
  currency: any
  totalValue: number = 0.0

  selectedExecutionOrder: MainItem[] = []
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
  selectedCurrency: string = "";
  //
  public rowIndex = 0;

  mainItemsRecords: MainItem[] = [];

  constructor(private cdr: ChangeDetectorRef, private router: Router, private _ApiService: ApiService, private _ExecutionOrderService: ExecutionOrderService, private messageService: MessageService, private confirmationService: ConfirmationService) {

    this.documentNumber = this.router.getCurrentNavigation()?.extras.state?.['documentNumber'];
    this.itemNumber = this.router.getCurrentNavigation()?.extras.state?.['itemNumber'];
    this.customerId = this.router.getCurrentNavigation()?.extras.state?.['customerId'];
    this.referenceSDDocument = this.router.getCurrentNavigation()?.extras.state?.['referenceSDDocument'];
    console.log(this.documentNumber, this.itemNumber, this.customerId, this.referenceSDDocument);
    this.cloudCurrency = this.router.getCurrentNavigation()?.extras.state?.['currency'];
    console.log(this.documentNumber, this.itemNumber, this.customerId,this.cloudCurrency);

  }
  ngOnInit() {
    this._ApiService.get<ServiceMaster[]>('servicenumbers').subscribe(response => {
      this.recordsServiceNumber = response;
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
    this._ApiService.get<any[]>('measurements').subscribe(response => {
      this.recordsUnitOfMeasure = response;
    });
    if (this.savedInMemory) {
      this.mainItemsRecords = [...this._ExecutionOrderService.getMainItems()];
      console.log(this.mainItemsRecords);
    }
    //localhost:8080/executionordermain/referenceid?referenceId=6&salesOrderItem=10
    this._ApiService.get<MainItem[]>(`executionordermain/referenceid?referenceId=${this.documentNumber}&salesOrderItem=${this.itemNumber}`).subscribe({
      next: (res) => {
        this.mainItemsRecords = res.sort((a, b) => a.executionOrderMainCode - b.executionOrderMainCode);
        this.itemText = this.mainItemsRecords[0].salesOrderItemText ? this.mainItemsRecords[0].salesOrderItemText : "";
        console.log(this.itemText);
        console.log(this.mainItemsRecords);
        this.loading = false;
        const filteredRecords = this.mainItemsRecords.filter(record => record.lineTypeCode != "Contingency line");
        this.totalValue = filteredRecords.reduce((sum, record) => sum + record.total, 0);
        console.log('Total Value:', this.totalValue);
      }, error: (err) => {
        console.log(err);
        console.log(err.status);
        if (err.status == 404) {
          this.mainItemsRecords = [];
          this.loading = false;
          this.totalValue = this.mainItemsRecords.reduce((sum, record) => sum + record.total, 0);
          console.log('Total Value:', this.totalValue);
        }
      },
      complete: () => {
      }
    });

    // this._ApiService.get<MainItemSalesQuotation[]>(`mainitems/all`).subscribe({
    //   next: (res) => {
    //     this.salesQuotations = res.sort((a, b) => a.invoiceMainItemCode - b.invoiceMainItemCode);
    //     console.log(this.mainItemsRecords);
    //   }, error: (err) => {
    //     console.log(err);
    //   },
    //   complete: () => {
    //   }
    // });
  }

  showImportsDialog() {
    this.displayImportsDialog = true;

  }
  showExcelDialog() {
    this.displayExcelDialog = true;

  }

  showSalesQuotationDialog() {
    this.displayTenderingDocumentDialog = true;
    // localhost:8080/mainitems?salesOrder=12&salesOrderItem=10

    this._ApiService.get<MainItemSalesQuotation[]>(`mainitems?salesOrder=${this.documentNumber}&salesOrderItem=${this.itemNumber}`).subscribe({
      // next: (res) => {
      //   this.salesQuotations = res.sort((a, b) => a.invoiceMainItemCode - b.invoiceMainItemCode);
      //   console.log(this.mainItemsRecords);
      // }
      next: (res) => {
        const uniqueRecords = res.filter(newRecord =>
          !this.mainItemsRecords.some(existingRecord =>
            existingRecord.invoiceMainItemCode === newRecord.invoiceMainItemCode
          )
        );
        this.salesQuotations = uniqueRecords.sort((a, b) => a.invoiceMainItemCode - b.invoiceMainItemCode);
        console.log(this.mainItemsRecords);
      }
      , error: (err) => {
        console.log(err);
      },
      complete: () => {
      }
    });
  }
  showModelSpecsDialog() {
    this.displayModelSpecsDialog = true;
    this._ApiService.get<ModelEntity[]>(`modelspecs`).subscribe({
      next: (res) => {
        // const uniqueRecords = res.filter(newRecord => 
        //   !this.mainItemsRecords.some(existingRecord => 
        //     existingRecord.invoiceMainItemCode === newRecord.modelSpecCode
        //   )
        // );
        this.models = res.sort((a, b) => a.modelSpecCode - b.modelSpecCode);
        console.log(this.models);
      }
      , error: (err) => {
        console.log(err);
      },
      complete: () => {
      }
    });
  }
  showModelSpecsDetailsDialog(model: ModelEntity) {
    this.displayModelSpecsDetailsDialog = true;
    const detailObservables = model.modelSpecDetailsCode.map(code =>
      this._ApiService.getID<ModelSpecDetails>('modelspecdetails', code)
    );
    forkJoin(detailObservables).subscribe(records => {
      this.modelSpecsDetails = records.sort((a, b) => b.modelSpecDetailsCode - a.modelSpecDetailsCode);
    });
  }
  saveSelectionModelSpecsDetails() {
    console.log('Selected items:', this.selectedModelSpecsDetails);
    this.displayModelSpecsDetailsDialog = false;
    this.displayModelSpecsDialog = false;
    this.displayImportsDialog = false;
  }

  openDocumentDialog() {
    this.displayTenderingDocumentDialog = true;
  }
  saveSelection() {
    console.log('Selected items:', this.selectedSalesQuotations);
    this.displayTenderingDocumentDialog = false;
  }
  cancelMainItemSalesQuotation(item: any): void {
    this.selectedSalesQuotations = this.selectedSalesQuotations.filter(i => i !== item);
  }
  // for selected sales quotation:
  saveMainItem(mainItem: MainItemSalesQuotation) {
    console.log(mainItem);
    const newRecord: MainItem = {
      //
      invoiceMainItemCode: mainItem.invoiceMainItemCode,
      //
      serviceNumberCode: mainItem.serviceNumberCode,
      unitOfMeasurementCode: mainItem.unitOfMeasurementCode,
      //this.selectedServiceNumberRecord?.baseUnitOfMeasurement,
      currencyCode: mainItem.currencyCode,
      description: mainItem.description,
      materialGroupCode: mainItem.materialGroupCode,
      serviceTypeCode: mainItem.serviceTypeCode,
      personnelNumberCode: mainItem.personnelNumberCode,
      lineTypeCode: mainItem.lineTypeCode,
      totalQuantity: mainItem.quantity,
      amountPerUnit: mainItem.amountPerUnitWithProfit,
      total: mainItem.totalWithProfit,
      actualQuantity: mainItem.actualQuantity,
      actualPercentage: mainItem.actualPercentage,
      overFulfillmentPercentage: mainItem.overFulfillmentPercentage,
      unlimitedOverFulfillment: mainItem.unlimitedOverFulfillment,
      manualPriceEntryAllowed: mainItem.manualPriceEntryAllowed,
      externalServiceNumber: mainItem.externalServiceNumber,
      serviceText: mainItem.serviceText,
      lineText: mainItem.lineText,
      lineNumber: mainItem.lineNumber,
      biddersLine: mainItem.biddersLine,
      supplementaryLine: mainItem.supplementaryLine,
      lotCostOne: mainItem.lotCostOne,
      doNotPrint: mainItem.doNotPrint,
      Type: '',
      executionOrderMainCode: 0
    }
    console.log(newRecord);
    if (newRecord.totalQuantity === 0) {
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: ' Quantity is required',
        life: 3000
      });
    }
    else {
      console.log(newRecord);
      //................
      const bodyRequest: any = {
        quantity: newRecord.totalQuantity,
        amountPerUnit: newRecord.amountPerUnit,
      };
      this._ApiService.post<any>(`/total`, bodyRequest).subscribe({
        next: (res) => {
          console.log('mainitem with total:', res);
          newRecord.total = res.total;
          console.log(' Record:', newRecord);
          const filteredRecord = Object.fromEntries(
            Object.entries(newRecord).filter(([_, value]) => {
              return value !== '' && value !== 0 && value !== undefined && value !== null;
            })
          ) as MainItem;
          console.log(filteredRecord);
          this._ExecutionOrderService.addMainItem(filteredRecord);
          this.savedInMemory = true;
          // this.cdr.detectChanges();
          const newMainItems = this._ExecutionOrderService.getMainItems();
          // Combine the current mainItemsRecords with the new list, ensuring no duplicates
          this.mainItemsRecords = [
            ...this.mainItemsRecords.filter(item => !newMainItems.some(newItem => newItem.executionOrderMainCode === item.executionOrderMainCode)), // Remove existing items
            ...newMainItems
          ];
          this.updateTotalValueAfterAction();
          console.log(this.mainItemsRecords);
          this.resetNewMainItem();
          const index = this.selectedSalesQuotations.findIndex(item => item.invoiceMainItemCode === mainItem.invoiceMainItemCode);
          if (index !== -1) {
            this.selectedSalesQuotations.splice(index, 1);
          }
        }, error: (err) => {
          console.log(err);
        },
        complete: () => {
        }
      });
      //................
    }
  }
  cancelModelSpecsDetails(item: any): void {
    this.selectedModelSpecsDetails = this.selectedModelSpecsDetails.filter(i => i !== item);
  }
  // for selected models specs details:
  saveModelSpecsDetails(item: ModelSpecDetails) {
    console.log(item);
    const newRecord: MainItem = {
      //
      invoiceMainItemCode: item.modelSpecDetailsCode,
      //
      serviceNumberCode: item.serviceNumberCode,
      unitOfMeasurementCode: item.unitOfMeasurementCode,
      //this.selectedServiceNumberRecord?.baseUnitOfMeasurement,
      currencyCode: item.currencyCode,
      description: item.shortText,
      materialGroupCode: item.materialGroupCode,
      serviceTypeCode: item.serviceTypeCode,
      personnelNumberCode: item.personnelNumberCode,
      lineTypeCode: item.lineTypeCode,
      totalQuantity: item.quantity,
      amountPerUnit: item.grossPrice,
      total: item.netValue,
      actualQuantity: item.actualQuantity,
      actualPercentage: item.actualPercentage,
      overFulfillmentPercentage: item.overFulfilmentPercentage,
      unlimitedOverFulfillment: item.unlimitedOverFulfillment,
      manualPriceEntryAllowed: item.manualPriceEntryAllowed,
      externalServiceNumber: item.externalServiceNumber,
      serviceText: item.serviceText,
      lineText: item.lineText,
      lineNumber: item.lineNumber,
      biddersLine: item.biddersLine,
      supplementaryLine: item.supplementaryLine,
      lotCostOne: item.lotSizeForCostingIsOne,
      // doNotPrint: item.doNotPrint,
      Type: '',
      executionOrderMainCode: 0
    }
    console.log(newRecord);
    if (newRecord.totalQuantity === 0) {
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: ' Quantity is required',
        life: 3000
      });
    }
    else {
      console.log(newRecord);
      //................
      const bodyRequest: any = {
        quantity: newRecord.totalQuantity,
        amountPerUnit: newRecord.amountPerUnit,
      };
      this._ApiService.post<any>(`/total`, bodyRequest).subscribe({
        next: (res) => {
          console.log('mainitem with total:', res);
          newRecord.total = res.total;
          console.log(' Record:', newRecord);
          const filteredRecord = Object.fromEntries(
            Object.entries(newRecord).filter(([_, value]) => {
              return value !== '' && value !== 0 && value !== undefined && value !== null;
            })
          ) as MainItem;
          console.log(filteredRecord);
          this._ExecutionOrderService.addMainItem(filteredRecord);
          this.savedInMemory = true;
          // this.cdr.detectChanges();
          const newMainItems = this._ExecutionOrderService.getMainItems();
          // Combine the current mainItemsRecords with the new list, ensuring no duplicates
          this.mainItemsRecords = [
            ...this.mainItemsRecords.filter(item => !newMainItems.some(newItem => newItem.executionOrderMainCode === item.executionOrderMainCode)), // Remove existing items
            ...newMainItems
          ];
          this.updateTotalValueAfterAction();
          console.log(this.mainItemsRecords);
          this.resetNewMainItem();
          const index = this.selectedModelSpecsDetails.findIndex(item => item.modelSpecDetailsCode === item.modelSpecDetailsCode);
          if (index !== -1) {
            this.selectedModelSpecsDetails.splice(index, 1);
          }
        }, error: (err) => {
          console.log(err);
        },
        complete: () => {
        }
      });
      //................
    }
  }
  // for selected from excel sheet:
  saveMainItemFromExcel(mainItem: MainItem) {
    console.log(mainItem);
    const newRecord: MainItem = {
      //
      invoiceMainItemCode: mainItem.invoiceMainItemCode,
      //
      serviceNumberCode: mainItem.serviceNumberCode,
      unitOfMeasurementCode: mainItem.unitOfMeasurementCode,
      //this.selectedServiceNumberRecord?.baseUnitOfMeasurement,
      currencyCode: mainItem.currencyCode,
      description: mainItem.description,
      materialGroupCode: mainItem.materialGroupCode,
      serviceTypeCode: mainItem.serviceTypeCode,
      personnelNumberCode: mainItem.personnelNumberCode,
      lineTypeCode: mainItem.lineTypeCode,
      totalQuantity: mainItem.totalQuantity,
      amountPerUnit: mainItem.amountPerUnit,
      total: mainItem.total,
      actualQuantity: mainItem.actualQuantity,
      actualPercentage: mainItem.actualPercentage,
      overFulfillmentPercentage: mainItem.overFulfillmentPercentage,
      unlimitedOverFulfillment: mainItem.unlimitedOverFulfillment,
      manualPriceEntryAllowed: mainItem.manualPriceEntryAllowed,
      externalServiceNumber: mainItem.externalServiceNumber,
      serviceText: mainItem.serviceText,
      lineText: mainItem.lineText,
      lineNumber: mainItem.lineNumber,
      biddersLine: mainItem.biddersLine,
      supplementaryLine: mainItem.supplementaryLine,
      lotCostOne: mainItem.lotCostOne,
      doNotPrint: mainItem.doNotPrint,
      Type: '',
      executionOrderMainCode: 0
    }
    console.log(newRecord);
    if (newRecord.totalQuantity === 0) {
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: ' Quantity is required',
        life: 3000
      });
    }
    else {
      console.log(newRecord);
      //................
      const bodyRequest: any = {
        quantity: newRecord.totalQuantity,
        amountPerUnit: newRecord.amountPerUnit,
      };
      this._ApiService.post<any>(`/total`, bodyRequest).subscribe({
        next: (res) => {
          console.log('mainitem with total:', res);
          newRecord.total = res.total;
          console.log(' Record:', newRecord);
          const filteredRecord = Object.fromEntries(
            Object.entries(newRecord).filter(([_, value]) => {
              return value !== '' && value !== 0 && value !== undefined && value !== null;
            })
          ) as MainItem;
          console.log(filteredRecord);
          this._ExecutionOrderService.addMainItem(filteredRecord);
          this.savedInMemory = true;
          // this.cdr.detectChanges();
          const newMainItems = this._ExecutionOrderService.getMainItems();
          // Combine the current mainItemsRecords with the new list, ensuring no duplicates
          this.mainItemsRecords = [
            ...this.mainItemsRecords.filter(item => !newMainItems.some(newItem => newItem.executionOrderMainCode === item.executionOrderMainCode)), // Remove existing items
            ...newMainItems
          ];
          this.updateTotalValueAfterAction();
          console.log(this.mainItemsRecords);
          this.resetNewMainItem();
          const index = this.parsedData.findIndex(item => item.executionOrderMainCode === mainItem.executionOrderMainCode);
          if (index !== -1) {
            this.parsedData.splice(index, 1);
          }
        }, error: (err) => {
          console.log(err);
        },
        complete: () => {
        }
      });
      //................
    }
  }
  cancelFromExcel(item: any): void {
    this.parsedData = this.parsedData.filter(i => i !== item);
  }
  calculateTotalValue(): void {
    this.totalValue = this.mainItemsRecords.reduce((sum, item) => sum + (item.total || 0), 0);
  }
  updateTotalValueAfterAction(): void {
    this.calculateTotalValue();
    console.log('Updated Total Value:', this.totalValue);
  }

  // Excel Import:
  parsedData: MainItem[] = []; // Parsed data from the Excel file
  displayedColumns: string[] = []; // Column headers from the Excel file

  onFileSelect(event: any, fileUploader: any) {

    console.log('Records before :', this.parsedData);
    
    const file = event.files[0];
    const reader = new FileReader();
  
    reader.onload = (e: any) => {
      const binaryData = e.target.result;
      const workbook = XLSX.read(binaryData, { type: 'binary' });
  
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
  
      const jsonData: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  
      if (jsonData.length > 0) {
        this.displayedColumns = jsonData[0].filter((col: any) => typeof col === 'string' && col.trim() !== '') as string[];
        this.parsedData = jsonData.slice(1).map((row: any[]) => {
          const rowData: any = {};
          this.displayedColumns.forEach((col, index) => {
            rowData[col] = row[index] !== undefined ? row[index] : '';
          });
          return rowData;
        });
        console.log('Records :', this.parsedData);
      this.messageService.add({
        severity: 'success',
        summary: 'Success',
        detail: 'Records copied from the excel sheet successfully!',
        life: 4000
      });
      } else {
        this.displayedColumns = [];
        this.parsedData = [];
      }
  
      // Reset the file input using the PrimeNG method
      fileUploader.clear();
    };
  
    reader.readAsBinaryString(file);
  }
  
  importSelectedRecords() {
    console.log('Records to Import:', this.parsedData);
    // You can now iterate over the parsed data and copy it to your main table
    this.parsedData.forEach(record => {
      console.log('Importing record:', record);
      // Copy or process each record as needed
    });
    this.displayImportsDialog = false;
    this.displayExcelDialog = false;
  }



  //End Excel Import:


  // For Edit  MainItem
  clonedMainItem: { [s: number]: MainItem } = {};
  onMainItemEditInit(record: MainItem) {
    console.log(record);

    this.clonedMainItem[record.executionOrderMainCode] = { ...record };
  }
  onMainItemEditSave(index: number, record: MainItem) {
    console.log(record);
    const updatedMainItem = this.removePropertiesFrom(record, ['executionOrderMainCode']);
    console.log(updatedMainItem);
    console.log(this.updateSelectedServiceNumber);

    if (this.updateSelectedServiceNumberRecord) {
      const newRecord: MainItem = {
        // ...record, // Copy all properties from the original record
        ...updatedMainItem,
        unitOfMeasurementCode: this.updateSelectedServiceNumberRecord.unitOfMeasurementCode,
        //this.updateSelectedServiceNumberRecord.baseUnitOfMeasurement,
        description: this.updateSelectedServiceNumberRecord.description,
        materialGroupCode: this.updateSelectedServiceNumberRecord.materialGroupCode,
        serviceTypeCode: this.updateSelectedServiceNumberRecord.serviceTypeCode,
      };
      console.log(newRecord);

      // Remove properties with empty or default values
      const filteredRecord = Object.fromEntries(
        Object.entries(newRecord).filter(([_, value]) => {
          return value !== '' && value !== 0 && value !== undefined && value !== null;
        })
      );
      console.log(filteredRecord);
      //....................
      const bodyRequest: any = {
        quantity: newRecord.totalQuantity,
        amountPerUnit: newRecord.amountPerUnit
      };
      this._ApiService.post<any>(`/total`, bodyRequest).subscribe({
        next: (res) => {
          console.log('mainitem with total:', res);
          newRecord.total = res.total;
          const mainItemIndex = this.mainItemsRecords.findIndex(item => item.executionOrderMainCode === index);
          if (mainItemIndex > -1) {
            console.log(newRecord);

            // Replace the object entirely to ensure Angular detects the change
            this.mainItemsRecords[mainItemIndex] = {
              ...this.mainItemsRecords[mainItemIndex],
              ...newRecord,
            };

            // Ensure the array itself updates its reference
            this.mainItemsRecords = [...this.mainItemsRecords];

            this.updateTotalValueAfterAction();
          }
          //this.cdr.detectChanges();
          console.log(this.mainItemsRecords);
        }, error: (err) => {
          console.log(err);
        },
        complete: () => {
        }
      });
      ///...................

      // this._ApiService.patch<MainItem>('executionordermain', record.executionOrderMainCode, filteredRecord).subscribe({
      //   next: (res) => {
      //     console.log('executionordermain  updated:', res);
      //     this.totalValue = 0;
      //     this.ngOnInit();
      //   }, error: (err) => {
      //     console.log(err);
      //     this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Invalid Data' });
      //   },
      //   complete: () => {
      //     this.updateSelectedServiceNumberRecord = undefined;
      //     this.messageService.add({ severity: 'success', summary: 'Success', detail: 'Record updated successfully ' });
      //     // this.ngOnInit()
      //   }

      // });
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

      //....................
      const bodyRequest: any = {
        quantity: updatedMainItem.totalQuantity,
        amountPerUnit: updatedMainItem.amountPerUnit
      };
      this._ApiService.post<any>(`/total`, bodyRequest).subscribe({
        next: (res) => {
          console.log('mainitem with total:', res);
          updatedMainItem.total = res.total;
          const mainItemIndex = this.mainItemsRecords.findIndex(item => item.executionOrderMainCode === index);
          if (mainItemIndex > -1) {
            console.log(updatedMainItem);

            this.mainItemsRecords[mainItemIndex] = { ...this.mainItemsRecords[mainItemIndex], ...updatedMainItem };
            // Ensure the array itself updates its reference
            this.mainItemsRecords = [...this.mainItemsRecords];

            this.updateTotalValueAfterAction();
          }
          // this.cdr.detectChanges();
          console.log(this.mainItemsRecords);
        }, error: (err) => {
          console.log(err);
        },
        complete: () => {
        }
      });
      ///...................

      // this._ApiService.patch<MainItem>('executionordermain', record.executionOrderMainCode, filteredRecord).subscribe({
      //   next: (res) => {
      //     console.log('executionordermain  updated:', res);
      //     this.totalValue = 0;
      //     this.ngOnInit()
      //   }, error: (err) => {
      //     console.log(err);
      //     this.messageService.add({ severity: 'error', summary: 'Error', detail: 'Invalid Data' });
      //   },
      //   complete: () => {
      //     this.messageService.add({ severity: 'success', summary: 'Success', detail: 'Record updated successfully ' });
      //   }
      // });
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

            this.mainItemsRecords = this.mainItemsRecords.filter(item => item.executionOrderMainCode !== record.executionOrderMainCode);
            this.updateTotalValueAfterAction();
            this.cdr.detectChanges();
            console.log(this.mainItemsRecords);
          }
          this.messageService.add({
            severity: 'success',
            summary: 'Successfully',
            detail: 'Deleted',
            life: 3000,
          });
          this.selectedExecutionOrder=[];
        }
      });
    }
  }


  // For Add new  Main Item
  newMainItem: MainItem = {
    Type: '',
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

  addMainItemInMemory() {

    if(this.newMainItem.description =="" && this.newMainItem.totalQuantity ===0 && this.newMainItem.amountPerUnit === 0){
      console.log("hereee");
      console.log(this.newMainItem.description,this.newMainItem.totalQuantity,this.newMainItem.amountPerUnit)
      this.messageService.add({
        severity: 'error',
        summary: 'Error',
        detail: 'Description & Quantity and AmountPerUnit are required.',
        life: 3000,
      });
    }
   else if (!this.selectedServiceNumberRecord) { // if user didn't select serviceNumber 

      const newRecord: MainItem = {
        unitOfMeasurementCode: this.selectedUnitOfMeasure,
        currencyCode: this.cloudCurrency,
        //this.selectedCurrency,
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
        Type: '',
        executionOrderMainCode: 0
      }
      // if (this.newMainItem.totalQuantity === 0 || this.newMainItem.description === "" || this.selectedCurrency === "") {
      //   // || this.newMainItem.unitOfMeasurementCode === ""  // till retrieved from cloud correctly
      //   this.messageService.add({
      //     severity: 'error',
      //     summary: 'Error',
      //     detail: 'Description & Quantity & Currency and UnitOfMeasurement are required',
      //     life: 3000
      //   });
      // }
      // else {
      console.log(newRecord);
      const filteredRecord = Object.fromEntries(
        Object.entries(newRecord).filter(([_, value]) => {
          return value !== '' && value !== 0 && value !== undefined && value !== null;
        })
      ) as MainItem;
      console.log(filteredRecord);
      ///,,,,,,,,,,,,,,,,,,,,,,,,
      const bodyRequest: any = {
        quantity: newRecord.totalQuantity,
        amountPerUnit: newRecord.amountPerUnit
      };

      this._ApiService.post<any>(`/total`, bodyRequest).subscribe({
        next: (res) => {
          console.log('mainitem with total:', res);
          // this.totalValue = 0;
          newRecord.total = res.total;

          const filteredRecord = Object.fromEntries(
            Object.entries(newRecord).filter(([_, value]) => {
              return value !== '' && value !== 0 && value !== undefined && value !== null;
            })
          ) as MainItem;
          console.log(filteredRecord);

          this._ExecutionOrderService.addMainItem(filteredRecord);

          this.savedInMemory = true;
          // this.cdr.detectChanges();

          const newMainItems = this._ExecutionOrderService.getMainItems();

          // Combine the current mainItemsRecords with the new list, ensuring no duplicates
          this.mainItemsRecords = [
            ...this.mainItemsRecords.filter(item => !newMainItems.some(newItem => newItem.executionOrderMainCode === item.executionOrderMainCode)), // Remove existing items
            ...newMainItems
          ];
          console.log(this.mainItemsRecords);

          this.updateTotalValueAfterAction();

          this.resetNewMainItem();



        }, error: (err) => {
          console.log(err);
        },
        complete: () => {
        }
      });
      //,,,,,,,,,,,,,,,,,,,,,,,,

      //https://trial.cfapps.us10-001.hana.ondemand.com/executionordermain?salesOrder=6&salesOrderItem=10&customerNumber=591001

      //   this._ApiService.post<MainItem>(`executionordermain?salesOrder=${this.documentNumber}&salesOrderItem=${this.itemNumber}&customerNumber=${this.customerId}`, filteredRecord).subscribe({
      //   next: (res) => {
      //     console.log('executionordermain created:', res);
      //     this.totalValue = 0;
      //     this.ngOnInit()
      //   }, error: (err) => {
      //     console.log(err);
      //   },
      //   complete: () => {
      //     this.resetNewMainItem();
      //     this.selectedUnitOfMeasure = "";
      //     this.selectedCurrency = "";
      //     this.selectedMaterialGroup = "";
      //     this.selectedServiceType = "";
      //     this.selectedLineType = "";

      //     this.messageService.add({ severity: 'success', summary: 'Success', detail: 'Record added successfully ' });
      //   }
      // });

    }

    else if (this.selectedServiceNumberRecord) { // if user select serviceNumber 
      const newRecord: MainItem = {
        serviceNumberCode: this.selectedServiceNumber,
        unitOfMeasurementCode: this.selectedServiceNumberRecord?.unitOfMeasurementCode,
        //this.selectedServiceNumberRecord?.baseUnitOfMeasurement,
        currencyCode:this.cloudCurrency,
        // this.selectedCurrency,

        description: this.selectedServiceNumberRecord?.description,
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
        Type: '',
        executionOrderMainCode: 0
      }
      // if (this.newMainItem.totalQuantity === 0 || this.selectedServiceNumberRecord.description === "" || this.selectedCurrency === "") {
      //   // || this.newMainItem.unitOfMeasurementCode === ""  // till retrieved from cloud correctly
      //   this.messageService.add({
      //     severity: 'error',
      //     summary: 'Error',
      //     detail: 'Description & Quantity & Currency and UnitOfMeasurement are required',
      //     life: 3000
      //   });
      // }
      // else {

      const filteredRecord = Object.fromEntries(
        Object.entries(newRecord).filter(([_, value]) => {
          return value !== '' && value !== 0 && value !== undefined && value !== null;
        })
      ) as MainItem;
      console.log(filteredRecord);

      ///,,,,,,,,,,,,,,,,,,,,,,,,
      const bodyRequest: any = {
        quantity: newRecord.totalQuantity,
        amountPerUnit: newRecord.amountPerUnit
      };

      this._ApiService.post<any>(`/total`, bodyRequest).subscribe({
        next: (res) => {
          console.log('mainitem with total:', res);
          // this.totalValue = 0;
          newRecord.total = res.total;

          const filteredRecord = Object.fromEntries(
            Object.entries(newRecord).filter(([_, value]) => {
              return value !== '' && value !== 0 && value !== undefined && value !== null;
            })
          ) as MainItem;
          console.log(filteredRecord);

          this._ExecutionOrderService.addMainItem(filteredRecord);

          this.savedInMemory = true;
          //this.cdr.detectChanges();

          const newMainItems = this._ExecutionOrderService.getMainItems();

          // Combine the current mainItemsRecords with the new list, ensuring no duplicates
          this.mainItemsRecords = [
            ...this.mainItemsRecords.filter(item => !newMainItems.some(newItem => newItem.executionOrderMainCode === item.executionOrderMainCode)), // Remove existing items
            ...newMainItems
          ];
          console.log(this.mainItemsRecords);

          this.updateTotalValueAfterAction();

          this.resetNewMainItem();
          this.selectedServiceNumberRecord = undefined;
        }, error: (err) => {
          console.log(err);
        },
        complete: () => {
        }
      });
      ///,,,,,,,,,,,,,,,,,,,,,,,,

      // this._ApiService.post<MainItem>(`executionordermain?salesOrder=${this.documentNumber}&salesOrderItem=${this.itemNumber}&customerNumber=${this.customerId}`, filteredRecord).subscribe({
      //   next: (res) => {
      //     console.log('executionordermain created:', res);
      //     this.totalValue = 0;
      //     this.ngOnInit()
      //   }, error: (err) => {
      //     console.log(err);
      //   },
      //   complete: () => {
      //     this.resetNewMainItem();
      //     this.selectedServiceNumberRecord = undefined;
      //     //this.selectedServiceNumber = 0;
      //     this.selectedLineType = "";
      //     this.selectedCurrency = ""
      //     this.messageService.add({ severity: 'success', summary: 'Success', detail: 'Record added successfully ' });
      //   }
      // });

    }
  }


  saveDocument() {
    // if (this.selectedMainItems.length) {
    this.confirmationService.confirm({
      message: 'Are you sure you want to save the document?',
      header: 'Confirm Saving ',
      accept: () => {

        const saveRequests = this.mainItemsRecords.map((item) => ({
          // executionOrderMain:item.executionOrderMain,
          // executionOrderMainCode:item.executionOrderMainCode,

          refrenceId: this.documentNumber,

          serviceNumberCode: item.serviceNumberCode,
          description: item.description,
          unitOfMeasurementCode: item.unitOfMeasurementCode,
          currencyCode: item.currencyCode,
          materialGroupCode: item.materialGroupCode,
          serviceTypeCode: item.serviceTypeCode,
          personnelNumberCode: item.personnelNumberCode,
          lineTypeCode: item.lineTypeCode,

          totalQuantity: item.totalQuantity,
          amountPerUnit: item.amountPerUnit,
          //total: item.total,


          // quantities:
          actualQuantity: item.actualQuantity,
          actualPercentage: item.actualPercentage,
          // remainingQuantity:item.remainingQuantity,


          overFulfillmentPercentage: item.overFulfillmentPercentage,
          unlimitedOverFulfillment: item.unlimitedOverFulfillment,
          manualPriceEntryAllowed: item.manualPriceEntryAllowed,
          externalServiceNumber: item.externalServiceNumber,
          serviceText: item.serviceText,
          lineText: item.lineText,
          lineNumber: item.lineNumber,

          biddersLine: item.biddersLine,
          supplementaryLine: item.supplementaryLine,
          lotCostOne: item.lotCostOne,
          doNotPrint: item.doNotPrint,

        }));
        // https://trial.cfapps.us10-001.hana.ondemand.com/executionordermain?salesOrder=6&salesOrderItem=10&customerNumber=591001

        // Set dynamic parameters for URL
        //localhost:8080/executionordermain?salesOrder=6&salesOrderItem=10&pricingProcedureStep=20&pricingProcedureCounter=1&customerNumber=12000000

        const url = `executionordermain?salesOrder=${this.documentNumber}&salesOrderItem=${this.itemNumber}&pricingProcedureStep=20&pricingProcedureCounter=1&customerNumber=${this.customerId}`;

        // Send the array of bodyRequest objects to the server in a single POST request
        this._ApiService.post<MainItem[]>(url, saveRequests).subscribe({
          next: (res) => {
            console.log('All main items saved successfully:', res);
            this.mainItemsRecords = res;
            const lastRecord = res[res.length - 1];
            console.log(this.mainItemsRecords);

            this.updateTotalValueAfterAction();
            // this.savedDBApp =true;
            // this.totalValue = 0;
            // this.totalValue = lastRecord.totalHeader ? lastRecord.totalHeader : 0;
            // this.ngOnInit();
            this.messageService.add({
              severity: 'success',
              summary: 'Success',
              detail: 'The Document has been saved successfully',
              life: 3000
            });

          }, error: (err) => {
            console.error('Error saving main items:', err);
            this.messageService.add({
              severity: 'error',
              summary: 'Error',
              detail: 'Error saving The Document',
              life: 3000
            });
          },
          complete: () => {
            // this.ngOnInit();
          }

        })
      }, reject: () => {

      }
    });
  }

  resetNewMainItem() {
    this.newMainItem = {
      Type: '',
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
    this.selectedServiceNumber = 0;
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
    // Toggle MainItem selection
    mainItem.selected = event.checked;

    // Update selectedMainItems array
    if (mainItem.selected) {
      this.selectedMainItems.push(mainItem);
    } else {
      this.selectedMainItems = this.selectedMainItems.filter(item => item !== mainItem);
    }



    // Debugging logs
    console.log('Selected MainItems:', this.selectedMainItems);
  }

  //end new selection...............

  // onMainItemSelection(event: any, mainItem: MainItem) {
  //   console.log('Event:', event);
  //   console.log('MainItem before change:', mainItem);

  //   console.log('Event checked state:', event.checked);
  //   console.log('MainItem selected before update:', mainItem.selected);

  //   mainItem.selected = event.checked;
  //   console.log('MainItem selected after update:', mainItem.selected);

  //   if (event.checked.length) {
  //     this.selectedMainItem = mainItem;
  //     console.log('Selected MainItem:', this.selectedMainItem);

  //     // Check if mainItem is already in the list to avoid duplication
  //     if (!this.selectedMainItems.includes(mainItem)) {
  //       this.selectedMainItems.push(mainItem);
  //       console.log('Selected Main Items after addition:', this.selectedMainItems);
  //     }

  //   } else {
  //     console.log('Entering else block');
  //     this.selectedMainItem = undefined;
  //     console.log('Selected MainItem after deselection:', this.selectedMainItem);

  //     const index = this.selectedMainItems.indexOf(mainItem);
  //     console.log('Index of mainItem in selected list:', index);

  //     if (index !== -1) {
  //       this.selectedMainItems.splice(index, 1);
  //       console.log('Selected Main Items after removal:', this.selectedMainItems);
  //     }
  //   }
  // }

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
