export class MainItemSalesQuotation {

    Type:string='';

    isPersisted: boolean=false;

    temporaryDeletion?:string;
    referenceId?:number;

    invoiceMainItemCode: number=0;
    serviceNumberCode?: number;
    description?: string;
    quantity?: number;
    unitOfMeasurementCode?:string;
    formulaCode?:string;
    amountPerUnit?: number;
    currencyCode?: string;
    total: number=0;
    profitMargin?: number;
    totalWithProfit: number=0;
    selected?: boolean;
    // subItems?:SubItem[];
   // subItems: SubItem[] = [];

    totalHeader?:number;
    
    doNotPrint?:boolean;
    amountPerUnitWithProfit?: number;

    // extra:
    lineNumber?:string;
    actualQuantity?: number;
    actualPercentage?: number;
    overFulfillmentPercentage?: number;
    unlimitedOverFulfillment?: boolean;
    manualPriceEntryAllowed?: boolean;
    externalServiceNumber?: string;
    serviceText?:string;
    lineText?: string;
    biddersLine?:boolean;
    supplementaryLine?:boolean;
    lotCostOne?:boolean;
    personnelNumberCode?: string;
    materialGroupCode?:string;
    serviceTypeCode?:string;
    lineTypeCode?:string;
    
}