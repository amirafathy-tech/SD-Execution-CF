import { Injectable } from '@angular/core';
import { Observable, Subject } from 'rxjs';
import { ApiService } from '../shared/ApiService.service';
import { MainItem } from './execution-order.model';
    
@Injectable()
export class ExecutionOrderService {

    recordsChanged = new Subject<MainItem[]>();
    startedEditing = new Subject<number>();
    constructor(private apiService: ApiService) { }
    private recordsApi!: MainItem[]
  
    getRecords() {
      this.apiService.get<MainItem[]>('executionordermain').subscribe(response => {
        console.log(response);
        this.recordsApi = response;
        this.recordsChanged.next(this.recordsApi);
      });
    }
  
    getRecord(index: number): Observable<MainItem> {
      return this.apiService.getID<MainItem>('executionordermain', index);
    }
   
    addRecord(record: MainItem) {
      this.apiService.post<MainItem>('executionordermain', record).subscribe((response: MainItem) => {
        console.log('mainItem  created:', response);
        this.getRecords();
        return response
      });
    }
  
    updateRecord(index: number, newRecord: MainItem) {
      this.apiService.put<MainItem>('executionordermain', index, newRecord).subscribe(response => {
        console.log('mainitem updated:',response);
        this.getRecords()
      });
    }
  
    deleteRecord(index: any) {
      this.apiService.delete<MainItem>('executionordermain', index).subscribe(response => {
        console.log('mainitem deleted:',response);
        this.getRecords()
      });
    }
};