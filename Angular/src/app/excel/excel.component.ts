import { Component, OnInit, Output, EventEmitter } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { IExcelfields } from './model/excelfields';

@Component({
  selector: 'app-excel',
  templateUrl: './excel.component.html',
  styleUrls: ['./excel.component.scss']
})
export class ExcelComponent implements OnInit {
  
  excelfields: IExcelfields[] = [];
  gridApi:any;
  columnsApi:any;
  constructor(private http: HttpClient) { }
  
  columnDefs = [
    {headerName: 'Sno', field: 'Sno' ,editable:true,width:100},
    {headerName: 'Company', field: 'Company',editable:true, },
    {headerName: 'NoOfEmployee', field: 'NoOfEmployee',editable:true, },
    {headerName: 'Region', field: 'Region',editable:true, },
    {headerName: 'Sector', field: 'Sector' ,editable:true,},
    {headerName: 'SubSector', field: 'SubSector',editable:true, },
    {headerName: 'TotalRevenu', field: 'TotalRevenu',editable:true, },
    {Websites: 'TotalRevenu', field: 'Websites',editable:true, },

];

  ngOnInit() {
  }
  
  OnGridReady(params){
    this.gridApi=params.api;
    this.columnsApi=params.columnApi;
  }
  download(){
    var params={};
    this.gridApi.exportDataAsCsv(params);

  }
  
  public uploadFile = (files) => {
    if (files.length === 0) {
      return;
    }
 
    let fileToUpload = <File>files[0];
    const formData = new FormData(); 
    formData.append('file', fileToUpload, fileToUpload.name);
   
 
    this.http.post<any>('http://localhost:52590/api/Exceloperation/ExcelUpload', formData)
    .subscribe(data => {
      this.excelfields=data;
      this.excelfields.shift()
      });
  }   
}
