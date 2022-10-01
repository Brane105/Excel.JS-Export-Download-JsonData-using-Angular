import { Component } from '@angular/core';
import { ExcelService } from './excel.service';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'excelJS_Project';
  // title = 'angular-export-to-excel';

  dataForExcel : any = [];

  empPerformance = [
    { ID: 10011, NAME: "Lucifer", DEPARTMENT: "Sales", MONTH: "Jan", YEAR: 2022, SALES: 132412, CHANGE: 12, LEADS: 35 },
    { ID: 10012, NAME: "Asmodeous", DEPARTMENT: "Sales", MONTH: "Feb", YEAR: 2022, SALES: 232324, CHANGE: 2, LEADS: 443 },
    { ID: 10013, NAME: "Bilial", DEPARTMENT: "Sales", MONTH: "Mar", YEAR: 2022, SALES: 542234, CHANGE: 45, LEADS: 345 },
    { ID: 10014, NAME: "Lilith", DEPARTMENT: "Sales", MONTH: "Apr", YEAR: 2022, SALES: 223335, CHANGE: 32, LEADS: 234 },
    { ID: 10015, NAME: "Belphegor", DEPARTMENT: "Sales", MONTH: "May", YEAR: 2022, SALES: 455535, CHANGE: 21, LEADS: 12 },
  ];

  constructor(public ete: ExcelService) { }

  exportToExcel() {
    
    this.empPerformance.map(element => {
      this.dataForExcel.push(Object.values(element))
    })

    let reportData = {
      title: 'Employee Sales Report - Jan 2020',
      data: this.dataForExcel,
      headers: Object.keys(this.empPerformance[0])
    }
    this.ete.exportExcel(reportData);
  }
//   json_data=[{
// 		"name": "Raja",
// 		"age": 20
// 	},
// 	{
// 		"name": "Mano",
// 		"age": 40
// 	},
// 	{
// 		"name": "Tom",
// 		"age": 40
// 	},
// 	{
// 		"name": "Devi",
// 		"age": 40
// 	},
// 	{
// 		"name": "Mango",
// 		"age": 40
// 	}
// ]
// downloadExcel(){
//     //create new excel work book
//   let workbook = new Workbook();
//   //add name to sheet
  
// let worksheet = workbook.addWorksheet("Employee Data");
//   //add column name
// let header=["Name","Age"]
// let headerRow = worksheet.addRow(header);
// // for (let x1 of this.json_data)
// // {
// // let x2=Object.keys(x1);
// // let temp=[]
// // for(let y of x2)
// // {
// //   temp.push(x1[y])
// // }
// // worksheet.addRow(temp)
// // }
// //set downloadable file name
// let fname="Emp Data Sep 2020"

// //add data and file name and download
// workbook.xlsx.writeBuffer().then((data) => {
// let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
// fs.saveAs(blob, fname+'-'+new Date().valueOf()+'.xlsx');
// });
//   }
}
