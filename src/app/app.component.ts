import { Component } from '@angular/core';
import { Workbook } from 'exceljs/dist/exceljs.min.js';
import * as fs from 'file-saver';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {

  title = 'angular-exceljs-example';
  dataForExcel = [];
  logs = [
    {
        "time": "Wed, 29 Sep 2021 13:35:30 GMT",
        "log": "Command Transcripts 0: Scorpion  Result Confidence :0.9299548864364624"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:30 GMT",
        "log": "Command Transcripts 1: skorpion  Result Confidence :0.04444444179534912"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:30 GMT",
        "log": "Command Transcripts 2: scorpian  Result Confidence :0.04444444179534912"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:30 GMT",
        "log": "Command Transcripts 3: scope in  Result Confidence :0.04444444179534912"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:30 GMT",
        "log": "Command Transcripts 4: scorpene  Result Confidence :0.04444444179534912"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:30 GMT",
        "log": "Recognized command: scope in at index: 0 for input: scorpion. Annotation Dictionary hits: 0"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:44 GMT",
        "log": "Command Transcripts 0: take a pic  Result Confidence :0.9299548268318176"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:44 GMT",
        "log": "Command Transcripts 1: take a pick  Result Confidence :0.6347847580909729"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:44 GMT",
        "log": "Command Transcripts 2: take a peek  Result Confidence :0.6347847580909729"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:44 GMT",
        "log": "Command Transcripts 3: Tay K pic  Result Confidence :0.339614599943161"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:44 GMT",
        "log": "Command Transcripts 4: teke pic  Result Confidence :0.48719966411590576"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:44 GMT",
        "log": "Sending to dialogue flow: snapshot"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:44 GMT",
        "log": "Recognized command: snapshot at index: 0 for input: take a pic. Annotation Dictionary hits: 0"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:50 GMT",
        "log": "Command Transcripts 0: retroflex  Result Confidence :0.39838969707489014"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:50 GMT",
        "log": "Command Transcripts 1: red reflex  Result Confidence :0.17512714862823486"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:50 GMT",
        "log": "Command Transcripts 2: metroplex  Result Confidence :0.04444444179534912"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:50 GMT",
        "log": "Command Transcripts 3: retro flex  Result Confidence :0.04444444179534912"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:50 GMT",
        "log": "Command Transcripts 4: metroflex  Result Confidence :0.04444444179534912"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:50 GMT",
        "log": "Sending to dialogue flow: snapshot  retroflex"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:57 GMT",
        "log": "Command Transcripts 0: scope out  Result Confidence :0.9299548864364624"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:57 GMT",
        "log": "Command Transcripts 1: scout  Result Confidence :0.04444444179534912"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:57 GMT",
        "log": "Command Transcripts 2: scope aut  Result Confidence :0.48719966411590576"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:57 GMT",
        "log": "Command Transcripts 3: skout  Result Confidence :0.04444444179534912"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:57 GMT",
        "log": "Command Transcripts 4: scoreboard  Result Confidence :0.04444444179534912"
    },
    {
        "time": "Wed, 29 Sep 2021 13:35:57 GMT",
        "log": "Recognized command: scope out at index: 0 for input: scope out. Annotation Dictionary hits: 0"
    }
];

  exportToExcel() {

    this.logs = this.logs.filter(x => x.log.includes("Command Transcripts") || x.log.includes("Recognized command") || x.log.includes("Sending to dialogue flow"));

    this.logs.forEach((row: any, index: number) => {
      let data = row["log"].split(":");
      let cmd = data[0];
      let confidence = data[2];
      let textdata = data[1].split("Result");
      let text = textdata[0].trim();
      let hit = 0;
      
      if(row["log"].includes("Recognized command")){
        let hitData = row["log"].split(":");
        let hitText = hitData[3].split(".")[0].trim();
        let itemToUpdate = this.dataForExcel.find(item => item[1].toLowerCase() == hitText);
        if(itemToUpdate){
          itemToUpdate[3] = 1;
        }
      }
      else if(row["log"].includes("Sending to dialogue flow")){
        let hitData = row["log"].split(" ");
        let hitText = hitData[hitData.length-1].trim()
        let itemToUpdate = this.dataForExcel.find(item => item[1].toLowerCase() == hitText);
        if(itemToUpdate){
          itemToUpdate[3] = 1;
        }
      }
      else{
        let excelRowData = {"Command Transcripts":cmd, "Text":text, "Confidence":confidence, "Hit":hit};
        this.dataForExcel.push(Object.values(excelRowData));
      }
    })

    let reportData = {
      title: 'Logs Report',
      data: this.dataForExcel,
      headers: ['Command Transcripts', 'Text', 'Confidence', 'Hit']
    }

    //Title, Header & Data
    const title = reportData.title;
    const header = reportData.headers
    const data = reportData.data;

    //Create a workbook with a worksheet
    let workbook = new Workbook();
    let worksheet = workbook.addWorksheet('Logs Data');
    
    //Blank Row 
    // worksheet.addRow([]);

    //Adding Header Row
    let headerRow = worksheet.addRow(header);
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '4167B8' },
        bgColor: { argb: '' }
      }
      cell.font = {
        bold: true,
        color: { argb: 'FFFFFF' },
        size: 12
      }
    })

    // Adding Data
    data.forEach(d => {
      let row = worksheet.addRow(d);
    }
    );

    worksheet.getColumn(3).width = 20;
    worksheet.addRow([]);

   

    //Generate & Save Excel File
    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      fs.saveAs(blob, title + '.xlsx');
    })
    
  }
}
