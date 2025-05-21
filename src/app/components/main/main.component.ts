import { Component } from '@angular/core';
import { TemplatePlaceholder } from '../../services/docx-template.service';
import { CommonModule } from '@angular/common';
import { UploadTemplateComponent } from '../upload-template/upload-template.component';
import { TemplateDataComponent } from '../template-data/template-data.component';
import { GenerateDocxService } from '../../services/txt-done.service';
// import { GenerateDocxService } from '../../services/generate-docx.service';

@Component({
  selector: 'app-main',
  imports: [CommonModule, UploadTemplateComponent, TemplateDataComponent],
  templateUrl: './main.component.html',
  styleUrl: './main.component.scss',
})
export class MainComponent {
  placeholders: TemplatePlaceholder[] = [];
  jsonData: any = null;
  uploadedTemplateFile!: File;

  constructor(private generateDocxService: GenerateDocxService) {}

  onTemplateUploaded(event: {
    file: File;
    placeholders: TemplatePlaceholder[];
  }) {
    this.uploadedTemplateFile = event.file;
    this.placeholders = event.placeholders;
  }

  onJsonValid(json: any) {
    this.jsonData = json;
  }

  async onGenerateDocx() {
    if (!this.uploadedTemplateFile || !this.jsonData) {
      alert('Önce hem template hem de JSON verisi yüklenmelidir!');
      return;
    }

    // await this.generateDocxService.generateDocx(
    //   this.uploadedTemplateFile,
    //   this.jsonData
    // );

    // Header ve footer örneği
    const header = 'ABC Teknoloji | {invoice.number}';
    const footer = 'Sayfa {page} / {total} | {invoice.date}';

    // Dökümanı oluştur
    await this.generateDocxService.generateDocx(
      this.uploadedTemplateFile,
      this.jsonData,
      {
        fileName: `DENEME.docx`,
        header: header,
        footer: footer,
      }
    );

    console.log('Generate DOCX Çalıştı 🚀');
    console.log('Placeholders:', this.placeholders);
    console.log('Json Data:', this.jsonData);
  }
}
