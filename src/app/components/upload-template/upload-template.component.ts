import { Component, EventEmitter, Output } from '@angular/core';
import {
  DocxTemplateService,
  TemplatePlaceholder,
} from '../../services/docx-template.service';
import { CommonModule } from '@angular/common';

@Component({
  selector: 'app-upload-template',
  imports: [CommonModule],
  templateUrl: './upload-template.component.html',
  styleUrl: './upload-template.component.scss',
})
export class UploadTemplateComponent {
  placeholders: TemplatePlaceholder[] = [];
  fileName: string = '';

  @Output() templateUploaded = new EventEmitter<{
    file: File;
    placeholders: TemplatePlaceholder[];
  }>();

  constructor(private docxTemplateService: DocxTemplateService) {}

  async onFileSelected(event: Event) {
    const input = event.target as HTMLInputElement;
    if (!input.files || input.files.length === 0) {
      return;
    }

    const file = input.files[0];
    this.fileName = file.name;

    try {
      const placeholders = await this.docxTemplateService.extractPlaceholders(
        file
      );

      this.placeholders = placeholders;
      console.log('placeholders', placeholders);
      this.templateUploaded.emit({ file, placeholders });
    } catch (error) {
      console.error('Placeholder extraction failed:', error);
    }
  }
}
