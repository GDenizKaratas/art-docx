import { Injectable } from '@angular/core';
import PizZip from 'pizzip';

export interface TemplatePlaceholder {
  key: string;
  type: 'text' | 'image' | 'table';
}

@Injectable({
  providedIn: 'root',
})
export class DocxTemplateService {
  async extractPlaceholders(file: File): Promise<TemplatePlaceholder[]> {
    const arrayBuffer = await file.arrayBuffer();
    const zip = new PizZip(arrayBuffer);

    const documentXml = zip.file('word/document.xml')?.asText();
    if (!documentXml) {
      throw new Error('Document XML not found inside the docx.');
    }

    console.log('documentXml', documentXml);

    const cleanText = this.extractPlainText(documentXml);
    console.log('cleanText', cleanText);

    const placeholders: TemplatePlaceholder[] = [];

    // Regex: {placeholder} ve {%image} gibi alanları bulur
    const regex = /{(.*?)}/g;
    let match;
    while ((match = regex.exec(cleanText)) !== null) {
      const rawKey = match[1].trim();
      let type: 'text' | 'image' | 'table' = 'text';

      if (rawKey.startsWith('%')) {
        type = 'image';
      } else if (rawKey.startsWith('#')) {
        type = 'table';
      }

      const key = rawKey.replace(/^[%#]/, ''); // % veya # baştaysa temizle
      placeholders.push({ key, type });
    }

    return placeholders;
  }

  private extractPlainText(xml: string): string {
    // Tüm <w:t>...</w:t> içeriklerini yakalayıp birleştiriyoruz
    const regex = /<w:t[^>]*>(.*?)<\/w:t>/g;
    let match;
    let result = '';

    while ((match = regex.exec(xml)) !== null) {
      result += match[1]; // Sadece içerik kısmı
    }

    return result;
  }
}
