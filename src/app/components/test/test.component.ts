// Örnek Kullanım
import { Component } from '@angular/core';
import { GenerateDocxService } from '../../services/txt-done.service';

@Component({
  selector: 'app-document-generator',
  templateUrl: './test.component.html',
})
export class TestComponent {
  constructor(private docxService: GenerateDocxService) {}
  selectedFile!: File;

  onFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) {
      this.selectedFile = input.files[0];
    }
  }

  async generateDocument(templateFile: File): Promise<void> {
    // Örnek veri
    const exampleFoto =
      'https://cdn.pixabay.com/photo/2017/09/01/00/15/png-2702691_960_720.png';
    const data = {
      customer: {
        name: 'Ahmet Yılmaz',
        company: 'Yılmaz İnşaat',
      },
      invoice: {
        number: '2023-001',
        date: '2023-10-25',
        status: 'Ödenmedi',
        dueDate: '2023-11-10',
        subtotal: 1500.0,
        tax: 300.0,
        total: 1800.0,
      },
      items: [
        {
          description: 'İnşaat Malzemeleri',
          quantity: 10,
          unitPrice: 100.0,
          total: 1000.0,
        },
      ],
      logo: 'https://cdn.pixabay.com/photo/2017/09/01/00/15/png-2702691_960_720.png',
      'logo.width': 200,
      'logo.height': 150,
      showDiscount: true,
      discountAmount: 200.0,
      notes: [
        {
          id: 1,
          content: 'Ödeme 10 gün içinde yapılmalıdır.',
        },
      ],
    };

    // Header ve footer örneği
    const header = 'ABC Teknoloji | {invoice.number}';
    const footer = 'Sayfa {page} / {total} | {invoice.date}';

    // Dökümanı oluştur
    await this.docxService.generateDocx(templateFile, data, {
      fileName: `Fatura_${data.invoice.number}.docx`,
      header: header,
      footer: footer,
    });
  }
}
