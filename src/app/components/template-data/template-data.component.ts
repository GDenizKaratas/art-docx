import { CommonModule } from '@angular/common';
import { Component, EventEmitter, Output } from '@angular/core';
import { FormsModule } from '@angular/forms';
@Component({
  selector: 'app-template-data',
  imports: [CommonModule, FormsModule],
  templateUrl: './template-data.component.html',
  styleUrl: './template-data.component.scss',
})
export class TemplateDataComponent {
  exampleData = {
    customer: {
      name: 'Ahmet Yılmaz',
      company: 'ABC Teknoloji Ltd.',
      address: 'Atatürk Bulvarı No:123, Ankara',
      email: 'ahmet@example.com',
      phone: '+90 532 123 4567',
    },
    fatura: {
      number: 'INV-2025-0042',
      date: '06/05/2025',
      dueDate: '20/05/2025',
      total: '₺5.600,00',
      tax: '₺1.008,00',
      status: 'Ödendi',
      // status: 'Ödenmedi',
    },
    items: [
      { Ürün: 'Laptop', Adet: 1, Fiyat: '₺4.500,00' },
      { Ürün: 'Mouse', Adet: 2, Fiyat: '₺300,00' },
      { Ürün: 'Klavye', Adet: 1, Fiyat: '₺800,00' },
    ],
    items2: [
      { Ürün: 'Masa', Adet: 9, Fiyat: '₺4.500,00' },
      { Ürün: 'Sandalye', Adet: 20, Fiyat: '₺300,00' },
      { Ürün: 'Kamış', Adet: 11, Fiyat: '₺700,00' },
      { Ürün: 'Masa', Adet: 9, Fiyat: '₺4.500,00' },
      { Ürün: 'Sandalye', Adet: 20, Fiyat: '₺300,00' },
      { Ürün: 'Kamış', Adet: 11, Fiyat: '₺700,00' },
    ],
    showTable2: 'Göster',
    logo: 'https://cdn.pixabay.com/photo/2017/09/01/00/15/png-2702691_960_720.png',
  };
  jsonInput: string = JSON.stringify(this.exampleData);

  errorMessage: string = '';

  @Output() jsonValid = new EventEmitter<any>(); // JSON'u dışarıya veriyoruz

  onInputChange() {
    try {
      const parsed = JSON.parse(this.jsonInput);
      this.errorMessage = '';
      this.jsonValid.emit(parsed);
    } catch (err) {
      this.errorMessage = 'Geçersiz JSON formatı!';
    }
  }
}
