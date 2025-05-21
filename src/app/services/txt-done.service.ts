import { Injectable } from '@angular/core';
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';

@Injectable({ providedIn: 'root' })
export class GenerateDocxService {
  private imageCounter = 1;
  private contentTypesXml: Document | null = null;
  private tableCounter = 1;

  async generateDocx(
    templateFile: File,
    data: any,
    options: {
      fileName?: string;
      header?: string;
      footer?: string;
    } = {}
  ): Promise<void> {
    try {
      const zip = await this.loadZip(templateFile);
      this.imageCounter = 1; // Reset counter for each document

      // Load content types XML
      const contentTypesContent = zip.file('[Content_Types].xml')?.asText();
      if (contentTypesContent) {
        const parser = new DOMParser();
        this.contentTypesXml = parser.parseFromString(
          contentTypesContent,
          'application/xml'
        );
      }

      // Process header if provided
      if (options.header) {
        await this.updateHeaderFooter(zip, 'header', options.header, data);
      }

      // Process footer if provided
      if (options.footer) {
        await this.updateHeaderFooter(zip, 'footer', options.footer, data);
      }

      const documentXml = zip.file('word/document.xml')?.asText() || '';
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(documentXml, 'application/xml');

      // Process section directives first
      this.processSectionDirectives(xmlDoc, data);

      const paragraphs = Array.from(xmlDoc.getElementsByTagName('w:p'));
      for (const paragraph of paragraphs) {
        await this.processParagraph(zip, xmlDoc, paragraph, data);
      }

      // Process table placeholders
      const tables = Array.from(xmlDoc.getElementsByTagName('w:tbl'));
      for (const table of tables) {
        await this.processTable(xmlDoc, table, data);
      }

      // Update content types if needed
      if (this.contentTypesXml) {
        const serializer = new XMLSerializer();
        zip.file(
          '[Content_Types].xml',
          serializer.serializeToString(this.contentTypesXml)
        );
      }

      const serializer = new XMLSerializer();
      zip.file('word/document.xml', serializer.serializeToString(xmlDoc));
      this.saveDocument(zip, options.fileName || 'generated-document.docx');
      console.log('Document generation completed successfully');
    } catch (error) {
      console.error('Error generating document:', error);
    }
  }

  /**
   * Process special section directives like page breaks
   */
  private processSectionDirectives(doc: Document, data: any): void {
    const body = doc.getElementsByTagName('w:body')[0];
    if (!body) return;

    const paragraphs = Array.from(doc.getElementsByTagName('w:p'));

    // Find and process page break directives: {%pagebreak}
    for (let i = 0; i < paragraphs.length; i++) {
      const paragraph = paragraphs[i];
      const text = this.getParagraphText(paragraph);

      if (text.trim() === '{%pagebreak}') {
        // Replace with actual page break
        const pageBreak = this.createPageBreak(doc);
        if (paragraph.parentNode) {
          paragraph.parentNode.replaceChild(pageBreak, paragraph);
        }
      }
    }
  }

  private getParagraphText(paragraph: Element): string {
    let text = '';
    const runs = Array.from(paragraph.getElementsByTagName('w:r'));

    runs.forEach((run) => {
      const textElements = run.getElementsByTagName('w:t');
      Array.from(textElements).forEach((t) => {
        text += t.textContent || '';
      });
    });

    return text;
  }

  private async loadZip(file: File): Promise<PizZip> {
    try {
      const arrayBuffer = await file.arrayBuffer();
      return new PizZip(arrayBuffer);
    } catch (error) {
      console.error('Error loading zip:', error);
      throw error;
    }
  }

  private async processParagraph(
    zip: PizZip,
    doc: Document,
    paragraph: Element,
    data: any
  ) {
    const runs = Array.from(paragraph.getElementsByTagName('w:r'));
    let fullText = '';
    let textElements: Element[] = [];

    runs.forEach((run) => {
      const texts = run.getElementsByTagName('w:t');
      Array.from(texts).forEach((t) => {
        fullText += t.textContent || '';
        textElements.push(t);
      });
    });

    // Check for conditional content: {#if condition}content{/if}
    const conditionalRegex = /\{#if\s+([^}]+)\}([\s\S]*?)\{\/if\}/g;
    let conditionalMatch;
    let hasConditional = false;

    while ((conditionalMatch = conditionalRegex.exec(fullText)) !== null) {
      hasConditional = true;
      const condition = conditionalMatch[1].trim();
      const content = conditionalMatch[2];

      // Evaluate condition
      const conditionResult = this.evaluateCondition(condition, data);

      if (!conditionResult) {
        // Remove the paragraph if condition is false
        if (paragraph.parentNode) {
          paragraph.parentNode.removeChild(paragraph);
          return;
        }
      } else {
        // Replace the conditional syntax with just the content
        fullText = fullText.replace(conditionalMatch[0], content);
      }
    }

    if (hasConditional) {
      // Update the paragraph with processed conditional content
      if (textElements.length > 0) {
        textElements[0].textContent = fullText;
        for (let i = 1; i < textElements.length; i++) {
          textElements[i].textContent = '';
        }
      }
    }

    // Check for table placeholder: {%table:tableKey}
    const tableRegex = /\{%\s*table:(\w+)\s*\}/;
    const tableMatch = fullText.match(tableRegex);

    if (tableMatch) {
      const tableKey = tableMatch[1];
      const tableData = this.getDeepValue(tableKey, data);

      if (tableData && Array.isArray(tableData) && tableData.length > 0) {
        try {
          // Create table
          const table = this.generateTable(doc, tableData);

          // Replace the paragraph with the table
          if (paragraph.parentNode) {
            paragraph.parentNode.replaceChild(table, paragraph);
          }

          console.log(`Table added for key: ${tableKey}`);
        } catch (error) {
          console.error(`Error processing table ${tableKey}:`, error);
        }
      }
      return;
    }

    // Check for image placeholder: {%imageKey}
    const imageRegex = /\{%\s*(\w+)(\.(\w+))?\s*\}/;
    const imageMatch = fullText.match(imageRegex);

    if (imageMatch) {
      const imageKey = imageMatch[1];
      const property = imageMatch[3] || null; // Optional property like width, height
      const imageUrl = this.getDeepValue(imageKey, data);

      console.log('imageMatch', imageMatch);
      alert(1);
      alert(property);
      if (imageUrl) {
        try {
          // Extract image dimensions if specified
          alert(parseInt(this.getDeepValue(`${imageKey}.width`, data), 10));
          const width =
            property === 'width'
              ? parseInt(this.getDeepValue(`${imageKey}.width`, data), 10) ||
                200
              : this.getDeepValue(`${imageKey}.width`, data) || 200;

          const height =
            property === 'height'
              ? parseInt(this.getDeepValue(`${imageKey}.height`, data), 10) ||
                150
              : this.getDeepValue(`${imageKey}.height`, data) || 150;

          const extension =
            imageUrl.split('.').pop()?.split('?')[0]?.toLowerCase() || 'png';
          const imageBuffer = await this.fetchImageBuffer(imageUrl);
          const imageId = `rId${1000 + this.imageCounter}`; // Use a high starting number to avoid conflicts
          const filename = `image${this.imageCounter}.${extension}`;
          this.imageCounter++;

          // Ensure media folder exists
          if (!zip.folder('word/media')) {
            zip.folder('word/media');
          }

          // Add image to media folder
          zip.folder('word/media')?.file(filename, imageBuffer);

          // Add image relationship
          this.addImageRelation(zip, imageId, filename);

          // Update content types if needed
          this.ensureContentType(extension);

          // Create image paragraph
          const imageParagraph = this.generateImageParagraph(
            doc,
            imageId,
            width,
            height
          );

          // Replace the original paragraph with the image paragraph
          paragraph.parentNode?.replaceChild(imageParagraph, paragraph);

          console.log(`Image added: ${filename}, ID: ${imageId}`);
        } catch (error) {
          console.error(`Error processing image ${imageKey}:`, error);
        }
      }
      return;
    }

    // Process formatting commands: {format:key:style}
    const formatRegex = /\{format:(\w+):([\w-]+)\}/g;
    let formatMatch;
    let formattedText = fullText;

    while ((formatMatch = formatRegex.exec(fullText)) !== null) {
      const formatKey = formatMatch[1];
      const style = formatMatch[2];
      const value = this.getDeepValue(formatKey, data);

      // Apply formatting based on style
      const formattedValue = this.applyFormatting(value, style, doc);
      formattedText = formattedText.replace(formatMatch[0], formattedValue);
    }

    if (formattedText !== fullText) {
      // Update with formatted content
      if (textElements.length > 0) {
        textElements[0].textContent = formattedText;
        for (let i = 1; i < textElements.length; i++) {
          textElements[i].textContent = '';
        }
      }
      return;
    }

    // Process loops: {#each items as item}{/each}
    const loopRegex = /\{#each\s+(\w+)\s+as\s+(\w+)\}([\s\S]*?)\{\/each\}/;
    const loopMatch = fullText.match(loopRegex);

    if (loopMatch) {
      const arrayKey = loopMatch[1];
      const itemName = loopMatch[2];
      const template = loopMatch[3];
      const array = this.getDeepValue(arrayKey, data);

      if (array && Array.isArray(array) && array.length > 0) {
        let result = '';

        for (const item of array) {
          let itemContent = template;

          // Replace all occurrences of {item.property}
          const itemRegex = new RegExp(`{${itemName}\\.(\\w+)}`, 'g');
          let itemMatch;

          while ((itemMatch = itemRegex.exec(template)) !== null) {
            const prop = itemMatch[1];
            itemContent = itemContent.replace(itemMatch[0], item[prop] || '');
          }

          result += itemContent;
        }

        if (textElements.length > 0) {
          textElements[0].textContent = result;
          for (let i = 1; i < textElements.length; i++) {
            textElements[i].textContent = '';
          }
        }
        return;
      }
    }

    // Process text placeholders: {key}
    const placeholders = this.findPlaceholders(fullText);
    if (placeholders.length === 0) return;

    textElements.forEach((t) => (t.textContent = ''));
    let result = '';
    let lastIndex = 0;

    placeholders.forEach((ph) => {
      result += fullText.substring(lastIndex, ph.start);
      result += this.getDeepValue(ph.key, data);
      lastIndex = ph.end;
    });

    result += fullText.substring(lastIndex);
    if (textElements.length > 0) {
      textElements[0].textContent = result;
    }
  }

  private applyFormatting(value: any, style: string, doc: Document): string {
    // Return the value as is for now - formatting will be applied at run level later
    return value?.toString() || '';
  }

  private evaluateCondition(condition: string, data: any): boolean {
    try {
      // Parse simple conditions like "user.age > 18"
      const parts = condition.split(/\s*(===|==|!=|!==|>=|<=|>|<)\s*/);

      if (parts.length === 3) {
        const left = this.getDeepValue(parts[0], data);
        const operator = parts[1];
        const right = isNaN(Number(parts[2]))
          ? this.getDeepValue(parts[2], data) || parts[2].replace(/["']/g, '')
          : Number(parts[2]);

        switch (operator) {
          case '===':
          case '==':
            return left == right;
          case '!=':
          case '!==':
            return left != right;
          case '>':
            return left > right;
          case '<':
            return left < right;
          case '>=':
            return left >= right;
          case '<=':
            return left <= right;
          default:
            return Boolean(left);
        }
      } else {
        // Just check if the value exists or is truthy
        return Boolean(this.getDeepValue(condition, data));
      }
    } catch (error) {
      console.error('Error evaluating condition:', error);
      return false;
    }
  }

  private async processTable(
    doc: Document,
    table: Element,
    data: any
  ): Promise<void> {
    try {
      // Get all rows in the table
      const rows = Array.from(table.getElementsByTagName('w:tr'));
      if (rows.length === 0) return;

      // Check if the table has a conditional format directive
      const firstRow = rows[0];
      const firstCellText = this.getCellText(
        firstRow.getElementsByTagName('w:tc')[0]
      );

      const conditionalFormatRegex = /\{format-table:(\w+)?\}/;
      const conditionalMatch = firstCellText.match(conditionalFormatRegex);

      if (conditionalMatch) {
        const dataKey = conditionalMatch[1];
        const tableData = dataKey ? this.getDeepValue(dataKey, data) : null;

        if (tableData && Array.isArray(tableData)) {
          // Remove the directive row
          if (conditionalMatch[0] === firstCellText.trim()) {
            table.removeChild(firstRow);
          }

          // Apply table formatting based on data
          this.applyTableFormatting(table, tableData, data);
        }
      }

      // Process cells for text replacements
      const cells = Array.from(table.getElementsByTagName('w:tc'));
      for (const cell of cells) {
        const paragraphs = Array.from(cell.getElementsByTagName('w:p'));
        for (const paragraph of paragraphs) {
          // Process paragraphs within cells
          await this.processParagraph(
            null as unknown as PizZip,
            doc,
            paragraph,
            data
          );
        }
      }
    } catch (error) {
      console.error('Error processing table:', error);
    }
  }

  private getCellText(cell: Element): string {
    let text = '';
    const paragraphs = Array.from(cell.getElementsByTagName('w:p'));

    paragraphs.forEach((paragraph) => {
      const runs = Array.from(paragraph.getElementsByTagName('w:r'));
      runs.forEach((run) => {
        const textElements = run.getElementsByTagName('w:t');
        Array.from(textElements).forEach((t) => {
          text += t.textContent || '';
        });
      });
    });

    return text;
  }

  private applyTableFormatting(
    table: Element,
    tableData: any[],
    data: any
  ): void {
    // Apply conditional formatting to rows based on data values
    const rows = Array.from(table.getElementsByTagName('w:tr'));

    rows.forEach((row, index) => {
      if (index === 0) return; // Skip header row

      if (index <= tableData.length) {
        const rowData = tableData[index - 1];

        // Check if the row has a conditional formatting rule
        const cellText = this.getCellText(row.getElementsByTagName('w:tc')[0]);
        const condFormatRegex =
          /\{if:(\w+)(==|!=|>|<|>=|<=)([^}]+)\}(\w+)\{\/if\}/;
        const match = cellText.match(condFormatRegex);

        if (match) {
          const field = match[1];
          const operator = match[2];
          const value = match[3].trim();
          const style = match[4];

          const fieldValue = rowData[field];
          let condition = false;

          switch (operator) {
            case '==':
              condition = fieldValue == value;
              break;
            case '!=':
              condition = fieldValue != value;
              break;
            case '>':
              condition = fieldValue > parseFloat(value);
              break;
            case '<':
              condition = fieldValue < parseFloat(value);
              break;
            case '>=':
              condition = fieldValue >= parseFloat(value);
              break;
            case '<=':
              condition = fieldValue <= parseFloat(value);
              break;
          }

          if (condition) {
            // Apply the style to the row
            this.applyRowStyle(row, style);
          }
        }
      }
    });
  }

  private applyRowStyle(row: Element, style: string): void {
    const cells = Array.from(row.getElementsByTagName('w:tc'));

    cells.forEach((cell) => {
      const tcPr = this.getOrCreateChildElement(cell, 'w:tcPr');

      // Apply style based on name
      switch (style) {
        case 'highlight':
          // Add yellow highlighting
          const shd = this.createElementWithNS(cell.ownerDocument, 'w:shd');
          shd.setAttribute('w:val', 'clear');
          shd.setAttribute('w:color', 'auto');
          shd.setAttribute('w:fill', 'FFFF00'); // Yellow
          tcPr.appendChild(shd);
          break;

        case 'bold':
          // Make text bold
          const paragraphs = Array.from(cell.getElementsByTagName('w:p'));
          paragraphs.forEach((p) => {
            const runs = Array.from(p.getElementsByTagName('w:r'));
            runs.forEach((r) => {
              const rPr = this.getOrCreateChildElement(r, 'w:rPr');
              const b = this.createElementWithNS(r.ownerDocument, 'w:b');
              rPr.appendChild(b);
            });
          });
          break;

        case 'red':
          // Add red background
          const shdRed = this.createElementWithNS(cell.ownerDocument, 'w:shd');
          shdRed.setAttribute('w:val', 'clear');
          shdRed.setAttribute('w:color', 'auto');
          shdRed.setAttribute('w:fill', 'FF0000'); // Red
          tcPr.appendChild(shdRed);
          break;

        case 'green':
          // Add green background
          const shdGreen = this.createElementWithNS(
            cell.ownerDocument,
            'w:shd'
          );
          shdGreen.setAttribute('w:val', 'clear');
          shdGreen.setAttribute('w:color', 'auto');
          shdGreen.setAttribute('w:fill', '00FF00'); // Green
          tcPr.appendChild(shdGreen);
          break;
      }
    });
  }

  private getOrCreateChildElement(parent: Element, tagName: string): Element {
    let element = parent.getElementsByTagName(tagName)[0];
    if (!element) {
      element = this.createElementWithNS(parent.ownerDocument, tagName);
      parent.insertBefore(element, parent.firstChild);
    }
    return element;
  }

  private createElementWithNS(doc: Document, tagName: string): Element {
    return doc.createElementNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      tagName
    );
  }

  private generateTable(doc: Document, data: any[]): Element {
    if (!data || !data.length) {
      throw new Error('Table data is empty or invalid');
    }

    // Get column headers (keys from the first object)
    const columns = Object.keys(data[0]);

    // Create table element
    const table = doc.createElementNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'w:tbl'
    );

    // Add table properties
    const tblPr = doc.createElementNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'w:tblPr'
    );

    // Add table style
    const tblStyle = doc.createElementNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'w:tblStyle'
    );
    tblStyle.setAttribute('w:val', 'TableGrid');
    tblPr.appendChild(tblStyle);

    // Add table width
    const tblW = doc.createElementNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'w:tblW'
    );
    tblW.setAttribute('w:w', '0');
    tblW.setAttribute('w:type', 'auto');
    tblPr.appendChild(tblW);

    // Add table borders
    const tblBorders = doc.createElementNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'w:tblBorders'
    );

    const borderElements = [
      'top',
      'left',
      'bottom',
      'right',
      'insideH',
      'insideV',
    ];
    borderElements.forEach((border) => {
      const borderElement = doc.createElementNS(
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        `w:${border}`
      );
      borderElement.setAttribute('w:val', 'single');
      borderElement.setAttribute('w:sz', '4');
      borderElement.setAttribute('w:space', '0');
      borderElement.setAttribute('w:color', 'auto');
      tblBorders.appendChild(borderElement);
    });

    tblPr.appendChild(tblBorders);
    table.appendChild(tblPr);

    // Add table grid (column definitions)
    const tblGrid = doc.createElementNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'w:tblGrid'
    );

    columns.forEach(() => {
      const gridCol = doc.createElementNS(
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w:gridCol'
      );
      gridCol.setAttribute('w:w', '2500'); // Default width
      tblGrid.appendChild(gridCol);
    });

    table.appendChild(tblGrid);

    // Create header row
    const headerRow = this.createTableRow(doc, columns, true);
    table.appendChild(headerRow);

    // Create data rows
    data.forEach((rowData) => {
      const values = columns.map((col) => rowData[col]?.toString() || '');
      const dataRow = this.createTableRow(doc, values, false);
      table.appendChild(dataRow);
    });

    return table;
  }

  private createTableRow(
    doc: Document,
    cellValues: string[],
    isHeader: boolean
  ): Element {
    const tr = doc.createElementNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'w:tr'
    );

    cellValues.forEach((value) => {
      const tc = doc.createElementNS(
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w:tc'
      );

      // Add cell properties
      const tcPr = doc.createElementNS(
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w:tcPr'
      );

      // Add paragraph with text
      const p = doc.createElementNS(
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w:p'
      );

      // Add run with text
      const r = doc.createElementNS(
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w:r'
      );

      // For headers, make text bold
      if (isHeader) {
        const rPr = doc.createElementNS(
          'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
          'w:rPr'
        );
        const b = doc.createElementNS(
          'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
          'w:b'
        );
        rPr.appendChild(b);
        r.appendChild(rPr);
      }

      const t = doc.createElementNS(
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w:t'
      );

      t.textContent = value;
      r.appendChild(t);
      p.appendChild(r);
      tc.appendChild(tcPr);
      tc.appendChild(p);
      tr.appendChild(tc);
    });

    return tr;
  }

  private findPlaceholders(
    text: string
  ): Array<{ key: string; start: number; end: number }> {
    const regex = /{([^%#\/].*?)}/g; // Exclude special placeholders
    const matches = [];
    let match;
    while ((match = regex.exec(text)) !== null) {
      matches.push({
        key: match[1].trim(),
        start: match.index,
        end: regex.lastIndex,
      });
    }
    return matches;
  }

  private async fetchImageBuffer(url: string): Promise<Uint8Array> {
    try {
      const response = await fetch(url);
      if (!response.ok)
        throw new Error(
          `Failed to fetch image: ${response.status} ${response.statusText}`
        );
      const arrayBuffer = await response.arrayBuffer();
      if (arrayBuffer.byteLength === 0) throw new Error('Image is empty');
      return new Uint8Array(arrayBuffer);
    } catch (error) {
      console.error('Error fetching image:', error);
      throw error;
    }
  }

  private addImageRelation(zip: PizZip, id: string, target: string) {
    try {
      const relsPath = 'word/_rels/document.xml.rels';
      let relsXml = zip.file(relsPath)?.asText();

      if (!relsXml) {
        // Create relationships file if it doesn't exist
        relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          </Relationships>`;
      }

      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(relsXml, 'application/xml');

      // Check if relationship already exists
      const relationships = xmlDoc.getElementsByTagName('Relationship');
      for (let i = 0; i < relationships.length; i++) {
        if (relationships[i].getAttribute('Id') === id) {
          // Already exists, update it
          relationships[i].setAttribute('Target', `media/${target}`);
          const serializer = new XMLSerializer();
          zip.file(relsPath, serializer.serializeToString(xmlDoc));
          return;
        }
      }

      // Add new relationship
      const newRel = xmlDoc.createElementNS(
        'http://schemas.openxmlformats.org/package/2006/relationships',
        'Relationship'
      );
      newRel.setAttribute('Id', id);
      newRel.setAttribute(
        'Type',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
      );
      newRel.setAttribute('Target', `media/${target}`);

      xmlDoc.documentElement.appendChild(newRel);

      const serializer = new XMLSerializer();
      zip.file(relsPath, serializer.serializeToString(xmlDoc));
    } catch (error) {
      console.error('Error adding image relation:', error);
      throw error;
    }
  }

  private ensureContentType(extension: string) {
    if (!this.contentTypesXml) return;

    const contentType = this.getContentTypeForExtension(extension);
    if (!contentType) return;

    // Check if we already have this content type
    const defaultTags = this.contentTypesXml.getElementsByTagName('Default');
    for (let i = 0; i < defaultTags.length; i++) {
      if (defaultTags[i].getAttribute('Extension') === extension) {
        return; // Already exists
      }
    }

    // Add the content type
    const newDefault = this.contentTypesXml.createElementNS(
      'http://schemas.openxmlformats.org/package/2006/content-types',
      'Default'
    );
    newDefault.setAttribute('Extension', extension);
    newDefault.setAttribute('ContentType', contentType);
    this.contentTypesXml.documentElement.appendChild(newDefault);
  }

  private getContentTypeForExtension(extension: string): string | null {
    const contentTypes: Record<string, string> = {
      png: 'image/png',
      jpg: 'image/jpeg',
      jpeg: 'image/jpeg',
      gif: 'image/gif',
      bmp: 'image/bmp',
      svg: 'image/svg+xml',
    };

    return contentTypes[extension.toLowerCase()] || null;
  }

  private generateImageParagraph(
    doc: Document,
    rId: string,
    width: number,
    height: number
  ): Element {
    const cx = width * 9525; // Convert to EMU (English Metric Unit)
    const cy = height * 9525;

    // Create a namespace resolver to handle XML namespaces correctly
    const nsResolver = doc.createNSResolver(doc.documentElement);

    // Get the document's default namespace URI if available
    const defaultNS =
      nsResolver.lookupNamespaceURI('w') ||
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

    // Create a new paragraph element with the correct namespace
    const paragraph = doc.createElementNS(defaultNS, 'w:p');

    // Set the paragraph's inner XML with all required namespaces
    paragraph.innerHTML = `
      <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="${cx}" cy="${cy}"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:docPr id="${this.imageCounter}" name="Picture ${this.imageCounter}" descr="Image"/>
            <wp:cNvGraphicFramePr>
              <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
            </wp:cNvGraphicFramePr>
            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:nvPicPr>
                    <pic:cNvPr id="${this.imageCounter}" name="Image" descr="Image"/>
                    <pic:cNvPicPr>
                      <a:picLocks noChangeAspect="1" noChangeArrowheads="1"/>
                    </pic:cNvPicPr>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="${rId}">
                      <a:extLst>
                        <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
                          <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
                        </a:ext>
                      </a:extLst>
                    </a:blip>
                    <a:stretch>
                      <a:fillRect/>
                    </a:stretch>
                  </pic:blipFill>
                  <pic:spPr bwMode="auto">
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="${cx}" cy="${cy}"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                      <a:avLst/>
                    </a:prstGeom>
                    <a:noFill/>
                    <a:ln>
                      <a:noFill/>
                    </a:ln>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    `;

    return paragraph;
  }

  /**
   * Creates a page break element
   */
  private createPageBreak(doc: Document): Element {
    const paragraph = doc.createElementNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'w:p'
    );

    const run = doc.createElementNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'w:r'
    );

    const br = doc.createElementNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'w:br'
    );

    br.setAttribute('w:type', 'page');
    run.appendChild(br);
    paragraph.appendChild(run);

    return paragraph;
  }

  /**
   * Gets a value from the data object using dot notation path
   */
  private getDeepValue(path: string, data: any): any {
    if (!path || !data) return '';

    const parts = path.split('.');
    let current = data;

    for (const part of parts) {
      if (current === null || current === undefined) return '';
      current = current[part];
    }

    return current !== null && current !== undefined ? current : '';
  }

  /**
   * Update headers and footers in the document
   */
  private async updateHeaderFooter(
    zip: PizZip,
    type: 'header' | 'footer',
    content: string,
    data: any
  ): Promise<void> {
    try {
      // Find header/footer files in the zip
      const files = Object.keys(zip.files).filter(
        (name) => name.startsWith(`word/${type}`) && name.endsWith('.xml')
      );

      if (files.length === 0) {
        console.warn(`No ${type} files found in the template`);
        return;
      }

      // Process each header/footer file
      for (const file of files) {
        const fileContent = zip.file(file)?.asText();
        if (!fileContent) continue;

        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(fileContent, 'application/xml');

        // Process paragraphs in header/footer
        const paragraphs = Array.from(xmlDoc.getElementsByTagName('w:p'));
        for (const paragraph of paragraphs) {
          await this.processParagraph(zip, xmlDoc, paragraph, data);
        }

        // Save updated header/footer
        const serializer = new XMLSerializer();
        zip.file(file, serializer.serializeToString(xmlDoc));
      }
    } catch (error) {
      console.error(`Error updating ${type}:`, error);
    }
  }

  /**
   * Save the generated document
   */
  private saveDocument(zip: PizZip, fileName: string): void {
    try {
      const blob = zip.generate({ type: 'blob' });
      saveAs(blob, fileName);
    } catch (error) {
      console.error('Error saving document:', error);
      throw error;
    }
  }

  /**
   * Adds a chart to the document
   */
  private async addChart(
    zip: PizZip,
    doc: Document,
    paragraph: Element,
    chartType: 'bar' | 'line' | 'pie',
    chartData: any[],
    chartOptions: {
      width?: number;
      height?: number;
      title?: string;
      xAxis?: string;
      yAxis?: string;
    } = {}
  ): Promise<void> {
    try {
      // Default options
      const width = chartOptions.width || 400;
      const height = chartOptions.height || 300;
      const title = chartOptions.title || 'Chart';

      // Generate unique IDs for chart components
      const chartId = `chart${Date.now()}`;
      const rId = `rId${2000 + this.tableCounter++}`;

      // Create chart XML
      const chartXml = this.generateChartXml(
        chartType,
        chartData,
        chartOptions
      );

      // Add chart to the zip file
      const chartPath = `word/charts/${chartId}.xml`;
      zip.file(chartPath, chartXml);

      // Add chart relationship to document.xml.rels
      this.addChartRelation(zip, rId, chartId);

      // Ensure chart content type is defined
      this.ensureChartContentType();

      // Create chart drawing paragraph
      const chartParagraph = this.generateChartParagraph(
        doc,
        rId,
        width,
        height,
        title
      );

      // Replace the original paragraph with the chart paragraph
      if (paragraph.parentNode) {
        paragraph.parentNode.replaceChild(chartParagraph, paragraph);
      }

      console.log(`Chart added: ${chartId}, ID: ${rId}`);
    } catch (error) {
      console.error('Error adding chart:', error);
    }
  }

  /**
   * Generate the XML for a chart
   */
  private generateChartXml(
    chartType: 'bar' | 'line' | 'pie',
    data: any[],
    options: any = {}
  ): string {
    if (!data || data.length === 0) {
      throw new Error('Chart data is empty');
    }

    // Extract category and series data
    const categories = data.map((item) => item.label || '');
    const seriesData = data.map((item) => item.value || 0);

    // Determine chart namespace and type-specific XML
    let chartSpecificXml = '';

    switch (chartType) {
      case 'bar':
        chartSpecificXml = `
        <c:barChart>
          <c:barDir val="col"/>
          <c:grouping val="clustered"/>
          <c:series>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx>
              <c:strRef>
                <c:strCache>
                  <c:ptCount val="1"/>
                  <c:pt idx="0">
                    <c:v>${options.seriesName || 'Series 1'}</c:v>
                  </c:pt>
                </c:strCache>
              </c:strRef>
            </c:tx>
            <c:cat>
              <c:strRef>
                <c:strCache>
                  <c:ptCount val="${categories.length}"/>
                  ${categories
                    .map(
                      (cat, i) => `<c:pt idx="${i}"><c:v>${cat}</c:v></c:pt>`
                    )
                    .join('')}
                </c:strCache>
              </c:strRef>
            </c:cat>
            <c:val>
              <c:numRef>
                <c:numCache>
                  <c:formatCode>General</c:formatCode>
                  <c:ptCount val="${seriesData.length}"/>
                  ${seriesData
                    .map(
                      (val, i) => `<c:pt idx="${i}"><c:v>${val}</c:v></c:pt>`
                    )
                    .join('')}
                </c:numCache>
              </c:numRef>
            </c:val>
          </c:series>
          <c:dLbls>
            <c:showVal val="0"/>
            <c:showSerName val="0"/>
            <c:showCatName val="0"/>
            <c:showLegendKey val="0"/>
          </c:dLbls>
          <c:axId val="42"/>
          <c:axId val="43"/>
        </c:barChart>`;
        break;

      case 'line':
        chartSpecificXml = `
        <c:lineChart>
          <c:grouping val="standard"/>
          <c:series>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx>
              <c:strRef>
                <c:strCache>
                  <c:ptCount val="1"/>
                  <c:pt idx="0">
                    <c:v>${options.seriesName || 'Series 1'}</c:v>
                  </c:pt>
                </c:strCache>
              </c:strRef>
            </c:tx>
            <c:marker>
              <c:symbol val="circle"/>
              <c:size val="5"/>
            </c:marker>
            <c:cat>
              <c:strRef>
                <c:strCache>
                  <c:ptCount val="${categories.length}"/>
                  ${categories
                    .map(
                      (cat, i) => `<c:pt idx="${i}"><c:v>${cat}</c:v></c:pt>`
                    )
                    .join('')}
                </c:strCache>
              </c:strRef>
            </c:cat>
            <c:val>
              <c:numRef>
                <c:numCache>
                  <c:formatCode>General</c:formatCode>
                  <c:ptCount val="${seriesData.length}"/>
                  ${seriesData
                    .map(
                      (val, i) => `<c:pt idx="${i}"><c:v>${val}</c:v></c:pt>`
                    )
                    .join('')}
                </c:numCache>
              </c:numRef>
            </c:val>
          </c:series>
          <c:axId val="42"/>
          <c:axId val="43"/>
        </c:lineChart>`;
        break;

      case 'pie':
        chartSpecificXml = `
        <c:pieChart>
          <c:series>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx>
              <c:strRef>
                <c:strCache>
                  <c:ptCount val="1"/>
                  <c:pt idx="0">
                    <c:v>${options.seriesName || 'Series 1'}</c:v>
                  </c:pt>
                </c:strCache>
              </c:strRef>
            </c:tx>
            <c:cat>
              <c:strRef>
                <c:strCache>
                  <c:ptCount val="${categories.length}"/>
                  ${categories
                    .map(
                      (cat, i) => `<c:pt idx="${i}"><c:v>${cat}</c:v></c:pt>`
                    )
                    .join('')}
                </c:strCache>
              </c:strRef>
            </c:cat>
            <c:val>
              <c:numRef>
                <c:numCache>
                  <c:formatCode>General</c:formatCode>
                  <c:ptCount val="${seriesData.length}"/>
                  ${seriesData
                    .map(
                      (val, i) => `<c:pt idx="${i}"><c:v>${val}</c:v></c:pt>`
                    )
                    .join('')}
                </c:numCache>
              </c:numRef>
            </c:val>
          </c:series>
          <c:dLbls>
            <c:showVal val="0"/>
            <c:showSerName val="0"/>
            <c:showCatName val="1"/>
            <c:showLegendKey val="0"/>
          </c:dLbls>
        </c:pieChart>`;
        break;
    }

    // Full chart XML
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" 
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <c:chart>
        <c:title>
          <c:tx>
            <c:rich>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:pPr>
                  <a:defRPr/>
                </a:pPr>
                <a:r>
                  <a:rPr lang="en-US"/>
                  <a:t>${options.title || 'Chart'}</a:t>
                </a:r>
              </a:p>
            </c:rich>
          </c:tx>
          <c:layout/>
          <c:overlay val="0"/>
        </c:title>
        <c:plotArea>
          <c:layout/>
          ${chartSpecificXml}
          <c:catAx>
            <c:axId val="42"/>
            <c:scaling>
              <c:orientation val="minMax"/>
            </c:scaling>
            <c:delete val="0"/>
            <c:axPos val="b"/>
            <c:title>
              <c:tx>
                <c:rich>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p>
                    <a:pPr>
                      <a:defRPr/>
                    </a:pPr>
                    <a:r>
                      <a:rPr lang="en-US"/>
                      <a:t>${options.xAxis || 'X Axis'}</a:t>
                    </a:r>
                  </a:p>
                </c:rich>
              </c:tx>
              <c:layout/>
              <c:overlay val="0"/>
            </c:title>
            <c:numFmt formatCode="General" sourceLinked="1"/>
            <c:majorTickMark val="out"/>
            <c:minorTickMark val="none"/>
            <c:tickLblPos val="nextTo"/>
            <c:crossAx val="43"/>
            <c:crosses val="autoZero"/>
            <c:auto val="1"/>
            <c:lblAlgn val="ctr"/>
            <c:lblOffset val="100"/>
          </c:catAx>
          <c:valAx>
            <c:axId val="43"/>
            <c:scaling>
              <c:orientation val="minMax"/>
            </c:scaling>
            <c:delete val="0"/>
            <c:axPos val="l"/>
            <c:title>
              <c:tx>
                <c:rich>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p>
                    <a:pPr>
                      <a:defRPr/>
                    </a:pPr>
                    <a:r>
                      <a:rPr lang="en-US"/>
                      <a:t>${options.yAxis || 'Y Axis'}</a:t>
                    </a:r>
                  </a:p>
                </c:rich>
              </c:tx>
              <c:layout/>
              <c:overlay val="0"/>
            </c:title>
            <c:numFmt formatCode="General" sourceLinked="1"/>
            <c:majorTickMark val="out"/>
            <c:minorTickMark val="none"/>
            <c:tickLblPos val="nextTo"/>
            <c:crossAx val="42"/>
            <c:crosses val="autoZero"/>
            <c:crossBetween val="between"/>
          </c:valAx>
        </c:plotArea>
        <c:legend>
          <c:legendPos val="r"/>
          <c:layout/>
          <c:overlay val="0"/>
        </c:legend>
        <c:plotVisOnly val="1"/>
      </c:chart>
    </c:chartSpace>`;
  }

  private addChartRelation(zip: PizZip, id: string, chartId: string): void {
    try {
      const relsPath = 'word/_rels/document.xml.rels';
      let relsXml = zip.file(relsPath)?.asText();

      if (!relsXml) {
        // Create relationships file if it doesn't exist
        relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          </Relationships>`;
      }

      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(relsXml, 'application/xml');

      // Add new relationship for chart
      const newRel = xmlDoc.createElementNS(
        'http://schemas.openxmlformats.org/package/2006/relationships',
        'Relationship'
      );
      newRel.setAttribute('Id', id);
      newRel.setAttribute(
        'Type',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
      );
      newRel.setAttribute('Target', `charts/${chartId}.xml`);

      xmlDoc.documentElement.appendChild(newRel);

      const serializer = new XMLSerializer();
      zip.file(relsPath, serializer.serializeToString(xmlDoc));
    } catch (error) {
      console.error('Error adding chart relation:', error);
      throw error;
    }
  }

  private ensureChartContentType(): void {
    if (!this.contentTypesXml) return;

    // Check if chart content type is already defined
    const overrideTags = this.contentTypesXml.getElementsByTagName('Override');
    for (let i = 0; i < overrideTags.length; i++) {
      const contentType = overrideTags[i].getAttribute('ContentType');
      if (
        contentType ===
        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
      ) {
        return; // Already exists
      }
    }

    // Add chart content type
    const newOverride = this.contentTypesXml.createElementNS(
      'http://schemas.openxmlformats.org/package/2006/content-types',
      'Override'
    );
    newOverride.setAttribute('PartName', '/word/charts/chart1.xml');
    newOverride.setAttribute(
      'ContentType',
      'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
    );
    this.contentTypesXml.documentElement.appendChild(newOverride);
  }

  private generateChartParagraph(
    doc: Document,
    rId: string,
    width: number,
    height: number,
    title: string
  ): Element {
    const cx = width * 9525; // Convert to EMU (English Metric Unit)
    const cy = height * 9525;

    const paragraph = doc.createElementNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'w:p'
    );

    paragraph.innerHTML = `
      <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="${cx}" cy="${cy}"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:docPr id="${this.tableCounter}" name="Chart ${this.tableCounter}" descr="${title}"/>
            <wp:cNvGraphicFramePr>
              <a:graphicFrameLocks noChangeAspect="1"/>
            </wp:cNvGraphicFramePr>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" 
                         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
                         r:id="${rId}"/>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    `;

    return paragraph;
  }

  /**
   * Add watermark text to the document
   */
  async addWatermark(
    zip: PizZip,
    text: string,
    options: {
      color?: string;
      opacity?: number;
      fontSize?: number;
      angle?: number;
    } = {}
  ): Promise<void> {
    try {
      // Default options
      const color = options.color || 'CCCCCC';
      const opacity = options.opacity || 0.5;
      const fontSize = options.fontSize || 60;
      const angle = options.angle || 45;

      // Get the header files
      const headerFiles = Object.keys(zip.files).filter(
        (name) => name.startsWith('word/header') && name.endsWith('.xml')
      );

      if (headerFiles.length === 0) {
        console.warn('No header files found to add watermark');
        return;
      }

      // Process each header file
      for (const headerFile of headerFiles) {
        const headerContent = zip.file(headerFile)?.asText();
        if (!headerContent) continue;

        const parser = new DOMParser();
        const headerDoc = parser.parseFromString(
          headerContent,
          'application/xml'
        );

        // Create watermark element
        const watermark = this.createWatermarkElement(
          headerDoc,
          text,
          color,
          opacity,
          fontSize,
          angle
        );

        // Insert watermark at the beginning of the header
        const headerRoot = headerDoc.getElementsByTagName('w:hdr')[0];

        if (headerRoot) {
          if (headerRoot.firstChild) {
            headerRoot.insertBefore(watermark, headerRoot.firstChild);
          } else {
            headerRoot.appendChild(watermark);
          }

          // Save updated header
          const serializer = new XMLSerializer();
          zip.file(headerFile, serializer.serializeToString(headerDoc));
        }
      }

      console.log('Watermark added successfully');
    } catch (error) {
      console.error('Error adding watermark:', error);
    }
  }

  private createWatermarkElement(
    doc: Document,
    text: string,
    color: string,
    opacity: number,
    fontSize: number,
    angle: number
  ): Element {
    // Create paragraph for watermark
    const paragraph = doc.createElementNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'w:p'
    );

    // Set the watermark XML
    paragraph.innerHTML = `
      <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
           xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <w:rPr>
          <w:noProof/>
        </w:rPr>
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251658240" 
                     behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="page">
              <wp:posOffset>0</wp:posOffset>
            </wp:positionH>
            <wp:positionV relativeFrom="page">
              <wp:posOffset>0</wp:posOffset>
            </wp:positionV>
            <wp:extent cx="5400000" cy="3600000"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapNone/>
            <wp:docPr id="1" name="Watermark" descr="Watermark"/>
            <wp:cNvGraphicFramePr>
              <a:graphicFrameLocks noChangeAspect="0"/>
            </wp:cNvGraphicFramePr>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
                <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr>
                    <a:xfrm rot="${angle * 60000}">
                      <a:off x="0" y="0"/>
                      <a:ext cx="5400000" cy="3600000"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                      <a:avLst/>
                    </a:prstGeom>
                    <a:noFill/>
                    <a:ln>
                      <a:noFill/>
                    </a:ln>
                  </wps:spPr>
                  <wps:style>
                    <a:lnRef idx="0">
                      <a:schemeClr val="accent1"/>
                    </a:lnRef>
                    <a:fillRef idx="0">
                      <a:schemeClr val="accent1"/>
                    </a:fillRef>
                    <a:effectRef idx="0">
                      <a:schemeClr val="accent1"/>
                    </a:effectRef>
                    <a:fontRef idx="minor">
                      <a:sch                    <a:fontRef idx="minor">
                      <a:schemeClr val="lt1"/>
                    </a:fontRef>
                  </wps:style>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p>
                        <w:r>
                          <w:rPr>
                            <w:color w:val="${color}" w:themeColor="background1"/>
                            <w:sz w:val="${fontSize * 2}"/>
                            <w:szCs w:val="${fontSize * 2}"/>
                          </w:rPr>
                          <w:t>${text}</w:t>
                        </w:r>
                      </w:p>
                    </w:txbxContent>
                  </wps:txbx>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </w:r>
    `;

    return paragraph;
  }
}
