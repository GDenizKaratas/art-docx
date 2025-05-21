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

      if (imageUrl) {
        try {
          // Extract image dimensions if specified
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

  // Continuing from the existing generateImageParagraph method:

  private createPageBreak(doc: Document): Element {
    const paragraph = doc.createElementNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'w:p'
    );

    const r = doc.createElementNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'w:r'
    );

    const br = doc.createElementNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'w:br'
    );

    br.setAttribute('w:type', 'page');

    r.appendChild(br);
    paragraph.appendChild(r);

    return paragraph;
  }

  private getDeepValue(key: string, data: any): any {
    try {
      // Handle null data
      if (!data) return '';

      // Handle direct access to a property
      if (key in data) return data[key] || '';

      // Split the key by dots for nested access
      const keys = key.split('.');
      let value = data;

      for (const k of keys) {
        if (value === undefined || value === null) return '';
        value = value[k];
      }

      return value !== undefined && value !== null ? value : '';
    } catch (error) {
      console.error(`Error getting value for key ${key}:`, error);
      return '';
    }
  }

  private async updateHeaderFooter(
    zip: PizZip,
    type: 'header' | 'footer',
    content: string,
    data: any
  ): Promise<void> {
    try {
      // Find all header/footer files in the zip
      const filePattern = new RegExp(`word/${type}\\d+\\.xml$`);
      const files = Object.keys(zip.files).filter((name) =>
        filePattern.test(name)
      );

      if (files.length === 0) {
        console.warn(`No ${type} files found in the template`);
        return;
      }

      // Process each header/footer file
      for (const filePath of files) {
        const fileContent = zip.file(filePath)?.asText();
        if (!fileContent) continue;

        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(fileContent, 'application/xml');

        // Process all paragraphs in the header/footer
        const paragraphs = Array.from(xmlDoc.getElementsByTagName('w:p'));
        for (const paragraph of paragraphs) {
          await this.processParagraph(zip, xmlDoc, paragraph, data);
        }

        // Save the modified header/footer
        const serializer = new XMLSerializer();
        zip.file(filePath, serializer.serializeToString(xmlDoc));
      }
    } catch (error) {
      console.error(`Error updating ${type}:`, error);
    }
  }

  private saveDocument(zip: PizZip, filename: string): void {
    try {
      // Generate the zip file
      const blob = zip.generate({
        type: 'blob',
        mimeType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        compression: 'DEFLATE',
      });

      // Save the file
      saveAs(blob, filename);
    } catch (error) {
      console.error('Error saving document:', error);
      throw error;
    }
  }

  // Add this helper method to handle complex content with mixed formatting
  private processComplexContent(
    doc: Document,
    content: string,
    data: any
  ): DocumentFragment {
    // Create a document fragment to hold the processed content
    const fragment = doc.createDocumentFragment();

    // Process different types of formatting and placeholders
    // This is a more advanced version that could handle mixed content
    // TODO: Implement advanced parsing and formatting if needed

    // For now, just create a simple text node
    const textNode = doc.createTextNode(content);
    fragment.appendChild(textNode);

    return fragment;
  }

  // Method for handling repeating sections
  public processRepeatingSection(
    doc: Document,
    sectionTemplate: Element,
    dataArray: any[]
  ): DocumentFragment {
    const fragment = doc.createDocumentFragment();

    if (!dataArray || !Array.isArray(dataArray) || dataArray.length === 0) {
      return fragment;
    }

    // Clone the template for each data item
    dataArray.forEach((item) => {
      const clonedSection = sectionTemplate.cloneNode(true) as Element;

      // Process placeholders in the cloned section
      const paragraphs = Array.from(clonedSection.getElementsByTagName('w:p'));
      paragraphs.forEach((paragraph) => {
        // Process with the current item data
        this.processParagraph(null as unknown as PizZip, doc, paragraph, item);
      });

      fragment.appendChild(clonedSection);
    });

    return fragment;
  }

  // Method to reset counters when starting a new document
  public resetCounters(): void {
    this.imageCounter = 1;
    this.tableCounter = 1;
  }

  // Utility method to create a new Word document from scratch
  public createBlankDocument(): PizZip {
    const zip = new PizZip();

    // Create the minimum required structure for a valid DOCX

    // [Content_Types].xml
    zip.file(
      '[Content_Types].xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
        <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
        <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
      </Types>`
    );

    // _rels/.rels
    zip.file(
      '_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
        <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
      </Relationships>`
    );

    // word/_rels/document.xml.rels
    zip.file(
      'word/_rels/document.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
      </Relationships>`
    );

    // word/document.xml
    zip.file(
      'word/document.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r>
              <w:t>New Document</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>`
    );

    // word/styles.xml
    zip.file(
      'word/styles.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:type="paragraph" w:styleId="Normal">
          <w:name w:val="Normal"/>
          <w:pPr/>
          <w:rPr/>
        </w:style>
      </w:styles>`
    );

    // docProps/core.xml
    const now = new Date().toISOString();
    zip.file(
      'docProps/core.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
                        xmlns:dc="http://purl.org/dc/elements/1.1/" 
                        xmlns:dcterms="http://purl.org/dc/terms/">
        <dc:creator>GenerateDocxService</dc:creator>
        <cp:lastModifiedBy>GenerateDocxService</cp:lastModifiedBy>
        <dcterms:created xsi:type="dcterms:W3CDTF" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">${now}</dcterms:created>
        <dcterms:modified xsi:type="dcterms:W3CDTF" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">${now}</dcterms:modified>
      </cp:coreProperties>`
    );

    // docProps/app.xml
    zip.file(
      'docProps/app.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
        <Application>GenerateDocxService</Application>
        <Company>Document Generator</Company>
      </Properties>`
    );

    return zip;
  }

  // Enhanced method to handle various date formats
  private formatDate(value: any, format: string = 'dd/MM/yyyy'): string {
    if (!value) return '';

    try {
      const date = new Date(value);

      if (isNaN(date.getTime())) {
        return value.toString();
      }

      // Replace format tokens with actual values
      return format
        .replace('yyyy', date.getFullYear().toString())
        .replace('MM', (date.getMonth() + 1).toString().padStart(2, '0'))
        .replace('dd', date.getDate().toString().padStart(2, '0'))
        .replace('HH', date.getHours().toString().padStart(2, '0'))
        .replace('mm', date.getMinutes().toString().padStart(2, '0'))
        .replace('ss', date.getSeconds().toString().padStart(2, '0'));
    } catch (error) {
      console.error('Error formatting date:', error);
      return value.toString();
    }
  }

  // Format numbers with proper decimal and thousand separators
  private formatNumber(value: any, format: string = '#,##0.00'): string {
    if (value === null || value === undefined) return '';

    try {
      const num = Number(value);

      if (isNaN(num)) {
        return value.toString();
      }

      // Simple implementation of number formatting
      const parts = format.split('.');
      const hasDecimal = parts.length > 1;
      const decimalPlaces = hasDecimal ? parts[1].length : 0;

      const formattedNum = new Intl.NumberFormat('tr-TR', {
        minimumFractionDigits: decimalPlaces,
        maximumFractionDigits: decimalPlaces,
        useGrouping: format.includes(','),
      }).format(num);

      return formattedNum;
    } catch (error) {
      console.error('Error formatting number:', error);
      return value.toString();
    }
  }
}
