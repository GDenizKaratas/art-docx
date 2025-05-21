import { Injectable } from '@angular/core';
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';

@Injectable({ providedIn: 'root' })
export class GenerateDocxService {
  private imageCounter = 1;
  private contentTypesXml: Document | null = null;

  async generateDocx(templateFile: File, data: any): Promise<void> {
    try {
      const zip = await this.loadZip(templateFile);

      // Load content types XML
      const contentTypesContent = zip.file('[Content_Types].xml')?.asText();
      if (contentTypesContent) {
        const parser = new DOMParser();
        this.contentTypesXml = parser.parseFromString(
          contentTypesContent,
          'application/xml'
        );
      }

      const documentXml = zip.file('word/document.xml')?.asText() || '';
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(documentXml, 'application/xml');

      const paragraphs = Array.from(xmlDoc.getElementsByTagName('w:p'));
      for (const paragraph of paragraphs) {
        await this.processParagraph(zip, xmlDoc, paragraph, data);
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
      this.saveDocument(zip);
      console.log('Document generation completed successfully');
    } catch (error) {
      console.error('Error generating document:', error);
    }
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
              : 200;
          const height =
            property === 'height'
              ? parseInt(this.getDeepValue(`${imageKey}.height`, data), 10) ||
                150
              : 150;

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

  private findPlaceholders(
    text: string
  ): Array<{ key: string; start: number; end: number }> {
    const regex = /{([^%].*?)}/g; // Exclude image placeholders
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

  private getDeepValue(path: string, obj: any): any {
    if (!path) return '';
    try {
      return path.split('.').reduce((acc, part) => acc?.[part], obj) ?? '';
    } catch {
      return '';
    }
  }

  private saveDocument(zip: PizZip): void {
    try {
      const blob = zip.generate({
        type: 'blob',
        mimeType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        compression: 'DEFLATE',
      });
      saveAs(blob, 'generated-document.docx');
    } catch (error) {
      console.error('Error saving document:', error);
      throw error;
    }
  }
}
