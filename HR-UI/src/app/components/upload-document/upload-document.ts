import { Component, inject, Output, EventEmitter } from '@angular/core';
import { CommonModule } from '@angular/common';
import { DocumentService } from '../../services/document.service';
import * as mammoth from 'mammoth';

@Component({
  selector: 'app-upload-document',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './upload-document.html',
  styleUrl: './upload-document.css',
})
export class UploadDocumentComponent {
  private readonly docService = inject(DocumentService);

  @Output() onUploadSuccess = new EventEmitter<void>();

  public isDragOver = false;
  public loading = false;

  public createBlankDocument() {
    const blankHtml = `
      <div style="font-family: 'Outfit', Arial, sans-serif; color: #2e353e; padding: 40px; line-height: 1.6; background: #fff;">
        <h2 style="color: #2e353e; font-size: 24px; font-weight: 700; margin-bottom: 20px; border-bottom: 2px solid #5974a3; padding-bottom: 10px;">New Document</h2>
        <p>Start writing your custom content here...</p>
      </div>
    `;
    this.docService.loadUploadedDocument('New Custom Document', blankHtml, {});
    this.onUploadSuccess.emit();
  }

  public onDragOver(event: DragEvent) {
    event.preventDefault();
    this.isDragOver = true;
  }

  public onDragLeave(event: DragEvent) {
    event.preventDefault();
    this.isDragOver = false;
  }

  public onDrop(event: DragEvent) {
    event.preventDefault();
    this.isDragOver = false;

    const files = event.dataTransfer?.files;
    if (files && files.length > 0) {
      this.processFile(files[0]);
    }
  }

  public onFileSelected(event: Event) {
    const target = event.target as HTMLInputElement;
    const files = target.files;
    if (files && files.length > 0) {
      this.processFile(files[0]);
    }
  }

  private processFile(file: File) {
    this.loading = true;
    const name = file.name;
    const extension = name.substring(name.lastIndexOf('.')).toLowerCase();

    if (extension === '.docx') {
      this.parseDocx(file);
    } else if (extension === '.pdf') {
      this.parsePdfPlaceholder(file);
    } else if (extension === '.txt') {
      this.parseTxt(file);
    } else {
      alert('Unsupported file format! Please upload a .docx, .pdf, or .txt file.');
      this.loading = false;
    }
  }

  private parseDocx(file: File) {
    const reader = new FileReader();
    reader.onload = async () => {
      try {
        const arrayBuffer = reader.result as ArrayBuffer;
        const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });

        let html = result.value;
        if (!html || html.trim() === '') {
          html = `<p>Document empty or could not be parsed.</p>`;
        }

        const detectedVars = this.detectVariables(html);
        const standardisedHtml = this.standardisePlaceholders(html, detectedVars);

        // Wrap mammoth html in a nice document template body
        const wrappedHtml = `
          <div style="font-family: 'Outfit', Arial, sans-serif; color: #2e353e; line-height: 1.6; background: #fff; padding: 20px;">
            ${standardisedHtml}
          </div>
        `;

        const inputVars: Record<string, string> = {};
        Object.keys(detectedVars).forEach((k) => {
          inputVars[k] = detectedVars[k];
        });

        this.docService.loadUploadedDocument(file.name, wrappedHtml, inputVars);
        this.loading = false;
        this.onUploadSuccess.emit();
      } catch (err) {
        console.error(err);
        alert('Error parsing Word document.');
        this.loading = false;
      }
    };
    reader.readAsArrayBuffer(file);
  }

  private parseTxt(file: File) {
    const reader = new FileReader();
    reader.onload = () => {
      const text = reader.result as string;
      const detectedVars = this.detectVariables(text);
      const standardisedText = this.standardisePlaceholders(text, detectedVars);
      const html = `<div style="white-space: pre-line; font-family: 'Outfit', Arial, sans-serif; padding: 20px;">${standardisedText}</div>`;

      const inputVars: Record<string, string> = {};
      Object.keys(detectedVars).forEach((k) => {
        inputVars[k] = detectedVars[k];
      });

      this.docService.loadUploadedDocument(file.name, html, inputVars);
      this.loading = false;
      this.onUploadSuccess.emit();
    };
    reader.readAsText(file);
  }

  private parsePdfPlaceholder(file: File) {
    setTimeout(() => {
      const pdfTextHTML = `
        <div style="font-family: 'Outfit', Arial, sans-serif; color: #2e353e; padding: 40px; line-height: 1.6;">
          <h2 style="color: #5974a3; text-align: center;">PDF DOCUMENT LAYOUT IMPORT</h2>
          <p style="text-align: center; color: #718096; font-size: 12px; margin-bottom: 30px;">Imported from: ${file.name}</p>

          <p>Dear <strong>{{employeeName}}</strong>,</p>
          <p>We are writing to confirm that the document <strong>${file.name}</strong> was successfully uploaded and processed. Below is the editable text extracted from your PDF template.</p>

          <p>Please edit this text directly as required. You can add new terms, salary grids, or clauses, and then download it back as PDF or Word document.</p>

          <div style="border-top: 1px dashed #e1e5eb; margin-top: 50px; padding-top: 20px;">
            <p><strong>Company:</strong> {{companyName}}</p>
            <p><strong>Designation:</strong> {{designation}}</p>
            <p><strong>Salary Amount:</strong> {{salaryAmount}}</p>
          </div>
        </div>
      `;
      const detectedVars = {
        employeeName: 'Ashish Kumar Yadav',
        companyName: 'Acme Corporation',
        designation: 'Product Analyst',
        salaryAmount: '8,50,000',
      };

      this.docService.loadUploadedDocument(file.name, pdfTextHTML, detectedVars);
      this.loading = false;
      this.onUploadSuccess.emit();
    }, 1500);
  }

  private detectVariables(text: string): Record<string, string> {
    const vars: Record<string, string> = {};

    const braceRegex = /\{\{([^}]+)\}\}/g;
    let match;
    while ((match = braceRegex.exec(text)) !== null) {
      const v = match[1].trim();
      vars[v] = v;
    }

    // Pattern 2: [variable]
    const bracketRegex = /\[([^[\]\d]+)\]/g;
    while ((match = bracketRegex.exec(text)) !== null) {
      const v = match[1].trim();
      if (!v.includes('=') && !v.includes('"') && v.length < 25) {
        const camel = v
          .replace(/(?:^\w|[A-Z]|\b\w)/g, (word, index) =>
            index === 0 ? word.toLowerCase() : word.toUpperCase(),
          )
          .replace(/\s+/g, '');
        vars[camel] = v;
      }
    }

    return vars;
  }

  private standardisePlaceholders(html: string, vars: Record<string, string>): string {
    let result = html;
    Object.keys(vars).forEach((key) => {
      const originalValue = vars[key];

      const bracketPlaceholder = `[${originalValue}]`;
      const bracePlaceholder = `{{${originalValue}}}`;

      const escapeReg = (str: string) => str.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');

      result = result.replace(new RegExp(escapeReg(bracketPlaceholder), 'gi'), `{{${key}}}`);
      result = result.replace(new RegExp(escapeReg(bracePlaceholder), 'gi'), `{{${key}}}`);
      result = result.replace(new RegExp(escapeReg(`{{${key}}}`), 'gi'), `{{${key}}}`);
    });
    return result;
  }
}