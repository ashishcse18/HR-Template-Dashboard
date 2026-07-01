import { Component, inject, Output, EventEmitter } from '@angular/core';
import { CommonModule } from '@angular/common';
import { DocumentService } from '../../services/document.service';

@Component({
  selector: 'app-template-list',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './template-list.html',
  styleUrls: ['./template-list.css'],
})
export class TemplateListComponent {
  public readonly docService = inject(DocumentService);

  @Output() selectTemplate = new EventEmitter<void>();
  @Output() triggerUploadNavigation = new EventEmitter<void>();

  openTemplate(templateId: string): void {
    this.docService.loadTemplate(templateId);
    this.selectTemplate.emit();
  }

  getObjectKeys(obj: any): string[] {
    return Object.keys(obj || {});
  }
}