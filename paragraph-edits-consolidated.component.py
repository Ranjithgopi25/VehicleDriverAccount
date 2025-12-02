import { Component, Input, Output, EventEmitter } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ParagraphEdit } from '../../../../core/models/message.model';
import { allParagraphsDecided } from '../../../../core/utils/paragraph-edit.utils';

@Component({
  selector: 'app-paragraph-edits',
  standalone: true,
  imports: [CommonModule],
  template: `
    <div class="result-section">
      <h4 class="result-title">Paragraph Edits</h4>
      <p class="paragraph-instructions" *ngIf="!showFinalOutput">
        Review each paragraph edit below. Click the buttons to approve (✓) or decline (✗) each edit.
      </p>
      <p class="paragraph-instructions" *ngIf="showFinalOutput">
        Below are the paragraph-by-paragraph edits. The revised article is shown below.
      </p>
      
      <div class="paragraph-edits-container">
        <div *ngFor="let paragraph of paragraphEdits; let i = index" 
             class="paragraph-edit-item"
             [class.approved]="paragraph.approved === true"
             [class.declined]="paragraph.approved === false">
          
          <div class="paragraph-header">
            <span class="paragraph-number">Paragraph {{ i + 1 }}</span>
            <div class="approval-buttons" *ngIf="!showFinalOutput">
              <button 
                type="button"
                class="approve-btn"
                [class.active]="paragraph.approved === true"
                (click)="onApprove(paragraph.index); $event.stopPropagation()"
                [title]="paragraph.approved === true ? 'Approved' : 'Click to approve'">
                ✓ Approve
              </button>
              <button 
                type="button"
                class="decline-btn"
                [class.active]="paragraph.approved === false"
                (click)="onDecline(paragraph.index); $event.stopPropagation()"
                [title]="paragraph.approved === false ? 'Declined' : 'Click to decline'">
                ✗ Decline
              </button>
            </div>
            <div class="approval-status" *ngIf="showFinalOutput">
              <span *ngIf="paragraph.approved === true" class="status-badge approved-badge">✓ Approved</span>
              <span *ngIf="paragraph.approved === false" class="status-badge declined-badge">✗ Declined</span>
              <span *ngIf="paragraph.approved === null" class="status-badge undecided-badge">○ Not Used</span>
            </div>
          </div>
          
          <div class="paragraph-comparison-boxes">
            <div class="paragraph-box paragraph-box-original">
              <h5>Original</h5>
              <div class="paragraph-text-box">
                <span *ngIf="paragraph.original && paragraph.original.trim(); else noOriginal">{{ paragraph.original }}</span>
                <ng-template #noOriginal>
                  <span class="no-content-placeholder">(No original content)</span>
                </ng-template>
              </div>
            </div>
            
            <div class="paragraph-box paragraph-box-edited"
                 [class.approved-box]="paragraph.approved === true"
                 [class.declined-box]="paragraph.approved === false">
              <h5>Edited</h5>
              <div class="paragraph-text-box" [class.declined-text]="paragraph.approved === false">
                <span *ngIf="paragraph.edited && paragraph.edited.trim(); else noEdited">{{ paragraph.edited }}</span>
                <ng-template #noEdited>
                  <span class="no-content-placeholder">(No edited content)</span>
                </ng-template>
              </div>
            </div>
          </div>
          
          <div *ngIf="paragraph.tags && paragraph.tags.length > 0" class="paragraph-tags">
            <strong>Services Used:</strong>
            <span *ngFor="let tag of paragraph.tags" class="tag-badge">{{ tag }}</span>
          </div>
        </div>
      </div>
      
      <div class="final-output-actions" *ngIf="!showFinalOutput">
        <button 
          type="button"
          class="final-output-btn"
          (click)="onGenerateFinal(); $event.stopPropagation()"
          [disabled]="!allParagraphsDecided || isGeneratingFinal">
          <span *ngIf="isGeneratingFinal" class="spinner"></span>
          {{ isGeneratingFinal ? 'Generating...' : 'Run Final Output' }}
        </button>
        <p *ngIf="!allParagraphsDecided" class="final-output-hint">
          Please approve or decline all paragraph edits to generate the final article.
        </p>
      </div>
    </div>
  `,
  styles: [`
    :host {
      display: block;
      position: relative;
      pointer-events: auto;
    }

    .result-section {
      margin-top: 16px;
      position: relative;
      pointer-events: auto;
    }

    .result-title {
      font-size: 14px;
      font-weight: 600;
      color: var(--text-primary, #1F2937);
      margin-bottom: 8px;
    }

    .paragraph-instructions {
      font-size: 13px;
      color: #6B7280;
      margin-bottom: 16px;
    }

    .paragraph-edits-container {
      display: flex;
      flex-direction: column;
      gap: 20px;
      margin-top: 16px;
    }

    .paragraph-edit-item {
      border: 1px solid var(--border-color, #E5E7EB);
      border-radius: 8px;
      padding: 16px;
      background: var(--bg-primary, #FFFFFF);
      position: relative;
      pointer-events: auto;
    }

    .paragraph-edit-item.approved {
      border-color: #10b981;
      background: #F0FDF4;
    }

    .paragraph-edit-item.declined {
      border-color: #EF4444;
      background: #FEF2F2;
    }

    .paragraph-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 12px;
      padding-bottom: 12px;
      border-bottom: 1px solid var(--border-color, #E5E7EB);
    }

    .paragraph-number {
      font-weight: 600;
      font-size: 14px;
      color: var(--text-primary, #1F2937);
    }

    .approval-buttons {
      display: flex;
      gap: 8px;
      position: relative;
      z-index: 20;
      pointer-events: auto;
    }

    .approve-btn,
    .decline-btn {
      padding: 6px 16px;
      border-radius: 6px;
      font-size: 13px;
      font-weight: 500;
      cursor: pointer;
      transition: all 0.2s ease;
      border: 2px solid transparent;
      display: inline-block;
      text-align: center;
      text-decoration: none;
      -webkit-appearance: none;
      -moz-appearance: none;
      appearance: none;
      user-select: none;
      margin: 0;
      font-family: inherit;
      position: relative;
      z-index: 25;
      pointer-events: auto;
      touch-action: manipulation;
    }

    .approve-btn:hover:not(:disabled),
    .decline-btn:hover:not(:disabled) {
      transform: translateY(-1px);
      box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }

    .approve-btn:active:not(:disabled),
    .decline-btn:active:not(:disabled) {
      transform: translateY(0);
    }

    .approve-btn:disabled,
    .decline-btn:disabled {
      opacity: 0.6;
      cursor: not-allowed;
      pointer-events: none;
    }

    .approve-btn:focus,
    .decline-btn:focus {
      outline: 2px solid #D04A02;
      outline-offset: 2px;
    }

    .approve-btn {
      background-color: #F0FDF4;
      color: #059669;
      border-color: #10b981;
    }

    .approve-btn:hover:not(:disabled) {
      background-color: #D1FAE5;
      border-color: #059669;
    }

    .approve-btn.active {
      background-color: #10b981;
      color: white;
      border-color: #10b981;
    }

    .decline-btn {
      background-color: #FEF2F2;
      color: #DC2626;
      border-color: #EF4444;
    }

    .decline-btn:hover:not(:disabled) {
      background-color: #FEE2E2;
      border-color: #DC2626;
    }

    .decline-btn.active {
      background-color: #EF4444;
      color: white;
      border-color: #EF4444;
    }

    .approval-status {
      display: flex;
      align-items: center;
    }

    .status-badge {
      padding: 4px 12px;
      border-radius: 12px;
      font-size: 12px;
      font-weight: 500;
    }

    .approved-badge {
      background-color: #F0FDF4;
      color: #059669;
      border: 1px solid #10b981;
    }

    .declined-badge {
      background-color: #FEF2F2;
      color: #DC2626;
      border: 1px solid #EF4444;
    }

    .undecided-badge {
      background-color: #F5F5F5;
      color: #6B7280;
      border: 1px solid #E5E7EB;
    }

    .paragraph-comparison-boxes {
      display: flex;
      flex-direction: row;
      gap: 16px;
      margin-bottom: 12px;
      width: 100%;
    }

    @media (max-width: 768px) {
      .paragraph-comparison-boxes {
        flex-direction: column;
      }
    }

    .paragraph-box {
      flex: 1 1 0;
      min-width: 0;
      border: 2px solid var(--border-color, #E5E7EB);
      border-radius: 8px;
      padding: 16px;
      background: white;
      min-height: 150px;
      display: flex;
      flex-direction: column;
    }

    .paragraph-box h5 {
      margin: 0 0 12px 0;
      font-size: 13px;
      font-weight: 600;
      color: #6B7280;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      flex-shrink: 0;
    }

    .paragraph-box-original {
      border-color: #E5E7EB;
    }

    .paragraph-box-original h5 {
      color: #6B7280;
    }

    .paragraph-box-edited {
      border-color: #D1D5DB;
      transition: border-color 0.2s ease, background-color 0.2s ease;
    }

    .paragraph-box-edited.approved-box {
      border-color: #10b981 !important;
      background: #F0FDF4 !important;
    }

    .paragraph-box-edited.declined-box {
      border-color: #EF4444 !important;
      background: #FEF2F2 !important;
    }

    .paragraph-box-edited h5 {
      color: #1F2937;
    }

    .paragraph-text-box {
      font-size: 14px;
      line-height: 1.6;
      color: var(--text-primary, #1F2937);
      white-space: pre-wrap;
      word-wrap: break-word;
      flex: 1;
      min-height: 50px;
    }

    .paragraph-text-box:empty::before {
      content: '(No content)';
      color: #9CA3AF;
      font-style: italic;
    }

    .no-content-placeholder {
      color: #9CA3AF;
      font-style: italic;
      display: block;
    }

    .paragraph-text-box.declined-text {
      text-decoration: line-through;
      opacity: 0.7;
    }

    .paragraph-tags {
      margin-top: 12px;
      padding-top: 12px;
      border-top: 1px solid var(--border-color, #E5E7EB);
      font-size: 12px;
    }

    .paragraph-tags strong {
      color: #6B7280;
      margin-right: 8px;
    }

    .tag-badge {
      display: inline-block;
      padding: 4px 10px;
      margin: 4px 4px 4px 0;
      background: #E0E7FF;
      color: #4338CA;
      border-radius: 12px;
      font-size: 11px;
      font-weight: 500;
    }

    .final-output-actions {
      margin-top: 24px;
      padding-top: 16px;
      border-top: 2px solid var(--border-color, #E5E7EB);
    }

    .final-output-btn {
      padding: 12px 24px;
      background-color: #D04A02;
      color: white;
      border: none;
      border-radius: 8px;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.2s ease;
      display: inline-flex;
      align-items: center;
      gap: 8px;
      text-align: center;
      text-decoration: none;
      -webkit-appearance: none;
      -moz-appearance: none;
      appearance: none;
      user-select: none;
      margin: 0;
      font-family: inherit;
    }

    .final-output-btn:hover:not(:disabled):not(.disabled) {
      background-color: #b83d01;
      transform: translateY(-1px);
      box-shadow: 0 4px 8px rgba(208, 74, 2, 0.3);
    }

    .final-output-btn:disabled,
    .final-output-btn.disabled {
      opacity: 0.6;
      cursor: not-allowed;
      pointer-events: none;
    }

    .final-output-btn:focus:not(:disabled):not(.disabled) {
      outline: 2px solid #D04A02;
      outline-offset: 2px;
    }

    .final-output-hint {
      margin-top: 12px;
      font-size: 13px;
      color: #6B7280;
      font-style: italic;
    }

    .spinner {
      display: inline-block;
      width: 14px;
      height: 14px;
      border: 2px solid rgba(255, 255, 255, 0.3);
      border-radius: 50%;
      border-top-color: white;
      animation: spin 0.6s linear infinite;
      margin-right: 8px;
    }

    @keyframes spin {
      to {
        transform: rotate(360deg);
      }
    }
  `]
})
export class ParagraphEditsConsolidatedComponent {
  @Input() paragraphEdits: ParagraphEdit[] = [];
  @Input() showFinalOutput: boolean = false;
  @Input() isGeneratingFinal: boolean = false;
  
  @Output() paragraphApproved = new EventEmitter<number>();
  @Output() paragraphDeclined = new EventEmitter<number>();
  @Output() generateFinal = new EventEmitter<void>();
  
  get allParagraphsDecided(): boolean {
    return allParagraphsDecided(this.paragraphEdits);
  }
  
  onApprove(index: number): void {
    if (index === undefined || index === null) {
      return;
    }
    
    const paragraph = this.paragraphEdits.find(p => p.index === index);
    if (!paragraph) {
      return;
    }
    
    this.paragraphApproved.emit(index);
  }
  
  onDecline(index: number): void {
    if (index === undefined || index === null) {
      return;
    }
    
    const paragraph = this.paragraphEdits.find(p => p.index === index);
    if (!paragraph) {
      return;
    }
    
    this.paragraphDeclined.emit(index);
  }
  
  onGenerateFinal(): void {
    this.generateFinal.emit();
  }
}

