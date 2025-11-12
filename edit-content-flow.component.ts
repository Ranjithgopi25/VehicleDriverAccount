import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { TlFlowService } from '../../../core/services/tl-flow.service';
import { ChatService } from '../../../core/services/chat.service';
import { GuidedDialogComponent } from '../../../shared/components/guided-dialog/guided-dialog.component';

type EditorType = 'brand-alignment' | 'copy' | 'line' | 'content' | 'development';

interface EditForm {
  selectedEditors: EditorType[];
  uploadedFile: File | null;
}

@Component({
  selector: 'app-edit-content-flow',
  standalone: true,
  imports: [CommonModule, FormsModule, GuidedDialogComponent],
  templateUrl: './edit-content-flow.component.html',
  styleUrls: ['./edit-content-flow.component.scss']
})
export class EditContentFlowComponent implements OnInit {
  showResults: boolean = false;
  isGenerating: boolean = false;
  editFeedback: string = '';
  revisedContent: string = '';  // HTML version for display
  revisedContentRaw: string = '';  // Raw markdown version for document generation
  originalContent: string = '';
  
  formData: EditForm = {
    selectedEditors: ['development', 'content', 'line', 'copy'],
    uploadedFile: null
  };
  
  fileReadError: string = '';

  editorTypes = [
    { 
      id: 'development' as EditorType, 
      name: 'Development Editor', 
      icon: '🚀', 
      description: 'Reviews and restructures content for alignment and coherence',
      details: 'Reviews: thought leadership quality, competitive differentiation, risk words (guarantee/promise/always), China terminology'
    },
    { 
      id: 'content' as EditorType, 
      name: 'Content Editor', 
      icon: '📄', 
      description: "Refines language to align with author's key objectives",
      details: 'Validates: mutually exclusive/collectively exhaustive structure, source citations, evidence quality, argument logic'
    },
    { 
      id: 'line' as EditorType, 
      name: 'Line Editor', 
      icon: '📝', 
      description: 'Improves sentence flow, readability and style preserving voice',
      details: 'Improves: active voice throughout, sentence length, precise word choice, paragraph structure, transitional phrases'
    },
    { 
      id: 'copy' as EditorType, 
      name: 'Copy Editor', 
      icon: '✏️', 
      description: 'Corrects grammar, punctuation and typos',
      details: 'Enforces: Oxford commas, apostrophes, em dashes, sentence case headlines, date formats, abbreviations, active voice'
    },
    { 
      id: 'brand-alignment' as EditorType, 
      name: 'PwC Brand Alignment Editor', 
      icon: '🎯', 
      description: 'Aligns content writing standards with PwC brand',
      details: 'Checks: we/you language, contractions, active voice, prohibited words (catalyst, PwC Network), China references, brand messaging'
    }
  ];

  constructor(
    public tlFlowService: TlFlowService,
    private chatService: ChatService
  ) {}

  ngOnInit(): void {}

  get isOpen(): boolean {
    return this.tlFlowService.currentFlow === 'edit-content';
  }

  onClose(): void {
    this.resetForm();
    this.tlFlowService.closeFlow();
  }

  resetForm(): void {
    this.showResults = false;
    this.isGenerating = false;
    this.editFeedback = '';
    this.revisedContent = '';
    this.revisedContentRaw = '';
    this.originalContent = '';
    this.fileReadError = '';
    this.formData = {
      selectedEditors: ['development', 'content', 'line', 'copy'],
      uploadedFile: null
    };
  }

  backToForm(): void {
    this.showResults = false;
    this.editFeedback = '';
    this.revisedContent = '';
    this.revisedContentRaw = '';
  }

  canEdit(): boolean {
    return this.formData.uploadedFile !== null && this.formData.selectedEditors.length > 0;
  }
  
  onFileSelect(event: any): void {
    const file = event.target.files?.[0];
    if (file) {
      this.formData.uploadedFile = file;
      this.fileReadError = '';
    }
  }

  removeFile(): void {
    this.formData.uploadedFile = null;
  }

  toggleEditor(type: EditorType): void {
    const index = this.formData.selectedEditors.indexOf(type);
    if (index > -1) {
      this.formData.selectedEditors.splice(index, 1);
    } else {
      this.formData.selectedEditors.push(type);
    }
  }

  isEditorSelected(type: EditorType): boolean {
    return this.formData.selectedEditors.includes(type);
  }

  getEditorNames(): string {
    if (this.formData.selectedEditors.length === 0) return '';
    if (this.formData.selectedEditors.length === 1) {
      const editor = this.editorTypes.find(e => e.id === this.formData.selectedEditors[0]);
      return editor ? editor.name : '';
    }
    return `${this.formData.selectedEditors.length} editors`;
  }

  async editContent(): Promise<void> {
    this.isGenerating = true;
    this.showResults = true;
    this.fileReadError = '';
    this.editFeedback = '';
    this.revisedContent = '';
    this.revisedContentRaw = '';
    
    let contentText = '';
    
    // Extract text from uploaded file
    if (this.formData.uploadedFile) {
      try {
        const extractedText = await this.extractFileText(this.formData.uploadedFile);
        contentText = extractedText;
        this.originalContent = contentText;
      } catch (error) {
        console.error('Error extracting file:', error);
        this.fileReadError = 'Error reading uploaded file. Please try again.';
        this.isGenerating = false;
        return;
      }
    }
    
    // Build message with selected editors
    const selectedEditorNames = this.formData.selectedEditors
      .map(id => this.editorTypes.find(e => e.id === id)?.name)
      .filter(name => name)
      .join(', ');
    
    const wordCount = contentText.split(/\s+/).filter(w => w.length > 0).length;
    
    let contentMessage = `You are a professional editorial service. Please analyze the following content using these editor types: ${selectedEditorNames}.

IMPORTANT INSTRUCTIONS:
1. Provide your response in TWO clear sections:
   - Section 1: FEEDBACK - Editorial comments, suggestions, and observations
   - Section 2: REVISED ARTICLE - The fully edited version of the content
   
2. For the REVISED ARTICLE section:
   - Apply all edits you recommended in the feedback
   - Maintain approximately the same length (~${wordCount} words)
   - Preserve the core message and structure
   - Make it publication-ready
   - Use proper markdown formatting: # for main headings, ## for subheadings, ### for sub-subheadings, - or * for bullet points, 1. for numbered lists
   - Ensure all headings use markdown format (e.g., # Heading Title, ## Subheading) so formatting is preserved

Format your response exactly like this:
=== FEEDBACK ===
[Your editorial feedback and suggestions here]

=== REVISED ARTICLE ===
[The fully revised and edited content here]

Content to Review:
${contentText}`;
    
    const messages = [{
      role: 'user' as const,
      content: contentMessage
    }];

    let fullResponse = '';

    // Pass selected editor types to backend
    this.chatService.streamEditContent(messages, this.formData.selectedEditors).subscribe({
      next: (data: any) => {
        let chunk = '';
        if (typeof data === 'string') {
          chunk = data;
        } else if (data.type === 'content' && data.content) {
          chunk = data.content;
        }
        
        fullResponse += chunk;
        this.parseEditResponse(fullResponse);
      },
      error: (error: any) => {
        console.error('Error editing content:', error);
        this.editFeedback = 'Sorry, there was an error editing your content. Please try again.';
        this.isGenerating = false;
      },
      complete: () => {
        this.parseEditResponse(fullResponse);
        this.isGenerating = false;
      }
    });
  }

  private parseEditResponse(response: string): void {
    const feedbackMatch = response.match(/===\s*FEEDBACK\s*===\s*([\s\S]*?)(?====\s*REVISED ARTICLE\s*===|$)/i);
    const revisedMatch = response.match(/===\s*REVISED ARTICLE\s*===\s*([\s\S]*?)$/i);
    
    if (feedbackMatch && feedbackMatch[1]) {
      this.editFeedback = this.formatMarkdown(feedbackMatch[1].trim());
    } else if (!revisedMatch) {
      this.editFeedback = this.formatMarkdown(response);
    }
    
    if (revisedMatch && revisedMatch[1]) {
      // Store raw markdown content for document generation (preserves #, ##, -, etc.)
      this.revisedContentRaw = revisedMatch[1].trim();
      // Store HTML version for display
      this.revisedContent = this.revisedContentRaw.replace(/\n/g, '<br>');
    }
  }

  private formatMarkdown(text: string): string {
    // Convert **text** to <strong>text</strong>
    let formatted = text.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
    // Convert *text* to <em>text</em>
    formatted = formatted.replace(/\*(.+?)\*/g, '<em>$1</em>');
    // Convert line breaks
    formatted = formatted.replace(/\n/g, '<br>');
    return formatted;
  }
  
  private async extractFileText(file: File): Promise<string> {
    const formData = new FormData();
    formData.append('file', file);
    
    const response = await fetch('/api/extract-text', {
      method: 'POST',
      body: formData
    });
    
    if (!response.ok) {
      throw new Error('Failed to extract text from file');
    }
    
    const data = await response.json();
    return data.text || '';
  }

  async downloadSection(section: 'feedback' | 'revised', format: 'docx' | 'pdf'): Promise<void> {
    const filename = section === 'feedback' ? 'editorial-feedback' : 'revised-article';
    
    // For revised content, use raw markdown to preserve formatting (#, ##, -, etc.)
    // For feedback, strip HTML tags
    let plainText: string;
    if (section === 'revised' && this.revisedContentRaw) {
      plainText = this.revisedContentRaw;  // Use raw markdown content
    } else {
      const content = section === 'feedback' ? this.editFeedback : this.revisedContent;
      plainText = content.replace(/<br>/g, '\n').replace(/<[^>]+>/g, '');
    }
    
    if (format === 'docx') {
      try {
        const response = await fetch('/api/generate-document', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            content: plainText,
            format: 'docx',
            filename: filename
          })
        });
        
        if (!response.ok) throw new Error('Failed to generate document');
        
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${filename}.docx`;
        link.click();
        window.URL.revokeObjectURL(url);
      } catch (error) {
        console.error('Error generating DOCX:', error);
        alert('Failed to generate DOCX file. Please try again.');
      }
    } else if (format === 'pdf') {
      try {
        const response = await fetch('/api/generate-document', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            content: plainText,
            format: 'pdf',
            filename: filename
          })
        });
        
        if (!response.ok) throw new Error('Failed to generate document');
        
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${filename}.pdf`;
        link.click();
        window.URL.revokeObjectURL(url);
      } catch (error) {
        console.error('Error generating PDF:', error);
        alert('Failed to generate PDF file. Please try again.');
      }
    }
  }

  copyFeedback(): void {
    const plainText = this.editFeedback.replace(/<br>/g, '\n').replace(/<[^>]+>/g, '');
    navigator.clipboard.writeText(plainText);
  }

  copyRevised(): void {
    // Use raw markdown content if available, otherwise strip HTML
    const plainText = this.revisedContentRaw || this.revisedContent.replace(/<br>/g, '\n').replace(/<[^>]+>/g, '');
    navigator.clipboard.writeText(plainText);
  }
}
