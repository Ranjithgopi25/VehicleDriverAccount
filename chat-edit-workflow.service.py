import { Injectable, inject } from '@angular/core';
import { DomSanitizer, SafeHtml } from '@angular/platform-browser';
import { BehaviorSubject, Observable, Subject, firstValueFrom } from 'rxjs';
import { Message, EditWorkflowMetadata, ParagraphEdit } from '../models';
import { ChatService } from './chat.service';
import { normalizeEditorOrder, normalizeContent, extractDocumentTitle, getEditorDisplayName, formatMarkdown, convertMarkdownToHtml, extractFileText } from '../utils/edit-content.utils';
import { 
  splitIntoParagraphs, 
  createParagraphEditsFromComparison, 
  allParagraphsDecided 
} from '../utils/paragraph-edit.utils';

export type EditWorkflowStep = 'idle' | 'awaiting_editors' | 'awaiting_content' | 'processing' | 'awaiting_approval';

export interface EditWorkflowState {
  step: EditWorkflowStep;
  uploadedFile: File | null;
  selectedEditors: string[];
  originalContent: string;
  paragraphEdits: ParagraphEdit[];
}

export interface EditWorkflowMessage {
  type: 'prompt' | 'result' | 'update';
  message: Message;
  metadata?: any;
}

export interface EditorOption {
  id: string;
  name: string;
  icon: string;
  description: string;
  selected: boolean;
  disabled?: boolean;
  alwaysSelected?: boolean;
}

@Injectable({
  providedIn: 'root'
})
export class ChatEditWorkflowService {
  private chatService = inject(ChatService);
  private sanitizer = inject(DomSanitizer);

  private stateSubject = new BehaviorSubject<EditWorkflowState>({
    step: 'idle',
    uploadedFile: null,
    selectedEditors: ['brand-alignment'],
    originalContent: '',
    paragraphEdits: []
  });

  public state$: Observable<EditWorkflowState> = this.stateSubject.asObservable();

  private messageSubject = new Subject<EditWorkflowMessage>();
  public message$: Observable<EditWorkflowMessage> = this.messageSubject.asObservable();

  private workflowCompletedSubject = new Subject<void>();
  public workflowCompleted$: Observable<void> = this.workflowCompletedSubject.asObservable();

  private workflowStartedSubject = new Subject<void>();
  public workflowStarted$: Observable<void> = this.workflowStartedSubject.asObservable();

  // Track final article generation state
  private isGeneratingFinalSubject = new BehaviorSubject<boolean>(false);
  public isGeneratingFinal$: Observable<boolean> = this.isGeneratingFinalSubject.asObservable();
  public get isGeneratingFinal(): boolean {
    return this.isGeneratingFinalSubject.value;
  }

  readonly editorOptions: EditorOption[] = [
    { 
      id: 'development', 
      name: 'Development Editor', 
      icon: '🚀', 
      description: 'Reviews and restructures content for alignment and coherence',
      selected: false
    },
    { 
      id: 'content', 
      name: 'Content Editor', 
      icon: '📄', 
      description: "Refines language to align with the author's objectives",
      selected: false
    },
    { 
      id: 'line', 
      name: 'Line Editor', 
      icon: '📝', 
      description: 'Improves sentence flow, readability and style preserving voice',
      selected: false
    },
    { 
      id: 'copy', 
      name: 'Copy Editor', 
      icon: '✏️', 
      description: 'Corrects grammar, punctuation and typos',
      selected: false
    },
    { 
      id: 'brand-alignment', 
      name: 'PwC Brand Alignment Editor', 
      icon: '🎯', 
      description: 'Aligns content writing standards with PwC brand',
      selected: true
    }
  ];

  get currentState(): EditWorkflowState {
    return this.stateSubject.value;
  }

  get isActive(): boolean {
    return this.currentState.step !== 'idle';
  }

  /** Detect edit intent using LLM agent via backend API */
  async detectEditIntent(input: string): Promise<{hasEditIntent: boolean, detectedEditors?: string[]}> {
    if (!input || !input.trim()) {
      return { hasEditIntent: false };
    }

    try {
      const result = await firstValueFrom(
        this.chatService.detectEditIntent(input.trim())
      );
      
      const hasEditIntent = result.is_edit_intent && result.confidence >= 0.7;
      const detectedEditors = result.detected_editors && result.detected_editors.length > 0 
        ? result.detected_editors 
        : undefined;
      
      return { 
        hasEditIntent, 
        detectedEditors 
      };
    } catch (error) {
      console.error('Error in LLM intent detection:', error);
      return { hasEditIntent: false };
    }
  }

  beginWorkflow(): void {
    const defaultState = this.getDefaultState();
    this.updateState({
      ...defaultState,
      step: 'awaiting_editors'
    });

    // Emit workflow started event to clear previous state
    this.workflowStartedSubject.next();

    const promptMessage = this.createEditorSelectionMessage(
      `I'll help you edit your content! 📝\n\n**Select the editing services you'd like to use:**`
    );

    this.messageSubject.next({
      type: 'prompt',
      message: promptMessage
    });
  }

  /** Begin workflow with pre-selected editors (Path 1: Direct Editor Detection) */
  beginWorkflowWithEditors(editorIds: string[]): void {
    if (!editorIds || editorIds.length === 0) {
      this.beginWorkflow();
      return;
    }

    const validEditorIds = this.editorOptions.map(e => e.id);
    const validatedEditors = editorIds.filter(id => validEditorIds.includes(id));

    if (validatedEditors.length === 0) {
      this.beginWorkflow();
      return;
    }

    const editorsWithBrand = [...validatedEditors];
    if (!editorsWithBrand.includes('brand-alignment')) {
      editorsWithBrand.push('brand-alignment');
    }

    const defaultState = this.getDefaultState();
    this.updateState({
      ...defaultState,
      step: 'awaiting_content',
      selectedEditors: editorsWithBrand
    });

    // Emit workflow started event to clear previous state
    this.workflowStartedSubject.next();

    const editorNamesText = this.getSelectedEditorNames(validatedEditors);

    const editWorkflowMetadata: EditWorkflowMetadata = {
      step: 'awaiting_content',
      showFileUpload: true,
      showCancelButton: false,
      showSimpleCancelButton: true
    };

    const contentRequestMessage: Message = {
      role: 'assistant',
      content: `✅ **Using ${editorNamesText} to edit your content**\n\n**Now, please upload your document:**`,
      timestamp: new Date(),
      editWorkflow: editWorkflowMetadata
    };

    this.messageSubject.next({
      type: 'prompt',
      message: contentRequestMessage
    });
  }

  /** Get editor names from editor IDs */
  private getEditorNamesFromIds(editorIds: string[]): string[] {
    return editorIds
      .map(id => this.editorOptions.find(e => e.id === id)?.name)
      .filter((name): name is string => !!name);
  }

  /** Get selected editor names as a formatted string */
  private getSelectedEditorNames(editorIds: string[]): string {
    const names = this.getEditorNamesFromIds(editorIds);
    if (names.length === 0) {
      return '';
    }
    if (names.length === 1) {
      return names[0];
    }
    if (names.length === 2) {
      return names.join(' and ');
    }
    return names.slice(0, -1).join(', ') + ', and ' + names[names.length - 1];
  }

  private getNumberedEditorList(editorOptions?: EditorOption[]): string {
    const editors = editorOptions || this.editorOptions;
    const currentSelectedIds = this.currentState.selectedEditors;
    
    return editors.map((editor, index) => {
      const num = index + 1;
      const isSelected = currentSelectedIds.includes(editor.id);
      const selected = isSelected ? ' ✓' : '';
      return `${num}. **${editor.name}** — ${editor.description}${selected}`;
    }).join('\n');
  }

  handleFileUpload(file: File): void {
    if (this.currentState.step !== 'awaiting_content') {
      return;
    }

    this.updateState({
      ...this.currentState,
      uploadedFile: file
    });
    
    this.processWithContent();
  }

  async handleChatInput(input: string, file?: File): Promise<void> {
    const trimmedInput = input.trim();
    const workflowActive = this.isActive;

    if (!workflowActive) {
      const intentResult = await this.detectEditIntent(trimmedInput);
      if (intentResult.hasEditIntent) {
        // Path 1: Direct Editor Detection - editors detected
        if (intentResult.detectedEditors && intentResult.detectedEditors.length > 0) {
          this.beginWorkflowWithEditors(intentResult.detectedEditors);
        } else {
          // Path 2: Standard Flow - show editor selection
          this.beginWorkflow();
        }
      return;
      }
    }

    if (!workflowActive) {
      return;
    }

    if (this.currentState.step === 'awaiting_editors') {
      if (trimmedInput) {
        const lowerInput = trimmedInput.toLowerCase();
        if (lowerInput.includes('proceed') || lowerInput.includes('continue') || lowerInput.includes('yes') || lowerInput === 'ok' || lowerInput === 'done') {
          this.proceedToContentStep();
          return;
        }
        
        if (lowerInput.includes('cancel')) {
          this.cancelWorkflow();
          return;
        }
        
        const selectionResult = this.parseNumericSelection(trimmedInput);
        if (selectionResult.selectedIndices.length > 0 || selectionResult.hasInput) {
          this.handleNumericSelection(selectionResult);
          return;
        }
        
        if (trimmedInput.trim().length > 0) {
          this.showInvalidSelectionError();
          return;
        }
      }
      return;
    }

    if (this.currentState.step === 'awaiting_content') {
      if (file) {
        this.handleFileUpload(file);
        return;
      }
      
      if (trimmedInput) {
        const errorMessage: Message = {
          role: 'assistant',
          content: '⚠️ **Please upload a document file** (Word, PDF, Text, or Markdown). Text pasting is not available in this workflow.',
          timestamp: new Date(),
          editWorkflow: {
            step: 'awaiting_content',
            showCancelButton: false,
            showSimpleCancelButton: true
          }
        };
        this.messageSubject.next({ type: 'prompt', message: errorMessage });
        return;
      }
    }
  }

  private parseNumericSelection(input: string): { selectedIndices: number[], invalidIndices: number[], hasInput: boolean } {
    const selectedIndices: number[] = [];
    const invalidIndices: number[] = [];
    let hasInput = false;
    
    const cleanedInput = input.replace(/(?:select|choose|pick|use|want|need|editor|editors)/gi, '').trim();
    
    if (!/\d/.test(cleanedInput)) {
      return { selectedIndices: [], invalidIndices: [], hasInput: cleanedInput.length > 0 };
    }
    
    hasInput = true;
    const parts = cleanedInput.split(/[,;\s]+/).filter(part => part.trim().length > 0);
    
    for (const part of parts) {
      const trimmedPart = part.trim();
      if (!trimmedPart) continue;
      
      const rangeMatch = trimmedPart.match(/^(\d+)\s*-\s*(\d+)$/);
      if (rangeMatch) {
        const start = parseInt(rangeMatch[1]);
        const end = parseInt(rangeMatch[2]);
        
        if (start > end) {
          continue;
        }
        
        for (let i = start; i <= end; i++) {
          if (i >= 1 && i <= 5) {
            if (!selectedIndices.includes(i)) {
              selectedIndices.push(i);
            }
          } else {
            if (!invalidIndices.includes(i)) {
              invalidIndices.push(i);
            }
          }
        }
        continue;
      }
      
      const numberMatch = trimmedPart.match(/^(\d+)$/);
      if (numberMatch) {
        const num = parseInt(numberMatch[1]);
        if (num >= 1 && num <= 5) {
          if (!selectedIndices.includes(num)) {
            selectedIndices.push(num);
          }
        } else {
          if (!invalidIndices.includes(num)) {
            invalidIndices.push(num);
          }
        }
        continue;
      }
    }
    
    selectedIndices.sort((a, b) => a - b);
    invalidIndices.sort((a, b) => a - b);
    
    return { selectedIndices, invalidIndices, hasInput };
  }

  private handleNumericSelection(result: { selectedIndices: number[], invalidIndices: number[], hasInput: boolean }): void {
    if (result.invalidIndices.length > 0) {
      const editorList = this.getNumberedEditorList();
      const errorMessage = this.createEditorSelectionMessage(
        `⚠️ **Invalid editor number(s):** ${result.invalidIndices.join(', ')}\n\n**Valid editor numbers are 1-5.**\n\n**Editor List:**\n\n${editorList}\n\nPlease provide valid editor numbers (1-5) or type "proceed" to continue with defaults.`
      );
      this.messageSubject.next({ type: 'prompt', message: errorMessage });
      return;
    }
    
    if (result.selectedIndices.length === 0 && result.hasInput) {
      this.showInvalidSelectionError();
      return;
    }
    
    if (result.selectedIndices.length > 0) {
      const updatedEditors = this.editorOptions.map((editor, index) => {
        const editorNum = index + 1;
        return {
          ...editor,
          selected: result.selectedIndices.includes(editorNum)
        };
      });
      
      const selectedIds = updatedEditors.filter(e => e.selected).map(e => e.id);
      
      // Ensure brand-alignment is always included
      if (!selectedIds.includes('brand-alignment')) {
        selectedIds.push('brand-alignment');
      }
      
      this.updateState({
        ...this.currentState,
        selectedEditors: selectedIds
      });
      
      const selectedNames = updatedEditors
        .filter(e => e.selected)
        .map((e, idx) => {
          const num = this.editorOptions.findIndex(opt => opt.id === e.id) + 1;
          return `${num}. ${e.name}`;
        })
        .join(', ');
      
      const confirmMessage = this.createEditorSelectionMessage(
        `✅ **Selected editors:** ${selectedNames}\n\nType "proceed" to continue or select different editors.`,
        updatedEditors
      );
      
      this.messageSubject.next({ type: 'prompt', message: confirmMessage });
    }
  }

  private showInvalidSelectionError(): void {
    const editorList = this.getNumberedEditorList();
    const errorMessage = this.createEditorSelectionMessage(
      `⚠️ **Please provide valid editor numbers (1-5).**\n\n**Editor List:**\n\n${editorList}\n\nOr type "proceed" to continue with defaults.`
    );
    this.messageSubject.next({ type: 'prompt', message: errorMessage });
  }

  private parseOptOutInput(input: string): { optedOut: number[], sections: string[] } {
    const lowerInput = input.toLowerCase();
    const optedOut: number[] = [];
    const sections: string[] = [];
    
    const optOutPattern = /(?:remove|skip|exclude|without|opt\s*out|deselect|don't\s*use|do\s*not\s*use)\s+(\d+(?:\s*[,\s]?\s*(?:and\s*)?\d+)*)/gi;
    
    let match;
    while ((match = optOutPattern.exec(lowerInput)) !== null) {
      const numbersStr = match[1];
      const numberMatches = numbersStr.match(/\d+/g);
      if (numberMatches) {
        numberMatches.forEach(numStr => {
          const num = parseInt(numStr);
          if (num >= 1 && num <= 5 && !optedOut.includes(num)) {
            optedOut.push(num);
          }
        });
      }
    }
    
    const sectionPatterns = [
      /(?:edit|review|focus\s*on)\s+(?:pages?|sections?)\s+(\d+(?:\s*-\s*\d+)?)/gi,
      /(?:edit|review)\s+(?:the\s+)?(introduction|conclusion|summary|abstract|body|content)/gi
    ];
    
    sectionPatterns.forEach(pattern => {
      let match;
      while ((match = pattern.exec(input)) !== null) {
        if (match[1] && !sections.includes(match[1])) {
          sections.push(match[1]);
        }
      }
    });
    
    return { optedOut, sections };
  }

  private handleOptOutAndProceed(result: { optedOut: number[], sections: string[] }): void {
    const currentEditors = [...this.editorOptions];
    // Find brand-alignment editor index to prevent it from being opted out
    const brandAlignmentIndex = currentEditors.findIndex(e => e.id === 'brand-alignment');
    const brandAlignmentNum = brandAlignmentIndex >= 0 ? brandAlignmentIndex + 1 : -1;
    
    const selectedEditors = currentEditors.map((editor, index) => {
      const editorNum = index + 1;
      // Brand alignment is always selected, cannot be opted out
      if (editor.id === 'brand-alignment') {
        return {
          ...editor,
          selected: true
        };
      }
      return {
        ...editor,
        selected: !result.optedOut.includes(editorNum)
      };
    });
    
    const selectedIds = selectedEditors.filter(e => e.selected).map(e => e.id);
    
    // Ensure brand-alignment is always included
    if (!selectedIds.includes('brand-alignment')) {
      selectedIds.push('brand-alignment');
    }
    
    this.updateState({
      ...this.currentState,
      selectedEditors: selectedIds
    });
    
    let responseMessage = '';
    if (result.optedOut.length > 0) {
      const optedOutNames = result.optedOut.map(num => {
        const editor = this.editorOptions[num - 1];
        return `${num}. ${editor.name}`;
      }).join(', ');
      responseMessage += `✅ **Opted out:** ${optedOutNames}\n\n`;
    }
    
    if (result.sections.length > 0) {
      responseMessage += `📄 **Sections to edit:** ${result.sections.join(', ')}\n\n`;
    }
    
    const remainingEditors = selectedEditors.filter(e => e.selected);
    if (remainingEditors.length === 0) {
      responseMessage += `⚠️ **No editors selected.** Please keep at least one editor active.`;
      const errorMessage = this.createEditorSelectionMessage(responseMessage, selectedEditors);
      this.messageSubject.next({ type: 'prompt', message: errorMessage });
      return;
    }
    
    responseMessage += `**Selected ${remainingEditors.length} editor${remainingEditors.length > 1 ? 's' : ''}:** ${remainingEditors.map(e => e.name).join(', ')}\n\nWhen you're ready, click "Continue" or type "proceed" to move to the next step.`;
    
    const confirmMessage = this.createEditorSelectionMessage(responseMessage, selectedEditors);
    
    this.messageSubject.next({
      type: 'prompt',
      message: confirmMessage
    });
  }

  private proceedToContentStep(): void {
    // Ensure brand-alignment is always included
    const selectedIds = [...this.currentState.selectedEditors];
    if (!selectedIds.includes('brand-alignment')) {
      selectedIds.push('brand-alignment');
    }
    
    if (selectedIds.length === 0) {
      this.createNoEditorsErrorMessage();
      return;
    }
    
    // Update state to ensure brand-alignment is included
    this.updateState({
      ...this.currentState,
      selectedEditors: selectedIds
    });

    this.updateState({
      ...this.currentState,
      step: 'awaiting_content'
    });

    const editorNamesText = this.getSelectedEditorNames(selectedIds);

    const editWorkflowMetadata: EditWorkflowMetadata = {
      step: 'awaiting_content',
      showFileUpload: true,  // Show file upload component
      showCancelButton: false,
      showSimpleCancelButton: true
    };

    const contentRequestMessage: Message = {
      role: 'assistant',
      content: `✅ **Using ${editorNamesText} to edit your content**\n\n**Now, please upload your document:**`,
      timestamp: new Date(),
      editWorkflow: editWorkflowMetadata
    };

    this.messageSubject.next({
      type: 'prompt',
      message: contentRequestMessage
    });
  }

  private async processWithContent(): Promise<void> {
    // Ensure brand-alignment is always included
    const selectedIds = [...this.currentState.selectedEditors];
    if (!selectedIds.includes('brand-alignment')) {
      selectedIds.push('brand-alignment');
    }
    const selectedNames = this.getSelectedEditorNames(selectedIds);

    try {
      let contentText = this.currentState.originalContent;
      
      if (this.currentState.uploadedFile && !contentText) {
        contentText = await extractFileText(this.currentState.uploadedFile);
        contentText = normalizeContent(contentText);
        this.updateState({
          ...this.currentState,
          originalContent: contentText
        });
      }

      if (!contentText || !contentText.trim()) {
        throw new Error('No content to process');
      }

      this.updateState({
        ...this.currentState,
        step: 'processing'
      });

      const processingMessage: Message = {
        role: 'assistant',
        content: `Processing your content with: **${selectedNames}**\n\nPlease wait while I analyze and edit your content...`,
        timestamp: new Date(),
      editWorkflow: {
        step: 'processing',
        showCancelButton: false
      }
      };

      this.messageSubject.next({
        type: 'prompt',
        message: processingMessage
      });

      await this.processContent(contentText, selectedIds, selectedNames);
    } catch (error) {
      const errorMessage: Message = {
        role: 'assistant',
        content: 'Sorry, there was an error processing your content. Please try again.',
        timestamp: new Date()
      };
      this.messageSubject.next({ type: 'result', message: errorMessage });
      this.completeWorkflow();
    }
  }

  handleEditorSelection(selectedIds: string[]): void {
    if (this.currentState.step !== 'awaiting_editors') {
      return;
    }

    // Ensure brand-alignment is always included
    const editorsWithBrand = [...selectedIds];
    if (!editorsWithBrand.includes('brand-alignment')) {
      editorsWithBrand.push('brand-alignment');
    }

    if (editorsWithBrand.length === 0) {
      this.createNoEditorsErrorMessage();
      return;
    }

    this.updateState({
      ...this.currentState,
      selectedEditors: editorsWithBrand
    });

    this.proceedToContentStep();
  }

  private async processContent(contentText: string, selectedIds: string[], selectedNames: string): Promise<void> {
    const messages = [{
      role: 'user' as const,
      content: contentText
    }];

    const normalizedEditorIds = normalizeEditorOrder(selectedIds);

    let fullResponse = '';
    let combinedFeedback = '';
    let finalRevisedContent = '';
    let currentEditorProgress: {current: number, total: number, currentEditor: string} | null = null;
    let editorErrors: Array<{editor: string, error: string}> = [];
    
    const editorProgressList: Array<{editorId: string, editorName: string, status: 'pending' | 'processing' | 'completed' | 'error', current?: number, total?: number}> = normalizedEditorIds.map((id, index) => ({
      editorId: id,
      editorName: getEditorDisplayName(id),
      status: 'pending' as const,
      current: index + 1,
      total: normalizedEditorIds.length
    }));

    // Use default temperature (0.15) - optimal for editing: allows minor improvements while staying deterministic
    this.chatService.streamEditContent(messages, normalizedEditorIds).subscribe({
      next: (data: any) => {
        if (data.type === 'editor_progress') {
          currentEditorProgress = {
            current: data.current || 0,
            total: data.total || 0,
            currentEditor: data.editor || ''
          };
          
          const currentIndex = data.current || 0;
          editorProgressList.forEach((editor, index) => {
            const editorIndex = index + 1;
            if (editorIndex < currentIndex) {
              editor.status = 'completed';
            } else if (editorIndex === currentIndex) {
              editor.status = 'processing';
              editor.current = currentIndex;
              editor.total = data.total || selectedIds.length;
            } else {
              editor.status = 'pending';
            }
          });
          
          const progressMessage: Message = {
            role: 'assistant',
            content: '',
            timestamp: new Date(),
            editWorkflow: {
              step: 'processing',
              showCancelButton: false,
              editorProgress: currentEditorProgress || undefined,
              editorProgressList: [...editorProgressList]
            }
          };
          this.messageSubject.next({ type: 'prompt', message: progressMessage });
        } else if (data.type === 'editor_content') {
          if (data.content) {
            fullResponse += data.content;
          }
        } else if (data.type === 'editor_complete') {
          const completedEditor = editorProgressList.find(e => e.editorId === data.editor);
          if (completedEditor) {
            completedEditor.status = 'completed';
          }
          
          if (data.revised_content) {
            fullResponse = data.revised_content;
          }
          
          const progressMessage: Message = {
            role: 'assistant',
            content: '',
            timestamp: new Date(),
            editWorkflow: {
              step: 'processing',
              showCancelButton: false,
              editorProgress: currentEditorProgress || undefined,
              editorProgressList: [...editorProgressList]
            }
          };
          this.messageSubject.next({ type: 'prompt', message: progressMessage });
        } else if (data.type === 'editor_error') {
          const errorEditor = editorProgressList.find(e => e.editorId === data.editor);
          if (errorEditor) {
            errorEditor.status = 'error';
          }
          
          editorErrors.push({
            editor: data.editor || 'Unknown',
            error: data.error || 'Unknown error'
          });
          
          const editorName = getEditorDisplayName(data.editor);
          const errorMessage: Message = {
            role: 'assistant',
            content: `⚠️ **${editorName} encountered an error:** ${data.error}\n\nContinuing with remaining editors...`,
            timestamp: new Date(),
            editWorkflow: {
              step: 'processing',
              showCancelButton: false,
              editorProgress: currentEditorProgress || undefined,
              editorProgressList: [...editorProgressList]
            }
          };
          this.messageSubject.next({ type: 'prompt', message: errorMessage });
        } else if (data.type === 'final_complete') {
          combinedFeedback = data.combined_feedback || '';
          finalRevisedContent = data.final_revised || '';
          
          let paragraphEdits: ParagraphEdit[] = [];
          if (data.paragraph_edits && Array.isArray(data.paragraph_edits)) {
            const allEditorNames = selectedIds.map(editorId => {
              return getEditorDisplayName(editorId);
            });
            
            // Get original content - prioritize data.original_content, then currentState
            const originalContent = data.original_content || this.currentState.originalContent || '';
            const originalParagraphs = originalContent ? splitIntoParagraphs(originalContent) : [];
            
            paragraphEdits = data.paragraph_edits.map((edit: any, arrayIndex: number) => {
              const existingTags = edit.tags || [];
              
              const existingEditorNames = new Set<string>(
                existingTags.map((tag: string) => {
                  const match = tag.match(/^(.+?)\s*\(/);
                  return match ? match[1].trim() : tag;
                })
              );
              
              const allTags = [...existingTags];
              allEditorNames.forEach(editorName => {
                const existingNamesArray = Array.from(existingEditorNames) as string[];
                if (!existingNamesArray.some((existing: string) => 
                  existing.toLowerCase().includes(editorName.toLowerCase()) || 
                  editorName.toLowerCase().includes(existing.toLowerCase())
                )) {
                  allTags.push(`${editorName} (Reviewed)`);
                }
              });
              
              // Use edit.index if provided, otherwise use array index
              // This ensures each paragraph has a unique index
              const paragraphIndex = (edit.index !== undefined && edit.index !== null) ? edit.index : arrayIndex;
              
              // Get original text - prioritize edit.original, then try to get from original content by index
              let originalText = (edit.original && edit.original.trim()) || '';
              if (!originalText && originalParagraphs.length > paragraphIndex && paragraphIndex >= 0) {
                originalText = (originalParagraphs[paragraphIndex] && originalParagraphs[paragraphIndex].trim()) || '';
              }
              
              // Ensure edited text is available
              const editedText = (edit.edited && edit.edited.trim()) || '';
              
              return {
                index: paragraphIndex,
                original: originalText,
                edited: editedText,
                tags: allTags,
                approved: null as boolean | null
              };
            });
          } else if (data.final_revised && data.original_content) {
            paragraphEdits = this.createParagraphEditsFromComparison(
              data.original_content,
              data.final_revised,
              selectedIds
            );
          }
          
          // Ensure originalContent is preserved - prioritize data.original_content, but keep existing if not provided
          const preservedOriginalContent = data.original_content || this.currentState.originalContent || '';
          
          this.updateState({
            ...this.currentState,
            paragraphEdits: paragraphEdits,
            originalContent: preservedOriginalContent
          });
          
          editorProgressList.forEach(editor => {
            if (editor.status !== 'error') {
              editor.status = 'completed';
            }
          });
          
          const completedProgress: {current: number, total: number, currentEditor: string} = {
            current: currentEditorProgress?.total || editorProgressList.length,
            total: currentEditorProgress?.total || editorProgressList.length,
            currentEditor: 'completed'
          };
          
          const completionMessage: Message = {
            role: 'assistant',
            content: '',
            timestamp: new Date(),
            editWorkflow: {
              step: 'processing',
              showCancelButton: false,
              editorProgress: completedProgress,
              editorProgressList: [...editorProgressList]
            }
          };
          this.messageSubject.next({ type: 'prompt', message: completionMessage });
          
          if (editorErrors.length > 0) {
            const errorSummary = editorErrors.map(e => {
              const editorName = getEditorDisplayName(e.editor);
              return `⚠️ ${editorName} encountered an error: ${e.error}. Processing continued with previous editor's output.`;
            }).join('\n\n');
            
            if (combinedFeedback) {
              combinedFeedback = errorSummary + '\n\n' + combinedFeedback;
            } else {
              combinedFeedback = errorSummary;
            }
          }
          
          this.dispatchResultsToChat('', selectedIds, selectedNames, combinedFeedback, finalRevisedContent, paragraphEdits);
        } else if (data.type === 'content' && data.content) {
          fullResponse += data.content;
        } else if (typeof data === 'string') {
          fullResponse += data;
        }
      },
      error: (error: any) => {
        const errorMsg: Message = {
          role: 'assistant',
          content: 'Sorry, there was an error editing your content. Please try again.',
          timestamp: new Date()
        };
        this.messageSubject.next({ type: 'result', message: errorMsg });
        this.completeWorkflow();
      },
      complete: () => {
        this.completeWorkflow();
      }
    });
  }

  /** Create paragraph edits by comparing original and edited content */
  private createParagraphEditsFromComparison(original: string, edited: string, editorIds?: string[]): ParagraphEdit[] {
    const editorIdsToUse = editorIds || this.currentState.selectedEditors;
    const allEditorNames = editorIdsToUse.map(editorId => {
      return getEditorDisplayName(editorId);
    });
    
    return createParagraphEditsFromComparison(original, edited, allEditorNames);
  }
  
  private dispatchResultsToChat(
    rawResponse: string,
    selectedEditorIds: string[],
    selectedEditorNames: string,
    combinedFeedback?: string,
    finalRevisedContent?: string,
    paragraphEdits?: ParagraphEdit[],
    extractedTitle?: string
  ): void {
    let feedbackMatch: RegExpMatchArray | null = null;
    if (combinedFeedback) {
      feedbackMatch = [null, combinedFeedback] as any;
    } else {
      feedbackMatch = rawResponse.match(/===\s*FEEDBACK\s*===\s*([\s\S]*?)(?====\s*REVISED ARTICLE\s*===|$)/i);
    }
    
    let revisedContent = '';
    if (finalRevisedContent && finalRevisedContent.trim()) {
      revisedContent = finalRevisedContent.trim();
    }
    const uploadedFileName = this.currentState.uploadedFile?.name;
    
    // Extract title from original content (use provided extractedTitle or extract from content)
    const documentTitle = extractedTitle || extractDocumentTitle(
      this.currentState.originalContent || '',
      uploadedFileName
    );
    const cleanTopic = documentTitle.trim() || 'Revised Article';
    
    let cleanFullContent = revisedContent || 'No revised article returned.';
    cleanFullContent = cleanFullContent.replace(/^```[\w]*\n?/gm, '').replace(/\n?```$/gm, '').trim();
    
    const metadata = {
      contentType: 'article' as const,
      topic: cleanTopic,
      fullContent: cleanFullContent,
      showActions: !!revisedContent && cleanFullContent.length > 0
    };
    
    // Send editorial feedback FIRST (matches Guided Journey display order)
    if (feedbackMatch && feedbackMatch[1]) {
      const feedbackPlainText = feedbackMatch[1].trim();
      const feedbackTitle = '**📝 Editorial Feedback**';
      const feedbackContent = feedbackPlainText;
      const combinedFeedback = `${feedbackTitle}\n\n${feedbackContent}`;
      const feedbackHtml = formatMarkdown(combinedFeedback);
      
      const feedbackMessage: Message = {
        role: 'assistant',
        content: feedbackHtml,
        timestamp: new Date(),
        isHtml: true, // Flag to indicate content is already HTML
        thoughtLeadership: {
          contentType: 'article',
          topic: 'Editorial Feedback',
          fullContent: feedbackPlainText,
          showActions: true
        }
      };
      this.messageSubject.next({ type: 'result', message: feedbackMessage });
    }
    
    // Send paragraph-by-paragraph comparison AFTER editorial feedback (matches Guided Journey display order)
    if (paragraphEdits && paragraphEdits.length > 0) {
      const paragraphMessage: Message = {
        role: 'assistant',
        content: '', // Content will be rendered by Angular component
        timestamp: new Date(),
        isHtml: false,
        editWorkflow: {
          step: 'awaiting_approval',
          paragraphEdits: paragraphEdits,
          showCancelButton: false,
          showSimpleCancelButton: true
        }
      };
      this.messageSubject.next({ type: 'result', message: paragraphMessage });
    } else if (revisedContent && !paragraphEdits) {
      const headerLines: string[] = [
        '### Quick Start Thought Leadership – Edit Content'
      ];
      
      if (uploadedFileName) {
        headerLines.push(`_Source: ${uploadedFileName}_`);
      }
      
      if (selectedEditorNames) {
        headerLines.push(`_Editors Applied: ${selectedEditorNames}_`);
      }
      
      headerLines.push('', '**Revised Article**', '');
      
      // If we have an extracted title, add it in bold before the content
      if (extractedTitle && extractedTitle !== 'Revised Article') {
        headerLines.push(`**${extractedTitle}**`, '');
      }
      
      const headerHtml = convertMarkdownToHtml(headerLines.join('\n'));
      const revisedHtml = convertMarkdownToHtml(revisedContent);
      const combinedHtml = `${headerHtml}${revisedHtml}`;
      
      const revisedMessage: Message = {
        role: 'assistant',
        content: combinedHtml,
        timestamp: new Date(),
        isHtml: true,
        thoughtLeadership: metadata
      };
      
      this.messageSubject.next({ type: 'result', message: revisedMessage });
    } else {
      const errorContent = '### Quick Start Thought Leadership – Edit Content\n\n**Revised Article**\n\n_No revised article was returned. Please try again._';
      const errorMessage: Message = {
        role: 'assistant',
        content: convertMarkdownToHtml(errorContent),
        timestamp: new Date(),
        isHtml: true,
        thoughtLeadership: metadata
      };
      this.messageSubject.next({ type: 'result', message: errorMessage });
    }
  }

  cancelWorkflow(): void {
    if (this.currentState.step === 'idle') {
      return;
    }

    if (this.currentState.step === 'processing') {
      return;
    }

    this.updateState(this.getDefaultState());
    this.workflowCompletedSubject.next();

    const cancelMessage: Message = {
      role: 'assistant',
      content: 'Edit workflow cancelled. How else can I help you?',
      timestamp: new Date()
    };

    this.messageSubject.next({
      type: 'prompt',
      message: cancelMessage
    });
  }

  completeWorkflow(): void {
    this.updateState(this.getDefaultState());
    this.workflowCompletedSubject.next();
  }

  private updateState(newState: EditWorkflowState): void {
    this.stateSubject.next(newState);
  }

  
  /** Sanitize HTML content using Angular's DomSanitizer */
  private sanitizeHtml(html: string): SafeHtml {
    return this.sanitizer.bypassSecurityTrustHtml(html);
  }
  
  private getDefaultState(): EditWorkflowState {
    return {
      step: 'idle',
      uploadedFile: null,
      selectedEditors: ['brand-alignment'],
      originalContent: '',
      paragraphEdits: []
    };
  }

  private cloneEditorOptions(): EditorOption[] {
    return this.editorOptions.map(opt => ({ ...opt }));
  }

  private createEditorSelectionMessage(content: string, editorOptions?: EditorOption[]): Message {
    const editors = editorOptions || this.cloneEditorOptions();
    // Set default selection state (only brand-alignment selected by default, and always selected)
    const defaultSelectedIds = ['brand-alignment'];
    const editorsWithSelection = editors.map(editor => ({
      ...editor,
      selected: defaultSelectedIds.includes(editor.id),
      // Mark brand-alignment as always selected and disabled
      disabled: editor.id === 'brand-alignment',
      alwaysSelected: editor.id === 'brand-alignment'
    }));

    return {
      role: 'assistant',
      content,
      timestamp: new Date(),
      editWorkflow: {
        step: 'awaiting_editors',
        showEditorSelection: true, // Enable visual UI component
        showCancelButton: false,
        showSimpleCancelButton: false,
        editorOptions: editorsWithSelection
      }
    };
  }

  private createNoEditorsErrorMessage(editorOptions?: EditorOption[]): void {
    const errorMessage = this.createEditorSelectionMessage(
      `⚠️ **Please select at least one editing service** before proceeding.`,
      editorOptions
    );
    this.messageSubject.next({ type: 'prompt', message: errorMessage });
  }
  
  /** Approve a paragraph edit */
  approveParagraph(index: number): void {
    const paragraphIndex = this.currentState.paragraphEdits.findIndex(p => p.index === index);
    
    if (paragraphIndex === -1) {
      return;
    }
    
    // Create new array with updated paragraph (new object reference for Angular change detection)
    const updatedParagraphEdits = this.currentState.paragraphEdits.map((p, i) => 
      i === paragraphIndex 
        ? { ...p, approved: true as boolean | null }
        : p
    );
    
    this.updateState({
      ...this.currentState,
      paragraphEdits: updatedParagraphEdits
    });
    
    // Emit update message to notify chat component
    this.emitParagraphUpdateMessage();
  }
  
  /** Decline a paragraph edit */
  declineParagraph(index: number): void {
    const paragraphIndex = this.currentState.paragraphEdits.findIndex(p => p.index === index);
    
    if (paragraphIndex === -1) {
      return;
    }
    
    // Create new array with updated paragraph (new object reference for Angular change detection)
    const updatedParagraphEdits = this.currentState.paragraphEdits.map((p, i) => 
      i === paragraphIndex 
        ? { ...p, approved: false as boolean | null }
        : p
    );
    
    this.updateState({
      ...this.currentState,
      paragraphEdits: updatedParagraphEdits
    });
    
    // Emit update message to notify chat component
    this.emitParagraphUpdateMessage();
  }
  
  /** Emit update message for paragraph edits */
  private emitParagraphUpdateMessage(): void {
    const updateMessage: Message = {
      role: 'assistant',
      content: '', // Content rendered by Angular component
      timestamp: new Date(),
      isHtml: false,
      editWorkflow: {
        step: 'awaiting_approval',
        paragraphEdits: [...this.currentState.paragraphEdits],
        showCancelButton: false,
        showSimpleCancelButton: true
      }
    };
    
    this.messageSubject.next({ type: 'update', message: updateMessage });
  }
  
  /** Sync paragraph edits from message to service state (for final article generation) */
  syncParagraphEditsFromMessage(paragraphEdits: ParagraphEdit[]): void {
    if (paragraphEdits && paragraphEdits.length > 0) {
      // Reconstruct originalContent from paragraphEdits if service state doesn't have it
      let originalContent = this.currentState.originalContent;
      if (!originalContent || !originalContent.trim()) {
        originalContent = this.reconstructOriginalContent(paragraphEdits);
      }
      
      this.updateState({
        ...this.currentState,
        paragraphEdits: [...paragraphEdits],
        originalContent: originalContent || this.currentState.originalContent
      });
    }
  }
  
  /** Check if all paragraphs have been decided */
  get allParagraphsDecided(): boolean {
    return allParagraphsDecided(this.currentState.paragraphEdits);
  }
  
  /** Reconstruct original content from paragraph edits (like Guided Journey) */
  private reconstructOriginalContent(paragraphEdits: ParagraphEdit[]): string {
    if (!paragraphEdits || paragraphEdits.length === 0) {
      return '';
    }
    
    // Sort by index to ensure correct order
    const sortedEdits = [...paragraphEdits].sort((a, b) => a.index - b.index);
    
    // Combine all original paragraphs
    return sortedEdits.map(p => p.original).filter(p => p && p.trim()).join('\n\n');
  }
  
  /** Generate final article using approved edits */
  async generateFinalArticle(): Promise<void> {
    if (!this.allParagraphsDecided) {
      const errorMessage: Message = {
        role: 'assistant',
        content: '⚠️ **Please approve or decline all paragraph edits** before generating the final article.',
        timestamp: new Date()
      };
      this.messageSubject.next({ type: 'prompt', message: errorMessage });
      return;
    }
    
    this.isGeneratingFinalSubject.next(true);
    
    try {
      const decisions = this.currentState.paragraphEdits.map(p => ({
        index: p.index,
        approved: p.approved === true
      }));
      
      // Get originalContent - use service state if available, otherwise reconstruct from paragraphEdits
      let originalContent = this.currentState.originalContent;
      
      if (!originalContent || !originalContent.trim()) {
        originalContent = this.reconstructOriginalContent(this.currentState.paragraphEdits);
      }
      
      if (!originalContent || !originalContent.trim()) {
        throw new Error('Original content cannot be empty. Unable to reconstruct from paragraph edits.');
      }
      
      const response = await fetch('/api/v1/tl/edit-content/final', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          original_content: originalContent,
          paragraph_edits: this.currentState.paragraphEdits.map(p => ({
            index: p.index,
            original: p.original,
            edited: p.edited,
            tags: p.tags
          })),
          decisions: decisions
        })
      });
      
      if (!response.ok) {
        const errorText = await response.text().catch(() => 'Unknown error');
        throw new Error(`Failed to generate final article: ${response.status} ${errorText}`);
      }
      
      const data = await response.json();
      const finalArticle = data.final_article || '';
      
      if (!finalArticle) {
        throw new Error('No final article returned from server');
      }
      
      // Update paragraph message to show final output (component will handle display)
      const updatedParagraphEdits = [...this.currentState.paragraphEdits];
      
      // Create final article HTML for separate message
      const uploadedFileName = this.currentState.uploadedFile?.name;
      const selectedEditorNames = this.getSelectedEditorNames(this.currentState.selectedEditors);
      
      const headerLines: string[] = [
        '### Quick Start Thought Leadership – Edit Content'
      ];
      
      if (uploadedFileName) {
        headerLines.push(`_Source: ${uploadedFileName}_`);
      }
      
      if (selectedEditorNames) {
        headerLines.push(`_Editors Applied: ${selectedEditorNames}_`);
      }
      
      headerLines.push('');
      
      const headerHtml = convertMarkdownToHtml(headerLines.join('\n'));
      const finalHtml = convertMarkdownToHtml(finalArticle);
      const finalArticleHtml = `${headerHtml}<div class="result-section"><h4 class="result-title">Final Revised Article</h4><div class="assistant-message revised-content-formatted">${finalHtml}</div></div>`;
      
      const finalMessage: Message = {
        role: 'assistant',
        content: finalArticleHtml,
        timestamp: new Date(),
        isHtml: true,
        thoughtLeadership: {
          contentType: 'article',
          topic: 'Final Revised Article',
          fullContent: finalArticle,
          showActions: true
        }
      };
      
      // Send final article message (paragraph edits remain visible in previous message)
      this.messageSubject.next({ type: 'result', message: finalMessage });
      this.completeWorkflow();
      
    } catch (error) {
      console.error('Error generating final article:', error);
      const errorMessage: Message = {
        role: 'assistant',
        content: `⚠️ **Failed to generate final article.** ${error instanceof Error ? error.message : 'Please try again.'}`,
        timestamp: new Date()
      };
      this.messageSubject.next({ type: 'prompt', message: errorMessage });
    } finally {
      this.isGeneratingFinalSubject.next(false);
    }
  }
}
