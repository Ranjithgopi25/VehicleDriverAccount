import { Component, OnInit, OnDestroy, HostListener, ViewChild, ElementRef, AfterViewChecked, ChangeDetectorRef, inject } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DomSanitizer, SafeHtml } from '@angular/platform-browser';
import { ChatService, ThemeService, ThemeMode, TlChatBridgeService, ChatEditWorkflowService } from '../../core/services';
import { Message, ChatSession, ThoughtLeadershipRequest, ThoughtLeadershipMetadata, EditorOption } from '../../core/models';
import { SourceCitationPipe } from '../../core/pipes';
import { TlFlowService } from '../../core/services/tl-flow.service';
import { DdcFlowService } from '../../core/services/ddc-flow.service';
import { DDC_WORKFLOWS } from '../../core/models/guided-journey.models';
import { DraftContentFlowComponent } from '../../features/thought-leadership/draft-content-flow/draft-content-flow.component';
import { ConductResearchFlowComponent } from '../../features/thought-leadership/conduct-research-flow/conduct-research-flow.component';
import { EditContentFlowComponent } from '../../features/thought-leadership/edit-content-flow/edit-content-flow.component';
import { RefineContentFlowComponent } from '../../features/thought-leadership/refine-content-flow/refine-content-flow.component';
import { FormatTranslatorFlowComponent } from '../../features/thought-leadership/format-translator-flow/format-translator-flow.component';
import { GeneratePodcastFlowComponent } from '../../features/thought-leadership/generate-podcast-flow/generate-podcast-flow.component';
import { BrandFormatFlowComponent } from '../../features/ddc/brand-format-flow/brand-format-flow.component';
import { ProfessionalPolishFlowComponent } from '../../features/ddc/professional-polish/professional-polish-flow.component';
import { SanitizationFlowComponent } from '../../features/ddc/sanitization/sanitization-flow.component';
import { ClientCustomizationFlowComponent } from '../../features/ddc/client-customization/client-customization-flow.component';
import { RfpResponseFlowComponent } from '../../features/ddc/rfp-response/rfp-response-flow.component';
import { FormatTranslatorFlowComponent as DdcFormatTranslatorFlowComponent } from '../../features/ddc/format-translator/format-translator-flow.component';
import { SlideCreationFlowComponent } from '../../features/ddc/slide-creation/slide-creation-flow.component';
import { GuidedDialogComponent } from '../../shared/components/guided-dialog/guided-dialog.component';
import { TlActionButtonsComponent } from '../../features/chat/components/message-list/tl-action-buttons/tl-action-buttons.component';
import { EditorSelectionComponent } from '../../features/chat/components/editor-selection/editor-selection.component';
import { EditorProgressComponent } from '../../shared/ui/components/editor-progress/editor-progress.component';
import { ParagraphEditsConsolidatedComponent } from '../../shared/ui/components/paragraph-edits/paragraph-edits-consolidated.component';
import { CanvasEditorComponent } from '../../features/thought-leadership/canvas-editor/canvas-editor.component';
import { CanvasStateService } from '../../core/services/canvas-state.service';
import { VoiceInputComponent } from '../../shared/components/voice-input/voice-input.component';
import { FileUploadComponent } from '../../shared/components/file-upload/file-upload.component';
import { Subject } from 'rxjs';
import { takeUntil } from 'rxjs/operators';

@Component({
  selector: 'app-chat',
  standalone: true,
  imports: [
    CommonModule, 
    FormsModule, 
    SourceCitationPipe,
    DraftContentFlowComponent,
    ConductResearchFlowComponent,
    EditContentFlowComponent,
    RefineContentFlowComponent,
    FormatTranslatorFlowComponent,
    GeneratePodcastFlowComponent,
    BrandFormatFlowComponent,
    ProfessionalPolishFlowComponent,
    SanitizationFlowComponent,
    ClientCustomizationFlowComponent,
    RfpResponseFlowComponent,
    DdcFormatTranslatorFlowComponent,
    SlideCreationFlowComponent,
    GuidedDialogComponent,
    TlActionButtonsComponent,
    EditorSelectionComponent,
    CanvasEditorComponent,
    VoiceInputComponent,
    FileUploadComponent,
    EditorProgressComponent,
    ParagraphEditsConsolidatedComponent
  ],
  templateUrl: './chat.component.html',
  styleUrls: ['./chat.component.scss']
})
export class ChatComponent implements OnInit, OnDestroy, AfterViewChecked {
  @ViewChild('messagesContainer') private messagesContainer?: ElementRef;
  @ViewChild('quickStartBtn') private quickStartBtn?: ElementRef;
  @ViewChild(VoiceInputComponent) voiceInput?: VoiceInputComponent;
  @ViewChild(RefineContentFlowComponent) refineContentFlow?: RefineContentFlowComponent;
  
  private shouldScrollToBottom = false;
  private destroy$ = new Subject<void>();
  private sanitizer = inject(DomSanitizer);
  messages: Message[] = [];
  userInput: string = '';
  isLoading: boolean = false;
  showDraftForm: boolean = false;
  showGuidedDialog: boolean = false;
  showPromptSuggestions: boolean = false;
  selectedActionCategory: string = '';
  selectedFlow: 'ppt' | 'thought-leadership' = 'ppt';
  selectedTLOperation: string = 'generate';
  selectedPPTOperation: string = 'draft';
  originalPPTFile: File | null = null;
  referencePPTFile: File | null = null;
  sanitizePPTFile: File | null = null;
  uploadedPPTFile: File | null = null;
  uploadedEditDocumentFile: File | null = null; // For Edit Content workflow
  referenceDocument: File | null = null;
  editorialDocumentFile: File | null = null;
  referenceLink: string = '';
  currentAction: string = '';
  selectedDownloadFormat: string = 'word';
  showAttachmentArea: boolean = false;
  
  // Dropdown state
  openDropdown: string | null = null;
  
  // Chat history persistence
  currentSessionId: string | null = null;
  savedSessions: ChatSession[] = [];
  private readonly STORAGE_KEY = 'pwc_chat_sessions';
  private readonly MAX_SESSIONS = 20;
  
  // Search functionality
  searchQuery: string = '';
  offeringVisibility = {
    'ppt': true,
    'thought-leadership': true
  };
  

  // Mobile menu state
  mobileMenuOpen: boolean = false;
  
  // Sidebar collapse state
  sidebarExpanded: boolean = false;
  
  // Theme dropdown state
  showThemeDropdown: boolean = false;
  prefersDark: boolean = window.matchMedia('(prefers-color-scheme: dark)').matches;
  
  // History panel state
  showHistoryPanel: boolean = false;

  
  // PPT Quick Actions
  pptQuickActions: string[] = ['Digital Document Development Center', 'Fix Formatting', 'Sanitize Documents', 'Validate Best Practices'];
  
  // NEW: Thought Leadership Quick Actions (5 Sections)
  tlQuickActions: string[] = ['Draft Content', 'Conduct Research', 'Edit Content', 'Refine Content', 'Format Translator'];
  
  // Dynamic quick actions based on selected flow
  get quickActions(): string[] {
    return this.selectedFlow === 'ppt' ? this.pptQuickActions : this.tlQuickActions;
  }
  
  promptCategories: any = {
    // PPT Categories
    draft: {
      title: 'Create Draft',
      prompts: [
        'Create a presentation on digital transformation strategy',
        'Draft slides about cloud migration benefits',
        'Build a deck on AI implementation roadmap',
        'Create an executive summary presentation'
      ]
    },
    improve: {
      title: 'Fix Formatting',
      prompts: [
        'Fix spelling and grammar in my presentation',
        'Align all shapes and text boxes',
        'Rebrand my deck with new colors',
        'Clean up slide formatting'
      ]
    },
    sanitize: {
      title: 'Sanitize Documents',
      prompts: [
        'Remove all client-specific data from my deck',
        'Sanitize numbers and metrics',
        'Clear all metadata and notes',
        'Remove logos and branding'
      ]
    },
    bestPractices: {
      title: 'Validate Best Practices',
      prompts: [
        'Validate my presentation against PwC best practices',
        'Check slide design and formatting standards',
        'Review chart and visual guidelines',
        'Ensure MECE framework compliance'
      ]
    },
    
    // NEW: Thought Leadership Categories (5 Sections)
    draftContent: {
      title: 'Draft Content',
      prompts: [
        'Draft an article on digital transformation trends',
        'Create a white paper on AI in business',
        'Write an executive brief on market insights',
        'Draft a blog post about future of work'
      ]
    },
    conductResearch: {
      title: 'Conduct Research',
      prompts: [
        'Research industry trends with multiple sources',
        'Analyze competitive landscape with citations',
        'Gather insights from PwC resources and external data',
        'Synthesize findings across documents and web sources'
      ]
    },
    editContent: {
      title: 'Edit Content',
      prompts: [
        'Apply brand alignment review to my article',
        'Perform copy editing on my white paper',
        'Get line editing suggestions for clarity',
        'Request content editor feedback on structure'
      ]
    },
    refineContent: {
      title: 'Refine Content',
      prompts: [
        'Expand my article to 2500 words with research',
        'Compress my white paper to executive brief format',
        'Adjust tone for C-suite audience',
        'Get suggestions to improve my content'
      ]
    },
    formatTranslator: {
      title: 'Format Translator',
      prompts: [
        'Convert my article to a blog post',
        'Transform this white paper into an executive brief',
        'Translate blog content to formal article',
        'Convert executive brief to comprehensive white paper'
      ]
    },
    generatePodcast: {
      title: 'Generate Podcast',
      prompts: [
        'Create a podcast episode about digital transformation',
        'Generate a podcast discussing industry trends',
        'Convert my article into a podcast script',
        'Create an audio version of my thought leadership content'
      ]
    },
    
    // Legacy TL Categories (kept for compatibility)
    generate: {
      title: 'Generate Article',
      prompts: [
        'Write an article on future of work',
        'Create thought leadership on sustainability',
        'Draft insights on digital innovation',
        'Generate content on industry trends'
      ]
    },
    research: {
      title: 'Research Assistant',
      prompts: [
        'Research trends in digital transformation',
        'Find competitive insights in my industry',
        'Analyze market opportunities and challenges',
        'Gather data on innovation best practices'
      ]
    },
    draftArticle: {
      title: 'Draft Article',
      prompts: [
        'Draft a case study on successful transformation',
        'Create an executive brief on industry trends',
        'Write a blog post about innovation',
        'Generate a white paper on technology adoption'
      ]
    },
    editorial: {
      title: 'Editorial Support',
      prompts: [
        'Review and improve my article structure',
        'Enhance clarity and readability',
        'Add professional touches to my draft',
        'Provide editorial feedback'
      ]
    }
  };

  draftData = {
    topic: '',
    objective: '',
    audience: '',
    additional_context: '',
    reference_document: '',
    reference_link: ''
  };

  sanitizeData = {
    clientName: '',
    products: '',
    options: {
      numericData: true,
      personalInfo: true,
      financialData: true,
      locations: true,
      identifiers: true,
      names: true,
      logos: true,
      metadata: true,
      llmDetection: true,
      hyperlinks: true,
      embeddedObjects: true
    }
  };

  thoughtLeadershipData = {
    topic: '',
    perspective: '',
    target_audience: '',
    document_text: '',
    target_format: '',
    additional_context: '',
    reference_document: '',
    reference_link: ''
  };

  researchData = {
    query: '',
    focus_areas: '',
    additional_context: '',
    links: ['']
  };
  researchFiles: File[] = [];

  articleData = {
    topic: '',
    content_type: 'Article',
    desired_length: 1000,
    tone: 'Professional',
    outline_text: '',
    additional_context: ''
  };

  bestPracticesData = {
    categories: {
      structure: true,
      visuals: true,
      design: true,
      charts: true,
      formatting: true,
      content: true
    }
  };

  outlineFile: File | null = null;
  supportingDocFiles: File[] = [];
  bestPracticesPPTFile: File | null = null;

  podcastData = {
    contentText: '',
    customization: '',
    podcastStyle: 'dialogue'
  };
  podcastFiles: File[] = [];

  // DDC Guided Journey support
  ddcWorkflows = DDC_WORKFLOWS;
  showDdcGuidedDialog: boolean = false;

  constructor(
    private chatService: ChatService,
    public themeService: ThemeService,
    private cdr: ChangeDetectorRef,
    public tlFlowService: TlFlowService,
    public ddcFlowService: DdcFlowService,
    private tlChatBridge: TlChatBridgeService,
    private canvasStateService: CanvasStateService,
    public editWorkflowService: ChatEditWorkflowService
  ) {}

  ngOnInit(): void {
    this.loadSavedSessions();
    this.subscribeToThoughtLeadership();
    this.subscribeToCanvasUpdates();
    this.subscribeToEditWorkflow();
    this.subscribeToDdcGuidedDialog();
    this.subscribeToTLGuidedDialog();
    this.messages.push({
      role: 'assistant',
      content: 'Welcome to PwC Presentation Assistant! I can help you with:\n\n**Digital Document Development Center:**\n• Create new digital documents with structured outlines\n• Improve existing documents: fix spelling/grammar, align shapes, rebrand colors\n• Sanitize presentations: remove ALL sensitive data for reuse\n• Apply MECE framework and PwC best practices\n• Create client-ready slide structures\n\n**Thought Leadership:**\n• Generate draft articles and insights\n• Research additional perspectives\n• Provide editorial support and improvements\n• Translate content between formats\n\nHow can I assist you today?',
      timestamp: new Date()
    });
    
    // Focus quick start button after view init
    setTimeout(() => {
      this.quickStartBtn?.nativeElement?.focus();
    }, 100);
  }
  
  ngOnDestroy(): void {
    this.destroy$.next();
    this.destroy$.complete();
  }

  ngAfterViewChecked(): void {
    if (this.shouldScrollToBottom) {
      this.scrollToBottom();
      this.shouldScrollToBottom = false;
    }
  }
  
  private subscribeToThoughtLeadership(): void {
    this.tlChatBridge.message$
      .pipe(takeUntil(this.destroy$))
      .subscribe({
        next: (message) => {
          console.log('[ChatComponent] Received message from TL bridge:', message);
          console.log('[ChatComponent] Message has thoughtLeadership metadata:', !!message.thoughtLeadership);
          if (message.thoughtLeadership) {
            console.log('[ChatComponent] TL metadata:', message.thoughtLeadership);
            console.log('[ChatComponent] Content type:', message.thoughtLeadership.contentType);
            console.log('[ChatComponent] Has podcast audio URL:', !!message.thoughtLeadership.podcastAudioUrl);
          }
          this.messages.push(message);
          this.saveCurrentSession();
          this.triggerScrollToBottom();
        },
        error: (err) => {
          console.error('[ChatComponent] Error in TL subscription:', err);
        }
      });
  }
  
  private subscribeToEditWorkflow(): void {
    this.editWorkflowService.message$
      .pipe(takeUntil(this.destroy$))
      .subscribe({
        next: (workflowMessage) => {
          // Handle message updates (e.g., paragraph approval state changes)
          if (workflowMessage.type === 'update') {
            // Find existing paragraph edit message to update
            const existingIndex = this.messages.findIndex(m => 
              m.editWorkflow?.step === 'awaiting_approval' && 
              m.editWorkflow?.paragraphEdits &&
              m.editWorkflow.paragraphEdits.length > 0
            );
            
            if (existingIndex !== -1) {
              // Update existing paragraph edit message with new state (create new array reference for change detection)
              if (workflowMessage.message.editWorkflow?.paragraphEdits) {
                this.messages[existingIndex].editWorkflow!.paragraphEdits = [...workflowMessage.message.editWorkflow.paragraphEdits];
              }
              
              this.saveCurrentSession();
              this.cdr.detectChanges();
              return;
            }
          }
          
          // If this is a progress message, update the existing one instead of creating new ones
          if (workflowMessage.message.editWorkflow?.step === 'processing' && 
              workflowMessage.message.editWorkflow?.editorProgress) {
            // Find and update existing progress message
            const existingIndex = this.messages.findIndex(m => 
              m.editWorkflow?.step === 'processing' && 
              m.editWorkflow?.editorProgress &&
              m.content === '' // Progress messages have empty content
            );
            
            if (existingIndex !== -1) {
              // Update existing progress message
              this.messages[existingIndex] = workflowMessage.message;
            } else {
              // First progress message, add it
              this.messages.push(workflowMessage.message);
            }
          } else {
            // Regular message, add it
            this.messages.push(workflowMessage.message);
          }
          
          this.saveCurrentSession();
          setTimeout(() => {
            this.triggerScrollToBottom();
          }, 100);
        },
        error: (err) => {
          console.error('[ChatComponent] Error in Edit Workflow subscription:', err);
        }
      });
    
    // Subscribe to workflow completion to clear state
    this.editWorkflowService.workflowCompleted$
      .pipe(takeUntil(this.destroy$))
      .subscribe({
        next: () => {
          console.log('[ChatComponent] Workflow completed - clearing state');
          this.clearWorkflowState();
        }
      });
    
    // Subscribe to workflow started to clear previous state when new workflow begins
    this.editWorkflowService.workflowStarted$
      .pipe(takeUntil(this.destroy$))
      .subscribe({
        next: () => {
          console.log('[ChatComponent] Workflow started - clearing previous state');
          this.clearWorkflowState();
        }
      });
  }
  
  private subscribeToCanvasUpdates(): void {
    this.canvasStateService.contentUpdate$
      .pipe(takeUntil(this.destroy$))
      .subscribe({
        next: (update) => {
          // Find the message by extracting index from messageId
          const messageIndex = parseInt(update.messageId.replace('msg_', ''));
          if (messageIndex >= 0 && messageIndex < this.messages.length) {
            const message = this.messages[messageIndex];
            // Update message content
            message.content = update.updatedContent;
            // Update thoughtLeadership metadata if it exists
            if (message.thoughtLeadership) {
              message.thoughtLeadership.fullContent = update.updatedContent;
            }
            this.saveCurrentSession();
            this.cdr.detectChanges();
          }
        },
        error: (err) => {
          console.error('[ChatComponent] Error in Canvas update subscription:', err);
        }
      });
  }
  

  private subscribeToDdcGuidedDialog(): void {
    this.ddcFlowService.guidedDialog$
      .pipe(takeUntil(this.destroy$))
      .subscribe({
        next: (isOpen) => {
          this.showDdcGuidedDialog = isOpen;
          this.cdr.detectChanges();
        },
        error: (err) => {
          console.error('[ChatComponent] Error in DDC Guided Dialog subscription:', err);
        }
      });
  }

  private subscribeToTLGuidedDialog(): void {
    this.tlFlowService.guidedDialog$
      .pipe(takeUntil(this.destroy$))
      .subscribe({
        next: (isOpen) => {
          this.showGuidedDialog = isOpen;
          this.cdr.detectChanges();
        },
        error: (err) => {
          console.error('[ChatComponent] Error in TL Guided Dialog subscription:', err);
        }
      });
  }

  private scrollToBottom(): void {
    try {
      if (this.messagesContainer) {
        const element = this.messagesContainer.nativeElement;
        element.scrollTop = element.scrollHeight;
      }
    } catch (err) {
      console.error('Error scrolling to bottom:', err);
    }
  }
  
  private triggerScrollToBottom(): void {
    this.shouldScrollToBottom = true;
    this.cdr.detectChanges();
  }
  
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: Event): void {
    // Close dropdown if click is outside
    const target = event.target as HTMLElement;
    if (!target.closest('.dropdown-wrapper')) {
      this.openDropdown = null;
    }
  }
  
  @HostListener('document:keydown', ['$event'])
  onKeyDown(event: KeyboardEvent): void {
    // Keyboard shortcuts
    if (event.metaKey || event.ctrlKey) {
      switch (event.key) {
        case 'k':
          event.preventDefault();
          this.focusInput();
          break;
        case 'n':
          event.preventDefault();
          this.goHome();
          break;
      }
    }
    
    // Escape to close dialogs
    if (event.key === 'Escape') {
      if (this.showGuidedDialog) {
        this.closeGuidedDialog();
      }
      if (this.openDropdown) {
        this.openDropdown = null;
      }
    }
  }
  
  private focusInput(): void {
    setTimeout(() => {
      const inputElement = document.querySelector('.composer-textarea') as HTMLTextAreaElement;
      if (inputElement) {
        inputElement.focus();
      }
    }, 50);
  }

  private handleEditWorkflowFlow(trimmedInput: string): void {
    // Add user message to chat
    const messageContent = trimmedInput || (this.uploadedEditDocumentFile ? `Uploaded document: ${this.uploadedEditDocumentFile.name}` : '');
    if (messageContent) {
      const workflowUserMessage: Message = {
        role: 'user',
        content: messageContent,
        timestamp: new Date()
      };
      this.messages.push(workflowUserMessage);
      this.triggerScrollToBottom();
    }

    const fileToUpload = this.uploadedEditDocumentFile || undefined;
    
    // Let handleChatInput manage the workflow - it will detect intent and start workflow if needed
    // This prevents double-triggering and ensures proper flow
    this.editWorkflowService.handleChatInput(trimmedInput, fileToUpload).catch(error => {
      console.error('Error in edit workflow:', error);
    });

    this.userInput = '';
    if (fileToUpload) {
      this.uploadedEditDocumentFile = null;
    }
    this.saveCurrentSession();
  }

  sendMessage(): void {
    const trimmedInput = this.userInput.trim();

    if ((!trimmedInput && !this.uploadedPPTFile && !this.uploadedEditDocumentFile) || this.isLoading) {
      return;
    }

    const isThoughtLeadershipFlow = this.selectedFlow === 'thought-leadership';

    // Quick Start Thought Leadership - Edit Content workflow
    const workflowActive = this.editWorkflowService.isActive;
    const hasEditWorkflowFile = !!this.uploadedEditDocumentFile;

    // Check for edit intent asynchronously (hybrid approach: keyword + LLM)
    if (isThoughtLeadershipFlow && (workflowActive || hasEditWorkflowFile)) {
      // Workflow already active or file uploaded - proceed
      this.editWorkflowService.handleChatInput(trimmedInput);
      return;
    }

    // Check for edit intent if not already in workflow
    if (isThoughtLeadershipFlow && !workflowActive && trimmedInput) {
      // Show loading message "Analyzing your request..."
      const loadingMessage: Message = {
        role: 'assistant',
        content: 'Analyzing your request...',
        timestamp: new Date(),
        actionInProgress: 'Analyzing your request...'
      };
      this.messages.push(loadingMessage);
      this.triggerScrollToBottom();

      // Add user message
      const userMessage: Message = {
        role: 'user',
        content: trimmedInput,
        timestamp: new Date()
      };
      this.messages.push(userMessage);
      this.userInput = '';
      this.triggerScrollToBottom();

      // Use async intent detection (LLM-based)
      this.editWorkflowService.detectEditIntent(trimmedInput).then(intentResult => {
        // Remove loading message
        const loadingIndex = this.messages.indexOf(loadingMessage);
        if (loadingIndex !== -1) {
          this.messages.splice(loadingIndex, 1);
        }

        if (intentResult.hasEditIntent) {
          // Start workflow - workflow service handles Path 1 (direct editor) vs Path 2 (selection)
          this.editWorkflowService.handleChatInput(trimmedInput);
        } else {
          // No edit intent - continue with normal chat flow
          this.proceedWithNormalChat(trimmedInput);
        }
      }).catch(error => {
        console.error('Error detecting edit intent:', error);
        // Remove loading message
        const loadingIndex = this.messages.indexOf(loadingMessage);
        if (loadingIndex !== -1) {
          this.messages.splice(loadingIndex, 1);
        }
        // Fallback to normal chat flow on error
        this.proceedWithNormalChat(trimmedInput);
      });
      return;
    }

    // No edit intent detected or not in TL flow - continue with normal chat
    this.proceedWithNormalChat(trimmedInput);
  }

  private proceedWithNormalChat(trimmedInput: string): void {
    const userInputLower = trimmedInput.toLowerCase();
    const isThoughtLeadershipFlow = this.selectedFlow === 'thought-leadership';
    
    // Check if user is requesting sanitization
    const sanitizationKeywords = ['sanitize', 'sanitise', 'sanitization', 'sanitation', 'remove sensitive', 'clean up', 'strip data', 'anonymize', 'anonymise'];
    const isSanitizationRequest = sanitizationKeywords.some(keyword => userInputLower.includes(keyword));

    // Check if user is requesting draft/create presentation
    const draftKeywords = ['create presentation', 'draft presentation', 'create a deck', 'draft a deck', 'build presentation', 'make presentation', 'new presentation', 'create slides'];
    const isDraftRequest = draftKeywords.some(keyword => userInputLower.includes(keyword));
    
    // Check if user is requesting podcast generation (ONLY in TL mode)
    const podcastKeywords = ['podcast', 'generate podcast', 'create podcast', 'make podcast', 'convert to podcast', 'audio version', 'turn into podcast', 'audio narration'];
    const isPodcastRequest = isThoughtLeadershipFlow && podcastKeywords.some(keyword => userInputLower.includes(keyword));

    // If there's an uploaded PPT file and NOT a sanitization request, process it
    if (this.uploadedPPTFile && !isSanitizationRequest) {
      this.processPPTUpload();
      return;
    }
    
    // If user asks for podcast generation in TL mode, open podcast flow
    if (isPodcastRequest) {
      this.openPodcastFlow(trimmedInput);
      return;
    }

    // If user asks to sanitize, start conversational workflow
    if (isSanitizationRequest) {
      this.startSanitizationConversation();
      return;
    }

    // If user asks to create/draft presentation
    if (isDraftRequest) {
      const userMessage: Message = {
        role: 'user',
        content: trimmedInput,
        timestamp: new Date()
      };
      this.messages.push(userMessage);

      const assistantMessage: Message = {
        role: 'assistant',
        content: '📝 I\'d be happy to help you create a presentation! To provide the best draft, please tell me:\n\n1. **Topic**: What is the main subject?\n2. **Objective**: What do you want to achieve?\n3. **Audience**: Who will view this presentation?\n\nYou can describe these in your next message, or click the "Guided Journey" button above for a structured form.',
        timestamp: new Date()
      };
      this.messages.push(assistantMessage);

      this.userInput = '';
      this.saveCurrentSession();
      return;
    }

    const userMessage: Message = {
      role: 'user',
      content: this.userInput,
      timestamp: new Date()
    };

    this.messages.push(userMessage);
    this.triggerScrollToBottom();
    this.userInput = '';
    this.isLoading = true;

    const assistantMessage: Message = {
      role: 'assistant',
      content: '',
      timestamp: new Date()
    };
    this.messages.push(assistantMessage);
    this.triggerScrollToBottom();

    const messagesToSend = this.messages
      .filter(m => m.role !== 'system')
      .map(m => ({ role: m.role, content: m.content }));

    this.chatService.streamChat(messagesToSend).subscribe({
      next: (content: string) => {
        assistantMessage.content += content;
        this.triggerScrollToBottom();
      },
      error: (error: any) => {
        console.error('Error:', error);
        assistantMessage.content = 'Sorry, I encountered an error. Please make sure the AI service is configured correctly.';
        this.isLoading = false;
        this.triggerScrollToBottom();
      },
      complete: () => {
        this.isLoading = false;
        this.saveCurrentSession();
        this.triggerScrollToBottom();
      }
    });
  }
  
  processPPTUpload(): void {
    if (!this.uploadedPPTFile) return;
    
    const userPrompt = this.userInput.trim() || 'Improve my presentation';
    const userMessage: Message = {
      role: 'user',
      content: `${userPrompt}: ${this.uploadedPPTFile.name}`,
      timestamp: new Date()
    };
    this.messages.push(userMessage);
    this.triggerScrollToBottom();

    const assistantMessage: Message = {
      role: 'assistant',
      content: '',
      timestamp: new Date(),
      actionInProgress: 'Improving presentation...'
    };
    this.messages.push(assistantMessage);
    this.triggerScrollToBottom();

    this.userInput = '';
    this.isLoading = true;
    this.currentAction = 'Improving presentation...';

    const pptFile = this.uploadedPPTFile;
    this.uploadedPPTFile = null;

    this.chatService.improvePPT(pptFile, null).subscribe({
      next: (blob) => {
        assistantMessage.actionInProgress = undefined;
        assistantMessage.content = `I've successfully improved your presentation "${pptFile.name}". Here's what was done:\n\n• Fixed spelling and grammar errors\n• Aligned text and shapes\n• Applied consistent formatting\n\nYou can download the improved version below.`;
        
        // Create download URL from blob
        const url = window.URL.createObjectURL(blob);
        const filename = pptFile.name.replace('.pptx', '_improved.pptx');
        assistantMessage.downloadUrl = url;
        assistantMessage.downloadFilename = filename;
      },
      error: (error) => {
        console.error('Error improving PPT:', error);
        assistantMessage.actionInProgress = undefined;
        assistantMessage.content = 'Sorry, I encountered an error while improving the presentation. Please try again.';
        this.isLoading = false;
        this.currentAction = '';
      },
      complete: () => {
        this.isLoading = false;
        this.currentAction = '';
        this.saveCurrentSession();
        this.triggerScrollToBottom();
      }
    });
  }

  startSanitizationConversation(): void {
    const userMessage: Message = {
      role: 'user',
      content: this.userInput,
      timestamp: new Date()
    };
    this.messages.push(userMessage);

    const assistantMessage: Message = {
      role: 'assistant',
      content: '',
      timestamp: new Date(),
      isStreaming: true
    };
    this.messages.push(assistantMessage);

    this.userInput = '';
    this.isLoading = true;
    this.triggerScrollToBottom();

    // Include file name if uploaded
    const fileName = this.uploadedPPTFile ? this.uploadedPPTFile.name : undefined;

    this.chatService.streamSanitizationConversation(
      this.messages.filter(m => !m.isStreaming),
      fileName
    ).subscribe({
      next: (chunk: string) => {
        assistantMessage.content += chunk;
        this.triggerScrollToBottom();
      },
      error: (error: any) => {
        console.error('Error:', error);
        assistantMessage.content = 'Sorry, I encountered an error. Please try again.';
        assistantMessage.isStreaming = false;
        this.isLoading = false;
      },
      complete: () => {
        assistantMessage.isStreaming = false;
        this.isLoading = false;
        this.saveCurrentSession();
      }
    });
  }

  processSanitizePPT(): void {
    if (!this.uploadedPPTFile) return;
    
    const userPrompt = this.userInput.trim() || 'Sanitize my presentation';
    const userMessage: Message = {
      role: 'user',
      content: `${userPrompt}: ${this.uploadedPPTFile.name}`,
      timestamp: new Date()
    };
    this.messages.push(userMessage);

    const assistantMessage: Message = {
      role: 'assistant',
      content: '',
      timestamp: new Date(),
      actionInProgress: 'Sanitizing presentation...'
    };
    this.messages.push(assistantMessage);

    this.userInput = '';
    this.isLoading = true;
    this.currentAction = 'Sanitizing presentation: removing sensitive data, client names, numbers, and metadata...';

    const pptFile = this.uploadedPPTFile;
    this.uploadedPPTFile = null;

    // Use empty strings for client name and products since we're in free text mode
    this.chatService.sanitizePPT(pptFile, '', '').subscribe({
      next: (response) => {
        const url = window.URL.createObjectURL(response.blob);

        let statsMessage = '';
        if (response.stats) {
          statsMessage = `\n\nSanitization Statistics:\n• Numeric replacements: ${response.stats.numeric_replacements}\n• Name replacements: ${response.stats.name_replacements}\n• Hyperlinks removed: ${response.stats.hyperlinks_removed}\n• Notes removed: ${response.stats.notes_removed}\n• Logos removed: ${response.stats.logos_removed}\n• Slides processed: ${response.stats.slides_processed}`;
          
          if (response.stats.llm_replacements) {
            statsMessage += `\n• LLM-detected items: ${response.stats.llm_replacements}`;
          }
        }

        assistantMessage.content = `✅ Your presentation has been sanitized!\n\nSanitization complete:\n• All numeric data replaced with X patterns\n• Personal information removed\n• Client/product names replaced with placeholders\n• Logos and watermarks removed\n• Speaker notes cleared\n• Metadata sanitized` + statsMessage + '\n\nYou can download your sanitized presentation below.';
        assistantMessage.downloadUrl = url;
        assistantMessage.downloadFilename = 'sanitized_presentation.pptx';
        assistantMessage.previewUrl = url;
        assistantMessage.actionInProgress = undefined;
        this.isLoading = false;
        this.currentAction = '';
      },
      error: (error: any) => {
        console.error('Error:', error);
        assistantMessage.content = 'Sorry, I encountered an error while sanitizing your presentation. Please make sure the file is a valid PowerPoint file (.pptx).';
        assistantMessage.actionInProgress = undefined;
        this.isLoading = false;
        this.currentAction = '';
      },
      complete: () => {
        this.saveCurrentSession();
      }
    });
  }

  toggleDraftForm(): void {
    this.showDraftForm = !this.showDraftForm;
  }

  selectFlow(flow: 'ppt' | 'thought-leadership'): void {
    this.selectedFlow = flow;
    // Close any open forms/dialogs and go to chat home
    this.showDraftForm = false;
    this.showGuidedDialog = false;
    this.showPromptSuggestions = false;
    this.closeMobileSidebar();
    // Clear uploaded files when switching flows
    this.uploadedEditDocumentFile = null;
    this.uploadedPPTFile = null;
    // Reset edit workflow if active
    if (this.editWorkflowService.isActive) {
      this.editWorkflowService.cancelWorkflow();
    }
    // Reset to initial state - just show welcome with only the initial assistant message
    if (this.messages.length > 1) {
      this.messages = this.messages.slice(0, 1);
    }
  }
  
  goHome(): void {
    // Reset to home state
    this.showDraftForm = false;
    this.showGuidedDialog = false;
    this.showPromptSuggestions = false;
    this.showAttachmentArea = false;
    this.userInput = '';
    this.referenceDocument = null;
    this.closeMobileSidebar();
    
    // Clear chat history and reset to initial assistant message
    if (this.messages.length > 1) {
      this.messages = this.messages.slice(0, 1);
    }
    
    // Reset all form data
    this.draftData = {
      topic: '',
      objective: '',
      audience: '',
      additional_context: '',
      reference_document: '',
      reference_link: ''
    };
    
    this.thoughtLeadershipData = {
      topic: '',
      perspective: '',
      target_audience: '',
      document_text: '',
      target_format: '',
      additional_context: '',
      reference_document: '',
      reference_link: ''
    };
    
    this.originalPPTFile = null;
    this.referencePPTFile = null;
    this.sanitizePPTFile = null;
    this.uploadedPPTFile = null;
    this.uploadedEditDocumentFile = null;
    this.editorialDocumentFile = null;
    // Reset edit workflow if active
    if (this.editWorkflowService.isActive) {
      this.editWorkflowService.cancelWorkflow();
    }
    this.currentSessionId = null;
    this.isLoading = false;
  }
  

  toggleMobileMenu(): void {
    this.mobileMenuOpen = !this.mobileMenuOpen;
  }
  
  closeMobileSidebar(): void {
    this.mobileMenuOpen = false;
  }
  
  toggleSidebar(): void {
    this.sidebarExpanded = !this.sidebarExpanded;
  }
  
  toggleThemeDropdown(): void {
    this.showThemeDropdown = !this.showThemeDropdown;
  }
  
  getFeatureName(): string {
    if (this.selectedFlow === 'ppt') {
      return 'Digital Document Development Center';
    } else if (this.selectedFlow === 'thought-leadership') {
      return 'Thought Leadership';
    }
    return 'MCX AI';

  }
  
  openGuidedDialog(): void {
    // Context-aware: Show DDC workflows for ppt flow, TL workflows for thought-leadership flow
    if (this.selectedFlow === 'ppt') {
      this.showDdcGuidedDialog = true;
    } else if (this.selectedFlow === 'thought-leadership') {
      this.showGuidedDialog = true;
    }
  }
  
  onWorkflowSelected(workflowId: string): void {
    console.log('[ChatComponent] DDC Workflow selected:', workflowId);
    this.showDdcGuidedDialog = false;
    this.ddcFlowService.openFlow(workflowId as any);
  }
  
  closeDdcGuidedDialog(): void {
    this.showDdcGuidedDialog = false;
  }
  
  // Chat history methods
  loadSavedSessions(): void {
    try {
      const stored = localStorage.getItem(this.STORAGE_KEY);
      if (stored) {
        const sessions = JSON.parse(stored);
        // Convert string dates back to Date objects
        this.savedSessions = sessions.map((s: any) => ({
          ...s,
          timestamp: new Date(s.timestamp),
          lastModified: new Date(s.lastModified),
          messages: s.messages.map((m: any) => ({
            ...m,
            timestamp: m.timestamp ? new Date(m.timestamp) : undefined
          }))
        }));
      }
    } catch (error) {
      console.error('Error loading saved sessions:', error);
      this.savedSessions = [];
    }
  }
  
  saveCurrentSession(): void {
    // Don't save if we only have the welcome message
    if (this.messages.length <= 1) {
      return;
    }
    
    // Generate title from first user message or use default
    let title = 'New Chat';
    const firstUserMessage = this.messages.find(m => m.role === 'user');
    if (firstUserMessage) {
      title = firstUserMessage.content.slice(0, 50);
      if (firstUserMessage.content.length > 50) {
        title += '...';
      }
    }
    
    const now = new Date();
    
    if (this.currentSessionId) {
      // Update existing session
      const index = this.savedSessions.findIndex(s => s.id === this.currentSessionId);
      if (index !== -1) {
        this.savedSessions[index] = {
          ...this.savedSessions[index],
          messages: [...this.messages],
          lastModified: now
        };
      }
    } else {
      // Create new session
      this.currentSessionId = `session_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
      const newSession: ChatSession = {
        id: this.currentSessionId,
        title: title,
        messages: [...this.messages],
        timestamp: now,
        lastModified: now
      };
      
      this.savedSessions.unshift(newSession);
      
      // Limit number of saved sessions
      if (this.savedSessions.length > this.MAX_SESSIONS) {
        this.savedSessions = this.savedSessions.slice(0, this.MAX_SESSIONS);
      }
    }
    
    // Save to localStorage
    try {
      localStorage.setItem(this.STORAGE_KEY, JSON.stringify(this.savedSessions));
    } catch (error) {
      console.error('Error saving session:', error);
    }
  }
  
  loadSession(sessionId: string): void {
    const session = this.savedSessions.find(s => s.id === sessionId);
    if (session) {
      this.currentSessionId = sessionId;
      this.messages = [...session.messages];
      this.showGuidedDialog = false;
      this.showDraftForm = false;
      this.showPromptSuggestions = false;
    }
  }
  
  deleteSession(sessionId: string, event: Event): void {
    event.stopPropagation();
    this.savedSessions = this.savedSessions.filter(s => s.id !== sessionId);
    
    try {
      localStorage.setItem(this.STORAGE_KEY, JSON.stringify(this.savedSessions));
    } catch (error) {
      console.error('Error deleting session:', error);
    }
    
    // If we deleted the current session, go home
    if (this.currentSessionId === sessionId) {
      this.goHome();
    }
  }
  
  // Search/filter methods
  filterOfferings(): void {
    const query = this.searchQuery.toLowerCase().trim();
    
    if (!query) {
      this.offeringVisibility['ppt'] = true;
      this.offeringVisibility['thought-leadership'] = true;
      return;
    }
    
    // Check if "presentation drafting" or related keywords match
    const pptKeywords = ['presentation', 'drafting', 'ppt', 'slides', 'deck', 'powerpoint', 'improve', 'sanitize', 'create'];
    const tlKeywords = ['thought', 'leadership', 'article', 'research', 'insights', 'editorial', 'review', 'generate'];
    
    this.offeringVisibility['ppt'] = pptKeywords.some(keyword => keyword.includes(query) || query.includes(keyword));
    this.offeringVisibility['thought-leadership'] = tlKeywords.some(keyword => keyword.includes(query) || query.includes(keyword));
  }
  
  isOfferingVisible(offering: string): boolean {
    return this.offeringVisibility[offering as keyof typeof this.offeringVisibility];
  }
  
  getFilteredSessions(): ChatSession[] {
    const query = this.searchQuery.toLowerCase().trim();
    
    if (!query) {
      return this.savedSessions;
    }
    
    return this.savedSessions.filter(session => 
      session.title.toLowerCase().includes(query)
    );
  }
  
  closeGuidedDialog(): void {
    this.showGuidedDialog = false;
  }
  
  onTLActionCardClick(flowType: string): void {
    this.closeGuidedDialog();
    this.tlFlowService.openFlow(flowType as 'draft-content' | 'conduct-research' | 'edit-content' | 'refine-content' | 'format-translator' | 'generate-podcast');
  }
  
  showActionPrompts(category: string): void {
    this.selectedActionCategory = category;
    this.showPromptSuggestions = true;
  }
  
  usePrompt(prompt: string): void {
    this.showPromptSuggestions = false;
    this.userInput = prompt;
    // Auto-send the message
    this.sendMessage();
  }
  
  triggerFileUpload(type: 'improve' | 'sanitize'): void {
    // Create a file input element dynamically
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.pptx';
    fileInput.onchange = (event: any) => {
      const file = event.target.files[0];
      if (file) {
        if (type === 'improve') {
          this.originalPPTFile = file;
          this.selectedPPTOperation = 'improve';
          this.userInput = `Improve my presentation: ${file.name}`;
        } else {
          this.sanitizePPTFile = file;
          this.selectedPPTOperation = 'sanitize';
          this.userInput = `Sanitize my presentation: ${file.name}`;
        }
        // Let the user review and send
      }
    };
    fileInput.click();
  }

  createThoughtLeadership(): void {
    this.isLoading = true;
    this.showDraftForm = false;

    let userMessageContent = '';
    const tlData = this.thoughtLeadershipData;

    switch (this.selectedTLOperation) {
      case 'generate':
        userMessageContent = `Generate thought leadership article:\n\nTopic: ${tlData.topic}\nPerspective: ${tlData.perspective}\nTarget Audience: ${tlData.target_audience}${tlData.additional_context ? '\nAdditional Context: ' + tlData.additional_context : ''}`;
        if (this.referenceDocument) {
          userMessageContent += `\n\nReference Document: ${this.referenceDocument.name} (Note: File content integration requires backend support)`;
        }
        if (tlData.reference_link) {
          userMessageContent += `\nReference Link: ${tlData.reference_link}`;
        }
        break;
      case 'research':
        userMessageContent = `Research additional insights:\n\nTopic: ${tlData.topic}\nCurrent Perspective: ${tlData.perspective}${tlData.additional_context ? '\nAdditional Context: ' + tlData.additional_context : ''}`;
        break;
      case 'editorial':
        if (this.editorialDocumentFile) {
          userMessageContent = `Provide editorial support:\n\nDocument File: ${this.editorialDocumentFile.name} (Note: File content integration requires backend support)${tlData.additional_context ? '\n\nAdditional Instructions: ' + tlData.additional_context : ''}`;
        } else if (tlData.document_text) {
          userMessageContent = `Provide editorial support:\n\nDocument:\n${tlData.document_text}${tlData.additional_context ? '\n\nAdditional Instructions: ' + tlData.additional_context : ''}`;
        }
        break;
      case 'improve':
        userMessageContent = `Recommend improvements:\n\nDocument:\n${tlData.document_text}${tlData.additional_context ? '\n\nFocus Areas: ' + tlData.additional_context : ''}`;
        break;
      case 'translate':
        userMessageContent = `Translate document format:\n\nOriginal Document:\n${tlData.document_text}\n\nTarget Format: ${tlData.target_format}${tlData.additional_context ? '\nAdditional Requirements: ' + tlData.additional_context : ''}`;
        break;
    }

    const userMessage: Message = {
      role: 'user',
      content: userMessageContent,
      timestamp: new Date()
    };
    this.messages.push(userMessage);

    const assistantMessage: Message = {
      role: 'assistant',
      content: '',
      timestamp: new Date()
    };
    this.messages.push(assistantMessage);

    // Convert reference_link to reference_urls array for backend
    const requestPayload: ThoughtLeadershipRequest = {
      operation: this.selectedTLOperation,
      topic: tlData.topic,
      perspective: tlData.perspective,
      target_audience: tlData.target_audience,
      document_text: tlData.document_text,
      target_format: tlData.target_format,
      additional_context: tlData.additional_context,
      reference_urls: tlData.reference_link ? [tlData.reference_link] : undefined
    };

    this.chatService.streamThoughtLeadership(requestPayload).subscribe({
      next: (content: string) => {
        assistantMessage.content += content;
      },
      error: (error: any) => {
        console.error('Error:', error);
        assistantMessage.content = 'Sorry, I encountered an error. Please make sure the AI service is configured correctly.';
        this.isLoading = false;
      },
      complete: () => {
        this.isLoading = false;
        this.thoughtLeadershipData = {
          topic: '',
          perspective: '',
          target_audience: '',
          document_text: '',
          target_format: '',
          additional_context: '',
          reference_document: '',
          reference_link: ''
        };
        this.referenceDocument = null;
        this.editorialDocumentFile = null;
      }
    });
  }

  createDraft(): void {
    if (!this.draftData.topic || !this.draftData.objective || !this.draftData.audience) {
      return;
    }

    this.isLoading = true;
    this.showDraftForm = false;

    // Prepare user message with reference information
    let messageContent = `Create a presentation draft:\n\nTopic: ${this.draftData.topic}\nObjective: ${this.draftData.objective}\nAudience: ${this.draftData.audience}`;
    if (this.draftData.additional_context) {
      messageContent += `\nAdditional Context: ${this.draftData.additional_context}`;
    }
    if (this.referenceDocument) {
      messageContent += `\n\nReference Document: ${this.referenceDocument.name} (Note: File content integration requires backend support)`;
    }
    if (this.draftData.reference_link) {
      messageContent += `\nReference Link: ${this.draftData.reference_link}`;
    }
    
    const userMessage: Message = {
      role: 'user',
      content: messageContent,
      timestamp: new Date()
    };
    this.messages.push(userMessage);

    const assistantMessage: Message = {
      role: 'assistant',
      content: '',
      timestamp: new Date()
    };
    this.messages.push(assistantMessage);

    // TODO: For file upload support, convert to FormData and update backend endpoint
    this.chatService.streamDraft(this.draftData).subscribe({
      next: (content: string) => {
        assistantMessage.content += content;
      },
      error: (error) => {
        console.error('Error:', error);
        assistantMessage.content = 'Sorry, I encountered an error while creating the draft. Please make sure the GROQ_API_KEY is configured correctly.';
        this.isLoading = false;
      },
      complete: () => {
        this.isLoading = false;
        this.draftData = {
          topic: '',
          objective: '',
          audience: '',
          additional_context: '',
          reference_document: '',
          reference_link: ''
        };
        this.referenceDocument = null;
      }
    });
  }

  handleKeyPress(event: KeyboardEvent): void {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault();
      this.sendMessage();
    }
  }

  onOriginalFileSelected(event: any): void {
    const file = event.target.files[0];
    if (file && file.name.endsWith('.pptx')) {
      this.originalPPTFile = file;
    }
  }

  onReferenceFileSelected(event: any): void {
    const file = event.target.files[0];
    if (file && file.name.endsWith('.pptx')) {
      this.referencePPTFile = file;
    }
  }

  formatFileSize(bytes: number): string {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  }

  improvePPT(): void {
    if (!this.originalPPTFile) {
      return;
    }

    this.isLoading = true;
    this.showDraftForm = false;
    this.currentAction = 'Improving presentation: correcting spelling, aligning shapes, rebranding colors...';

    const userMessage: Message = {
      role: 'user',
      content: `Improve PowerPoint presentation:\n\nOriginal File: ${this.originalPPTFile.name}${this.referencePPTFile ? '\nReference File: ' + this.referencePPTFile.name : ''}\n\nOperations: Correct spelling/grammar, align shapes, rebrand colors${this.referencePPTFile ? ' (using reference PPT)' : ''}`,
      timestamp: new Date()
    };
    this.messages.push(userMessage);

    const assistantMessage: Message = {
      role: 'assistant',
      content: '',
      timestamp: new Date(),
      actionInProgress: 'Processing your presentation...'
    };
    this.messages.push(assistantMessage);

    this.chatService.improvePPT(this.originalPPTFile, this.referencePPTFile).subscribe({
      next: (blob) => {
        const url = window.URL.createObjectURL(blob);
        
        assistantMessage.content = '✅ Your presentation has been improved!\n\nChanges made:\n• Spelling and grammar corrections\n• Text and shape alignment\n' + (this.referencePPTFile ? '• Color rebranding applied\n' : '') + '\nYou can download your presentation below.';
        assistantMessage.downloadUrl = url;
        assistantMessage.downloadFilename = 'improved_presentation.pptx';
        assistantMessage.previewUrl = url; // Preview will trigger download for PPTX files
        assistantMessage.actionInProgress = undefined;
        this.isLoading = false;
        this.currentAction = '';
        this.originalPPTFile = null;
        this.referencePPTFile = null;
      },
      error: (error) => {
        console.error('Error:', error);
        assistantMessage.content = 'Sorry, I encountered an error while improving your presentation. Please make sure both files are valid PowerPoint files (.pptx).';
        assistantMessage.actionInProgress = undefined;
        this.isLoading = false;
        this.currentAction = '';
      }
    });
  }

  onSanitizeFileSelected(event: any): void {
    const file = event.target.files[0];
    if (file && file.name.endsWith('.pptx')) {
      this.sanitizePPTFile = file;
    }
  }

  sanitizePPT(): void {
    if (!this.sanitizePPTFile) {
      return;
    }

    this.isLoading = true;
    this.showDraftForm = false;
    this.currentAction = 'Sanitizing presentation: removing sensitive data, client names, numbers, and metadata...';

    const userMessage: Message = {
      role: 'user',
      content: `Sanitize PowerPoint presentation:\n\nFile: ${this.sanitizePPTFile.name}${this.sanitizeData.clientName ? '\nClient Name: ' + this.sanitizeData.clientName : ''}${this.sanitizeData.products ? '\nProducts: ' + this.sanitizeData.products : ''}\n\nRemoving: All sensitive data, numbers, client names, personal info, logos, and metadata`,
      timestamp: new Date()
    };
    this.messages.push(userMessage);

    const assistantMessage: Message = {
      role: 'assistant',
      content: '',
      timestamp: new Date(),
      actionInProgress: 'Sanitizing your presentation...'
    };
    this.messages.push(assistantMessage);

    this.chatService.sanitizePPT(this.sanitizePPTFile, this.sanitizeData.clientName, this.sanitizeData.products, this.sanitizeData.options).subscribe({
      next: (response) => {
        const url = window.URL.createObjectURL(response.blob);

        let statsMessage = '';
        if (response.stats) {
          statsMessage = `\n\nSanitization Statistics:\n• Numeric replacements: ${response.stats.numeric_replacements}\n• Name replacements: ${response.stats.name_replacements}\n• Hyperlinks removed: ${response.stats.hyperlinks_removed}\n• Notes removed: ${response.stats.notes_removed}\n• Logos removed: ${response.stats.logos_removed}\n• Slides processed: ${response.stats.slides_processed}`;
          if (response.stats.llm_replacements) {
            statsMessage += `\n• LLM-detected items: ${response.stats.llm_replacements}`;
          }
        }

        assistantMessage.content = '✅ Your presentation has been sanitized!\n\nSanitization complete:\n• All numeric data replaced with X patterns\n• Personal information removed\n• Client/product names replaced with placeholders\n• Logos and watermarks removed\n• Speaker notes cleared\n• Metadata sanitized' + statsMessage + '\n\nYou can download your sanitized presentation below.';
        assistantMessage.downloadUrl = url;
        assistantMessage.downloadFilename = 'sanitized_presentation.pptx';
        assistantMessage.previewUrl = url; // Preview will trigger download for PPTX files
        assistantMessage.actionInProgress = undefined;
        this.isLoading = false;
        this.currentAction = '';
        this.sanitizePPTFile = null;
        this.sanitizeData = { 
          clientName: '', 
          products: '',
          options: {
            numericData: true,
            personalInfo: true,
            financialData: true,
            locations: true,
            identifiers: true,
            names: true,
            logos: true,
            metadata: true,
            llmDetection: true,
            hyperlinks: true,
            embeddedObjects: true
          }
        };
      },
      error: (error: any) => {
        console.error('Error:', error);
        assistantMessage.content = 'Sorry, I encountered an error while sanitizing your presentation. Please make sure the file is a valid PowerPoint file (.pptx).';
        assistantMessage.actionInProgress = undefined;
        this.isLoading = false;
        this.currentAction = '';
      }
    });
  }

  setTheme(theme: ThemeMode): void {
    this.themeService.setTheme(theme);
  }

  showChat(): void {
    this.showDraftForm = false;
  }

  startQuickChat(): void {
    // Quick Start goes directly to chat without showing the form
    this.showDraftForm = false;
    this.showAttachmentArea = true;
    // Add a message from assistant to start the conversation
    if (this.messages.length === 1) {
      this.messages.push({
        role: 'assistant',
        content: 'I\'m ready to help! What would you like to create today?\n\n💡 **Tip:** Upload a PowerPoint file to improve or sanitize it, or start typing to create new content.',
        timestamp: new Date()
      });
    }
  }
  
  quickStart(): void {
    // Check if Quick Start message has already been shown (avoid duplicates)
    const hasQuickStartMessage = this.messages.some(msg => 
      msg.role === 'assistant' && (
        msg.content.includes('Here\'s what I can help you with in the Digital Document Development Center') ||
        msg.content.includes('Here\'s what I can help you with in Thought Leadership')
      )
    );
    
    if (hasQuickStartMessage) {
      // Already shown, just scroll to bottom
      this.triggerScrollToBottom();
      return;
    }
    
    // Create flow-specific welcome message
    let welcomeMessage = '';
    
    if (this.selectedFlow === 'ppt') {
      welcomeMessage = `👋 Welcome! Here's what I can help you with in the **Digital Document Development Center**:

📝 **Create New Presentations**
• AI-generated slide outlines with MECE framework
• Structured content following PwC consulting best practices
• Client-ready presentation templates

🔧 **Improve Existing Presentations**
• Fix spelling and grammar errors
• Align all text boxes and shapes
• Rebrand colors and formatting
• Apply consistent PwC styling

🔒 **Sanitize Documents**
• Remove ALL client-specific data for reuse
• Clear metadata and speaker notes
• Sanitize numbers, metrics, and identifiers
• Multi-tier sanitization options

✅ **Validate Best Practices**
• MECE framework compliance checking
• PwC slide design standards
• Chart and visual guidelines review

💡 **Tips:** Upload a PowerPoint file using the attachment button, or simply describe what you need and I'll guide you through the process!`;
    } else {
      welcomeMessage = `👋 Welcome! Here's what I can help you with in **Thought Leadership**:

✍️ **Draft Content**
• Articles (2,000-3,000 words)
• Blog Posts (800-1,500 words)
• White Papers (5,000+ words)
• Executive Briefs (500-1,000 words)
• Podcasts with AI-generated audio

🔍 **Conduct Research**
• Multi-document synthesis with citations
• Upload PDFs, DOCX, TXT files for analysis
• Reference external URLs and sources
• Executive summaries and insights

✏️ **Edit Content**
• Brand Alignment Editor
• Copy Editor (grammar, style)
• Line Editor (clarity, flow)
• Content Editor (structure, messaging)
• Development Editor (strategic improvements)

📄 **Refine Content**
• Expand or compress content
• Adjust tone for different audiences
• Research enhancement
• Improvement suggestions

🔄 **Format Translator**
• Article ↔ Blog Post
• White Paper ↔ Executive Brief
• Long-form ↔ Social Media

🎙️ **Generate Podcast**
• Convert any content to audio
• Dialogue or monologue styles
• Professional voice synthesis
• Downloadable MP3 files

💡 **Tips:** Type your request naturally, or click "Guided Journey" for a step-by-step wizard to create comprehensive content!`;
    }
    
    // Add the welcome message to chat
    this.messages.push({
      role: 'assistant',
      content: welcomeMessage,
      timestamp: new Date()
    });
    
    // Save session and scroll to bottom
    this.saveCurrentSession();
    this.triggerScrollToBottom();
  }
  
  toggleDropdown(dropdownId: string, event?: Event): void {
    if (event) {
      event.stopPropagation();
    }
    this.openDropdown = this.openDropdown === dropdownId ? null : dropdownId;
  }
  
  selectPrompt(prompt: string, event?: Event): void {
    if (event) {
      event.stopPropagation();
    }
    this.userInput = prompt;
    this.openDropdown = null;
    // Focus the input after selection
    setTimeout(() => {
      const inputElement = document.querySelector('.chat-input-area textarea') as HTMLTextAreaElement;
      if (inputElement) {
        inputElement.focus();
      }
    }, 100);
  }
  
  getDropdownPrompts(dropdownId: string): string[] {
    const promptMap: {[key: string]: string[]} = {
      // PPT prompts
      'draft': this.promptCategories.draft.prompts,
      'fix': this.promptCategories.improve.prompts,
      'sanitize': this.promptCategories.sanitize.prompts,
      'bestPractices': this.promptCategories.bestPractices.prompts,
      // NEW: TL Section prompts
      'draftContent': this.promptCategories.draftContent.prompts,
      'conductResearch': this.promptCategories.conductResearch.prompts,
      'editContent': this.promptCategories.editContent.prompts,
      'refineContent': this.promptCategories.refineContent.prompts,
      'formatTranslator': this.promptCategories.formatTranslator.prompts,
      // Legacy TL prompts
      'generate': this.promptCategories.generate.prompts,
      'research': this.promptCategories.research.prompts,
      'draftArticle': this.promptCategories.draftArticle.prompts,
      'review': this.promptCategories.editorial.prompts
    };
    return promptMap[dropdownId] || [];
  }
  
  quickActionClick(action: string): void {
    // For PPT actions, set prompt in chat
    if (this.selectedFlow === 'ppt') {
      const pptPrompts: {[key: string]: string} = {
        'Digital Document Development Center': 'Help me create a new digital document',
        'Fix Formatting': 'I need to fix formatting in my presentation',
        'Sanitize Documents': 'I need to sanitize sensitive data from my presentation',
        'Validate Best Practices': 'Validate my presentation against PwC best practices'
      };
      this.userInput = pptPrompts[action] || action;
    } else {
      // For TL actions, open the appropriate guided flow
      const flowMapping: {[key: string]: any} = {
        'Draft Content': 'draft-content',
        'Conduct Research': 'conduct-research',
        'Edit Content': 'edit-content',
        'Refine Content': 'refine-content',
        'Format Translator': 'format-translator'
      };
      
      const flowType = flowMapping[action];
      if (flowType) {
        this.tlFlowService.openFlow(flowType);
      }
    }
  }
  
  openDdcWorkflow(workflowId: string): void {
  console.log('[ChatComponent] Opening DDC workflow:', workflowId);
  this.ddcFlowService.openFlow(workflowId as any);
}
 

  
  onReferenceDocumentSelected(event: any): void {
    const file = event.target.files[0];
    if (file) {
      this.referenceDocument = file;
    }
  }
  
  onEditorialDocumentSelected(event: any): void {
    const file = event.target.files[0];
    if (file && (file.name.endsWith('.pdf') || file.name.endsWith('.docx') || file.name.endsWith('.doc'))) {
      this.editorialDocumentFile = file;
    }
  }
  
  triggerReferenceUpload(): void {
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.pptx';
    fileInput.onchange = (event: any) => {
      const file = event.target.files[0];
      if (file && file.name.endsWith('.pptx')) {
        this.uploadedPPTFile = file;
      }
    };
    fileInput.click();
  }
  
  removeUploadedPPT(): void {
    this.uploadedPPTFile = null;
  }

  onEditDocumentSelected(event: any): void {
    const file = event.target.files[0];
    if (file) {
      // Accept Word, PDF, Text, Markdown files
      const validExtensions = ['.doc', '.docx', '.pdf', '.txt', '.md', '.markdown'];
      const fileName = file.name.toLowerCase();
      const isValid = validExtensions.some(ext => fileName.endsWith(ext));
      
      if (isValid) {
        this.uploadedEditDocumentFile = file;
        console.log('[ChatComponent] Edit document selected:', file.name);
        
        // Auto-trigger workflow if in Thought Leadership mode
        if (this.selectedFlow === 'thought-leadership') {
          // Small delay to ensure file is set before sendMessage processes it
          setTimeout(() => {
            this.sendMessage();
          }, 100);
        }
      } else {
        alert('Please upload a Word (.doc, .docx), PDF (.pdf), Text (.txt), or Markdown (.md, .markdown) file.');
      }
    }
  }

  removeUploadedEditDocument(): void {
    this.uploadedEditDocumentFile = null;
  }

  triggerEditDocumentUpload(): void {
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.doc,.docx,.pdf,.txt,.md,.markdown';
    fileInput.onchange = (event: any) => {
      this.onEditDocumentSelected(event);
    };
    fileInput.click();
  }

  onWorkflowEditorsSubmitted(selectedIds: string[]): void {
    this.editWorkflowService.handleEditorSelection(selectedIds);
  }

  onWorkflowEditorsSelectionChanged(message: Message, editors: EditorOption[]): void {
    if (message.editWorkflow?.editorOptions) {
      message.editWorkflow.editorOptions = editors;
    }
  }

  onWorkflowCancelled(): void {
    this.editWorkflowService.cancelWorkflow();
  }

  onWorkflowFileSelected(file: File): void {
    if (this.editWorkflowService.currentState.step === 'awaiting_content') {
      // Store the file so it can be displayed in the upload component
      this.uploadedEditDocumentFile = file;
      // Handle the file upload through the workflow service
      this.editWorkflowService.handleFileUpload(file);
    }
  }

  onWorkflowFileRemoved(): void {
    // File removed - clear the uploaded file
    this.uploadedEditDocumentFile = null;
    // Note: Workflow continues even if file is removed - user can upload again
  }

  getUploadedFileForMessage(message: Message): File | null {
    // Only return the uploaded file if we're in awaiting_content step AND workflow is active
    // This prevents showing old files when workflow is idle or starting new workflow
    if (message.editWorkflow?.step === 'awaiting_content' && 
        this.editWorkflowService.isActive && 
        this.uploadedEditDocumentFile) {
      return this.uploadedEditDocumentFile;
    }
    return null;
  }

  onParagraphApproved(message: Message, index: number): void {
    if (!message.editWorkflow?.paragraphEdits) {
      return;
    }
    
    const paragraph = message.editWorkflow.paragraphEdits.find(p => p.index === index);
    if (!paragraph) {
      return;
    }
    
    // Update the paragraph directly (like Guided Journey)
    paragraph.approved = true;
    
    // Also sync with service state for final article generation
    this.editWorkflowService.syncParagraphEditsFromMessage(message.editWorkflow.paragraphEdits);
    
    // Save session and trigger change detection
    this.saveCurrentSession();
    this.cdr.detectChanges();
  }

  onParagraphDeclined(message: Message, index: number): void {
    if (!message.editWorkflow?.paragraphEdits) {
      return;
    }
    
    const paragraph = message.editWorkflow.paragraphEdits.find(p => p.index === index);
    if (!paragraph) {
      return;
    }
    
    // Update the paragraph directly (like Guided Journey)
    paragraph.approved = false;
    
    // Also sync with service state for final article generation
    this.editWorkflowService.syncParagraphEditsFromMessage(message.editWorkflow.paragraphEdits);
    
    // Save session and trigger change detection
    this.saveCurrentSession();
    this.cdr.detectChanges();
  }

  onGenerateFinalArticle(message: Message): void {
    // Sync paragraphEdits from message to service before generating final article
    if (message.editWorkflow?.paragraphEdits && message.editWorkflow.paragraphEdits.length > 0) {
      this.editWorkflowService.syncParagraphEditsFromMessage(message.editWorkflow.paragraphEdits);
    }
    
    // Call the service to generate final article
    this.editWorkflowService.generateFinalArticle();
  }

  getParagraphEditsGeneratingState(message: Message): boolean {
    // Return the service's generating state
    return this.editWorkflowService.isGeneratingFinal;
  }

  private clearWorkflowState(): void {
    this.userInput = '';
    this.uploadedEditDocumentFile = null;
    // Clear file input elements in workflow file upload components
    setTimeout(() => {
      const workflowFileInputs = document.querySelectorAll('.workflow-file-upload input[type="file"]');
      workflowFileInputs.forEach((input: any) => {
        if (input.value) {
          input.value = '';
        }
      });
      // Also clear any file inputs in chat input area
      const chatFileInputs = document.querySelectorAll('.chat-composer input[type="file"]');
      chatFileInputs.forEach((input: any) => {
        if (input.value) {
          input.value = '';
        }
      });
    }, 0);
    // Trigger change detection to update FileUploadComponent bindings
    this.cdr.detectChanges();
  }

  // Check if we're in step 2 (awaiting_content) - now optional since we show upload component
  get isAwaitingContent(): boolean {
    return this.editWorkflowService.isActive && 
           this.editWorkflowService.currentState.step === 'awaiting_content';
  }

  isEditWorkflowResult(message: Message): boolean {
    // Only show action buttons for Editorial Feedback and Revised Article results
    // These are messages with thoughtLeadership metadata from edit workflow
    if (!message.thoughtLeadership || !message.thoughtLeadership.showActions) {
      return false;
    }
    
    // Check if content indicates it's a result (Editorial Feedback or Revised Article)
    const content = message.content.toLowerCase();
    return content.includes('editorial feedback') || 
           content.includes('revised article') || 
           content.includes('quick start thought leadership');
  }


  shouldHideEditorialFeedback(message: Message, messageIndex: number): boolean {
    // Check if this message is editorial feedback
    const isEditorialFeedback = message.thoughtLeadership?.topic === 'Editorial Feedback' ||
                                (message.content && message.content.toLowerCase().includes('editorial feedback'));
    
    if (!isEditorialFeedback) {
      return false;
    }
    
    // Only hide editorial feedback if it's in the SAME message as paragraph edits
    // (Separate messages should both be shown - editorial feedback first, then paragraph edits)
    if (message.editWorkflow?.paragraphEdits && message.editWorkflow.paragraphEdits.length > 0) {
      return true;
    }
    
    return false;
  }
  
  downloadGeneratedDocument(format: string, content: string, filename: string): void {
    if (format === 'txt') {
      const blob = new Blob([content], { type: 'text/plain' });
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${filename}.txt`;
      link.click();
      window.URL.revokeObjectURL(url);
    } else if (format === 'pdf' || format === 'word') {
      this.chatService.exportDocument(content, filename, format).subscribe({
        next: (blob: Blob) => {
          const url = window.URL.createObjectURL(blob);
          const link = document.createElement('a');
          link.href = url;
          link.download = `${filename}.${format === 'word' ? 'docx' : 'pdf'}`;
          link.click();
          window.URL.revokeObjectURL(url);
        },
        error: (error: any) => {
          console.error(`Error downloading ${format}:`, error);
          alert(`Failed to download ${format === 'word' ? 'Word document' : 'PDF'}. Please try again.`);
        }
      });
    }
  }

  copyToClipboard(content: string): void {
    navigator.clipboard.writeText(content).then(() => {
      alert('Content copied to clipboard!');
    }).catch(err => {
      console.error('Failed to copy:', err);
      alert('Failed to copy content. Please try again.');
    });
  }

  downloadAsWord(content: string): void {
    // Extract title from content (first line or "Refined Content")
    const lines = content.split('\n');
    let title = 'Refined Content';
    
    // Try to extract title from markdown heading or first line
    const titleMatch = content.match(/\*\*(.+?)\*\*/);
    if (titleMatch) {
      title = titleMatch[1].trim();
    } else if (lines[0] && lines[0].trim()) {
      title = lines[0].trim().replace(/^#+\s*/, '').substring(0, 50);
    }
    
    // Clean title for filename
    const filename = title.replace(/[^a-z0-9]/gi, '_').toLowerCase() || 'refined_content';
    
    this.downloadGeneratedDocument('word', content, filename);
  }

  downloadAsPDF(content: string): void {
    // Extract title from content (first line or "Refined Content")
    const lines = content.split('\n');
    let title = 'Refined Content';
    
    // Try to extract title from markdown heading or first line
    const titleMatch = content.match(/\*\*(.+?)\*\*/);
    if (titleMatch) {
      title = titleMatch[1].trim();
    } else if (lines[0] && lines[0].trim()) {
      title = lines[0].trim().replace(/^#+\s*/, '').substring(0, 50);
    }
    
    // Clean title for filename
    const filename = title.replace(/[^a-z0-9]/gi, '_').toLowerCase() || 'refined_content';
    
    this.downloadGeneratedDocument('pdf', content, filename);
  }
  
  // Helper method to get TL metadata for any assistant message in TL mode
  getTLMetadata(message: Message): ThoughtLeadershipMetadata | undefined {
    // If message already has TL metadata, return it
    if (message.thoughtLeadership) {
      return message.thoughtLeadership;
    }
    
    // If we're in TL mode and this is an assistant message with content, create default metadata
    if (this.selectedFlow === 'thought-leadership' && message.role === 'assistant' && message.content) {
      return {
        contentType: 'article', // Default type
        topic: 'Generated Content',
        fullContent: message.content,
        showActions: true
      };
    }
    
    return undefined;
  }
  
  // Helper to detect if message is a welcome/instructional message (not actual generated content)
  private isWelcomeMessage(message: Message): boolean {
    if (!message.content || message.role !== 'assistant') return false;
    
    const content = message.content.toLowerCase();
    const welcomePatterns = [
      'welcome to',
      'how can i assist',
      'how can i help',
      'i\'ll help you',
      'please provide:',
      'you can also use'
    ];
    
    // Check if content starts with or contains welcome patterns
    return welcomePatterns.some(pattern => content.includes(pattern));
  }
  
  // Check if message should show TL action buttons
  shouldShowTLActions(message: Message): boolean {
    // Don't show action buttons for welcome/instructional messages
    if (this.isWelcomeMessage(message)) {
      return false;
    }
    
    // Show TL actions only for messages with thoughtLeadership metadata and showActions flag
    return !!(message.thoughtLeadership && message.thoughtLeadership.showActions);
  }
  
  openPodcastFlow(userQuery: string): void {
    // Add user message
    const userMessage: Message = {
      role: 'user',
      content: userQuery,
      timestamp: new Date()
    };
    this.messages.push(userMessage);
    
    // Add assistant response suggesting podcast generation
    const assistantMessage: Message = {
      role: 'assistant',
      content: `I'll help you generate a podcast! Please provide:\n\n1. **Topic or Content**: What should the podcast be about?\n2. **Style**: Dialogue (2 hosts) or Monologue (1 narrator)?\n3. **Additional Context** (optional): Any specific points or customization?\n\nYou can also use the **Guided Journey** button above to open the full podcast creation wizard, or type your requirements here and I'll generate it for you.`,
      timestamp: new Date()
    };
    this.messages.push(assistantMessage);
    
    this.userInput = '';
    this.saveCurrentSession();
    this.triggerScrollToBottom();
    
    // Optionally, open the guided dialog directly to the podcast workflow
    this.selectedTLOperation = 'generate-podcast';
    this.showGuidedDialog = true;
  }

  startGuidedJourney(): void {
    // Guided Journey shows the form first, then goes to chat after submission
    this.showDraftForm = true;
    this.selectedPPTOperation = 'draft'; // Default to draft operation
    this.selectedTLOperation = 'generate'; // Default to generate operation
  }

  selectAction(action: string): void {
    if (this.selectedFlow === 'ppt') {
      this.selectedPPTOperation = action;
    } else {
      this.selectedTLOperation = action;
    }
    this.showDraftForm = true;
  }

  getFormTitle(): string {
    if (this.selectedFlow === 'ppt') {
      switch (this.selectedPPTOperation) {
        case 'draft': return 'Digital Document Development Center';
        case 'improve': return 'Improve Existing Presentation';
        case 'sanitize': return 'Sanitize Presentation';
        default: return 'Document Development Operations';
      }
    } else {
      switch (this.selectedTLOperation) {
        case 'generate': return 'Generate Thought Leadership Article';
        case 'research': return 'Research Additional Insights';
        case 'editorial': return 'Editorial Support';
        case 'improve': return 'Improve Document';
        case 'translate': return 'Translate Document Format';
        default: return 'Thought Leadership Operations';
      }
    }
  }

  downloadFile(url: string, filename: string): void {
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.click();
    window.URL.revokeObjectURL(url);
  }

  previewFile(url: string): void {
    // For PPTX files, browsers will trigger download since they cannot preview natively
    // For true preview, we would need to convert PPTX to PDF or images on the backend
    window.open(url, '_blank');
  }
  
  getPromptKeys(): string[] {
    if (this.selectedFlow === 'ppt') {
      return ['draft', 'improve', 'sanitize'];
    } else {
      return ['generate', 'editorial'];
    }
  }
  
  onEnterPress(event: Event): void {
    const keyboardEvent = event as KeyboardEvent;
    
    // Note: Step 2 now shows file upload component, so text input can be enabled
    // But we can still optionally prevent sending if needed
    
    if (!keyboardEvent.shiftKey) {
      event.preventDefault();
      this.sendMessage();
    }
  }

  private showStep2ErrorNotification(): void {
    // Show error message via the workflow service
    const errorMessage: Message = {
      role: 'assistant',
      content: '⚠️ **Please upload a document file** (Word, PDF, Text, or Markdown). Text input is disabled in this step - only file uploads are accepted.',
      timestamp: new Date(),
      editWorkflow: {
        step: 'awaiting_content',
        showCancelButton: false,
        showSimpleCancelButton: true
      }
    };
    this.messages.push(errorMessage);
    this.saveCurrentSession();
    this.triggerScrollToBottom();
  }

  submitResearchForm(): void {
    if (!this.researchData.query.trim() || this.isLoading) {
      return;
    }

    this.isLoading = true;
    this.showGuidedDialog = false;

    const validLinks = this.researchData.links.filter(link => link.trim().length > 0);
    const userMessage: Message = {
      role: 'user',
      content: `Research Assistant: ${this.researchData.query}\n${this.researchFiles.length > 0 ? 'Files: ' + this.researchFiles.map(f => f.name).join(', ') + '\n' : ''}${validLinks.length > 0 ? 'Links: ' + validLinks.join(', ') + '\n' : ''}${this.researchData.focus_areas ? 'Focus Areas: ' + this.researchData.focus_areas + '\n' : ''}${this.researchData.additional_context ? 'Additional Context: ' + this.researchData.additional_context : ''}`,
      timestamp: new Date()
    };
    this.messages.push(userMessage);

    const assistantMessage: Message = {
      role: 'assistant',
      content: '',
      timestamp: new Date(),
      actionInProgress: 'Analyzing materials and researching...'
    };
    this.messages.push(assistantMessage);
    this.saveCurrentSession();

    this.chatService.streamResearchWithMaterials(
      this.researchFiles.length > 0 ? this.researchFiles : null,
      validLinks.length > 0 ? validLinks : null,
      this.researchData.query,
      this.researchData.focus_areas ? this.researchData.focus_areas.split(',').map(a => a.trim()) : [],
      this.researchData.additional_context
    ).subscribe({
      next: (data) => {
        if (data.type === 'progress') {
          assistantMessage.actionInProgress = data.message;
          this.saveCurrentSession();
        } else if (data.type === 'content') {
          assistantMessage.content += data.content;
          this.saveCurrentSession();
        } else if (data.type === 'sources') {
          // Store source metadata for rendering clickable citations
          assistantMessage.sources = data.sources;
          this.saveCurrentSession();
        } else if (data.type === 'complete') {
          assistantMessage.actionInProgress = undefined;
          this.isLoading = false;
          this.saveCurrentSession();
          this.resetResearchForm();
        } else if (data.type === 'error') {
          assistantMessage.content = `❌ Error: ${data.message}`;
          assistantMessage.actionInProgress = undefined;
          this.isLoading = false;
          this.saveCurrentSession();
        }
      },
      error: (error) => {
        console.error('Error:', error);
        assistantMessage.actionInProgress = undefined;
        assistantMessage.content = 'Sorry, I encountered an error while researching. Please try again.';
        this.isLoading = false;
        this.saveCurrentSession();
      },
      complete: () => {
        assistantMessage.actionInProgress = undefined;
        this.isLoading = false;
        this.saveCurrentSession();
      }
    });
  }

  submitArticleForm(): void {
    if (!this.articleData.topic.trim() || this.isLoading) {
      return;
    }

    this.isLoading = true;
    this.showGuidedDialog = false;

    const userMessage: Message = {
      role: 'user',
      content: `Draft Article: ${this.articleData.topic}\nType: ${this.articleData.content_type}\nLength: ${this.articleData.desired_length} words\nTone: ${this.articleData.tone}${this.articleData.outline_text ? '\nOutline: ' + this.articleData.outline_text : ''}${this.outlineFile ? '\nOutline File: ' + this.outlineFile.name : ''}${this.supportingDocFiles.length > 0 ? '\nSupporting Documents: ' + this.supportingDocFiles.map(f => f.name).join(', ') : ''}${this.articleData.additional_context ? '\nAdditional Context: ' + this.articleData.additional_context : ''}`,
      timestamp: new Date()
    };
    this.messages.push(userMessage);

    const assistantMessage: Message = {
      role: 'assistant',
      content: '',
      timestamp: new Date(),
      actionInProgress: 'Drafting article...'
    };
    this.messages.push(assistantMessage);

    this.chatService.draftArticle(this.articleData, this.outlineFile || undefined, this.supportingDocFiles.length > 0 ? this.supportingDocFiles : undefined).subscribe({
      next: (content: string) => {
        assistantMessage.content += content;
      },
      error: (error) => {
        console.error('Error:', error);
        assistantMessage.actionInProgress = undefined;
        assistantMessage.content = 'Sorry, I encountered an error while drafting the article. Please try again.';
        this.isLoading = false;
      },
      complete: () => {
        assistantMessage.actionInProgress = undefined;
        assistantMessage.downloadUrl = 'generated';
        this.isLoading = false;
        this.saveCurrentSession();
        this.resetArticleForm();
      }
    });
  }

  submitBestPracticesForm(): void {
    if (!this.bestPracticesPPTFile || this.isLoading) {
      return;
    }

    this.isLoading = true;
    this.showGuidedDialog = false;

    const selectedCategories = Object.keys(this.bestPracticesData.categories)
      .filter(key => this.bestPracticesData.categories[key as keyof typeof this.bestPracticesData.categories]);

    const userMessage: Message = {
      role: 'user',
      content: `Validate Best Practices: ${this.bestPracticesPPTFile.name}\nCategories: ${selectedCategories.join(', ')}`,
      timestamp: new Date()
    };
    this.messages.push(userMessage);

    const assistantMessage: Message = {
      role: 'assistant',
      content: '',
      timestamp: new Date(),
      actionInProgress: 'Analyzing presentation against best practices...'
    };
    this.messages.push(assistantMessage);

    this.chatService.streamBestPractices(this.bestPracticesPPTFile, selectedCategories).subscribe({
      next: (content: string) => {
        assistantMessage.content += content;
      },
      error: (error) => {
        console.error('Error:', error);
        assistantMessage.actionInProgress = undefined;
        assistantMessage.content = 'Sorry, I encountered an error while validating best practices. Please try again.';
        this.isLoading = false;
      },
      complete: () => {
        assistantMessage.actionInProgress = undefined;
        this.isLoading = false;
        this.saveCurrentSession();
        this.resetBestPracticesForm();
      }
    });
  }

  onOutlineFileSelected(event: any): void {
    const file = event.target.files[0];
    if (file) {
      this.outlineFile = file;
    }
  }

  onSupportingDocsSelected(event: any): void {
    const files = Array.from(event.target.files) as File[];
    this.supportingDocFiles = files;
  }

  onBestPracticesFileSelected(event: any): void {
    const file = event.target.files[0];
    if (file && file.name.endsWith('.pptx')) {
      this.bestPracticesPPTFile = file;
    }
  }

  resetResearchForm(): void {
    this.researchData = {
      query: '',
      focus_areas: '',
      additional_context: '',
      links: ['']
    };
    this.researchFiles = [];
  }
  
  onResearchFilesSelected(event: any): void {
    const files = Array.from(event.target.files) as File[];
    this.researchFiles = files.filter(file => {
      const name = file.name.toLowerCase();
      return name.endsWith('.pdf') || name.endsWith('.docx') || name.endsWith('.txt') || name.endsWith('.md');
    });
  }
  
  addResearchLink(): void {
    this.researchData.links.push('');
  }
  
  removeResearchLink(index: number): void {
    if (this.researchData.links.length > 1) {
      this.researchData.links.splice(index, 1);
    }
  }

  resetArticleForm(): void {
    this.articleData = {
      topic: '',
      content_type: 'Article',
      desired_length: 1000,
      tone: 'Professional',
      outline_text: '',
      additional_context: ''
    };
    this.outlineFile = null;
    this.supportingDocFiles = [];
  }

  resetBestPracticesForm(): void {
    this.bestPracticesData = {
      categories: {
        structure: true,
        visuals: true,
        design: true,
        charts: true,
        formatting: true,
        content: true
      }
    };
    this.bestPracticesPPTFile = null;
  }

  submitPodcastForm(): void {
    if ((this.podcastFiles.length === 0 && !this.podcastData.contentText.trim()) || this.isLoading) {
      return;
    }

    this.isLoading = true;

    const userMessage: Message = {
      role: 'user',
      content: `Generate Podcast (${this.podcastData.podcastStyle === 'dialogue' ? 'Dialogue' : 'Monologue'})\n\nFiles: ${this.podcastFiles.map(f => f.name).join(', ') || 'None'}\nContent: ${this.podcastData.contentText ? 'Provided' : 'None'}\nCustomization: ${this.podcastData.customization || 'None'}`,
      timestamp: new Date()
    };
    this.messages.push(userMessage);

    const assistantMessage: Message = {
      role: 'assistant',
      content: '',
      timestamp: new Date(),
      actionInProgress: 'Generating podcast...'
    };
    this.messages.push(assistantMessage);
    this.saveCurrentSession();
    
    // Close the guided dialog
    this.showGuidedDialog = false;

    let scriptContent = '';
    let audioBase64 = '';

    this.chatService.generatePodcast(
      this.podcastFiles.length > 0 ? this.podcastFiles : null,
      this.podcastData.contentText || null,
      this.podcastData.customization || null,
      this.podcastData.podcastStyle || 'dialogue'
    ).subscribe({
      next: (data) => {
        if (data.type === 'progress') {
          assistantMessage.actionInProgress = data.message;
          this.saveCurrentSession();
        } else if (data.type === 'script') {
          scriptContent = data.content;
          assistantMessage.content = `📻 **Podcast Generated Successfully!**\n\n**Script:**\n\n${scriptContent}\n\n`;
          this.saveCurrentSession();
        } else if (data.type === 'complete') {
          audioBase64 = data.audio;
          assistantMessage.content += `\n🎧 **Audio Ready!** Listen to your podcast below or download it as an MP3 file.\n\n`;
          
          // Convert base64 to blob and create download URL
          console.log('Audio base64 length:', audioBase64.length);
          const audioBlob = this.base64ToBlob(audioBase64, 'audio/mpeg');
          console.log('Audio blob size:', audioBlob.size, 'bytes');
          const audioUrl = URL.createObjectURL(audioBlob);
          console.log('Audio URL created:', audioUrl);
          
          assistantMessage.downloadUrl = audioUrl;
          assistantMessage.downloadFilename = 'podcast.mp3';
          
          assistantMessage.actionInProgress = undefined;
          this.isLoading = false;
          this.saveCurrentSession();
          this.resetPodcastForm();
        } else if (data.type === 'error') {
          assistantMessage.content = `❌ Error generating podcast: ${data.message}`;
          assistantMessage.actionInProgress = undefined;
          this.isLoading = false;
          this.saveCurrentSession();
        }
      },
      error: (error) => {
        console.error('Error generating podcast:', error);
        assistantMessage.content = `❌ Error generating podcast: ${error.message || 'Unknown error occurred'}`;
        assistantMessage.actionInProgress = undefined;
        this.isLoading = false;
        this.saveCurrentSession();
        this.resetPodcastForm();
      }
    });
  }

  onPodcastFilesSelected(event: any): void {
    const files = Array.from(event.target.files) as File[];
    this.podcastFiles = files.filter(file => {
      const name = file.name.toLowerCase();
      return name.endsWith('.pdf') || name.endsWith('.docx') || name.endsWith('.txt') || name.endsWith('.md');
    });
  }

  resetPodcastForm(): void {
    this.podcastData = {
      contentText: '',
      customization: '',
      podcastStyle: 'dialogue'
    };
    this.podcastFiles = [];
  }

  private base64ToBlob(base64: string, contentType: string = ''): Blob {
    const byteCharacters = atob(base64);
    const byteArrays = [];

    for (let offset = 0; offset < byteCharacters.length; offset += 512) {
      const slice = byteCharacters.slice(offset, offset + 512);
      const byteNumbers = new Array(slice.length);
      for (let i = 0; i < slice.length; i++) {
        byteNumbers[i] = slice.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      byteArrays.push(byteArray);
    }

    return new Blob(byteArrays, { type: contentType });
  }

  // Voice input methods
  startVoiceInput(): void {
    setTimeout(() => {
      this.voiceInput?.startListening();
    }, 100);
  }

  onVoiceTranscriptChange(transcript: string): void {
    this.userInput = transcript;
  }

  onVoiceListeningChange(isListening: boolean): void {
    // Optional: Handle listening state changes if needed
  }

  onRefinedContentGenerated(content: string): void {
    // Populate the chat input textarea with the refined content
    this.userInput = content;
    console.log('[ChatComponent] Refined content populated in chat input');
  }

  onRefineContentStreamToChat(event: {userMessage: string, streamObservable: any}): void {
    // Add user message
    const userMessage: Message = {
      role: 'user',
      content: event.userMessage,
      timestamp: new Date()
    };
    this.messages.push(userMessage);
    this.triggerScrollToBottom();

    // Create assistant message for streaming
    const assistantMessage: Message = {
      role: 'assistant',
      content: '',
      timestamp: new Date(),
      isStreaming: true
    };
    this.messages.push(assistantMessage);
    this.triggerScrollToBottom();

    this.isLoading = true;

    // Subscribe to the stream
    event.streamObservable.subscribe({
      next: (chunk: string) => {
        assistantMessage.content += chunk;
        this.triggerScrollToBottom();
      },
      error: (error: any) => {
        console.error('Error streaming refine content:', error);
        assistantMessage.content = 'Sorry, I encountered an error while refining content. Please try again.';
        assistantMessage.isStreaming = false;
        this.isLoading = false;
        this.triggerScrollToBottom();
      },
      complete: () => {
        assistantMessage.isStreaming = false;
        this.isLoading = false;
        this.saveCurrentSession();
        this.triggerScrollToBottom();
      }
    });
  }

  /**
   * Format simple text for display (convert newlines to <br> tags)
   * Used for messages that are not already HTML formatted
   */
  formatSimpleText(text: string): string {
    if (!text) return '';
    // Escape HTML first to prevent XSS, then convert newlines to <br>
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML.replace(/\n/g, '<br>');
  }

  /**
   * Get formatted content for display
   * If message is HTML, return as-is. Otherwise, format as simple text.
   */
  getFormattedContent(message: Message): string | SafeHtml {
    if (message.isHtml) {
      // Use DomSanitizer to bypass security for trusted HTML (allows buttons and interactive elements)
      return this.sanitizer.bypassSecurityTrustHtml(message.content);
    }
    if (message.role === 'assistant' && message.sources) {
      // Use source citation pipe logic inline
      return this.formatSimpleText(message.content);
    }
    return this.formatSimpleText(message.content);
  }
  // onRefineContentStreamToChat(event: {userMessage: string, streamObservable: any}): void {
  //   // Add user message to chat
  //   const userMessage: Message = {
  //     role: 'user',
  //     content: event.userMessage,
  //     timestamp: new Date()
  //   };
  //   this.messages.push(userMessage);

  //   // Create assistant message for streaming
  //   const assistantMessage: Message = {
  //     role: 'assistant',
  //     content: '',
  //     timestamp: new Date(),
  //     isStreaming: true
  //   };
  //   this.messages.push(assistantMessage);

  //   // Set loading state
  //   this.isLoading = true;
  //   this.triggerScrollToBottom();

  //   // Subscribe to stream and update assistant message
  //   event.streamObservable.subscribe({
  //     next: (data: any) => {
  //       if (typeof data === 'string') {
  //         assistantMessage.content += data;
  //       } else if (data.type === 'content' && data.content) {
  //         assistantMessage.content += data.content;
  //       }
  //       this.triggerScrollToBottom();
  //     },
  //     error: (error: Error) => {
  //       console.error('[ChatComponent] Refine content stream error:', error);
  //       assistantMessage.isStreaming = false;
  //       assistantMessage.content = 'I apologize, but I encountered an error refining your content. Please try again.';
  //       this.isLoading = false;
  //       this.triggerScrollToBottom();
  //     },
  //     complete: () => {
  //       console.log('[ChatComponent] Refine content stream complete');
  //       assistantMessage.isStreaming = false;
  //       this.isLoading = false;
  //       this.saveCurrentSession();
  //       this.triggerScrollToBottom();
  //     }
  //   });
  // }
}
